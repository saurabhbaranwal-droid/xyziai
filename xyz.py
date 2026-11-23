#!/usr/bin/env python3
"""Daily ADR headroom report for BNY, JPMC, DB and CITI -> one Excel file.

One sheet per depositary. Only ADRs are kept: a row survives only if the
depositary's OWN official DR directory shows a sponsored program of that
depositary for it trading on a US venue (NYSE, NASDAQ, NYSE American/AMEX,
OTC/OTCQX/Pink). The local shares can be anywhere.

Columns:
    Market | Ticker | Company | ISIN / Local Code |
    Headroom from Website | Headroom from Exchange | Unit |
    Website As-Of | Exchange As-Of

Sources (nothing else is contacted):

  Headroom from the depositary websites
    JPMC  https://www.adr.com/broker/headroom   (daily PDF, via adr.com's API)
    BNY   https://www.adrbny.com/resources/dr-tools/dr-availability.html
          (one JSON call per market listed on that page)
    CITI  https://depositaryreceipts.citi.com/adr/common/linkpage.aspx?linkFormat=M&pageId=5&subpageid=168
          (India PDF linked from that page)
    DB    https://adr.db.com/drwebrebrand/market-data/headr   (PDF linked from that page)

  Headroom from the exchanges
    Korea   https://www.ksd.or.kr/en/resource-center/dr/dr-conversion  (KSD API)
            https://www.ksd.or.kr/en/resource-center/dr/dr-issuers     (depositary per DR)
    Taiwan  https://emops.twse.com.tw/server-java/t47sb25_e  (TWSE filings; same
            logic/pacing/retries as twse_adr_headroom.py / _table.py)

  US-venue check + tickers: each bank's official DR directory
    BNY   adrbny.com DR Directory           (/bin/adr/export/drDirectoryList)
    CITI  Citi Global DR Directory           (depositaryreceipts.citi.com uig.aspx)
    DB    adr.db.com DR Universe             (/drwebrebrand/api/druniverse)
    JPMC  adr.com DR universe                (adr.api.wsod.com drUniverse)
    Fallback for a CUSIP's underlying ISIN: Citi pgm_d.aspx ("ORD ISIN").

Matching rules for the US check (always against the row's own depositary):
  * Rows that name specific DRs (KSD, BNY availability lines, DB PDF ISINs):
    that exact CUSIP/ISIN must be US-traded.
  * Company-level rows (JPM PDF, Citi India PDF, TWSE filing): the depositary
    must have a sponsored program for that underlying (ISIN / TWSE code; name
    as last resort) trading in the US. If one TWSE issuer files several programs
    for the same bank, the TWSE "place of listing" must be in the US too.
  * Programs named "Reg S"/"144A" in the source are GDR tranches -> excluded.
  * Tickers = tickers of the US-traded programs; blank ones are filled from
    the same CUSIP on the other banks' directories.

Units
  Every headroom number is in DRs, except DB, which publishes ordinary shares;
  the Unit column says which. KSD publishes Korean availability in ordinary
  shares. The script converts it to DRs with KSD's own "Ratio (ORD:DR)"
  (DRs = shares x DR / ORD, rounded down) so it can sit next to the depositary
  number. TWSE's figure is "Quota for Sponsor in Reissuance (K = D - J)" in DR units.

Usage
  python adr_headroom.py                    # writes ADR_Headroom_<YYYY-MM-DD>.xlsx
  python adr_headroom.py --out report.xlsx

Requires: requests, pypdf, openpyxl
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import json
import re
import sys
import threading
import time
from collections import defaultdict
from fractions import Fraction
from zoneinfo import ZoneInfo

import requests
from pypdf import PdfReader

# --------------------------------------------------------------------------- URLs
JPM_API = "https://adr.api.wsod.com/jpmadr-public/v1"
JPM_HEADERS = {"Origin": "https://www.adr.com", "Referer": "https://www.adr.com/"}

BNY_PAGE = "https://www.adrbny.com/resources/dr-tools/dr-availability.html"
BNY_API = "https://www.adrbny.com/bin/adr/export/drAvailability"
BNY_DIRECTORY_API = "https://www.adrbny.com/bin/adr/export/drDirectoryList"
BNY_DIRECTORY_PAGE = "https://www.adrbny.com/directory/dr-directory.html"

CITI_PAGE = ("https://depositaryreceipts.citi.com/adr/common/linkpage.aspx"
             "?linkFormat=M&pageId=5&subpageid=168")
CITI_PROGRAM_URL = "https://depositaryreceipts.citi.com/adr/guides/pgm_d.aspx"
CITI_DIRECTORY_URL = "https://depositaryreceipts.citi.com/adr/guides/uig.aspx"

DB_PAGE = "https://adr.db.com/drwebrebrand/market-data/headr"
DB_DIRECTORY_API = "https://adr.db.com/drwebrebrand/api/druniverse"

KSD_API = "https://www.ksd.or.kr/en/api"
KSD_REFERER = "https://www.ksd.or.kr/en/resource-center/dr/dr-conversion"

TWSE_BASE = "https://emops.twse.com.tw/server-java/t47sb25_e"

DEPOSITARIES = ["BNY", "JPMC", "DB", "CITI"]
COLUMNS = ["Market", "Ticker", "Company", "ISIN / Local Code",
           "Headroom from Website", "Headroom from Exchange", "Unit",
           "Website As-Of", "Exchange As-Of"]
UNIT_DR = "DRs"
UNIT_SHARES = "Ordinary shares"
ASIA = {"India", "South Korea", "Taiwan"}

# TWSE pacing / retry settings, taken as-is from twse_adr_headroom_table.py:
# TWSE's servlet is flaky under sustained load (load balancer with backends
# holding inconsistent data), so requests go out one at a time, with a pause
# between them, a fresh connection each time, and retries that ride out a streak.
TWSE_REQUEST_DELAY = 0.5
TWSE_RETRY_ATTEMPTS = 6
TWSE_RETRY_BACKOFF = 2.0  # seconds, multiplied by attempt number

WARNINGS: list[str] = []


def log(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


def warn(msg: str) -> None:
    WARNINGS.append(msg)
    log(f"  WARNING: {msg}")


# ----------------------------------------------------------------------- session
class AutoVerifySession(requests.Session):
    """TLS verification on; drops to unverified only after an actual cert error
    (TLS-intercepting corporate proxy). Same pattern as twse_adr_headroom.py."""

    _warned = False
    _lock = threading.Lock()

    def request(self, *args, **kwargs):
        try:
            return super().request(*args, **kwargs)
        except requests.exceptions.SSLError:
            if self.verify is False:
                raise
            with AutoVerifySession._lock:
                if not AutoVerifySession._warned:
                    log("note: TLS verification failed (corporate proxy) -- "
                        "continuing without verification")
                    AutoVerifySession._warned = True
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            self.verify = False
            return super().request(*args, **kwargs)


def make_session(close: bool = False) -> requests.Session:
    s = AutoVerifySession()
    s.headers.update({"User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                     "AppleWebKit/537.36 (KHTML, like Gecko) "
                                     "Chrome/120.0 Safari/537.36")})
    if close:
        # TWSE's load balancer pins kept-alive connections to backends holding
        # inconsistent data (see twse_adr_headroom.py) -- force fresh ones.
        s.headers.update({"Connection": "close"})
    return s


def get(sess: requests.Session, url: str, *, tries: int = 3, **kw) -> requests.Response:
    """GET with a couple of retries for transient network errors / 5xx."""
    kw.setdefault("timeout", 60)
    last = None
    for attempt in range(1, tries + 1):
        try:
            r = sess.get(url, **kw)
            if r.status_code >= 500:
                raise requests.HTTPError(f"HTTP {r.status_code} for {url}")
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last = e
            if attempt < tries:
                time.sleep(2 * attempt)
    raise last  # type: ignore[misc]


# ----------------------------------------------------------------------- helpers
def to_int(text) -> int | None:
    if text is None:
        return None
    t = str(text).strip().replace(",", "")
    if re.fullmatch(r"-?\d+(\.0+)?", t):
        return int(float(t)) if "." in t else int(t)
    return None


def clean(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


def dep_code(name: str) -> str | None:
    """Map any depositary label (TWSE, KSD, bank directories) to BNY/JPMC/DB/CITI."""
    u = clean(name).upper()
    if not u:
        return None
    if "MELLON" in u or "BNY" in u or "BANK OF NEW YORK" in u or u == "BONY":
        return "BNY"
    if "MORGAN" in u or u.startswith("JPM"):
        return "JPMC"
    if u.startswith("CIT"):          # Citi, CITI, CIT, Citibank N.A., CITIBANK
        return "CITI"
    if "DEUTSCHE" in u or re.fullmatch(r"DB\w*", u):
        return "DB"
    return None


def deps_of(label: str) -> set[str]:
    return {c for c in (dep_code(p) for p in re.split(r"[,/]", label or "")) if c}


def norm_country(c: str) -> str:
    u = clean(c).upper()
    if u.startswith("KOREA") or u == "SOUTH KOREA":
        return "South Korea"
    if u.startswith("TAIWAN"):
        return "Taiwan"
    return clean(c)


def is_us_venue(exchange: str) -> bool:
    u = clean(exchange).upper()
    if not u:
        return False
    if any(b in u for b in ("PORTAL", "PRTL", "DUBAI", "ALTERNEXT", "EURONEXT", "AMSTERDAM", "PARIS")):
        return False
    return (u.startswith("OTC") or u in ("PINK", "PINKS") or u.startswith("NASDAQ")
            or u.startswith("NYSE") or u.startswith("NEW YORK STOCK EXCHANGE") or u == "AMEX")


def is_isin(s: str) -> bool:
    return bool(re.fullmatch(r"[A-Z]{2}[A-Z0-9]{9}\d", s or ""))


def clean_ticker(t: str) -> str:
    t = clean(t)
    if t.upper() in ("", "--", "-", "N/A", "NA", "NONE"):
        return ""
    return re.sub(r"\s+US$", "", t)


_STOP = {"LTD", "LIMITED", "CO", "COMPANY", "CORP", "CORPORATION", "INC", "INCORPORATED",
         "PLC", "THE", "HOLDING", "HOLDINGS", "GROUP", "SA", "SAE", "NV", "AG", "BHD",
         "BERHAD", "TBK", "PT", "S"}


def norm_name(s: str) -> str:
    u = (s or "").upper().replace("&", " AND ").replace("+", " AND ")
    u = re.sub(r"\bREG\.?\s*S\b", " ", u)
    u = re.sub(r"\b144\s*A\b", " ", u)
    u = re.sub(r"\s-\s.*$", " ", u)            # "Compal Electronics - 144A" style suffixes
    toks = [t for t in re.split(r"[^A-Z0-9]+", u) if t and t not in _STOP]
    return "".join(toks)


def name_match(name: str, candidates: list[dict], key: str = "name") -> list[dict]:
    """Exact normalised-name match first; else prefix match (shorter side >= 6 chars).
    Deliberately strict: no match is better than a wrong one."""
    n = norm_name(name)
    if not n:
        return []
    exact = [c for c in candidates if norm_name(c.get(key, "")) == n]
    if exact:
        return exact
    out = []
    for c in candidates:
        m = norm_name(c.get(key, ""))
        if not m:
            continue
        short, long_ = (n, m) if len(n) <= len(m) else (m, n)
        if len(short) >= 6 and long_.startswith(short):
            out.append(c)
    return out


def is_gdr_tranche_name(name: str) -> bool:
    return bool(re.search(r"\bREG\.?\s*S\b|\b144\s*A\b", (name or "").upper()))


def join_unique(*lists) -> str:
    seen, out = set(), []
    for lst in lists:
        for t in lst or []:
            t = clean(t)
            if t and t not in ("--", "-") and t not in seen:
                seen.add(t)
                out.append(t)
    return ", ".join(out)


def pdf_pages(content: bytes, layout: bool = False) -> list[str]:
    r = PdfReader(io.BytesIO(content))
    if layout:
        return [p.extract_text(extraction_mode="layout") or "" for p in r.pages]
    return [p.extract_text() or "" for p in r.pages]


def row(dep, market, *, company="", code="", web=None, exch=None, unit=UNIT_DR,
        web_asof="", exch_asof="", key=None, cusips=(), dr_isins=(), und=(),
        co_id=None, listing_place="", venue_count=1, gdr_tranche=False,
        tickers=()) -> dict:
    """One headroom row plus the identifiers used for the US check."""
    return {"dep": dep, "market": market, "ticker": "", "company": clean(company),
            "code": code, "web": web, "exch": exch, "unit": unit,
            "web_asof": web_asof, "exch_asof": exch_asof, "key": key,
            "cusips": {c for c in cusips if c}, "dr_isins": {i for i in dr_isins if i},
            "und": {u for u in und if u}, "co_id": co_id, "listing_place": listing_place,
            "venue_count": venue_count, "gdr_tranche": gdr_tranche,
            "src_tickers": [t for t in tickers if clean_ticker(t)]}


# ================================================================ DR DIRECTORIES
def _prog(src, *, deps, sponsored, cusip, dr_isin="", und="", ticker="", exchange="",
          country="", name="") -> dict:
    cusip = clean(cusip).upper()
    return {"src": src, "deps": deps, "sponsored": sponsored, "cusip": cusip,
            "dr_isin": clean(dr_isin).upper(), "und": clean(und).upper() if is_isin(clean(und).upper()) else "",
            "ticker": clean(ticker), "exchange": clean(exchange),
            "country": norm_country(country), "name": clean(name)}


def load_bny_directory(sess) -> list[dict]:
    params = {"region": "", "countryCode": "", "industryCode": "", "exchange": "",
              "depositaryBank": "", "letter": "", "capitalRaised": "", "drType": "",
              "sponsorship": "", "drPage": "directory", "count": 1000, "start": 0}
    hdr = {"Referer": BNY_DIRECTORY_PAGE, "X-Requested-With": "XMLHttpRequest"}
    items, total = [], None
    while True:
        d = get(sess, BNY_DIRECTORY_API, params=params, headers=hdr, timeout=120).json()
        d = json.loads(d) if isinstance(d, str) else d
        batch = ((d or {}).get("data") or {}).get("drdirectory") or []
        total = d.get("count") if total is None else total
        items.extend(batch)
        if not batch or len(items) >= (total or 0):
            break
        params["start"] += len(batch)
    if not items:
        raise RuntimeError("BNY DR Directory returned no programs")
    return [_prog("BNY", deps=deps_of(x.get("drDepositaries")),
                  sponsored=(x.get("drSponsType") == "S"), cusip=x.get("cusip"),
                  dr_isin=x.get("drIsinNo"), und=x.get("isinNo"), ticker=x.get("drTicker"),
                  exchange=x.get("drExchange"), country=x.get("ctryNm"), name=x.get("drTx"))
            for x in items if x.get("cusip")]


def load_citi_directory(sess) -> list[dict]:
    out = []
    for letter in ["0-9"] + [chr(c) for c in range(ord("A"), ord("Z") + 1)]:
        html = get(sess, CITI_DIRECTORY_URL, params={"pageId": 8, "subPageId": 159, "active": "A",
                                                     "company_name": letter}, timeout=120).text
        header_ok = False
        for r in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S | re.I):
            c = [clean(re.sub(r"<[^>]+>", "", x).replace("&nbsp;", " ").replace("&amp;", "&"))
                 for x in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", r, re.S | re.I)]
            if c[:8] == ["Issuer", "Active", "CUSIP", "Ticker", "Country", "Depositary", "Exchange", "Sp/Un"]:
                header_ok = True
                continue
            if len(c) == 8 and re.fullmatch(r"[0-9A-Z]{9}", c[2]):
                out.append(_prog("CITI", deps=deps_of(c[5]), sponsored=(c[7] == "S"), cusip=c[2],
                                 ticker=c[3], exchange=c[6], country=c[4], name=c[0]))
        if not header_ok:
            raise RuntimeError(f"Citi Global DR Directory page {letter!r}: table header changed")
    if not out:
        raise RuntimeError("Citi Global DR Directory returned no programs")
    return out


def load_db_directory(sess) -> list[dict]:
    items = get(sess, DB_DIRECTORY_API, headers={"Accept": "application/json"}, timeout=180).json()
    if not items:
        raise RuntimeError("DB DR Universe returned no programs")
    return [_prog("DB", deps=deps_of(x.get("depositaryBankCode")),
                  sponsored=(x.get("drSponsored") == "S"), cusip=x.get("drCusIp"),
                  dr_isin=x.get("drIsin"), und=x.get("ordIsin"), ticker=x.get("companyTicker"),
                  exchange=x.get("drExchange"), country=x.get("countryName"),
                  name=x.get("companyShortName"))
            for x in items if x.get("drCusIp") and (x.get("activeFlag") or "Active") == "Active"]


def load_jpm_directory(sess) -> list[dict]:
    items, offset = [], 0
    while True:
        r = get(sess, f"{JPM_API}/drUniverse", params={"limit": 100, "offset": offset},
                headers=JPM_HEADERS)
        data = (r.json() or {}).get("data") or {}
        items.extend(data.get("items") or [])
        if not (data.get("pagination") or {}).get("hasMore"):
            break
        offset += 100
    if not items:
        raise RuntimeError("adr.com drUniverse returned no programs")
    out = []
    for x in items:
        deps = deps_of(x.get("depositary"))
        # adr.com has no sponsorship flag; programs listed under several banks
        # are the unsponsored ones, a single bank means a sponsored program.
        out.append(_prog("JPMC", deps=deps, sponsored=(len(deps) == 1), cusip=x.get("cusip"),
                         ticker=x.get("ticker"), exchange=x.get("exchange"),
                         country=x.get("country"), name=x.get("name")))
    return out


_ORD_ISIN_RE = re.compile(r"ORD ISIN</td>\s*<td[^>]*>\s*([A-Z]{2}[A-Z0-9]{9}\d)", re.I)


class Directories:
    """All four banks' directories plus lookups used by the US check."""

    def __init__(self, programs: list[dict], sess):
        self.programs = programs
        self.sess = sess
        self.by_cusip = defaultdict(list)
        self.by_dr_isin = defaultdict(list)
        self.und = {}
        for p in programs:
            self.by_cusip[p["cusip"]].append(p)
            if p["dr_isin"]:
                self.by_dr_isin[p["dr_isin"]].append(p)
            if p["und"]:
                self.und.setdefault(p["cusip"], p["und"])
        self._own = {}
        self._pgm_tried = set()

    def own(self, dep: str) -> list[dict]:
        """Sponsored programs of `dep` as listed on dep's own website."""
        if dep not in self._own:
            self._own[dep] = [p for p in self.programs
                              if p["src"] == dep and dep in p["deps"] and p["sponsored"]]
        return self._own[dep]

    def und_of(self, p: dict) -> str:
        u = self.und.get(p["cusip"], "")
        if u or p["country"] not in ASIA or p["cusip"] in self._pgm_tried:
            return u
        # Fallback: Citi's universal program page carries the ORD ISIN for any CUSIP.
        self._pgm_tried.add(p["cusip"])
        try:
            r = get(self.sess, CITI_PROGRAM_URL, timeout=30,
                    params={"idc": 1, "pageid": 15, "subpageid": 151, "cusip": p["cusip"]})
            m = _ORD_ISIN_RE.search(r.text)
            if m:
                self.und[p["cusip"]] = m.group(1).upper()
        except requests.RequestException:
            pass
        time.sleep(TWSE_REQUEST_DELAY)
        return self.und.get(p["cusip"], "")

    def ticker(self, p: dict) -> str:
        t = clean_ticker(p["ticker"])
        if t:
            return t
        for q in self.by_cusip.get(p["cusip"], []):   # same DR on another bank's site
            t = clean_ticker(q["ticker"])
            if t:
                return t
        return ""


def _us_place(place: str) -> bool:
    u = clean(place).upper()
    return any(k in u for k in ("NEW YORK", "OTC IN US", "NASDAQ", "NYSE", "UNITED STATES", "U.S."))


def adr_check(r: dict, dirs: Directories) -> tuple[list[dict], str]:
    """US-traded programs backing row r, per r's own depositary. Returns (us, reason)."""
    dep = r["dep"]
    if r["gdr_tranche"]:
        return [], "named as a Reg S/144A GDR tranche in the source"
    own = dirs.own(dep)
    cands = []
    if r["cusips"] or r["dr_isins"]:
        cands = [p for p in own if p["cusip"] in r["cusips"] or (p["dr_isin"] and p["dr_isin"] in r["dr_isins"])]
        if not cands:   # not on dep's own site: same DR, same depositary, on another bank's site
            cands = [p for p in dirs.programs if p["sponsored"] and dep in p["deps"]
                     and (p["cusip"] in r["cusips"] or (p["dr_isin"] and p["dr_isin"] in r["dr_isins"]))]
    else:
        if r["und"] or r["co_id"]:
            for p in own:
                u = dirs.und_of(p) if p["country"] == r["market"] or p["cusip"] in dirs.und else dirs.und.get(p["cusip"], "")
                if not u:
                    continue
                if u in r["und"] or (r["co_id"] and u.startswith("TW000") and u[5:9] == r["co_id"]):
                    cands.append(p)
        if not cands and r["company"]:
            cands = name_match(r["company"], [p for p in own if p["country"] == r["market"]])
        if r["co_id"] and r["venue_count"] > 1 and not _us_place(r["listing_place"]):
            return [], f"one of several TWSE programs, listed in {r['listing_place']}"
    if not cands:
        return [], f"no sponsored {dep} program found on {dep}'s DR directory"
    us = [p for p in cands if is_us_venue(p["exchange"])]
    if not us:
        venues = sorted({p["exchange"] or "no venue" for p in cands})
        return [], f"{dep} directory lists it only on {', '.join(venues)}"
    return us, ""


def apply_adr_filter(rows: list[dict], dirs: Directories, excluded: list[dict]) -> list[dict]:
    kept = []
    for r in rows:
        us, reason = adr_check(r, dirs)
        if not us:
            excluded.append({**r, "reason": reason})
            continue
        r["ticker"] = join_unique([dirs.ticker(p) for p in us])
        if not r["code"]:
            r["code"] = join_unique([dirs.und.get(p["cusip"], "") for p in us])
        if r["key"] is None or r["key"][0] in ("KRNAME", "TWNAME"):
            if r["market"] == "South Korea":
                cus = {p["cusip"] for p in us}
                if len(cus) == 1:
                    r["key"] = ("KR", cus.pop())
            elif r["market"] == "Taiwan":
                codes = {dirs.und.get(p["cusip"], "")[5:9] for p in us
                         if dirs.und.get(p["cusip"], "").startswith("TW000")}
                if len(codes) == 1:
                    r["key"] = ("TW", codes.pop())
        kept.append(r)
    return kept


# ================================================================ WEBSITES
# ---------------------------------------------------------------- JPMC
_JPM_ROW_RE = re.compile(r"^(.*?)\s*(\d+\s*:\s*\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s*$")
_JPM_MARKET = {"KOREA, SOUTH": "South Korea", "TAIWAN, PROVINCE OF CHINA": "Taiwan"}


def parse_jpm_pdf(text: str) -> tuple[list[dict], str]:
    m = re.search(r"As of Date:\s*(\d{1,2}/\d{1,2}/\d{4})", text)
    asof = m.group(1) if m else ""
    rows, pending, market = [], [], None
    for raw in text.splitlines():
        line = clean(raw)
        if not line:
            continue
        if line.startswith("Program Name"):
            if not re.search(r"Ratio \(ORD:DR\)\s+DR Headroom Limit\s+DR Headroom Available", line):
                raise RuntimeError(f"JPM PDF header changed: {line!r}")
            if pending and not re.search(r"\d", pending[-1]) and len(pending[-1]) < 60:
                market = pending[-1]
            pending = []
            continue
        if line.startswith("JPMorgan Chase Bank"):  # page footer: stop collecting names
            pending = []
            continue
        m = _JPM_ROW_RE.match(line)
        if m and market:
            name = m.group(1).strip()
            if not name:
                name = " ".join(pending)
            elif pending and rows:   # stray wrapped tail of the previous name
                rows[-1]["name"] += " " + " ".join(pending)
                warn(f"JPM PDF: attached {' '.join(pending)!r} to previous row {rows[-1]['name']!r}")
            pending = []
            rows.append({"market": _JPM_MARKET.get(market.upper(), market), "name": name,
                         "ratio": m.group(2).replace(" ", ""),
                         "limit": to_int(m.group(3)), "available": to_int(m.group(4))})
        else:
            pending.append(line)
    return rows, asof


def fetch_jpm_website(sess) -> list[dict]:
    meta = get(sess, f"{JPM_API}/headroom", headers=JPM_HEADERS).json()
    items = ((meta or {}).get("data") or {}).get("items") or []
    items = [i for i in items if i.get("docTag")]
    if not items:
        raise RuntimeError("adr.com headroom API returned no document")
    doc = max(items, key=lambda i: i.get("date") or "")
    tag = doc["docTag"]
    m = re.search(r"HeadroomAvailability([0-9]*)", tag)
    fname = m.group(0) if m else tag.replace("adr.com.", "")
    r = get(sess, f"{JPM_API}/dm/document", params={"docTag": tag, "fileName": fname},
            headers=JPM_HEADERS)
    if not r.content.startswith(b"%PDF"):
        raise RuntimeError("adr.com headroom document is not a PDF")
    parsed, asof = parse_jpm_pdf("\n".join(pdf_pages(r.content)))
    if not parsed:
        raise RuntimeError("no rows parsed from JPM headroom PDF")
    out = []
    for p in parsed:
        key = None
        if p["market"] == "South Korea":
            key = ("KRNAME", norm_name(p["name"]))
        elif p["market"] == "Taiwan":
            key = ("TWNAME", norm_name(p["name"]))
        out.append(row("JPMC", p["market"], company=p["name"], web=p["available"],
                       web_asof=asof, key=key, gdr_tranche=is_gdr_tranche_name(p["name"])))
    return out


# ---------------------------------------------------------------- BNY
def fetch_bny_website(sess) -> list[dict]:
    page = get(sess, BNY_PAGE).text
    markets = []
    for tag in re.findall(r"<a\b[^>]*data-attr-countrycode=[^>]*>", page, re.I):
        cc = re.search(r'data-attr-countrycode="([^"]+)"', tag, re.I)
        nm = re.search(r'data-attr-country="([^"]+)"', tag, re.I)
        if cc:
            markets.append((cc.group(1).strip(), clean(nm.group(1)) if nm else cc.group(1)))
    markets = list(dict.fromkeys(markets))
    if not markets:
        raise RuntimeError("no markets found on BNY DR Availability page")

    out = []
    for cc, mname in markets:
        r = get(sess, BNY_API, params={"countryCode": cc},
                headers={"Referer": BNY_PAGE, "X-Requested-With": "XMLHttpRequest"})
        data = r.json()
        asof = clean(data.get("asofDate"))
        for issuer, lines in data.items():
            if issuer == "asofDate" or not isinstance(lines, list) or not lines:
                continue
            by_avail = defaultdict(list)
            for ln in lines:
                by_avail[to_int(ln.get("availShrCt"))].append(ln)
            if len(by_avail) > 1:
                warn(f"BNY {mname} {clean(issuer)}: DR lines show different availability "
                     f"{sorted(k for k in by_avail if k is not None)} -- one row each")
            for avail, lns in by_avail.items():
                uisin = clean(lns[0].get("udrlyingISIN"))
                key = None
                if cc == "TW" and re.fullmatch(r"TW000\d{4}\d{2}\d", uisin):
                    key = ("TW", uisin[5:9])
                elif cc == "TW":   # e.g. Cayman-incorporated issuer (KY... ISIN)
                    key = ("TWNAME", norm_name(issuer))
                elif cc == "KR":
                    cus = {clean(l.get("cusipNo")) for l in lns}
                    key = ("KR", cus.pop()) if len(cus) == 1 else None
                out.append(row("BNY", norm_country(mname), company=issuer, code=uisin,
                               web=avail, web_asof=asof, key=key,
                               cusips=[clean(l.get("cusipNo")).upper() for l in lns],
                               dr_isins=[clean(l.get("drISIN")).upper() for l in lns]))
    return out


# ---------------------------------------------------------------- CITI
_CITI_ROW_RE = re.compile(
    r"^(.+?)\s+(IN[A-Z0-9]{10})\s+(\d+\s*:\s*\d+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s*$")


def fetch_citi_website(sess) -> list[dict]:
    page = get(sess, CITI_PAGE).text
    m = re.search(r"<a[^>]+href=['\"]([^'\"]+)['\"][^>]*>\s*(?:<[^>]+>\s*)*INDIA\s*:", page, re.I)
    if not m:
        raise RuntimeError("India link not found on Citi Limited Two-Way Markets page")
    url = requests.compat.urljoin(CITI_PAGE, m.group(1).replace("&amp;", "&"))
    r = get(sess, url)
    if not r.content.startswith(b"%PDF"):
        raise RuntimeError(f"Citi India file is not a PDF ({url})")
    text = "\n".join(pdf_pages(r.content))
    flat = clean(text)
    if not re.search(r"Head-Room\s*\(DRs\)\s*Head-Room\s*\(Shares\)\s*Share Reserved\s*\(DRs\)", flat):
        raise RuntimeError("Citi India PDF column layout changed")
    am = re.search(r"Information as of\s+(.+?\d{4})", flat)
    asof = am.group(1) if am else ""
    out = []
    for line in text.splitlines():
        mm = _CITI_ROW_RE.match(clean(line))
        if not mm:
            continue
        name, isin = mm.group(1), mm.group(2)
        out.append(row("CITI", "India", company=name, code=isin, web=to_int(mm.group(4)),
                       web_asof=asof, und=[isin]))
    if not out:
        raise RuntimeError("no rows parsed from Citi India PDF")
    return out


# ---------------------------------------------------------------- DB
_NUM_RE = re.compile(r"\d{1,3}(?:,\d{3})*")
_ISIN_RE = re.compile(r"\b([A-Z]{2}[A-Z0-9]{9,10})\b")


def parse_db_pdf(pages: list[str]) -> list[dict]:
    rows: list[dict] = []
    for page in pages:
        cols = None
        cur = None
        for line in page.splitlines():
            if "Security" in line and "ISIN" in line and "Current Headroom" in line:
                cols = (line.index("ISIN"), line.index("Current Headroom"), line.find("Contacts"))
                cur = None
                continue
            if cols is None:
                continue
            low = line.strip().lower()
            if (low.startswith("the local custodian") or low.startswith("certain of these")
                    or low.startswith("*") or low.startswith("local custodian")):
                cols, cur = None, None
                continue
            if not line.strip() or low.startswith("(ordinary shares)"):
                continue
            isin_c, hr_c, ct_c = cols
            name_part = line[:isin_c].strip()
            isin_part = line[isin_c:hr_c - 3]
            hr_part = line[hr_c - 3:ct_c - 1] if ct_c > 0 else line[hr_c - 3:]
            hr_txt = hr_part.strip()
            headroom = None
            if hr_txt:
                if _NUM_RE.fullmatch(hr_txt):
                    headroom = to_int(hr_txt)
                else:
                    raise RuntimeError(f"DB PDF: unexpected text in headroom column: {line!r}")
            isins = _ISIN_RE.findall(isin_part)
            if headroom is not None:
                cur = {"name": name_part, "isins": isins, "headroom": headroom}
                rows.append(cur)
            elif cur is not None:
                if name_part:
                    cur["name"] = (cur["name"] + " " + name_part).strip()
                cur["isins"].extend(isins)
    return rows


def fetch_db_website(sess) -> list[dict]:
    page = get(sess, DB_PAGE).text
    links = re.findall(r'href="(https?://[^"]+\.pdf)"', page, re.I)
    if not links:
        raise RuntimeError("no PDF link found on DB Headroom Availability page")
    am = re.search(r"Headroom Availability as of\s+([A-Za-z]+\s+\d{1,2},\s*\d{4})", page)
    asof = am.group(1) if am else ""
    r = get(sess, links[0].replace("&amp;", "&"))
    if not r.content.startswith(b"%PDF"):
        raise RuntimeError(f"DB headroom file is not a PDF ({links[0]})")
    parsed = parse_db_pdf(pdf_pages(r.content, layout=True))
    if not parsed:
        raise RuntimeError("no rows parsed from DB headroom PDF")
    out = []
    for p in parsed:
        cusips = []
        for isin in p["isins"]:
            if isin.startswith("US") and len(isin) in (11, 12):
                cusips.append(isin[2:11])     # an 11-char ISIN just lost its check digit
            if len(isin) != 12:
                warn(f"DB PDF lists a malformed ISIN {isin!r} for {p['name']!r} (kept as published)")
        out.append(row("DB", "India", company=p["name"].rstrip("."), code=", ".join(p["isins"]),
                       web=p["headroom"], unit=UNIT_SHARES, web_asof=asof,
                       cusips=cusips, dr_isins=[i for i in p["isins"] if len(i) == 12]))
    return out


# ================================================================ EXCHANGES
# ---------------------------------------------------------------- Korea / KSD
def fetch_ksd(sess, dirs: Directories) -> list[dict]:
    h = {"Referer": KSD_REFERER}
    conv = (get(sess, f"{KSD_API}/dr/drconversion", headers=h).json().get("body") or {}).get("list") or []
    issu = (get(sess, f"{KSD_API}/dr/drissue", params={"sort": 1}, headers=h).json()
            .get("body") or {}).get("list") or []
    if not conv:
        raise RuntimeError("KSD DR conversion list is empty")
    fetched = dt.datetime.now(ZoneInfo("Asia/Seoul")).strftime("%Y-%m-%d %H:%M KST")
    dep_by_isin = {clean(i.get("drIsinNo")): dep_code(i.get("custNm", "")) for i in issu}

    out = []
    for c in conv:
        dr_isin = clean(c.get("drIsinNo")).upper()
        cusip = dr_isin[2:11] if len(dr_isin) == 12 else ""
        dep = dep_by_isin.get(dr_isin)
        if not dep:   # not on KSD's issuer list: use the banks' directories
            deps = set()
            for p in dirs.by_cusip.get(cusip, []):
                if p["sponsored"]:
                    deps |= p["deps"]
            if len(deps) == 1:
                dep = deps.pop()
        if not dep:
            warn(f"KSD: no depositary for {c.get('custNm')} {dr_isin} -- not placed on any sheet")
            continue
        shares = to_int(c.get("convPosQty"))
        rm = re.fullmatch(r"\s*(\d+)\s*:\s*(\d+)\s*", c.get("ratio") or "")
        if shares is None or not rm or int(rm.group(1)) == 0:
            warn(f"KSD: cannot read availability/ratio for {c.get('custNm')} {dr_isin}")
            drs = None
        else:
            drs = int(Fraction(shares) * int(rm.group(2)) / int(rm.group(1)))  # floor
        out.append(row(dep, "South Korea", company=c.get("custNm"), code=dr_isin, exch=drs,
                       exch_asof=f"KSD live, fetched {fetched}", key=("KR", cusip),
                       cusips=[cusip], dr_isins=[dr_isin]))
    return out


# ---------------------------------------------------------------- Taiwan / TWSE
_ROW_RE = re.compile(r"<tr[^>]*>(.*?)</tr>", re.S | re.I)
_CELL_RE = re.compile(r"<td[^>]*>(.*?)</td>", re.S | re.I)
_FIELD_LABELS = {
    "total_issued_to_date": "Total Number of Issued Shares as of this Period",
    "outstanding_current_period": "Outstanding Balance as of this Period",
    "headroom": "Quota for Sponsor in Reissuance as of this Period",
}


def _cells(row_html: str) -> list[str]:
    return [re.sub(r"<[^>]+>", "", c).strip() for c in _CELL_RE.findall(row_html)]


def twse_list_issuers(sess) -> list[dict]:
    html = get(sess, f"{TWSE_BASE}?step=summary1", timeout=30).text
    out = []
    for r in _ROW_RE.finditer(html):
        hm = re.search(r"href=['\"][^'\"]*co_id=(\d+)&(?:amp;)?location_id=(\d+)['\"]", r.group(1))
        if not hm:
            continue
        c = _cells(r.group(1))
        if len(c) < 4:
            continue
        out.append({"co_id": hm.group(1), "location_id": hm.group(2), "market": c[0],
                    "name": c[2], "listing_place": c[3]})
    return out


def twse_available_months(sess, co_id, location_id) -> list[tuple[int, int]]:
    html = get(sess, f"{TWSE_BASE}?step=summary2",
               params={"co_id": co_id, "location_id": location_id}, timeout=30).text
    months = {(int(y), int(m)) for y, m in re.findall(r"year=(\d{4})&(?:amp;)?month=(\d{1,2})", html)}
    return sorted(months, reverse=True)


def twse_fetch_filing(sess, co_id, location_id, year, month) -> dict:
    html = get(sess, f"{TWSE_BASE}?step=detail", params={
        "co_id": co_id, "location_id": location_id, "year": year, "month": month}, timeout=30).text

    def field(label):
        return re.search(rf"{re.escape(label)}[^<]*</td>\s*"
                         r"(?:<td[^>]*>[^<]*</td>\s*){0,1}"
                         r"<td[^>]*(?:colspan=['\"]?\d['\"]?)?[^>]*>([^<]*)</td>\s*"
                         r"<td[^>]*>([^<]*)</td>", html)

    out = {}
    m = re.search(r"Provided by:\s*([^<]+)", html)
    out["company_name"] = clean(m.group(1)) if m else ""
    m = re.search(r"Period:\s*(.*?)</td>", html, re.S)
    out["period"] = clean(re.sub(r"<[^>]+>", "", m.group(1))) if m else ""
    m = re.search(r"Name of the Depositary Institution</td>\s*<td[^>]*>([^<]*)</td>\s*</tr>"
                  r"(?:\s*<tr>\s*<td[^>]*>([^<]*)</td>\s*</tr>)?"
                  r"(?:\s*<tr>\s*<td[^>]*>([^<]*)</td>\s*</tr>)?", html)
    out["depositaries"] = [clean(g) for g in (m.groups() if m else ()) if g and clean(g)]
    for key, label in _FIELD_LABELS.items():
        fm = field(label)
        out[key] = {"dr_units": to_int(fm.group(1)), "shares": to_int(fm.group(2))} if fm else None
    return out


# --- retry wrappers: same as twse_adr_headroom_table.py -------------------------
def _available_months_with_retry(sess, co_id, location_id):
    """Retried if it comes back with zero months. A genuine zero-filings issuer
    still returns zero after retrying; a flaky one recovers within 1-2 attempts."""
    result = []
    for attempt in range(1, TWSE_RETRY_ATTEMPTS + 1):
        result = twse_available_months(sess, co_id, location_id)
        if result:
            return result
        if attempt < TWSE_RETRY_ATTEMPTS:
            time.sleep(TWSE_RETRY_BACKOFF * attempt)
    return result


def _list_issuers_with_retry(sess):
    """Retried if it comes back empty (the listing page hits the same flaky backend)."""
    result = []
    for attempt in range(1, TWSE_RETRY_ATTEMPTS + 1):
        result = twse_list_issuers(sess)
        if result:
            return result
        if attempt < TWSE_RETRY_ATTEMPTS:
            time.sleep(TWSE_RETRY_BACKOFF * attempt)
    return result


def _fetch_filing_with_retry(sess, co_id, location_id, year, month):
    """Retried if the headroom field comes back unparsed (stub detail page).
    Also requires the depositary name, which this report needs to pick the sheet."""
    result = {}
    for attempt in range(1, TWSE_RETRY_ATTEMPTS + 1):
        result = twse_fetch_filing(sess, co_id, location_id, year, month)
        if (result.get("headroom") or {}).get("dr_units") is not None and result.get("depositaries"):
            return result
        if attempt < TWSE_RETRY_ATTEMPTS:
            time.sleep(TWSE_RETRY_BACKOFF * attempt)
    return result


def fetch_twse() -> list[dict]:
    """Same flow as build_table() in twse_adr_headroom_table.py: sequential,
    one paced request at a time; one bad issuer never aborts the table."""
    sess = make_session(close=True)
    issuers = _list_issuers_with_retry(sess)
    if not issuers:
        raise RuntimeError("TWSE's issuer listing came back empty even after retries -- "
                           "the site itself may be down; try again shortly")
    uniq = list({(i["co_id"], i["location_id"]): i for i in issuers}.values())
    log(f"  TWSE: {len(uniq)} DR programs on file, fetching latest filings "
        f"(sequential, takes a while)...")

    results = []
    for n, i in enumerate(uniq, 1):
        log(f"    [{n}/{len(uniq)}] {i['co_id']} {i['name']} ({i['listing_place']})")
        res = {"issuer": i, "error": "", "filing": None, "year": None, "month": None}
        try:
            months = _available_months_with_retry(sess, i["co_id"], i["location_id"])
            if not months:
                res["error"] = "no filings on record"
                results.append(res)
                time.sleep(TWSE_REQUEST_DELAY)
                continue
            y, m = months[0]
            time.sleep(TWSE_REQUEST_DELAY)
            filing = _fetch_filing_with_retry(sess, i["co_id"], i["location_id"], y, m)
            res.update(filing=filing, year=y, month=m)
            if (filing.get("headroom") or {}).get("dr_units") is None:
                res["error"] = "filing page returned without a usable headroom row (TWSE flakiness, exhausted retries)"
        except Exception as e:  # one bad issuer must not abort the whole table
            res["error"] = f"{type(e).__name__}: {e}"
        results.append(res)
        time.sleep(TWSE_REQUEST_DELAY)

    venues = defaultdict(int)
    places = defaultdict(int)
    for i in uniq:
        venues[i["co_id"]] += 1
        places[(i["co_id"], i["listing_place"])] += 1

    out = []
    for res in results:
        i, f = res["issuer"], res["filing"]
        label = f"{i['co_id']} {i['name']} ({i['listing_place']})"
        if res["error"] and not f:
            warn(f"TWSE {label}: {res['error']} -- not placed on any sheet")
            continue
        deps = list(dict.fromkeys(c for c in (dep_code(d) for d in f["depositaries"]) if c))
        if not deps:
            warn(f"TWSE {label}: depositary {f['depositaries'] or 'missing'} is not BNY/JPMC/DB/CITI "
                 f"-- not placed on any sheet")
            continue
        hd = (f.get("headroom") or {}).get("dr_units")
        tot = (f.get("total_issued_to_date") or {}).get("dr_units")
        outst = (f.get("outstanding_current_period") or {}).get("dr_units")
        if hd is not None and tot is not None and outst is not None and hd != tot - outst:
            warn(f"TWSE {label}: K ({hd:,}) != D - J ({tot:,} - {outst:,}); reporting K as filed")
        if res["error"]:
            warn(f"TWSE {label}: {res['error']}")
        period = re.sub(r"\s*days of the Current.*$", "", f.get("period") or "", flags=re.I)
        asof = f"TWSE filing {res['year']}/{res['month']:02d}" + (f", {period}" if period else "")
        if venues[i["co_id"]] == 1:
            code = i["co_id"]
        elif places[(i["co_id"], i["listing_place"])] == 1:
            code = f"{i['co_id']} ({i['listing_place']})"
        else:   # same company, same listing place, separate TWSE programs
            code = f"{i['co_id']} ({i['listing_place']}, TWSE location_id {i['location_id']})"
        for d in deps:
            out.append(row(d, "Taiwan", company=f.get("company_name") or i["name"], code=code,
                           exch=hd, exch_asof=asof, key=("TW", i["co_id"]), co_id=i["co_id"],
                           listing_place=i["listing_place"], venue_count=venues[i["co_id"]]))
    return out


# ================================================================ MERGE + EXCEL
def merge(web_rows: list[dict], exch_rows: list[dict]) -> dict[str, list[dict]]:
    sheets = {d: [] for d in DEPOSITARIES}
    for dep in DEPOSITARIES:
        ex = [r for r in exch_rows if r["dep"] == dep]
        idx = defaultdict(list)
        for r in ex:
            if r["key"]:
                idx[r["key"]].append(r)
        for r in ex:   # name-based fallback keys (exchange company names), same depositary only
            if r["key"] and r["key"][0] == "KR":
                idx[("KRNAME", norm_name(r["company"]))].append(r)
            if r["key"] and r["key"][0] == "TW":
                idx[("TWNAME", norm_name(r["company"]))].append(r)
        used = set()
        for w in (r for r in web_rows if r["dep"] == dep):
            cands = idx.get(w["key"], []) if w["key"] else []
            if w["key"] and w["key"][0] in ("KRNAME", "TWNAME") and not cands:
                kind = "KR" if w["key"][0] == "KRNAME" else "TW"
                pool = [r for r in ex if r["key"] and r["key"][0] == kind]
                cands = name_match(w["company"], pool, key="company")
            if len(cands) == 1 and id(cands[0]) not in used and cands[0]["unit"] == w["unit"]:
                e = cands[0]
                used.add(id(e))
                merged = dict(w)
                merged.update(ticker=join_unique(w["ticker"].split(", "), e["ticker"].split(", ")),
                              company=w["company"] or e["company"], code=e["code"] or w["code"],
                              exch=e["exch"], exch_asof=e["exch_asof"])
                sheets[dep].append(merged)
            else:
                if len(cands) > 1:
                    warn(f"{dep} {w['market']} {w['company']}: matches {len(cands)} exchange "
                         f"programs -- website and exchange rows left separate")
                sheets[dep].append(w)
        sheets[dep].extend(r for r in ex if id(r) not in used)
        sheets[dep].sort(key=lambda r: (r["market"], r["company"].upper()))
    return sheets


def write_excel(sheets: dict[str, list[dict]], path: str) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="1F4E78")
    for dep in DEPOSITARIES:
        ws = wb.create_sheet(dep)
        ws.append(COLUMNS)
        for c in ws[1]:
            c.font, c.fill = head_font, head_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in sheets[dep]:
            ws.append([r["market"], r["ticker"] or None, r["company"], r["code"] or None,
                       r["web"], r["exch"], r["unit"], r["web_asof"] or None, r["exch_asof"] or None])
        for row_cells in ws.iter_rows(min_row=2, min_col=5, max_col=6):
            for c in row_cells:
                c.number_format = "#,##0"
        widths = [14, 22, 42, 30, 22, 22, 16, 20, 34]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{max(ws.max_row, 1)}"
    wb.save(path)


# ================================================================ MAIN
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", help="output .xlsx path (default ADR_Headroom_<YYYY-MM-DD>.xlsx)")
    args = ap.parse_args(argv)
    out_path = args.out or f"ADR_Headroom_{dt.date.today():%Y-%m-%d}.xlsx"

    sess = make_session()
    failed: list[str] = []

    def step(label, fn, *a):
        log(f"[{label}]")
        try:
            res = fn(*a)
            log(f"  {len(res)} rows")
            return res
        except Exception as e:
            failed.append(f"{label}: {type(e).__name__}: {e}")
            log(f"  ERROR: {type(e).__name__}: {e}")
            return []

    programs = []
    dir_ok = {}
    for label, fn in [("BNY DR Directory", load_bny_directory), ("Citi Global DR Directory", load_citi_directory),
                      ("DB DR Universe", load_db_directory), ("JPM adr.com DR universe", load_jpm_directory)]:
        got = step(label, fn, sess)
        dir_ok[label] = bool(got)
        programs += got
    dirs = Directories(programs, sess)
    for dep, label in [("BNY", "BNY DR Directory"), ("CITI", "Citi Global DR Directory"),
                       ("DB", "DB DR Universe"), ("JPMC", "JPM adr.com DR universe")]:
        if not dir_ok[label]:
            log(f"  NOTE: {dep}'s own directory failed -- every {dep} row will be excluded "
                f"(cannot verify it is an ADR)")

    web = []
    web += step("JPMC website (adr.com headroom PDF)", fetch_jpm_website, sess)
    web += step("BNY website (DR Availability)", fetch_bny_website, sess)
    web += step("CITI website (India headroom PDF)", fetch_citi_website, sess)
    web += step("DB website (India headroom PDF)", fetch_db_website, sess)
    exch = []
    exch += step("Korea exchange (KSD)", fetch_ksd, sess, dirs)
    exch += step("Taiwan exchange (TWSE)", fetch_twse)

    log("[ADR check against each depositary's own DR directory]")
    excluded: list[dict] = []
    web = apply_adr_filter(web, dirs, excluded)
    exch = apply_adr_filter(exch, dirs, excluded)
    log(f"  kept {len(web)} website rows and {len(exch)} exchange rows; excluded {len(excluded)} non-ADR rows")

    sheets = merge(web, exch)
    write_excel(sheets, out_path)

    log("\nExcluded (not a US-traded DR on the depositary's own directory):")
    for x in sorted(excluded, key=lambda r: (r["dep"], r["market"], r["company"].upper())):
        src = "website" if x["web"] is not None else "exchange"
        log(f"  {x['dep']:5} {x['market']:12} {x['company'][:45]:45} [{src}] {x['reason']}")
    log("")
    for d in DEPOSITARIES:
        n_w = sum(1 for r in sheets[d] if r["web"] is not None)
        n_e = sum(1 for r in sheets[d] if r["exch"] is not None)
        n_b = sum(1 for r in sheets[d] if r["web"] is not None and r["exch"] is not None)
        n_t = sum(1 for r in sheets[d] if not r["ticker"])
        log(f"{d:5}: {len(sheets[d]):3} rows  (website {n_w}, exchange {n_e}, both {n_b}, no ticker {n_t})")
    if WARNINGS:
        log(f"\n{len(WARNINGS)} warning(s) -- see above.")
    if failed:
        log("\nSOURCES THAT FAILED (their rows are missing from the file):")
        for f in failed:
            log(f"  - {f}")
    log(f"\nWrote {out_path}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
