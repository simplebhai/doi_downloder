#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOI Paper Downloader
====================
Reads Record_ID + DOI from an Excel file, finds a LEGALLY and OPENLY accessible
PDF for each DOI (Unpaywall, OpenAlex, Crossref, Europe PMC, Semantic Scholar,
publisher landing page), saves it as <Record_ID>.pdf and writes a status report.

It never bypasses paywalls, logins, CAPTCHAs, DRM or any access control:
if a source refuses access, the paper is simply reported as not accessible.

Usage
-----
  GUI (default):  python paper_downloader.py
  Command line :  python paper_downloader.py --cli PAPER_LINK.xlsx Downloaded_Papers [--email you@x.org]
"""

import argparse
import datetime as _dt
import shutil
import tempfile
import json
import os
import queue
import re
import sys
import threading
import time
import traceback
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field, asdict

import pandas as pd
import requests
from bs4 import BeautifulSoup

# Use the Windows/macOS certificate store (works behind university/corporate SSL inspection)
try:
    import truststore
    truststore.inject_into_ssl()
except Exception:
    pass
# Prefer IPv4: broken IPv6 routes on campus networks cause "connection errors"
try:
    import urllib3.util.connection as _u3c
    _u3c.HAS_IPV6 = False
except Exception:
    pass

APP_NAME = "DOI Paper Downloader"
APP_VERSION = "2.1.0"
ENGINE_VERSION = 2          # results from older engines are re-checked on resume

# --------------------------------------------------------------------------- #
#  Statuses
# --------------------------------------------------------------------------- #
S_DOWNLOADED = "Downloaded"
S_EXISTS = "Already Exists"
S_MISSING = "Missing DOI"
S_INVALID = "DOI Invalid"
S_NOT_OA = "Not Open Access"
S_NOT_FOUND = "PDF Not Found"
S_FAILED = "Download Failed"
S_SERVER = "Server Error"
S_TIMEOUT = "Timeout"
S_NOT_PROCESSED = "Not Processed"

# Temporary problems -> retried automatically on the next run (resume)
TEMPORARY_STATUSES = {S_FAILED, S_SERVER, S_TIMEOUT, S_NOT_PROCESSED}
# Final "no legal PDF" results -> skipped on resume unless user asks to re-check
NOT_ACCESSIBLE_STATUSES = {S_NOT_OA, S_NOT_FOUND}

STATE_FILE = ".download_state.json"
STATUS_XLSX = "Download_Status.xlsx"
LOG_FILE = "download_log.txt"

BROWSER_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")

MIN_PDF_BYTES = 5 * 1024            # anything smaller is almost certainly not a paper
MAX_PDF_BYTES = 300 * 1024 * 1024   # safety cap
MAX_HTML_BYTES = 3 * 1024 * 1024
TIMEOUT = (15, 60)                  # (connect, read) seconds
MAX_ATTEMPTS = 3
RETRY_WAIT = 4                      # seconds, grows with each attempt
PUBLISHER_HOST_GAP = 1.5            # min seconds between hits to the same website
API_HOST_GAP = 0.15

# API endpoints that require a subscriber key - never used
BLOCKED_LINK_HOSTS = ("api.elsevier.com", "api.wiley.com", "onlinelibrary.wiley.com/doi/full-xml")

CHALLENGE_MARKERS = ("captcha", "cf-chl", "challenge-platform", "are you a robot",
                     "verify you are human", "just a moment...", "access denied",
                     "perimeterx", "hcaptcha", "recaptcha")


# --------------------------------------------------------------------------- #
#  Small helpers
# --------------------------------------------------------------------------- #
def now_str():
    return _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


_DOI_RE = re.compile(r"10\.\d{4,9}/\S+", re.I)


def normalize_doi(raw):
    """Return (doi, problem). problem is None, 'missing' or 'invalid'."""
    if raw is None:
        return None, "missing"
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", "null", "-", "na", "n/a", "#n/a"):
        return None, "missing"
    s = urllib.parse.unquote(s).strip()
    s = re.sub(r"^\s*doi\s*[:：]?\s*", "", s, flags=re.I)
    s = re.sub(r"^(https?://)?(www\.)?(dx\.)?doi\.org/", "", s, flags=re.I)
    m = _DOI_RE.search(s)
    if not m:
        return None, "invalid"
    doi = m.group(0)
    # trim trailing punctuation that is not part of the DOI
    while doi and doi[-1] in ".,;:'\"]}>":
        doi = doi[:-1]
    while doi.endswith(")") and doi.count(")") > doi.count("("):
        doi = doi[:-1]
    if "/" not in doi or len(doi.split("/", 1)[1]) == 0:
        return None, "invalid"
    return doi, None


_ILLEGAL = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
             *(f"LPT{i}" for i in range(1, 10))}


def clean_record_id(value):
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    if re.fullmatch(r"-?\d+\.0+", s):          # Excel number 12.0 -> "12"
        s = s.split(".")[0]
    return s


def safe_filename(name, max_len=150):
    s = _ILLEGAL.sub("_", str(name)).strip().rstrip(". ")
    s = re.sub(r"\s+", " ", s)
    if not s:
        s = "unnamed"
    if s.split(".")[0].upper() in _RESERVED:
        s = "_" + s
    return s[:max_len]


def is_valid_pdf_file(path):
    try:
        if os.path.getsize(path) < MIN_PDF_BYTES:
            return False
        with open(path, "rb") as fh:
            return b"%PDF" in fh.read(1024)
    except OSError:
        return False


ERROR_LOG_DIR = [os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "DOI_Paper_Downloader")]


def write_error_log(context=""):
    """Append the current exception with full traceback to error_log.txt (never raises)."""
    try:
        tb = traceback.format_exc()
        for d in ERROR_LOG_DIR:
            try:
                os.makedirs(d, exist_ok=True)
                with open(os.path.join(d, "error_log.txt"), "a", encoding="utf-8") as fh:
                    fh.write(f"\n===== {now_str()}  v{APP_VERSION}  {context}\n{tb}")
                return os.path.join(d, "error_log.txt")
            except OSError:
                continue
    except Exception:
        pass
    return ""


def open_path(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", path])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# --------------------------------------------------------------------------- #
#  Excel input
# --------------------------------------------------------------------------- #
def _norm_header(h):
    return re.sub(r"[^a-z0-9]", "", str(h).lower())


RECORD_ALIASES = ["recordid", "record", "recid", "recordno", "recordnumber", "paperid", "id"]
DOI_ALIASES = ["doi", "doiurl", "doilink", "doinumber", "doino"]


def _match_col(headers, aliases):
    normed = {h: _norm_header(h) for h in headers}
    for alias in aliases:
        for h, n in normed.items():
            if n == alias:
                return h
    return None


def load_workbook_tables(path):
    """Return {sheet_name: (DataFrame with real header row, record_col, doi_col)}."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        raw_sheets = {"CSV": pd.read_csv(path, header=None, dtype=str, keep_default_na=False)}
    else:
        engine = "xlrd" if ext == ".xls" else None
        raw_sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=str,
                                   engine=engine, keep_default_na=False)
    tables = {}
    for sheet, raw in raw_sheets.items():
        if raw.empty:
            continue
        header_row = 0
        best = None
        # the header may not be on row 1 (titles above the table) - scan 15 rows
        for i in range(min(15, len(raw))):
            vals = [str(v) for v in raw.iloc[i].tolist()]
            rc, dc = _match_col(vals, RECORD_ALIASES), _match_col(vals, DOI_ALIASES)
            if rc or dc:
                score = (rc is not None) + (dc is not None)
                if best is None or score > best[0]:
                    best, header_row = (score, rc, dc), i
                if score == 2:
                    break
        headers = []
        for j, v in enumerate(raw.iloc[header_row].tolist()):
            v = str(v).strip() or f"Column {j + 1}"
            while v in headers:
                v += "_"
            headers.append(v)
        df = raw.iloc[header_row + 1:].copy()
        df.columns = headers
        df = df[~(df.apply(lambda r: all(str(x).strip() == "" for x in r), axis=1))]
        df = df.reset_index(drop=True)
        rc = _match_col(headers, RECORD_ALIASES)
        dc = _match_col(headers, DOI_ALIASES)
        tables[sheet] = (df, rc, dc)
    return tables


def pick_default_sheet(tables):
    for name, (_, rc, dc) in tables.items():
        if rc and dc:
            return name
    for name, (_, rc, dc) in tables.items():
        if dc:
            return name
    return next(iter(tables), None)


# --------------------------------------------------------------------------- #
#  Result record
# --------------------------------------------------------------------------- #
@dataclass
class PaperResult:
    row: int
    record_id: str
    doi_input: str
    doi: str = ""
    status: str = ""
    downloaded_file: str = ""
    source: str = ""
    source_url: str = ""
    title: str = ""
    authors: str = ""
    journal: str = ""
    year: str = ""
    publisher: str = ""
    oa_status: str = ""
    error: str = ""
    checked_at: str = ""
    from_previous_run: bool = False
    notes: list = field(default_factory=list)

    def key(self):
        return f"{self.record_id}|{self.doi or self.doi_input}"


# --------------------------------------------------------------------------- #
#  Polite HTTP client with retries
# --------------------------------------------------------------------------- #
class FetchError(Exception):
    def __init__(self, kind, msg, status=None):
        super().__init__(msg)
        self.kind = kind        # timeout | server | network | denied | notfound | http | stopped
        self.status = status


class PoliteHttp:
    def __init__(self, email, stop_event):
        self.email = email
        self.stop_event = stop_event
        self._lock = threading.Lock()
        self._next_ok = {}
        self._local = threading.local()

    def session(self):
        s = getattr(self._local, "s", None)
        if s is None:
            s = requests.Session()
            s.headers.update({
                "User-Agent": BROWSER_UA,
                "Accept-Language": "en-US,en;q=0.9",
                "Upgrade-Insecure-Requests": "1",
                "Sec-Fetch-Dest": "document",
                "Sec-Fetch-Mode": "navigate",
                "Sec-Fetch-Site": "none",
                "Sec-Fetch-User": "?1",
            })
            adapter = requests.adapters.HTTPAdapter(pool_connections=20, pool_maxsize=20)
            s.mount("http://", adapter)
            s.mount("https://", adapter)
            self._local.s = s
        return s

    def new_paper(self):
        """Fresh cookie jar for every paper."""
        self.session().cookies.clear()

    def _wait_turn(self, url, api):
        host = urllib.parse.urlsplit(url).netloc.lower()
        gap = API_HOST_GAP if api else PUBLISHER_HOST_GAP
        while True:
            with self._lock:
                t = time.monotonic()
                ready = self._next_ok.get(host, 0)
                if t >= ready:
                    self._next_ok[host] = t + gap
                    return
                wait = ready - t
            self._sleep(min(wait, 1.0))

    def _sleep(self, secs):
        end = time.monotonic() + secs
        while time.monotonic() < end:
            if self.stop_event.is_set():
                raise FetchError("stopped", "Stopped by user")
            time.sleep(min(0.2, end - time.monotonic()) if end > time.monotonic() else 0)

    def get(self, url, *, api=False, stream=False, accept=None, attempts=MAX_ATTEMPTS,
            referer=None, extra_headers=None):
        headers = {}
        if api:
            headers["Accept"] = accept or "application/json"
            ua = f"{APP_NAME}/{APP_VERSION}"
            if self.email:
                ua += f" (mailto:{self.email})"
            headers["User-Agent"] = ua
        else:
            headers["Accept"] = accept or ("application/pdf,text/html;q=0.9,"
                                           "application/xhtml+xml;q=0.8,*/*;q=0.5")
        if referer:
            headers["Referer"] = referer
        if extra_headers:
            headers.update(extra_headers)
        last = None
        for attempt in range(1, attempts + 1):
            if self.stop_event.is_set():
                raise FetchError("stopped", "Stopped by user")
            self._wait_turn(url, api)
            try:
                r = self.session().get(url, headers=headers, timeout=TIMEOUT,
                                       stream=stream, allow_redirects=True)
            except requests.exceptions.Timeout:
                last = FetchError("timeout", f"Timeout contacting {urllib.parse.urlsplit(url).netloc}")
            except requests.exceptions.SSLError as e:
                raise FetchError("network", f"{_host(url)}: SSL/certificate error ({conn_cause(e)})")
            except requests.exceptions.TooManyRedirects:
                raise FetchError("http", "Too many redirects")
            except requests.exceptions.ConnectionError as e:
                last = FetchError("network", f"{_host(url)}: {conn_cause(e)}")
                if attempt >= 2:        # refused/reset connections rarely heal within seconds
                    break
            except requests.exceptions.RequestException as e:
                raise FetchError("network", f"{_host(url)}: {conn_cause(e)}")
            else:
                code = r.status_code
                if code == 429 or code >= 500:
                    kind = "server"
                    last = FetchError(kind, f"HTTP {code} from {urllib.parse.urlsplit(r.url).netloc}", code)
                    wait = RETRY_WAIT * attempt
                    ra = r.headers.get("Retry-After", "")
                    if ra.isdigit():
                        wait = min(max(int(ra), wait), 30)
                    r.close()
                    if attempt < attempts:
                        self._sleep(wait)
                    continue
                return r
            if attempt < attempts:
                self._sleep(RETRY_WAIT * attempt)
        raise last

    def get_json(self, url, attempts=MAX_ATTEMPTS, extra_headers=None):
        """Returns (json_or_None, http_status). 404 -> (None, 404)."""
        r = self.get(url, api=True, attempts=attempts, extra_headers=extra_headers)
        try:
            if r.status_code == 200:
                try:
                    return r.json(), 200
                except ValueError:
                    return None, 200
            return None, r.status_code
        finally:
            r.close()


# --------------------------------------------------------------------------- #
#  Open-access discovery
# --------------------------------------------------------------------------- #
class Candidate:
    __slots__ = ("url", "source", "kind")

    def __init__(self, url, source, kind="pdf"):
        self.url = url
        self.source = source
        self.kind = kind   # pdf | landing


_PMC_RE = re.compile(r"(?:ncbi\.nlm\.nih\.gov/(?:pmc/)?articles/|pmc\.ncbi\.nlm\.nih\.gov/articles/)(?:PMC)?(\d+)", re.I)


def expand_candidates(cands):
    """Add well-known, openly served PDF endpoints for repository links."""
    out = []
    for c in cands:
        m = _PMC_RE.search(c.url)
        if m:
            pmcid = "PMC" + m.group(1)
            out.append(Candidate(f"https://europepmc.org/articles/{pmcid}?pdf=render",
                                 "Europe PMC (open access copy of " + pmcid + ")"))
            out.append(Candidate(f"https://pmc.ncbi.nlm.nih.gov/articles/{pmcid}/", c.source, "landing"))
            continue
        u = c.url
        if "arxiv.org/abs/" in u:
            out.append(Candidate(u.replace("/abs/", "/pdf/"), c.source))
        elif re.search(r"hal\.[a-z.]+/(hal|tel|inria|cea|halshs|insu|ird|pasteur|in2p3|lirmm|mnhn|ineris|sde|meteo)-\d+(v\d+)?/?$", u):
            out.append(Candidate(u.rstrip("/") + "/document", c.source))
        out.append(c)
    seen, res = set(), []
    for c in out:
        if c.url not in seen:
            seen.add(c.url)
            res.append(c)
    return res


class Finder:
    """Collects metadata + candidate open-access URLs for a DOI."""

    def __init__(self, http, email):
        self.http = http
        self.email = email

    # ---- metadata / primary sources ------------------------------------- #
    def crossref(self, doi, res, info):
        url = f"https://api.crossref.org/works/{urllib.parse.quote(doi, safe='')}"
        if self.email:
            url += f"?mailto={urllib.parse.quote(self.email)}"
        data, code = self.http.get_json(url)
        info["crossref_status"] = code
        if not data:
            return []
        msg = data.get("message", {})
        info["known"] = True
        if not res.title:
            res.title = " ".join((msg.get("title") or [""])[0].split())
        if not res.authors:
            names = []
            for a in msg.get("author", []) or []:
                n = ", ".join(x for x in (a.get("family"), a.get("given")) if x) or a.get("name", "")
                if n:
                    names.append(n)
            res.authors = "; ".join(names)
        if not res.journal:
            res.journal = (msg.get("container-title") or [""])[0]
        if not res.year:
            for k in ("published-print", "published-online", "issued", "created"):
                try:
                    res.year = str(msg[k]["date-parts"][0][0])
                    break
                except (KeyError, IndexError, TypeError):
                    continue
        if not res.publisher:
            res.publisher = msg.get("publisher", "") or ""
        for lic in msg.get("license", []) or []:
            if "creativecommons.org" in (lic.get("URL") or ""):
                info["oa_votes"].append(True)
                break
        cands = []
        for ln in msg.get("link", []) or []:
            u = ln.get("URL") or ""
            ct = (ln.get("content-type") or "").lower()
            if not u or any(b in u for b in BLOCKED_LINK_HOSTS):
                continue
            if ct in ("application/pdf", "unspecified"):
                cands.append(Candidate(u, "Crossref (publisher full-text link)",
                                       "pdf" if ct == "application/pdf" else "landing"))
        return cands

    def openalex(self, doi, res, info):
        url = f"https://api.openalex.org/works/doi:{urllib.parse.quote(doi, safe='/')}"
        if self.email:
            url += f"?mailto={urllib.parse.quote(self.email)}"
        data, code = self.http.get_json(url)
        info["openalex_status"] = code
        if not data:
            return []
        info["known"] = True
        oa = data.get("open_access") or {}
        if oa.get("is_oa") is not None:
            info["oa_votes"].append(bool(oa.get("is_oa")))
        if oa.get("oa_status") and not res.oa_status:
            res.oa_status = oa.get("oa_status")
        if not res.title:
            res.title = data.get("display_name") or data.get("title") or ""
        if not res.authors:
            res.authors = "; ".join(a.get("author", {}).get("display_name", "")
                                    for a in data.get("authorships", []) or [])
        src = ((data.get("primary_location") or {}).get("source") or {})
        if not res.journal:
            res.journal = src.get("display_name") or ""
        if not res.year and data.get("publication_year"):
            res.year = str(data.get("publication_year"))
        if not res.publisher:
            res.publisher = src.get("host_organization_name") or ""
        cands = []
        best = data.get("best_oa_location") or {}
        if best.get("pdf_url"):
            cands.append(Candidate(best["pdf_url"], "OpenAlex (best OA location)"))
        for loc in data.get("locations", []) or []:
            if loc.get("is_oa") and loc.get("pdf_url"):
                name = ((loc.get("source") or {}).get("display_name")) or "repository"
                cands.append(Candidate(loc["pdf_url"], f"OpenAlex ({name})"))
        for loc in [best] + (data.get("locations", []) or []):
            if loc and loc.get("is_oa") and loc.get("landing_page_url"):
                cands.append(Candidate(loc["landing_page_url"], "OpenAlex (OA landing page)", "landing"))
        return cands

    def unpaywall(self, doi, res, info):
        if not self.email:
            return []
        url = (f"https://api.unpaywall.org/v2/{urllib.parse.quote(doi, safe='/')}"
               f"?email={urllib.parse.quote(self.email)}")
        data, code = self.http.get_json(url)
        info["unpaywall_status"] = code
        if not data:
            if code == 422:
                info["notes"].append("Unpaywall rejected the e-mail address - enter a real e-mail")
            return []
        info["known"] = True
        if data.get("is_oa") is not None:
            info["oa_votes"].append(bool(data.get("is_oa")))
        if data.get("oa_status"):
            res.oa_status = data["oa_status"]
        if not res.title:
            res.title = data.get("title") or ""
        if not res.journal:
            res.journal = data.get("journal_name") or ""
        if not res.year and data.get("year"):
            res.year = str(data["year"])
        if not res.publisher:
            res.publisher = data.get("publisher") or ""
        cands = []
        best = data.get("best_oa_location") or {}
        locs = [best] + [l for l in (data.get("oa_locations") or []) if l is not best]
        for loc in locs:
            if loc and loc.get("url_for_pdf"):
                where = loc.get("host_type") or "location"
                cands.append(Candidate(loc["url_for_pdf"], f"Unpaywall ({where})"))
        for loc in locs:
            if loc and loc.get("url_for_landing_page"):
                cands.append(Candidate(loc["url_for_landing_page"], "Unpaywall (OA landing page)", "landing"))
        return cands

    # ---- secondary sources ---------------------------------------------- #
    def europepmc(self, doi, res, info):
        q = urllib.parse.quote(f'DOI:"{doi}"')
        url = (f"https://www.ebi.ac.uk/europepmc/webservices/rest/search?query={q}"
               f"&format=json&resultType=lite&pageSize=1")
        data, _ = self.http.get_json(url, attempts=2)
        if not data:
            return []
        hits = (data.get("resultList") or {}).get("result") or []
        cands = []
        for h in hits:
            if (h.get("doi") or "").lower() != doi.lower():
                continue
            pmcid = h.get("pmcid")
            if pmcid and h.get("isOpenAccess") == "Y":
                info["oa_votes"].append(True)
                cands.append(Candidate(f"https://europepmc.org/articles/{pmcid}?pdf=render",
                                       "Europe PMC (open access)"))
        return cands

    def semantic_scholar(self, doi, res, info):
        url = (f"https://api.semanticscholar.org/graph/v1/paper/DOI:{urllib.parse.quote(doi, safe='/')}"
               f"?fields=isOpenAccess,openAccessPdf")
        try:
            data, _ = self.http.get_json(url, attempts=1)
        except FetchError:
            return []
        if not data:
            return []
        pdf = (data.get("openAccessPdf") or {}).get("url")
        if pdf:
            return [Candidate(pdf, "Semantic Scholar (open-access PDF)")]
        return []

    def doi_registered(self, doi):
        """Ask the DOI handle system directly. True / False / None(unknown)."""
        try:
            data, code = self.http.get_json(
                f"https://doi.org/api/handles/{urllib.parse.quote(doi, safe='/')}", attempts=2)
        except FetchError:
            return None
        if code == 404:
            return False
        if data is not None:
            return data.get("responseCode") == 1
        return None


# --------------------------------------------------------------------------- #
#  PDF downloader + validation
# --------------------------------------------------------------------------- #
def extract_pdf_links(html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for name in ("citation_pdf_url", "eprints.document_url", "bepress_citation_pdf_url",
                 "wkhealth_pdf_url", "dc.identifier.pdf"):
        for m in soup.find_all("meta", attrs={"name": re.compile(f"^{re.escape(name)}$", re.I)}):
            if m.get("content"):
                links.append(urllib.parse.urljoin(base_url, m["content"].strip()))
    for ln in soup.find_all("link", attrs={"type": re.compile("application/pdf", re.I)}):
        if ln.get("href"):
            links.append(urllib.parse.urljoin(base_url, ln["href"].strip()))
    # Repository pages (DSpace, EPrints, OJS): explicit "download PDF" anchors
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        text = (a.get_text(" ", strip=True) or "").lower()
        h = href.lower()
        if (h.endswith(".pdf") or "/bitstream/" in h or "/download/" in h) and \
                ("pdf" in text or "download" in text or h.endswith(".pdf")):
            links.append(urllib.parse.urljoin(base_url, href))
        if len(links) > 8:
            break
    out, seen = [], set()
    for u in links:
        if u.startswith("http") and u not in seen:
            seen.add(u)
            out.append(u)
    return out[:5]


def looks_like_challenge(text):
    t = text[:20000].lower()
    return any(m in t for m in CHALLENGE_MARKERS)


class Downloader:
    def __init__(self, http):
        self.http = http

    def fetch_pdf(self, url, dest, depth=0, referer=None, extra_headers=None):
        """Download url -> dest if (and only if) it is a real PDF.
        Returns final URL. Raises FetchError otherwise."""
        r = self.http.get(url, stream=True, referer=referer, extra_headers=extra_headers)
        try:
            code = r.status_code
            final_url = r.url
            if code in (401, 402, 403, 407, 451):
                raise FetchError("denied", f"HTTP {code} (access restricted) at {_host(final_url)}", code)
            if code in (404, 410):
                raise FetchError("notfound", f"HTTP {code} (not found) at {_host(final_url)}", code)
            if code >= 400:
                raise FetchError("http", f"HTTP {code} at {_host(final_url)}", code)

            ctype = (r.headers.get("Content-Type") or "").lower()
            it = r.iter_content(chunk_size=65536)
            first = b""
            try:
                while len(first) < 2048:
                    chunk = next(it)
                    if not chunk:
                        continue
                    first += chunk
            except StopIteration:
                pass
            except requests.exceptions.RequestException as e:
                raise FetchError("timeout" if "timed out" in str(e).lower() else "network",
                                 f"Transfer interrupted: {str(e)[:100]}")

            if b"%PDF" in first[:1024]:
                return self._stream_to_file(r, it, first, dest, final_url, ctype)

            # Not a PDF: an HTML page may point to the real (open) PDF
            if "html" in ctype or first.lstrip()[:1] == b"<":
                body = first
                try:
                    for chunk in it:
                        body += chunk
                        if len(body) > MAX_HTML_BYTES:
                            break
                except requests.exceptions.RequestException:
                    pass
                text = body.decode(r.encoding or "utf-8", errors="replace")
                links = extract_pdf_links(text, final_url)
                if depth == 0 and links:
                    last_err = None
                    for link in links:
                        if link == url:
                            continue
                        try:
                            return self.fetch_pdf(link, dest, depth=1, referer=final_url)
                        except FetchError as e:
                            if e.kind == "stopped":
                                raise
                            last_err = e
                    if last_err:
                        raise last_err
                if looks_like_challenge(text):
                    raise FetchError("denied", f"Access check / CAPTCHA page at {_host(final_url)} "
                                               f"(not bypassed)")
                raise FetchError("notpdf", f"HTML page returned, no accessible PDF at {_host(final_url)}")
            raise FetchError("notpdf", f"Not a PDF (Content-Type: {ctype or 'unknown'}) at {_host(final_url)}")
        finally:
            r.close()

    def _stream_to_file(self, r, it, first, dest, final_url, ctype):
        tmp = dest + ".part"
        size = 0
        try:
            with open(tmp, "wb") as fh:
                fh.write(first)
                size = len(first)
                for chunk in it:
                    if self.http.stop_event.is_set():
                        raise FetchError("stopped", "Stopped by user")
                    if chunk:
                        fh.write(chunk)
                        size += len(chunk)
                        if size > MAX_PDF_BYTES:
                            raise FetchError("http", "File larger than 300 MB - skipped")
        except requests.exceptions.RequestException as e:
            _silent_remove(tmp)
            raise FetchError("timeout" if "timed out" in str(e).lower() else "network",
                             f"Transfer interrupted: {str(e)[:100]}")
        except FetchError:
            _silent_remove(tmp)
            raise
        except OSError as e:
            _silent_remove(tmp)
            raise FetchError("http", f"Cannot write file: {e}")

        # ---- validation ----
        expected = r.headers.get("Content-Length")
        if expected and expected.isdigit() and "gzip" not in (r.headers.get("Content-Encoding") or "") \
                and int(expected) != size:
            _silent_remove(tmp)
            raise FetchError("network", f"Incomplete download ({size} of {expected} bytes)")
        if size < MIN_PDF_BYTES:
            _silent_remove(tmp)
            raise FetchError("notpdf", f"PDF too small ({size} bytes) - probably not the paper")
        with open(tmp, "rb") as fh:
            fh.seek(max(0, size - 65536))
            tail = fh.read()
        if b"%%EOF" not in tail:
            _silent_remove(tmp)
            raise FetchError("network", "PDF appears truncated (no %%EOF marker)")
        os.replace(tmp, dest)
        return final_url


def _host(url):
    return urllib.parse.urlsplit(url).netloc


def conn_cause(e):
    """Turn a long urllib3 error into a short human reason."""
    s = str(e)
    low = s.lower()
    table = [
        (("nameresolution", "getaddrinfo", "name or service not known", "11001"),
         "DNS lookup failed - no internet or DNS blocked"),
        (("proxyerror", "tunnel connection failed", "407"), "proxy refused the connection"),
        (("10054", "connectionreset", "connection reset", "remotedisconnected",
          "connection aborted", "10053"), "connection reset by the site or a firewall"),
        (("10061", "refused"), "connection refused (site or firewall blocks it)"),
        (("10060", "timed out", "timeout"), "connection timed out"),
        (("certificate", "ssl"), "SSL/certificate problem"),
        (("10051", "10065", "unreachable"), "network unreachable"),
    ]
    for keys, msg in table:
        if any(k in low for k in keys):
            return msg
    m = re.search(r"Caused by (\w+)\(", s)
    return (m.group(1) if m else s)[:120]


def _silent_remove(p):
    try:
        os.remove(p)
    except OSError:
        pass


# --------------------------------------------------------------------------- #
#  Real-browser engine (your installed Microsoft Edge / Google Chrome)
# --------------------------------------------------------------------------- #
# Many publishers and repositories (ScienceDirect, PubMed Central, HAL, Wiley,
# Taylor & Francis, AMS, ...) only serve their *open* PDFs to a real browser:
# they use JavaScript redirects or an automatic "checking your browser" step and
# refuse plain HTTP clients. For those, the PDF is opened in the user's own
# browser exactly as if the user had clicked it.
#   * No CAPTCHA is ever clicked or solved by the program.
#   * No login, cookie or credential of the user is used (separate empty profile).
#   * If a site demands human verification the paper is reported as not accessible
#     (with "Show browser window" ticked the USER may complete it personally).

BROWSER_URL_TIMEOUT = 45        # seconds allowed per URL in the browser
BROWSER_IDLE_GIVEUP = 10        # seconds on a normal page without any PDF link
BROWSER_PAPER_BUDGET = 150      # seconds of browser time per paper at most

CAPTCHA_MARKERS = ("g-recaptcha", "recaptcha/api", "hcaptcha.com", "h-captcha",
                   "captcha-delivery", "px-captcha", "arkoselabs", "funcaptcha")
WAIT_MARKERS = ("just a moment", "cf-chl", "challenge-platform", "checking your browser",
                "checking if the site connection is secure", "making sure you", "anubis",
                "preparing to download", "please wait", "redirecting", "verifying you are human",
                "verify you are human", "ddos-guard", "one moment, please", "loading...")


def check_pdf_file(path):
    """None if the file is a complete PDF, otherwise the reason it is not."""
    try:
        size = os.path.getsize(path)
        with open(path, "rb") as fh:
            head = fh.read(1024)
            fh.seek(max(0, size - 65536))
            tail = fh.read()
    except OSError as e:
        return f"cannot read file ({e})"
    if b"%PDF" not in head:
        return "not a PDF file"
    if size < MIN_PDF_BYTES:
        return f"PDF too small ({size} bytes)"
    if b"%%EOF" not in tail:
        return "PDF incomplete (no %%EOF)"
    return None


def browser_pdf_links(html, base_url):
    """PDF links a person would click on an article page."""
    links = extract_pdf_links(html, base_url)
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if href.startswith(("javascript:", "#", "mailto:")):
            continue
        h = href.lower()
        text = (a.get_text(" ", strip=True) or "") + " " + (a.get("aria-label") or "") + " " + (a.get("title") or "")
        text = text.lower()
        good_href = any(k in h for k in ("pdfft", "/pdf/", "/pdfdirect/", "/epdf/", "download=true",
                                         "/content/pdf/", "blobtype=pdf", "/doi/pdf", ".pdf"))
        good_text = "pdf" in text and not any(k in text for k in ("supplement", "purchase", "buy", "rent"))
        if good_href and (good_text or h.endswith(".pdf")) or (good_text and "download" in text):
            u = urllib.parse.urljoin(base_url, href)
            u = u.replace("/doi/epdf/", "/doi/pdfdirect/")
            links.append(u)
    out, seen = [], set()
    for u in links:
        if u.startswith("http") and u not in seen:
            seen.add(u)
            out.append(u)
    return out[:6]


def page_state(html):
    t = html[:80000].lower()
    if any(m in t for m in CAPTCHA_MARKERS):
        return "captcha"
    if any(m in t for m in WAIT_MARKERS):
        return "wait"
    try:
        text = BeautifulSoup(html, "html.parser").get_text(" ", strip=True)
    except Exception:
        text = t
    if len(text) < 200:
        return "wait"       # page still being built by JavaScript
    return "page"


def _local_appdata():
    return os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or os.path.expanduser("~")


class BrowserFetcher:
    """Owns one Edge/Chrome instance in a dedicated thread (Playwright is not thread-safe)."""

    def __init__(self, stop_event, show_window=False, log=None):
        self.stop_event = stop_event
        self.show = show_window
        self.log = log or (lambda msg, lvl="info": None)
        self.jobs = queue.Queue()
        self.thread = None
        self.ready = threading.Event()
        self.error = None
        self.name = ""
        self.internal_errors = 0
        self.log_attempts = True
        self._lock = threading.Lock()

    # ---- public (any thread) -------------------------------------------- #
    def start(self):
        with self._lock:
            if self.thread is None:
                self.thread = threading.Thread(target=self._loop, daemon=True, name="browser")
                self.thread.start()
        self.ready.wait(120)
        if not self.ready.is_set():
            self.error = self.error or "browser did not start within 2 minutes"

    def fetch(self, url, dest):
        self.start()
        if self.error:
            raise FetchError("nobrowser", f"Browser unavailable: {self.error}")
        if self.log_attempts:
            self.log(f"   … opening {_host(url)} in browser", "info")
        done, box = threading.Event(), {}
        self.jobs.put((url, dest, done, box))
        while not done.wait(0.5):
            if self.thread and not self.thread.is_alive():
                raise FetchError("nobrowser", "Browser closed unexpectedly")
        if "err" in box:
            raise box["err"]
        return box["url"]

    def close(self):
        if self.thread and self.thread.is_alive():
            self.jobs.put(None)
            self.thread.join(15)

    # ---- browser thread --------------------------------------------------- #
    def _launch(self, pw):
        profile = os.path.join(_local_appdata(), "DOI_Paper_Downloader", "browser_profile")
        os.makedirs(os.path.join(profile, "Default"), exist_ok=True)
        prefs_path = os.path.join(profile, "Default", "Preferences")
        try:
            prefs = json.load(open(prefs_path, encoding="utf-8")) if os.path.exists(prefs_path) else {}
        except (OSError, ValueError):
            prefs = {}
        if not os.environ.get("DOI_DOWNLOADER_TEST_VIEWER"):
            prefs.setdefault("plugins", {})["always_open_pdf_externally"] = True   # PDFs -> download
        prefs.setdefault("download", {})["prompt_for_download"] = False
        try:
            with open(prefs_path, "w", encoding="utf-8") as fh:
                json.dump(prefs, fh)
        except OSError:
            pass
        self.dl_dir = tempfile.mkdtemp(prefix="doi_dl_")
        args = ["--no-first-run", "--no-default-browser-check", "--disable-popup-blocking"]
        if not self.show:
            args += ["--window-position=-2400,-2400", "--window-size=1280,900"]
        tries = []
        custom = os.environ.get("DOI_DOWNLOADER_BROWSER")
        if custom:
            tries.append(("custom browser", dict(executable_path=custom)))
        tries += [("Microsoft Edge", dict(channel="msedge")), ("Google Chrome", dict(channel="chrome")),
                  ("Chromium", dict())]
        errors = []
        for name, kw in tries:
            try:
                ctx = pw.chromium.launch_persistent_context(
                    profile, headless=False, args=args, accept_downloads=True,
                    downloads_path=self.dl_dir, viewport={"width": 1280, "height": 860}, **kw)
                self.name = name
                return ctx
            except Exception as e:
                errors.append(f"{name}: {str(e).splitlines()[0][:100]}")
        raise RuntimeError("no usable browser found (install Microsoft Edge or Google Chrome). "
                           + " | ".join(errors))

    def _loop(self):
        try:
            from playwright.sync_api import sync_playwright
        except Exception:
            self.error = "browser support (playwright) is not installed"
            self.ready.set()
            return
        pw = ctx = None
        try:
            pw = sync_playwright().start()
            ctx = self._launch(pw)
        except Exception as e:
            self.error = str(e)[:400]
            self.ready.set()
            if pw:
                try:
                    pw.stop()
                except Exception:
                    pass
            return
        self.ctx = ctx
        self.ready.set()
        self.log(f"Browser engine ready: {self.name}", "info")
        while True:
            job = self.jobs.get()
            if job is None:
                break
            url, dest, done, box = job
            try:
                box["url"] = self._fetch(url, dest)
            except FetchError as e:
                box["err"] = e
            except Exception as e:
                msg = str(e).splitlines()[0][:160] if str(e) else type(e).__name__
                write_error_log(f"browser error on {url}")
                self.internal_errors += 1
                if self.internal_errors >= 5 or "closed" in msg.lower():
                    self.error = f"browser stopped working ({msg})"
                box["err"] = FetchError("network", f"{_host(url)}: browser error ({msg})")
            finally:
                done.set()
        for closer in (lambda: ctx.close(), lambda: pw.stop(),
                       lambda: shutil.rmtree(self.dl_dir, ignore_errors=True)):
            try:
                closer()
            except Exception:
                pass

    def _goto(self, page, url):
        try:
            resp = page.goto(url, wait_until="commit", timeout=BROWSER_URL_TIMEOUT * 1000)
        except Exception as e:
            msg = str(e)
            low = msg.lower()
            if "download is starting" in low or "err_aborted" in low:
                return None                      # the PDF is being downloaded
            if "timeout" in low:
                raise FetchError("timeout", f"{_host(url)}: page did not load in time")
            if "err_name_not_resolved" in low:
                raise FetchError("network", f"{_host(url)}: DNS lookup failed")
            m = re.search(r"net::(ERR_[A-Z_]+)", msg)
            raise FetchError("network", f"{_host(url)}: {m.group(1) if m else msg.splitlines()[0][:100]}")
        if resp is not None and resp.status in (404, 410):
            raise FetchError("notfound", f"HTTP {resp.status} at {_host(url)}")
        return resp

    def _save(self, src_path, dest):
        why = check_pdf_file(src_path)
        if why:
            raise FetchError("notpdf", why)
        tmp = dest + ".part"
        shutil.copyfile(src_path, tmp)
        os.replace(tmp, dest)

    def _page_fetch(self, page, url, dest):
        """Read a PDF with the browser's own fetch() (same site, same cookies)."""
        try:
            page.wait_for_timeout(300)
            b64 = page.evaluate("""async (u) => {
                const r = await fetch(u, {credentials: 'include'});
                if (!r.ok) return 'ERR:' + r.status;
                const b = new Uint8Array(await r.arrayBuffer());
                let s = '';
                for (let i = 0; i < b.length; i += 0x8000)
                    s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
                return btoa(s);
            }""", url)
        except Exception:
            return None
        if not b64 or b64.startswith("ERR:"):
            return None
        import base64
        data = base64.b64decode(b64)
        if b"%PDF" not in data[:1024]:
            return None
        tmp = os.path.join(self.dl_dir, "viewer.pdf")
        with open(tmp, "wb") as fh:
            fh.write(data)
        self._save(tmp, dest)
        return url

    def _fetch(self, url, dest):
        page = self.ctx.new_page()
        downloads, pdf_responses = [], []
        page.on("download", lambda d: downloads.append(d))

        pdf_urls = []

        def on_response(r):
            try:
                if "pdf" in (r.headers.get("content-type") or "").lower() and \
                        r.request.resource_type in ("document", "other") and r.status < 400:
                    pdf_responses.append(r)
                    pdf_urls.append(r.url)
            except Exception:
                pass
        page.on("response", on_response)
        try:
            self._goto(page, url)
            visited = {url}
            hops = 0
            last_nav = time.monotonic()
            deadline = last_nav + BROWSER_URL_TIMEOUT
            saw_captcha = saw_wait = False
            while True:
                if self.stop_event.is_set():
                    raise FetchError("stopped", "Stopped by user")
                # 1) a download started (PDF opened "externally")
                if downloads:
                    d = downloads[0]
                    path = d.path()                         # waits for completion
                    if d.failure() or not path:
                        raise FetchError("network", f"{_host(d.url)}: download failed ({d.failure()})")
                    self._save(path, dest)
                    return d.url
                # 2) a PDF displayed inline
                for r in list(pdf_responses):
                    pdf_responses.remove(r)
                    try:
                        body = r.body()
                    except Exception:
                        continue
                    if b"%PDF" in body[:1024]:
                        tmp = os.path.join(self.dl_dir, "inline.pdf")
                        with open(tmp, "wb") as fh:
                            fh.write(body)
                        self._save(tmp, dest)
                        return r.url
                # 3) the browser shows the PDF in its built-in viewer: read it through the page
                if pdf_urls:
                    target = pdf_urls[-1]
                    pdf_urls.clear()
                    got = self._page_fetch(page, target, dest)
                    if got:
                        return got
                now = time.monotonic()
                if now > deadline:
                    if saw_captcha:
                        raise FetchError("denied", f"{_host(page.url)} requires human verification "
                                                   f"(CAPTCHA) - not bypassed")
                    if saw_wait:
                        raise FetchError("denied", f"{_host(page.url)} kept showing a browser check "
                                                   f"page - not bypassed")
                    raise FetchError("notpdf", f"No accessible PDF at {_host(page.url)}")
                try:
                    html, cur = page.content(), page.url
                except Exception:
                    page.wait_for_timeout(400)              # navigation in progress
                    continue
                low_html = html[:5000].lower()
                if "<embed" in low_html and "application/pdf" in low_html:     # PDF viewer
                    got = self._page_fetch(page, cur, dest)
                    if got:
                        return got
                    raise FetchError("notpdf", f"PDF viewer at {_host(cur)} but file could not be read")
                state = page_state(html)
                if state == "captcha":
                    if not saw_captcha:
                        saw_captcha = True
                        if self.show:
                            self.log(f"Human verification shown at {_host(cur)} - you may complete it "
                                     f"in the browser window (waiting up to 2 minutes)", "warn")
                            deadline = max(deadline, now + 120)
                        else:
                            deadline = min(deadline, now + 5)
                elif state == "wait":
                    saw_wait = True
                elif now - last_nav > 2.5:                  # let JavaScript redirects happen first
                    links = [l for l in browser_pdf_links(html, cur) if l not in visited]
                    if links and hops < 4:
                        hops += 1
                        visited.add(links[0])
                        last_nav = time.monotonic()
                        deadline = max(deadline, last_nav + BROWSER_URL_TIMEOUT)
                        self._goto(page, links[0])
                        continue
                    if now - last_nav > BROWSER_IDLE_GIVEUP:
                        raise FetchError("notpdf", f"No accessible PDF at {_host(cur)}")
                page.wait_for_timeout(500)
        finally:
            try:
                page.close()
            except Exception:
                pass


def connection_test(email="", use_browser=True, show=False, log=print):
    """Checks every service the program relies on and reports the exact problem."""
    stop = threading.Event()
    http = PoliteHttp(email, stop)
    tests = [("OpenAlex", "https://api.openalex.org/works/doi:10.1371/journal.pone.0000308"),
             ("Crossref", "https://api.crossref.org/works/10.1371/journal.pone.0000308"),
             ("Unpaywall", f"https://api.unpaywall.org/v2/10.1371/journal.pone.0000308?email={email or 'x'}"),
             ("doi.org", "https://doi.org/api/handles/10.1371/journal.pone.0000308"),
             ("Europe PMC", "https://www.ebi.ac.uk/europepmc/webservices/rest/search?query=cancer&format=json&pageSize=1"),
             ("HAL repository", "https://hal.science/"),
             ("Publisher (PLOS)", "https://journals.plos.org/plosone/")]
    ok_all = True
    for name, url in tests:
        try:
            r = http.get(url, api="api" in url or "handles" in url or "webservices" in url, attempts=1)
            code = r.status_code
            r.close()
            good = code < 500 and code != 407
            log(f"{'OK ' if good else 'BAD'}  {name:<18} HTTP {code}", "ok" if good else "bad")
            ok_all &= good
        except FetchError as e:
            ok_all = False
            log(f"BAD  {name:<18} {e}", "bad")
    if use_browser:
        bf = BrowserFetcher(stop, show_window=show, log=log)
        bf.start()
        if bf.error:
            log(f"BAD  Browser engine       {bf.error}", "bad")
        else:
            tmp = os.path.join(tempfile.gettempdir(), "doi_dl_selftest.pdf")
            try:
                bf.fetch("https://europepmc.org/articles/PMC3517176?pdf=render", tmp)
                log(f"OK   Browser engine       {bf.name} - test PDF downloaded", "ok")
            except FetchError as e:
                log(f"WARN Browser engine       {bf.name} started, test PDF failed: {e}", "warn")
            finally:
                _silent_remove(tmp)
        bf.close()
    log("Connection test finished." + ("" if ok_all else
        " Some services are blocked on this network - try another network (e.g. mobile hotspot)"
        " or ask IT to allow them."), "info")
    return ok_all


# --------------------------------------------------------------------------- #
#  Engine
# --------------------------------------------------------------------------- #
@dataclass
class Options:
    excel_path: str
    output_dir: str
    sheet: str = None
    record_col: str = None
    doi_col: str = None
    email: str = ""
    redownload: bool = False
    recheck_unavailable: bool = False
    workers: int = 3
    elsevier_key: str = ""          # free key from dev.elsevier.com (official Elsevier API)
    wiley_token: str = ""           # free Wiley TDM token (official Wiley API)
    use_browser: bool = True        # use installed Edge/Chrome for sites that need a real browser
    show_browser: bool = False      # show that browser window (lets the USER pass a human check)
    institutional: bool = False     # also fetch papers the user's network is subscribed to


class Engine:
    """Runs the whole job. Communicates through callback(event, payload)."""

    def __init__(self, opts: Options, callback=None):
        self.opts = opts
        self.cb = callback or (lambda *a: None)
        self.stop_event = threading.Event()
        self.http = PoliteHttp(opts.email.strip(), self.stop_event)
        self.finder = Finder(self.http, opts.email.strip())
        self.downloader = Downloader(self.http)
        self.browser = BrowserFetcher(self.stop_event, opts.show_browser,
                                      log=lambda m, lvl="info": self.emit("log", (m, lvl))) \
            if opts.use_browser else None
        self._browser_warned = False
        self._lock = threading.Lock()
        self._log_lock = threading.Lock()
        self.results = {}
        self.state = {}
        self.counts = {}

    # ---- plumbing -------------------------------------------------------- #
    def stop(self):
        self.stop_event.set()

    def emit(self, event, payload=None):
        try:
            self.cb(event, payload)
        except Exception:
            pass

    def log_line(self, record_id, status, detail=""):
        line = f"{now_str()} | {record_id} | {status}"
        if detail:
            line += f" | {detail}"
        with self._log_lock:
            try:
                with open(os.path.join(self.opts.output_dir, LOG_FILE), "a", encoding="utf-8") as fh:
                    fh.write(line + "\n")
            except OSError:
                pass

    def _state_path(self):
        return os.path.join(self.opts.output_dir, STATE_FILE)

    def _load_state(self):
        try:
            with open(self._state_path(), "r", encoding="utf-8") as fh:
                return json.load(fh)
        except (OSError, ValueError):
            return {}

    def _save_state(self):
        tmp = self._state_path() + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as fh:
                json.dump(self.state, fh, ensure_ascii=False, indent=0)
            os.replace(tmp, self._state_path())
        except OSError:
            pass

    # ---- input ----------------------------------------------------------- #
    def read_rows(self):
        tables = load_workbook_tables(self.opts.excel_path)
        if not tables:
            raise ValueError("The Excel file contains no data.")
        sheet = self.opts.sheet if self.opts.sheet in tables else pick_default_sheet(tables)
        df, rc, dc = tables[sheet]
        rc = self.opts.record_col if self.opts.record_col in df.columns else rc
        dc = self.opts.doi_col if self.opts.doi_col in df.columns else dc
        if not dc:
            raise ValueError(
                f"Could not find a DOI column in sheet '{sheet}'.\n\n"
                f"Columns found: {', '.join(map(str, df.columns))}\n\n"
                "Please name the column 'DOI' or pick it from the DOI Column list.")
        if not rc:
            raise ValueError(
                f"Could not find a Record_ID column in sheet '{sheet}'.\n\n"
                f"Columns found: {', '.join(map(str, df.columns))}\n\n"
                "Please name the column 'Record_ID' or pick it from the Record ID Column list.")
        rows = []
        for i, r in df.iterrows():
            rows.append((i + 1, clean_record_id(r[rc]), str(r[dc]).strip()))
        return rows, sheet, rc, dc

    # ---- one paper ------------------------------------------------------- #
    def _target_name(self, res, used):
        base = safe_filename(res.record_id) if res.record_id else \
            safe_filename((res.doi or f"row_{res.row}").replace("/", "_"))
        name = base
        n = 2
        while name.lower() in used:
            name = f"{base}_{n}"
            n += 1
        used.add(name.lower())
        return name + ".pdf"

    def process(self, res: PaperResult, filename):
        """Fills in res; never raises."""
        try:
            self._process(res, filename)
        except FetchError as e:
            if e.kind == "stopped":
                res.status = S_NOT_PROCESSED
                res.error = "Stopped by user"
            else:
                res.status = S_TIMEOUT if e.kind == "timeout" else S_SERVER if e.kind == "server" else S_FAILED
                res.error = str(e)
        except Exception as e:  # absolutely never crash the batch
            res.status = S_FAILED
            res.error = f"Unexpected error: {type(e).__name__}: {e} (details in error_log.txt)"
            write_error_log(f"record {res.record_id} {res.doi_input}")
        res.checked_at = now_str()
        return res

    def _process(self, res: PaperResult, filename):
        dest = os.path.join(self.opts.output_dir, filename)
        doi, problem = normalize_doi(res.doi_input)
        res.doi = doi or ""
        if problem == "missing":
            res.status, res.error = S_MISSING, "No DOI in the Excel row"
            return
        if problem == "invalid":
            res.status, res.error = S_INVALID, f"'{res.doi_input}' is not a valid DOI"
            return

        # Already on disk?
        if os.path.exists(dest) and not self.opts.redownload:
            if is_valid_pdf_file(dest):
                res.status = S_EXISTS
                res.downloaded_file = filename
                prev = self.state.get(res.key()) or {}
                for k in ("source", "source_url", "title", "authors", "journal", "year",
                          "publisher", "oa_status"):
                    if prev.get(k):
                        setattr(res, k, prev[k])
                if not res.title:
                    try:
                        self.finder.crossref(doi, res, {"oa_votes": [], "notes": []})
                    except FetchError:
                        pass
                return
            res.notes.append("Existing file was not a valid PDF - replaced")

        self.http.new_paper()
        info = {"oa_votes": [], "notes": [], "known": False}
        failures = []          # FetchError list from download attempts (OA sources)
        landing_failures = []  # failures from the plain DOI landing-page fallback
        lookup_errors = []     # FetchError list from API lookups
        tried = set()

        def run_lookup(fn):
            try:
                return fn(doi, res, info)
            except FetchError as e:
                if e.kind == "stopped":
                    raise
                lookup_errors.append(e)
                return []
            except Exception:                   # a bad answer from one source must not stop the others
                write_error_log(f"lookup {getattr(fn, '__name__', fn)} for {doi}")
                return []

        def try_candidates(cands, sink=None):
            sink = failures if sink is None else sink
            for c in cands:
                if c.url in tried:
                    continue
                tried.add(c.url)
                self.emit("attempt", (res.record_id, c.source, c.url))
                try:
                    final = self.downloader.fetch_pdf(c.url, dest)
                except FetchError as e:
                    if e.kind == "stopped":
                        raise
                    sink.append(e)
                    continue
                res.status = S_DOWNLOADED
                res.downloaded_file = os.path.basename(dest)
                res.source = c.source
                res.source_url = final
                self.emit("pdf_found", (res.record_id, c.source))
                return True
            return False

        # Stage 1: main open-access indexes (+ metadata)
        cands = []
        cands += run_lookup(self.finder.unpaywall)
        cands += run_lookup(self.finder.openalex)
        cands += run_lookup(self.finder.crossref)
        votes = info["oa_votes"]
        is_oa = True if any(votes) else (False if votes else None)
        if is_oa is False and not self.opts.institutional:
            # publisher links of a subscription paper = subscription access, not open access
            cands = [c for c in cands if not c.source.startswith("Crossref")]
        cands = expand_candidates(cands)
        ordered = [c for c in cands if c.kind == "pdf"] + [c for c in cands if c.kind == "landing"]
        if try_candidates(ordered):
            return

        # DOI existence check when nobody knows it
        if not info["known"]:
            reg = self.finder.doi_registered(doi)
            if reg is False:
                res.status, res.error = S_INVALID, "DOI is not registered (doi.org returned 'not found')"
                return

        # Stage 1b: official publisher APIs (legal full text, needs the user's own free key)
        if self._publisher_api(doi, res, dest, is_oa, info, failures):
            return

        # Stage 2: secondary open repositories
        extra = expand_candidates(run_lookup(self.finder.europepmc) + run_lookup(self.finder.semantic_scholar))
        if try_candidates(extra):
            return
        votes = info["oa_votes"]
        is_oa = True if any(votes) else (False if votes else None)

        # Stage 3: DOI -> publisher page, plain HTTP (only for open / unknown papers)
        allow_publisher = is_oa is not False or self.opts.institutional
        pub_label = "Publisher website (via DOI)" if is_oa is not False else \
            "Publisher website (institutional access)"
        doi_landing = Candidate(os.environ.get("DOI_DOWNLOADER_RESOLVER", "https://doi.org/") + doi,
                                pub_label, "landing")
        if allow_publisher and try_candidates([doi_landing], sink=landing_failures):
            return

        # Stage 4: the same links in a real browser (JavaScript pages, browser checks)
        browser_failures = []
        if self.browser is not None:
            targets, seen = [], set()
            pool_ = [c for c in ordered + extra if c.kind == "pdf"] + \
                    [c for c in ordered + extra if c.kind == "landing"]
            if allow_publisher:
                pool_.append(doi_landing)
            for c in pool_:
                if c.url in seen:
                    continue
                seen.add(c.url)
                targets.append(c)
            t_end = time.monotonic() + BROWSER_PAPER_BUDGET
            for c in targets[:6]:
                if time.monotonic() > t_end:
                    break
                self.emit("attempt", (res.record_id, c.source + " [browser]", c.url))
                try:
                    final = self.browser.fetch(c.url, dest)
                except FetchError as e:
                    if e.kind == "stopped":
                        raise
                    if e.kind == "nobrowser":
                        if not self._browser_warned:
                            self._browser_warned = True
                            self.emit("log", (str(e) + " - continuing without it", "warn"))
                        info["notes"].append(str(e))
                        break
                    browser_failures.append(e)
                    continue
                res.status = S_DOWNLOADED
                res.downloaded_file = os.path.basename(dest)
                res.source = c.source + " [via browser]"
                res.source_url = final
                self.emit("pdf_found", (res.record_id, res.source))
                return

        if browser_failures:
            failures = browser_failures
        elif not failures:
            failures = landing_failures

        # ---- nothing worked: classify ----
        temp = [e for e in failures + lookup_errors if e.kind in ("timeout", "server", "network")]
        details = "; ".join(dict.fromkeys(str(e) for e in failures))[:600]
        notes = "; ".join(info["notes"])
        if is_oa is None and not info["known"] and lookup_errors and len(temp) == len(failures + lookup_errors):
            kinds = [e.kind for e in temp]
            res.status = S_TIMEOUT if kinds.count("timeout") * 2 >= len(kinds) else S_SERVER
            res.error = "Lookup services unreachable (check internet/firewall): " + "; ".join(dict.fromkeys(str(e) for e in temp))[:400]
        elif is_oa:
            kinds = [e.kind for e in failures]
            if kinds and all(k in ("timeout", "server", "network") for k in kinds):
                res.status = S_TIMEOUT if kinds.count("timeout") * 2 >= len(kinds) else S_SERVER
            elif any(k in ("denied", "network", "timeout", "server") for k in kinds):
                res.status = S_FAILED
            else:
                res.status = S_NOT_FOUND
            res.error = ("Open access, but the PDF could not be retrieved: " + details).strip()
        else:
            res.status = S_NOT_OA
            res.error = "No legal open-access PDF found (subscription paper - not bypassed)."
            if details:
                res.error += " Tried: " + details
        if notes:
            res.error = (res.error + " | " + notes).strip(" |")

    def _publisher_api(self, doi, res, dest, is_oa, info, failures):
        pub = (res.publisher or "").lower()
        allowed = is_oa is not False or self.opts.institutional
        if not allowed:
            return False
        key = (self.opts.elsevier_key or "").strip()
        if key and ("elsevier" in pub or doi.startswith("10.1016/")):
            q = urllib.parse.quote(doi, safe="/")
            hdr = {"X-ELS-APIKey": key}
            try:
                ok = is_oa is True
                if not ok:
                    data, _ = self.http.get_json(
                        f"https://api.elsevier.com/content/article/entitlement/doi/{q}?httpAccept=application/json",
                        attempts=2, extra_headers=hdr)
                    ent = ((data or {}).get("entitlement-response") or {}).get("document-entitlement") or {}
                    ok = str(ent.get("entitled")).lower() == "true"
                if ok:
                    url = f"https://api.elsevier.com/content/article/doi/{q}?httpAccept=application/pdf"
                    final = self.downloader.fetch_pdf(url, dest, depth=1,
                                                      extra_headers=dict(hdr, Accept="application/pdf"))
                    res.status, res.downloaded_file = S_DOWNLOADED, os.path.basename(dest)
                    res.source, res.source_url = "Elsevier Article API (official)", final.split("?")[0]
                    return True
            except FetchError as e:
                if e.kind == "stopped":
                    raise
                failures.append(e)
        tok = (self.opts.wiley_token or "").strip()
        if tok and ("wiley" in pub or doi.startswith(("10.1002/", "10.1111/", "10.1029/"))):
            url = f"https://api.wiley.com/onlinelibrary/tdm/v1/articles/{urllib.parse.quote(doi, safe='')}"
            try:
                final = self.downloader.fetch_pdf(url, dest, depth=1,
                                                  extra_headers={"Wiley-TDM-Client-Token": tok})
                res.status, res.downloaded_file = S_DOWNLOADED, os.path.basename(dest)
                res.source, res.source_url = "Wiley TDM API (official)", final
                return True
            except FetchError as e:
                if e.kind == "stopped":
                    raise
                failures.append(e)
        return False

    # ---- batch ----------------------------------------------------------- #
    def run(self):
        os.makedirs(self.opts.output_dir, exist_ok=True)
        if self.opts.output_dir not in ERROR_LOG_DIR:
            ERROR_LOG_DIR.insert(0, self.opts.output_dir)
        rows, sheet, rc, dc = self.read_rows()
        self.state = self._load_state()
        total = len(rows)
        self.emit("start", {"total": total, "sheet": sheet, "record_col": rc, "doi_col": dc})
        self.log_line("-", "RUN START",
                      f"{os.path.basename(self.opts.excel_path)} [{sheet}] {total} records -> {self.opts.output_dir}")

        used = set()
        jobs = []
        ordered_results = []
        for rownum, rid, doi_raw in rows:
            res = PaperResult(row=rownum, record_id=rid, doi_input=doi_raw)
            fname = self._target_name(res, used)
            ordered_results.append(res)
            # resume: reuse final "no legal PDF" / invalid verdicts from earlier runs
            prev = self.state.get(res.key())
            dest = os.path.join(self.opts.output_dir, fname)
            if prev and prev.get("engine") == ENGINE_VERSION and not self.opts.redownload \
                    and not os.path.exists(dest):
                pst = prev.get("status")
                reuse = pst in (S_MISSING, S_INVALID) or \
                    (pst in NOT_ACCESSIBLE_STATUSES and not self.opts.recheck_unavailable)
                if reuse:
                    for k, v in prev.items():
                        if hasattr(res, k) and k not in ("row", "notes"):
                            setattr(res, k, v)
                    res.row = rownum
                    res.from_previous_run = True
                    self._finish(res, total, reused=True)
                    continue
            jobs.append((res, fname))

        def work(job):
            res, fname = job
            if self.stop_event.is_set():
                res.status, res.error = S_NOT_PROCESSED, "Stopped by user"
                self._finish(res, total)
                return
            self.emit("current", (res.record_id, res.doi_input))
            self.process(res, fname)
            self._finish(res, total)

        workers = max(1, min(int(self.opts.workers or 1), 6))
        try:
            with ThreadPoolExecutor(max_workers=workers) as pool:
                list(pool.map(work, jobs))
        finally:
            if self.browser is not None:
                self.browser.close()

        for res in ordered_results:
            if not res.status:
                res.status, res.error = S_NOT_PROCESSED, "Stopped by user"
        report = self.write_report(ordered_results)
        summary = self.summary(ordered_results)
        self.log_line("-", "RUN END", ", ".join(f"{k}: {v}" for k, v in summary.items()))
        self.emit("done", {"summary": summary, "report": report, "stopped": self.stop_event.is_set()})
        return ordered_results, summary, report

    def _finish(self, res, total, reused=False):
        with self._lock:
            self.results[res.row] = res
            if res.status != S_NOT_PROCESSED:
                d = asdict(res)
                d.pop("notes", None)
                d.pop("from_previous_run", None)
                d["engine"] = ENGINE_VERSION
                self.state[res.key()] = d
                self._save_state()
            done = len(self.results)
            ordered = [self.results[k] for k in sorted(self.results)]
            if done % 10 == 0:
                try:
                    self.write_report(ordered, partial=True)
                except Exception:
                    pass
        detail = res.source if res.status == S_DOWNLOADED else res.error
        if reused:
            detail = "(result from previous run) " + (detail or "")
        if res.status != S_NOT_PROCESSED:
            self.log_line(res.record_id or f"row {res.row}", res.status, (detail or "")[:300])
        self.emit("result", {"result": res, "done": done, "total": total, "reused": reused})

    # ---- report ---------------------------------------------------------- #
    @staticmethod
    def summary(results):
        out = {"Total": len(results)}
        for s in (S_DOWNLOADED, S_EXISTS, S_NOT_OA, S_NOT_FOUND, S_FAILED, S_SERVER,
                  S_TIMEOUT, S_MISSING, S_INVALID, S_NOT_PROCESSED):
            out[s] = sum(1 for r in results if r.status == s)
        return out

    def write_report(self, results, partial=False):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter

        cols = [("Record_ID", "record_id", 14), ("DOI", "doi_input", 32), ("Status", "status", 17),
                ("Downloaded_File", "downloaded_file", 18), ("Source", "source", 30),
                ("Source_URL", "source_url", 45), ("Title", "title", 50), ("Authors", "authors", 35),
                ("Journal", "journal", 30), ("Year", "year", 7), ("Publisher", "publisher", 25),
                ("OA_Status", "oa_status", 11), ("Error", "error", 60), ("Checked_At", "checked_at", 19)]
        fills = {
            S_DOWNLOADED: "C6EFCE", S_EXISTS: "DDEBF7", S_NOT_OA: "FCE4D6", S_NOT_FOUND: "FCE4D6",
            S_FAILED: "FFC7CE", S_SERVER: "FFC7CE", S_TIMEOUT: "FFC7CE", S_MISSING: "EDEDED",
            S_INVALID: "EDEDED", S_NOT_PROCESSED: "FFF2CC",
        }
        wb = Workbook()
        ws = wb.active
        ws.title = "Download Status"
        hdr_font = Font(bold=True, color="FFFFFF")
        hdr_fill = PatternFill("solid", fgColor="1F4E78")
        for j, (h, _, w) in enumerate(cols, 1):
            c = ws.cell(row=1, column=j, value=h)
            c.font, c.fill = hdr_font, hdr_fill
            c.alignment = Alignment(vertical="center")
            ws.column_dimensions[get_column_letter(j)].width = w
        for i, r in enumerate(results, 2):
            for j, (_, attr, _) in enumerate(cols, 1):
                v = getattr(r, attr, "")
                if isinstance(v, str) and v[:1] in ("=", "+", "-", "@"):
                    v = "'" + v  # avoid formula injection
                cell = ws.cell(row=i, column=j, value=v)
                if attr == "downloaded_file" and v:
                    cell.hyperlink = v
                    cell.font = Font(color="0563C1", underline="single")
                elif attr == "source_url" and v:
                    cell.hyperlink = v
                    cell.font = Font(color="0563C1", underline="single")
            f = fills.get(r.status)
            if f:
                ws.cell(row=i, column=3).fill = PatternFill("solid", fgColor=f)
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(1, len(results) + 1)}"

        s = wb.create_sheet("Summary")
        s["A1"], s["A1"].font = f"{APP_NAME} - Summary", Font(bold=True, size=14)
        s["A2"] = f"Excel: {self.opts.excel_path}"
        s["A3"] = f"Generated: {now_str()}" + ("  (IN PROGRESS / PARTIAL)" if partial else "")
        row = 5
        s.cell(row=row, column=1, value="Status").font = Font(bold=True)
        s.cell(row=row, column=2, value="Count").font = Font(bold=True)
        for k, v in self.summary(results).items():
            row += 1
            s.cell(row=row, column=1, value=k)
            s.cell(row=row, column=2, value=v)
            if k in fills:
                s.cell(row=row, column=1).fill = PatternFill("solid", fgColor=fills[k])
        row += 2
        s.cell(row=row, column=1, value="Status meanings").font = Font(bold=True)
        meanings = [
            (S_DOWNLOADED, "Legal open PDF found, validated and saved as <Record_ID>.pdf"),
            (S_EXISTS, "File already in the output folder - not downloaded again"),
            (S_NOT_OA, "No legal open-access version exists / publisher restricts access"),
            (S_NOT_FOUND, "Listed as open access, but no downloadable PDF file could be located"),
            (S_FAILED, "A PDF link exists but the download was refused or broke - run again later"),
            (S_SERVER, "Server busy / error - temporary, retried automatically on next run"),
            (S_TIMEOUT, "Server did not respond in time - temporary, retried on next run"),
            (S_MISSING, "DOI cell is empty"),
            (S_INVALID, "DOI text is malformed or not registered at doi.org"),
            (S_NOT_PROCESSED, "Run was stopped before this record - processed on next run"),
        ]
        for k, v in meanings:
            row += 1
            s.cell(row=row, column=1, value=k)
            s.cell(row=row, column=2, value=v)
        s.column_dimensions["A"].width = 22
        s.column_dimensions["B"].width = 90

        path = os.path.join(self.opts.output_dir, STATUS_XLSX)
        try:
            wb.save(path)
        except PermissionError:
            # usually: the report is open in Excel
            alt = os.path.join(self.opts.output_dir,
                               f"Download_Status_{_dt.datetime.now():%Y%m%d_%H%M%S}.xlsx")
            wb.save(alt)
            if not partial:
                self.emit("log", (f"{STATUS_XLSX} is open in Excel - report saved as {os.path.basename(alt)}", "warn"))
            path = alt
        return path


# --------------------------------------------------------------------------- #
#  Settings
# --------------------------------------------------------------------------- #
def settings_path():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, ".doi_paper_downloader.json")


def load_settings():
    try:
        with open(settings_path(), "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def save_settings(d):
    try:
        with open(settings_path(), "w", encoding="utf-8") as fh:
            json.dump(d, fh, indent=1)
    except OSError:
        pass


# --------------------------------------------------------------------------- #
#  GUI
# --------------------------------------------------------------------------- #
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinter.scrolledtext import ScrolledText

    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    class App:
        def __init__(self, root):
            self.root = root
            self.q = queue.Queue()
            self.engine = None
            self.worker = None
            self.tables = {}
            self.settings = load_settings()
            self.counts = {}
            self.total = 0

            root.title(f"{APP_NAME} {APP_VERSION}")
            root.geometry("920x870")
            root.minsize(760, 640)
            style = ttk.Style()
            try:
                style.theme_use("vista" if sys.platform.startswith("win") else "clam")
            except tk.TclError:
                pass
            style.configure("Title.TLabel", font=("Segoe UI", 17, "bold"), foreground="#1F4E78")
            style.configure("Big.TButton", font=("Segoe UI", 12, "bold"), padding=10)
            style.configure("Stat.TLabel", font=("Segoe UI", 10))
            style.configure("StatNum.TLabel", font=("Segoe UI", 11, "bold"))

            outer = ttk.Frame(root, padding=14)
            outer.pack(fill="both", expand=True)
            ttk.Label(outer, text="DOI PAPER DOWNLOADER", style="Title.TLabel").pack(anchor="center")
            ttk.Label(outer, text="Downloads legally open-access PDFs and names them by Record_ID",
                      foreground="#555").pack(anchor="center", pady=(0, 8))

            form = ttk.Frame(outer)
            form.pack(fill="x")
            form.columnconfigure(1, weight=1)

            self.v_excel = tk.StringVar(value=self.settings.get("excel", ""))
            self.v_out = tk.StringVar(value=self.settings.get("output", ""))
            self.v_sheet = tk.StringVar()
            self.v_rc = tk.StringVar()
            self.v_dc = tk.StringVar()
            self.v_email = tk.StringVar(value=self.settings.get("email", ""))
            self.v_els = tk.StringVar(value=self.settings.get("elsevier_key", ""))
            self.v_wil = tk.StringVar(value=self.settings.get("wiley_token", ""))
            self.v_redl = tk.BooleanVar(value=False)
            self.v_recheck = tk.BooleanVar(value=False)
            self.v_workers = tk.IntVar(value=int(self.settings.get("workers", 3)))
            self.v_browser = tk.BooleanVar(value=bool(self.settings.get("use_browser", True)))
            self.v_show = tk.BooleanVar(value=bool(self.settings.get("show_browser", False)))
            self.v_inst = tk.BooleanVar(value=bool(self.settings.get("institutional", False)))

            r = 0
            ttk.Label(form, text="Excel File:").grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(form, textvariable=self.v_excel).grid(row=r, column=1, sticky="ew", padx=6)
            ttk.Button(form, text="Browse", command=self.browse_excel).grid(row=r, column=2)
            r += 1
            ttk.Label(form, text="Output Folder:").grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(form, textvariable=self.v_out).grid(row=r, column=1, sticky="ew", padx=6)
            ttk.Button(form, text="Browse", command=self.browse_out).grid(row=r, column=2)
            r += 1
            cols = ttk.Frame(form)
            cols.grid(row=r, column=0, columnspan=3, sticky="ew", pady=4)
            ttk.Label(cols, text="Sheet:").pack(side="left")
            self.cb_sheet = ttk.Combobox(cols, textvariable=self.v_sheet, state="readonly", width=18)
            self.cb_sheet.pack(side="left", padx=(4, 14))
            self.cb_sheet.bind("<<ComboboxSelected>>", lambda e: self.fill_columns())
            ttk.Label(cols, text="Record ID Column:").pack(side="left")
            self.cb_rc = ttk.Combobox(cols, textvariable=self.v_rc, state="readonly", width=18)
            self.cb_rc.pack(side="left", padx=(4, 14))
            ttk.Label(cols, text="DOI Column:").pack(side="left")
            self.cb_dc = ttk.Combobox(cols, textvariable=self.v_dc, state="readonly", width=18)
            self.cb_dc.pack(side="left", padx=4)
            r += 1
            ttk.Label(form, text="Your e-mail:").grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(form, textvariable=self.v_email).grid(row=r, column=1, sticky="ew", padx=6)
            ttk.Label(form, text="(needed by Unpaywall)", foreground="#777").grid(row=r, column=2, sticky="w")
            r += 1
            keys = ttk.Frame(form)
            keys.grid(row=r, column=0, columnspan=3, sticky="ew", pady=(2, 0))
            ttk.Label(keys, text="Elsevier API key (optional):").pack(side="left")
            ttk.Entry(keys, textvariable=self.v_els, width=26, show="•").pack(side="left", padx=(4, 14))
            ttk.Label(keys, text="Wiley TDM token (optional):").pack(side="left")
            ttk.Entry(keys, textvariable=self.v_wil, width=26, show="•").pack(side="left", padx=4)
            ttk.Label(keys, text="free - see README", foreground="#777").pack(side="left", padx=6)
            r += 1
            opt = ttk.Frame(form)
            opt.grid(row=r, column=0, columnspan=3, sticky="w", pady=(4, 0))
            ttk.Checkbutton(opt, text="Re-download existing files", variable=self.v_redl).pack(side="left")
            ttk.Checkbutton(opt, text="Re-check papers previously found not accessible",
                            variable=self.v_recheck).pack(side="left", padx=16)
            ttk.Label(opt, text="Parallel downloads:").pack(side="left", padx=(8, 2))
            ttk.Spinbox(opt, from_=1, to=5, width=3, textvariable=self.v_workers,
                        state="readonly").pack(side="left")
            r += 1
            opt2 = ttk.Frame(form)
            opt2.grid(row=r, column=0, columnspan=3, sticky="w", pady=(2, 0))
            ttk.Checkbutton(opt2, text="Use my web browser (Edge/Chrome) for sites that need it  [recommended]",
                            variable=self.v_browser).pack(side="left")
            ttk.Checkbutton(opt2, text="Show browser window", variable=self.v_show).pack(side="left", padx=16)
            r += 1
            ttk.Checkbutton(form, text="Also download papers my institution subscribes to "
                                       "(only on a campus network where you are entitled to them)",
                            variable=self.v_inst).grid(row=r, column=0, columnspan=3, sticky="w")

            ttk.Separator(outer).pack(fill="x", pady=10)
            btns = ttk.Frame(outer)
            btns.pack(fill="x")
            self.b_start = ttk.Button(btns, text="DOWNLOAD ALL PAPERS", style="Big.TButton",
                                      command=self.start)
            self.b_start.pack(side="left", expand=True, fill="x")
            self.b_stop = ttk.Button(btns, text="Stop", command=self.stop, state="disabled")
            self.b_stop.pack(side="left", padx=6, ipady=6)

            prog = ttk.Frame(outer)
            prog.pack(fill="x", pady=(10, 2))
            self.pb = ttk.Progressbar(prog, maximum=100)
            self.pb.pack(side="left", fill="x", expand=True)
            self.l_pct = ttk.Label(prog, text="0%", width=6, anchor="e", style="StatNum.TLabel")
            self.l_pct.pack(side="left")

            stats = ttk.Frame(outer)
            stats.pack(fill="x", pady=4)
            self.stat_labels = {}
            names = [("Total papers", "total"), ("Processed", "processed"), ("Remaining", "remaining"),
                     ("Downloaded", "downloaded"), ("Already Exists", "exists"),
                     ("Not Accessible", "notacc"), ("Failed / Temporary", "failed"),
                     ("Missing / Invalid DOI", "invalid")]
            colors = {"downloaded": "#2E7D32", "exists": "#1565C0", "notacc": "#E65100",
                      "failed": "#C62828", "invalid": "#616161"}
            for i, (label, key) in enumerate(names):
                f = ttk.Frame(stats)
                f.grid(row=i // 4, column=i % 4, sticky="w", padx=(0, 22), pady=1)
                ttk.Label(f, text=label + ":", style="Stat.TLabel").pack(side="left")
                lab = ttk.Label(f, text="0", style="StatNum.TLabel", foreground=colors.get(key, "#000"))
                lab.pack(side="left", padx=4)
                self.stat_labels[key] = lab

            self.l_current = ttk.Label(outer, text="Currently processing: -", foreground="#333",
                                       font=("Segoe UI", 9))
            self.l_current.pack(anchor="w", pady=(4, 0))

            ttk.Label(outer, text="Activity Log:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 2))
            self.logbox = ScrolledText(outer, height=14, font=("Consolas", 9), wrap="word", state="disabled")
            self.logbox.pack(fill="both", expand=True)
            self.logbox.tag_config("ok", foreground="#2E7D32")
            self.logbox.tag_config("exists", foreground="#1565C0")
            self.logbox.tag_config("bad", foreground="#C62828")
            self.logbox.tag_config("warn", foreground="#E65100")
            self.logbox.tag_config("info", foreground="#555555")

            bottom = ttk.Frame(outer)
            bottom.pack(fill="x", pady=(8, 0))
            ttk.Button(bottom, text="Open Output Folder", command=self.open_out).pack(side="left")
            ttk.Button(bottom, text="Open Download_Status.xlsx", command=self.open_report).pack(side="left", padx=6)
            self.b_test = ttk.Button(bottom, text="Test Connection", command=self.test_connection)
            self.b_test.pack(side="left")
            ttk.Label(bottom, text="Only legally open-access PDFs are downloaded.",
                      foreground="#777").pack(side="right")

            if self.v_excel.get() and os.path.exists(self.v_excel.get()):
                self.load_excel(self.v_excel.get(), quiet=True)
            root.protocol("WM_DELETE_WINDOW", self.on_close)
            root.after(100, self.poll)

        # ---- helpers ----
        def log(self, text, tag="info"):
            self.logbox.configure(state="normal")
            self.logbox.insert("end", text + "\n", tag)
            self.logbox.see("end")
            self.logbox.configure(state="disabled")

        def browse_excel(self):
            p = filedialog.askopenfilename(
                title="Select Excel file",
                filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("CSV", "*.csv"), ("All files", "*.*")])
            if p:
                self.v_excel.set(p)
                if not self.v_out.get():
                    self.v_out.set(os.path.join(os.path.dirname(p), "Downloaded_Papers"))
                self.load_excel(p)

        def browse_out(self):
            p = filedialog.askdirectory(title="Select output folder")
            if p:
                self.v_out.set(p)

        def load_excel(self, path, quiet=False):
            try:
                self.tables = load_workbook_tables(path)
            except Exception as e:
                self.tables = {}
                if not quiet:
                    messagebox.showerror(APP_NAME, f"Could not read the Excel file:\n\n{e}")
                return
            if not self.tables:
                messagebox.showerror(APP_NAME, "The Excel file contains no data.")
                return
            self.cb_sheet["values"] = list(self.tables)
            self.v_sheet.set(pick_default_sheet(self.tables))
            self.fill_columns()
            if not self.v_out.get():
                self.v_out.set(os.path.join(os.path.dirname(path), "Downloaded_Papers"))

        def fill_columns(self):
            t = self.tables.get(self.v_sheet.get())
            if not t:
                return
            df, rc, dc = t
            cols = [str(c) for c in df.columns]
            self.cb_rc["values"] = cols
            self.cb_dc["values"] = cols
            self.v_rc.set(rc or "")
            self.v_dc.set(dc or "")
            n = len(df)
            self.stat_labels["total"].configure(text=str(n))
            self.stat_labels["remaining"].configure(text=str(n))
            msg = f"Loaded sheet '{self.v_sheet.get()}' - {n} rows."
            if rc and dc:
                self.log(msg + f" Detected columns: Record ID = '{rc}', DOI = '{dc}'.")
            else:
                missing = [x for x, v in (("Record_ID", rc), ("DOI", dc)) if not v]
                self.log(msg + f" Could not detect: {', '.join(missing)} - please choose from the lists.", "warn")

        def open_out(self):
            if self.v_out.get() and os.path.isdir(self.v_out.get()):
                open_path(self.v_out.get())

        def open_report(self):
            p = os.path.join(self.v_out.get(), STATUS_XLSX)
            if os.path.exists(p):
                open_path(p)
            else:
                messagebox.showinfo(APP_NAME, "No status report yet - run the download first.")

        def test_connection(self):
            self.b_test.configure(state="disabled")
            self.log("=" * 60)
            self.log("Testing connection to all paper sources ...")
            email, ub, sh = self.v_email.get().strip(), self.v_browser.get(), self.v_show.get()

            def target():
                try:
                    connection_test(email, ub, sh, log=lambda m, lvl="info": self.q.put(("log", (m, lvl))))
                except Exception as e:
                    self.q.put(("log", (f"Test failed: {e}", "bad")))
                self.q.put(("test_done", None))
            threading.Thread(target=target, daemon=True).start()

        # ---- run ----
        def start(self):
            excel, out = self.v_excel.get().strip(), self.v_out.get().strip()
            if not excel or not os.path.exists(excel):
                messagebox.showerror(APP_NAME, "Please select a valid Excel file.")
                return
            if not out:
                messagebox.showerror(APP_NAME, "Please select an output folder.")
                return
            if not self.tables:
                self.load_excel(excel)
            if not self.v_rc.get() or not self.v_dc.get():
                messagebox.showerror(
                    APP_NAME, "The Record_ID and DOI columns could not be found.\n\n"
                              "Please choose them from the 'Record ID Column' and 'DOI Column' lists.")
                return
            if self.v_rc.get() == self.v_dc.get():
                messagebox.showerror(APP_NAME, "Record ID column and DOI column must be different.")
                return
            email = self.v_email.get().strip()
            if email and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
                messagebox.showerror(APP_NAME, "Please enter a valid e-mail address (or leave it empty).")
                return
            if not email:
                if not messagebox.askyesno(
                        APP_NAME, "No e-mail entered.\n\nUnpaywall (the best source of legal open-access "
                                  "PDFs) requires an e-mail address and will be skipped.\n\nContinue anyway?"):
                    return
            try:
                os.makedirs(out, exist_ok=True)
            except OSError as e:
                messagebox.showerror(APP_NAME, f"Cannot create output folder:\n{e}")
                return
            self.settings.update(excel=excel, output=out, email=email, workers=self.v_workers.get(),
                                 use_browser=self.v_browser.get(), show_browser=self.v_show.get(),
                                 institutional=self.v_inst.get(), elsevier_key=self.v_els.get().strip(),
                                 wiley_token=self.v_wil.get().strip())
            save_settings(self.settings)

            opts = Options(excel_path=excel, output_dir=out, sheet=self.v_sheet.get(),
                           record_col=self.v_rc.get(), doi_col=self.v_dc.get(), email=email,
                           redownload=self.v_redl.get(), recheck_unavailable=self.v_recheck.get(),
                           workers=self.v_workers.get(), use_browser=self.v_browser.get(),
                           show_browser=self.v_show.get(), institutional=self.v_inst.get(),
                           elsevier_key=self.v_els.get().strip(), wiley_token=self.v_wil.get().strip())
            self.engine = Engine(opts, callback=lambda ev, p: self.q.put((ev, p)))
            self.counts = {k: 0 for k in ("processed", "downloaded", "exists", "notacc", "failed", "invalid")}
            self.b_start.configure(state="disabled")
            self.b_stop.configure(state="normal")
            self.pb["value"] = 0
            self.log("=" * 60)
            self.log(f"Started {now_str()}")

            def target():
                try:
                    self.engine.run()
                except ValueError as e:              # e.g. columns not found
                    self.q.put(("fatal", str(e)))
                except Exception as e:
                    path = write_error_log("run")
                    self.q.put(("fatal", f"{type(e).__name__}: {e}\n\nFull details were saved to:\n{path}\n"
                                         f"Please send that file."))

            self.worker = threading.Thread(target=target, daemon=True)
            self.worker.start()

        def stop(self):
            if self.engine:
                self.engine.stop()
                self.b_stop.configure(state="disabled")
                self.log("Stopping after current requests... (progress is saved; run again to resume)", "warn")

        def on_close(self):
            if self.worker and self.worker.is_alive():
                if not messagebox.askyesno(APP_NAME, "Downloads are still running.\n\n"
                                                     "Stop and exit? (You can resume later.)"):
                    return
                self.engine.stop()
                self.worker.join(timeout=5)
            self.root.destroy()

        def update_stats(self):
            c = self.counts
            for k in ("processed", "downloaded", "exists", "notacc", "failed", "invalid"):
                self.stat_labels[k].configure(text=str(c.get(k, 0)))
            self.stat_labels["total"].configure(text=str(self.total))
            self.stat_labels["remaining"].configure(text=str(max(0, self.total - c.get("processed", 0))))
            pct = (c.get("processed", 0) / self.total * 100) if self.total else 0
            self.pb["value"] = pct
            self.l_pct.configure(text=f"{pct:.0f}%")

        def poll(self):
            try:
                while True:
                    ev, p = self.q.get_nowait()
                    self.handle(ev, p)
            except queue.Empty:
                pass
            self.root.after(120, self.poll)

        def handle(self, ev, p):
            if ev == "start":
                self.total = p["total"]
                self.update_stats()
                self.log(f"Sheet '{p['sheet']}': {self.total} records "
                         f"(Record ID = '{p['record_col']}', DOI = '{p['doi_col']}')")
            elif ev == "current":
                rid, doi = p
                self.l_current.configure(text=f"Currently processing:  {rid}     DOI: {doi}")
            elif ev == "pdf_found":
                pass
            elif ev == "log":
                self.log(p[0], p[1])
            elif ev == "result":
                res = p["result"]
                if res.status == S_NOT_PROCESSED:
                    return
                self.counts["processed"] += 1
                st = res.status
                name = res.record_id or f"(row {res.row})"
                prev = "  [previous run]" if p["reused"] else ""
                if st == S_DOWNLOADED:
                    self.counts["downloaded"] += 1
                    self.log(f"✓ {name} — Downloaded → {res.downloaded_file}  ({res.source})", "ok")
                elif st == S_EXISTS:
                    self.counts["exists"] += 1
                    self.log(f"• {name} — Already exists", "exists")
                elif st in NOT_ACCESSIBLE_STATUSES:
                    self.counts["notacc"] += 1
                    self.log(f"✗ {name} — {st}{prev}", "warn")
                elif st in (S_MISSING, S_INVALID):
                    self.counts["invalid"] += 1
                    self.log(f"✗ {name} — {st}{prev}", "bad")
                else:
                    self.counts["failed"] += 1
                    self.log(f"✗ {name} — {st}: {res.error[:220]}", "bad")
                self.update_stats()
            elif ev == "done":
                self.b_start.configure(state="normal")
                self.b_stop.configure(state="disabled")
                self.l_current.configure(text="Currently processing: -")
                s = p["summary"]
                title = "Stopped" if p["stopped"] else "Finished"
                lines = [f"{k}: {v}" for k, v in s.items() if v or k == "Total"]
                self.log(f"{title}. " + " | ".join(lines), "info")
                self.log(f"Status report: {p['report']}", "info")
                if messagebox.askyesno(APP_NAME, f"{title}!\n\n" + "\n".join(lines) +
                                       "\n\nOpen Download_Status.xlsx now?"):
                    open_path(p["report"])
            elif ev == "test_done":
                self.b_test.configure(state="normal")
            elif ev == "fatal":
                self.b_start.configure(state="normal")
                self.b_stop.configure(state="disabled")
                self.log(f"ERROR: {p}", "bad")
                messagebox.showerror(APP_NAME, p)

    root = tk.Tk()

    def report_tk_error(exc, val, tb):
        import traceback as _tb
        path = ""
        try:
            for d in ERROR_LOG_DIR:
                os.makedirs(d, exist_ok=True)
                path = os.path.join(d, "error_log.txt")
                with open(path, "a", encoding="utf-8") as fh:
                    fh.write(f"\n===== {now_str()} v{APP_VERSION} GUI\n" + "".join(_tb.format_exception(exc, val, tb)))
                break
        except Exception:
            pass
        messagebox.showerror(APP_NAME, f"Unexpected error: {val}\n\nDetails saved to:\n{path}")
    root.report_callback_exception = report_tk_error
    App(root)
    root.mainloop()


# --------------------------------------------------------------------------- #
#  Command line
# --------------------------------------------------------------------------- #
def run_cli(argv):
    ap = argparse.ArgumentParser(description=APP_NAME)
    ap.add_argument("--cli", action="store_true")
    ap.add_argument("excel", nargs="?")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--email", default="")
    ap.add_argument("--sheet")
    ap.add_argument("--record-col")
    ap.add_argument("--doi-col")
    ap.add_argument("--redownload", action="store_true")
    ap.add_argument("--recheck", action="store_true")
    ap.add_argument("--workers", type=int, default=3)
    ap.add_argument("--no-browser", action="store_true", help="never use Edge/Chrome")
    ap.add_argument("--show-browser", action="store_true")
    ap.add_argument("--institutional", action="store_true")
    ap.add_argument("--elsevier-key", default="")
    ap.add_argument("--wiley-token", default="")
    ap.add_argument("--test", action="store_true", help="only run the connection test")
    a = ap.parse_args(argv)
    if a.test:
        connection_test(a.email, not a.no_browser, a.show_browser, log=lambda m, lvl="info": print(m))
        return 0

    def cb(ev, p):
        if ev == "start":
            print(f"Sheet '{p['sheet']}' | {p['total']} records | Record ID='{p['record_col']}' DOI='{p['doi_col']}'")
        elif ev == "result":
            r = p["result"]
            extra = r.source if r.status == S_DOWNLOADED else r.error
            print(f"[{p['done']}/{p['total']}] {r.record_id or '(row %d)' % r.row:<12} {r.status:<16} {extra[:110]}")
        elif ev == "log":
            print(p[0])
        elif ev == "done":
            print("\nSUMMARY:", ", ".join(f"{k}: {v}" for k, v in p["summary"].items() if v))
            print("Report:", p["report"])

    if not a.excel or not a.output:
        ap.error("excel and output are required")
    eng = Engine(Options(excel_path=a.excel, output_dir=a.output, sheet=a.sheet, record_col=a.record_col,
                         doi_col=a.doi_col, email=a.email, redownload=a.redownload,
                         recheck_unavailable=a.recheck, workers=a.workers, use_browser=not a.no_browser,
                         show_browser=a.show_browser, institutional=a.institutional,
                         elsevier_key=a.elsevier_key, wiley_token=a.wiley_token), cb)
    try:
        eng.run()
    except KeyboardInterrupt:
        eng.stop()
        print("Stopped.")
    except ValueError as e:
        print("ERROR:", e)
        return 2
    return 0


if __name__ == "__main__":
    # a windowed .exe has no console: give libraries somewhere harmless to write
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")
    if "--cli" in sys.argv:
        sys.exit(run_cli(sys.argv[1:]))
    run_gui()
