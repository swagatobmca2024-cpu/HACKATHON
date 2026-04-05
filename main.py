import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
import json
import re
import time
import csv
import io
import xml.etree.ElementTree as ET
from urllib.parse import urljoin, urlparse, urlencode
from urllib.robotparser import RobotFileParser
from collections import defaultdict
from datetime import datetime
import hashlib
import trafilatura
from fake_useragent import UserAgent
import lxml

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="WebHarvest Pro",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# SVG ICONS (no emojis)
# ─────────────────────────────────────────────
ICON_SPIDER = """<svg width="28" height="28" viewBox="0 0 28 28" fill="none" xmlns="http://www.w3.org/2000/svg">
<circle cx="14" cy="12" r="5" fill="#00D4AA" opacity="0.9"/>
<circle cx="14" cy="12" r="2.5" fill="#001a12"/>
<line x1="14" y1="7" x2="7" y2="3" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round"/>
<line x1="14" y1="7" x2="21" y2="3" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round"/>
<line x1="9" y1="11" x2="2" y2="9" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round"/>
<line x1="19" y1="11" x2="26" y2="9" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round"/>
<line x1="9" y1="14" x2="2" y2="16" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round"/>
<line x1="19" y1="14" x2="26" y2="16" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round"/>
<line x1="12" y1="17" x2="10" y2="24" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round"/>
<line x1="16" y1="17" x2="18" y2="24" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round"/>
</svg>"""

ICON_LINK = """<svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M7.5 10.5C7.83 10.99 8.26 11.4 8.76 11.7C9.26 12 9.82 12.16 10.39 12.16C10.96 12.16 11.52 12 12.02 11.7L14.52 10.2C15.43 9.56 16 8.52 16 7.39C16 5.52 14.48 4 12.61 4C11.98 4 11.39 4.18 10.89 4.5L9.5 5.34" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
<path d="M10.5 7.5C10.17 7.01 9.74 6.6 9.24 6.3C8.74 6 8.18 5.84 7.61 5.84C7.04 5.84 6.48 6 5.98 6.3L3.48 7.8C2.57 8.44 2 9.48 2 10.61C2 12.48 3.52 14 5.39 14C6.02 14 6.61 13.82 7.11 13.5L8.5 12.66" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
</svg>"""

ICON_TABLE = """<svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
<rect x="2" y="2" width="14" height="14" rx="2" stroke="#00D4AA" stroke-width="1.5"/>
<line x1="2" y1="6.5" x2="16" y2="6.5" stroke="#00D4AA" stroke-width="1.5"/>
<line x1="7" y1="6.5" x2="7" y2="16" stroke="#00D4AA" stroke-width="1.2"/>
<line x1="11.5" y1="6.5" x2="11.5" y2="16" stroke="#00D4AA" stroke-width="1.2"/>
</svg>"""

ICON_IMAGE = """<svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
<rect x="2" y="3" width="14" height="12" rx="2" stroke="#00D4AA" stroke-width="1.5"/>
<circle cx="6.5" cy="7.5" r="1.5" fill="#00D4AA"/>
<path d="M2 12L6 8.5L9 11.5L12 9L16 13" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
</svg>"""

ICON_META = """<svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
<circle cx="9" cy="9" r="7" stroke="#00D4AA" stroke-width="1.5"/>
<line x1="9" y1="8" x2="9" y2="13" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round"/>
<circle cx="9" cy="5.5" r="1" fill="#00D4AA"/>
</svg>"""

ICON_CRAWL = """<svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
<circle cx="9" cy="9" r="7" stroke="#00D4AA" stroke-width="1.5"/>
<path d="M9 2C9 2 12 5.5 12 9C12 12.5 9 16 9 16" stroke="#00D4AA" stroke-width="1.2" stroke-linecap="round"/>
<path d="M9 2C9 2 6 5.5 6 9C6 12.5 9 16 9 16" stroke="#00D4AA" stroke-width="1.2" stroke-linecap="round"/>
<line x1="2" y1="9" x2="16" y2="9" stroke="#00D4AA" stroke-width="1.2"/>
</svg>"""

ICON_EXPORT = """<svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M9 2V11M9 11L6 8M9 11L12 8" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
<path d="M3 13V14C3 15.1 3.9 16 5 16H13C14.1 16 15 15.1 15 14V13" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round"/>
</svg>"""

ICON_SEARCH = """<svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
<circle cx="8" cy="8" r="5.5" stroke="#00D4AA" stroke-width="1.5"/>
<line x1="12.5" y1="12.5" x2="16" y2="16" stroke="#00D4AA" stroke-width="2" stroke-linecap="round"/>
</svg>"""

ICON_SETTINGS = """<svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
<circle cx="9" cy="9" r="2.5" stroke="#00D4AA" stroke-width="1.5"/>
<path d="M9 1.5V3M9 15V16.5M1.5 9H3M15 9H16.5M3.2 3.2L4.3 4.3M13.7 13.7L14.8 14.8M3.2 14.8L4.3 13.7M13.7 4.3L14.8 3.2" stroke="#00D4AA" stroke-width="1.5" stroke-linecap="round"/>
</svg>"""

# ─────────────────────────────────────────────
# CSS STYLING
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;600&display=swap');

:root {
    --bg: #060d0a;
    --surface: #0d1f18;
    --surface2: #112b20;
    --accent: #00D4AA;
    --accent2: #00ff88;
    --text: #d4ead6;
    --muted: #5a7a65;
    --border: #1a3828;
    --danger: #ff4d6d;
    --warn: #ffb347;
}

html, body, .stApp {
    background-color: var(--bg) !important;
    color: var(--text) !important;
    font-family: 'DM Sans', sans-serif;
}

/* Hide default streamlit chrome */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 1.5rem 2rem 3rem 2rem !important; max-width: 1400px; }

/* ── HEADER ── */
.wh-header {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 1.4rem 0 1rem 0;
    border-bottom: 1px solid var(--border);
    margin-bottom: 1.6rem;
}
.wh-header h1 {
    font-family: 'Space Mono', monospace;
    font-size: 1.7rem;
    color: var(--accent);
    margin: 0;
    letter-spacing: -0.5px;
}
.wh-header span.sub {
    font-size: 0.78rem;
    color: var(--muted);
    font-family: 'Space Mono', monospace;
    border: 1px solid var(--border);
    padding: 2px 8px;
    border-radius: 3px;
    margin-left: 6px;
}

/* ── CARDS ── */
.wh-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 1rem;
}
.wh-card h4 {
    font-family: 'Space Mono', monospace;
    font-size: 0.82rem;
    color: var(--accent);
    text-transform: uppercase;
    letter-spacing: 1px;
    margin: 0 0 0.8rem 0;
    display: flex;
    align-items: center;
    gap: 7px;
}

/* ── STAT PILLS ── */
.stat-row { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 1rem; }
.stat-pill {
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 0.55rem 1rem;
    font-family: 'Space Mono', monospace;
    font-size: 0.78rem;
    color: var(--text);
    display: flex;
    flex-direction: column;
    gap: 2px;
    min-width: 120px;
}
.stat-pill .val { font-size: 1.3rem; color: var(--accent2); font-weight: 700; }
.stat-pill .lbl { color: var(--muted); font-size: 0.68rem; text-transform: uppercase; letter-spacing: 0.5px; }

/* ── BADGE ── */
.badge {
    display: inline-block;
    font-family: 'Space Mono', monospace;
    font-size: 0.68rem;
    padding: 2px 7px;
    border-radius: 3px;
    font-weight: 700;
}
.badge-ok   { background: #003d2a; color: var(--accent2); border: 1px solid #00a060; }
.badge-warn { background: #3d2a00; color: var(--warn);    border: 1px solid #a06000; }
.badge-err  { background: #3d0010; color: var(--danger);  border: 1px solid #a00030; }

/* ── INPUTS ── */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea,
.stSelectbox > div > div > div,
.stMultiSelect > div > div > div {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    color: var(--text) !important;
    border-radius: 6px !important;
    font-family: 'DM Sans', sans-serif !important;
}
.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 2px rgba(0,212,170,0.15) !important;
}

/* ── BUTTON ── */
.stButton > button {
    background: var(--accent) !important;
    color: #000 !important;
    font-family: 'Space Mono', monospace !important;
    font-size: 0.82rem !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 6px !important;
    padding: 0.55rem 1.4rem !important;
    letter-spacing: 0.5px;
    transition: all 0.15s ease;
}
.stButton > button:hover {
    background: var(--accent2) !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 16px rgba(0,212,170,0.3) !important;
}
.stButton > button[kind="secondary"] {
    background: var(--surface2) !important;
    color: var(--accent) !important;
    border: 1px solid var(--border) !important;
}

/* ── TABS ── */
.stTabs [data-baseweb="tab-list"] {
    background: var(--surface) !important;
    border-radius: 8px 8px 0 0;
    border-bottom: 1px solid var(--border);
    gap: 0;
    padding: 0 0.5rem;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: var(--muted) !important;
    font-family: 'Space Mono', monospace !important;
    font-size: 0.75rem !important;
    border-radius: 0 !important;
    border-bottom: 2px solid transparent !important;
    padding: 0.7rem 1rem !important;
}
.stTabs [aria-selected="true"] {
    color: var(--accent) !important;
    border-bottom: 2px solid var(--accent) !important;
    background: transparent !important;
}
.stTabs [data-baseweb="tab-panel"] {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-top: none !important;
    border-radius: 0 0 8px 8px;
    padding: 1.2rem;
}

/* ── SIDEBAR ── */
section[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border) !important;
}
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stSlider label,
section[data-testid="stSidebar"] .stCheckbox label,
section[data-testid="stSidebar"] .stNumberInput label,
section[data-testid="stSidebar"] .stTextInput label {
    color: var(--text) !important;
    font-family: 'DM Sans', sans-serif;
    font-size: 0.85rem;
}

/* ── DATAFRAME ── */
.stDataFrame { border: 1px solid var(--border) !important; border-radius: 6px; overflow: hidden; }
.stDataFrame thead th {
    background: var(--surface2) !important;
    color: var(--accent) !important;
    font-family: 'Space Mono', monospace !important;
    font-size: 0.75rem !important;
}

/* ── LOG BOX ── */
.log-box {
    background: #030a06;
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 0.9rem 1rem;
    font-family: 'Space Mono', monospace;
    font-size: 0.72rem;
    color: var(--accent);
    max-height: 220px;
    overflow-y: auto;
    line-height: 1.7;
}
.log-box .log-ok   { color: var(--accent2); }
.log-box .log-warn { color: var(--warn); }
.log-box .log-err  { color: var(--danger); }
.log-box .log-info { color: var(--muted); }

/* ── EXPANDER ── */
.streamlit-expanderHeader {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    border-radius: 6px !important;
    color: var(--text) !important;
    font-family: 'DM Sans', sans-serif !important;
}
.streamlit-expanderContent {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-top: none !important;
}

/* ── PROGRESS ── */
.stProgress > div > div { background: var(--accent) !important; }

/* ── DIVIDER ── */
hr { border-color: var(--border) !important; }

/* ── ALERTS ── */
.stAlert { border-radius: 6px !important; }

/* ── MULTISELECT TAGS ── */
.stMultiSelect span[data-baseweb="tag"] {
    background: var(--surface2) !important;
    color: var(--accent) !important;
    border: 1px solid var(--border) !important;
}

/* ── NUMBER INPUT ── */
.stNumberInput input {
    background: var(--surface2) !important;
    color: var(--text) !important;
    border: 1px solid var(--border) !important;
}

/* ── SLIDER ── */
.stSlider [data-baseweb="slider"] div[role="slider"] {
    background: var(--accent) !important;
    border-color: var(--accent) !important;
}

/* ── CHECKBOX ── */
.stCheckbox > label > div:first-child {
    border-color: var(--accent) !important;
}

/* ── TOOLTIP ── */
.tooltip-wrap { position: relative; display: inline-block; }
.tooltip-wrap .tooltip-text {
    visibility: hidden;
    width: 200px;
    background: var(--surface2);
    color: var(--text);
    font-size: 0.72rem;
    text-align: center;
    border-radius: 5px;
    padding: 5px 8px;
    position: absolute;
    z-index: 1;
    bottom: 125%;
    left: 50%;
    transform: translateX(-50%);
    border: 1px solid var(--border);
}
.tooltip-wrap:hover .tooltip-text { visibility: visible; }

/* ── RESULT ITEM ── */
.result-item {
    padding: 0.65rem 0.9rem;
    border-left: 3px solid var(--accent);
    background: var(--surface2);
    border-radius: 0 5px 5px 0;
    margin-bottom: 0.5rem;
    font-size: 0.84rem;
}
.result-item a { color: var(--accent); text-decoration: none; font-family: 'Space Mono', monospace; font-size: 0.75rem; }
.result-item a:hover { color: var(--accent2); }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "results" not in st.session_state:
    st.session_state.results = {}
if "logs" not in st.session_state:
    st.session_state.logs = []
if "history" not in st.session_state:
    st.session_state.history = []
if "crawl_visited" not in st.session_state:
    st.session_state.crawl_visited = set()

# ─────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────
def log(msg: str, level: str = "ok"):
    ts = datetime.now().strftime("%H:%M:%S")
    st.session_state.logs.append({"ts": ts, "msg": msg, "level": level})

def get_session():
    ua = UserAgent()
    s = requests.Session()
    s.headers.update({
        "User-Agent": ua.random,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
        "Accept-Encoding": "gzip, deflate",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    })
    return s

def check_robots(url: str, ua: str = "*") -> bool:
    parsed = urlparse(url)
    robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    try:
        rp = RobotFileParser()
        rp.set_url(robots_url)
        rp.read()
        return rp.can_fetch(ua, url)
    except Exception:
        return True  # assume allowed if robots.txt unreadable

def fetch_page(url: str, timeout: int = 15, retries: int = 2, delay: float = 1.0):
    s = get_session()
    for attempt in range(retries):
        try:
            resp = s.get(url, timeout=timeout, allow_redirects=True)
            resp.raise_for_status()
            log(f"GET {url} -> {resp.status_code} ({len(resp.content)//1024}KB)", "ok")
            return resp
        except requests.exceptions.HTTPError as e:
            log(f"HTTP error ({e}) on attempt {attempt+1}", "err")
        except requests.exceptions.ConnectionError:
            log(f"Connection error on attempt {attempt+1}", "err")
        except requests.exceptions.Timeout:
            log(f"Timeout on attempt {attempt+1}", "warn")
        except Exception as e:
            log(f"Unexpected: {e}", "err")
        if attempt < retries - 1:
            time.sleep(delay)
    return None

def get_soup(resp) -> BeautifulSoup:
    encoding = resp.encoding or "utf-8"
    return BeautifulSoup(resp.content, "lxml", from_encoding=encoding)

def url_fingerprint(url: str) -> str:
    return hashlib.md5(url.strip().lower().encode()).hexdigest()

# ─────────────────────────────────────────────
# SCRAPER FUNCTIONS
# ─────────────────────────────────────────────
def scrape_links(soup: BeautifulSoup, base_url: str) -> list[dict]:
    links = []
    seen = set()
    for tag in soup.find_all("a", href=True):
        href = tag["href"].strip()
        abs_url = urljoin(base_url, href)
        if abs_url in seen:
            continue
        seen.add(abs_url)
        links.append({
            "text": tag.get_text(strip=True)[:100] or "(no text)",
            "url": abs_url,
            "rel": tag.get("rel", ["—"])[0] if tag.get("rel") else "—",
            "title": tag.get("title", "")[:60],
            "external": urlparse(abs_url).netloc != urlparse(base_url).netloc,
        })
    return links

def scrape_tables(soup: BeautifulSoup) -> list[pd.DataFrame]:
    tables = []
    for tbl in soup.find_all("table"):
        try:
            df = pd.read_html(str(tbl))[0]
            tables.append(df)
        except Exception:
            pass
    return tables

def scrape_images(soup: BeautifulSoup, base_url: str) -> list[dict]:
    imgs = []
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-src") or img.get("data-lazy-src") or ""
        if not src:
            continue
        imgs.append({
            "src": urljoin(base_url, src),
            "alt": img.get("alt", "")[:80],
            "width": img.get("width", "—"),
            "height": img.get("height", "—"),
            "loading": img.get("loading", "—"),
        })
    return imgs

def scrape_meta(soup: BeautifulSoup, url: str) -> dict:
    meta = {}
    meta["url"] = url
    meta["title"] = soup.title.string.strip() if soup.title else "—"

    for m in soup.find_all("meta"):
        name = (m.get("name") or m.get("property") or "").lower()
        content = m.get("content", "")
        if name:
            meta[name] = content[:300]

    # Open Graph
    og = {k: v for k, v in meta.items() if k.startswith("og:")}
    meta["_open_graph"] = og

    # Twitter Card
    tw = {k: v for k, v in meta.items() if k.startswith("twitter:")}
    meta["_twitter"] = tw

    # Canonical
    canonical = soup.find("link", rel="canonical")
    meta["canonical"] = canonical["href"] if canonical else "—"

    # Headings
    meta["h1"] = [h.get_text(strip=True) for h in soup.find_all("h1")]
    meta["h2"] = [h.get_text(strip=True) for h in soup.find_all("h2")]

    # Word count approximation
    body_text = soup.get_text(separator=" ", strip=True)
    meta["word_count"] = len(body_text.split())

    # Schema.org JSON-LD
    schemas = []
    for s in soup.find_all("script", type="application/ld+json"):
        try:
            schemas.append(json.loads(s.string))
        except Exception:
            pass
    meta["_schema_org"] = schemas

    return meta

def scrape_custom_css(soup: BeautifulSoup, selectors: list[str]) -> dict:
    results = {}
    for sel in selectors:
        sel = sel.strip()
        if not sel:
            continue
        try:
            found = soup.select(sel)
            results[sel] = [
                el.get_text(separator=" ", strip=True)[:400]
                for el in found
            ]
            log(f"Selector '{sel}' -> {len(found)} match(es)", "ok")
        except Exception as e:
            results[sel] = []
            log(f"Selector '{sel}' error: {e}", "err")
    return results

def scrape_article(url: str) -> str:
    try:
        downloaded = trafilatura.fetch_url(url)
        text = trafilatura.extract(downloaded, include_comments=False, include_tables=False)
        return text or "— No article text extracted —"
    except Exception as e:
        return f"Error: {e}"

def crawl_site(
    start_url: str,
    max_pages: int = 10,
    same_domain: bool = True,
    delay: float = 1.0,
    respect_robots: bool = True,
    progress_bar=None,
    status_text=None,
) -> list[dict]:
    visited = {}
    queue = [start_url]
    base_domain = urlparse(start_url).netloc

    page_num = 0
    while queue and page_num < max_pages:
        url = queue.pop(0)
        fp = url_fingerprint(url)
        if fp in visited:
            continue

        if respect_robots and not check_robots(url):
            log(f"Robots.txt disallows: {url}", "warn")
            visited[fp] = {"url": url, "status": "robots_blocked", "title": "—", "links": 0}
            continue

        if status_text:
            status_text.markdown(f'<div class="log-box"><span class="log-info">Crawling [{page_num+1}/{max_pages}]:</span> <span class="log-ok">{url[:80]}</span></div>', unsafe_allow_html=True)

        resp = fetch_page(url, delay=delay)
        page_num += 1

        if progress_bar:
            progress_bar.progress(page_num / max_pages)

        if resp is None:
            visited[fp] = {"url": url, "status": "error", "title": "—", "links": 0}
            continue

        soup = get_soup(resp)
        title = soup.title.string.strip() if soup.title else "—"
        links_found = scrape_links(soup, url)

        visited[fp] = {
            "url": url,
            "status": resp.status_code,
            "title": title[:80],
            "links": len(links_found),
            "content_type": resp.headers.get("Content-Type", "—"),
            "size_kb": len(resp.content) // 1024,
        }

        for lnk in links_found:
            lurl = lnk["url"]
            lp = urlparse(lurl)
            if url_fingerprint(lurl) in visited:
                continue
            if lp.scheme not in ("http", "https"):
                continue
            if same_domain and lp.netloc != base_domain:
                continue
            if lurl not in queue:
                queue.append(lurl)

        time.sleep(delay)

    return list(visited.values())

def extract_structured_data(soup: BeautifulSoup) -> dict:
    """Extract emails, phones, social links, addresses from page."""
    text = soup.get_text(separator=" ")
    emails = list(set(re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)))
    phones = list(set(re.findall(r"(?:\+?\d[\d\s\-().]{7,}\d)", text)))[:20]
    
    social_patterns = {
        "twitter": r"twitter\.com/([a-zA-Z0-9_]{1,50})",
        "linkedin": r"linkedin\.com/(?:in|company)/([a-zA-Z0-9_\-]{1,100})",
        "github": r"github\.com/([a-zA-Z0-9_\-]{1,100})",
        "instagram": r"instagram\.com/([a-zA-Z0-9_.]{1,50})",
        "facebook": r"facebook\.com/([a-zA-Z0-9_.]{1,100})",
        "youtube": r"youtube\.com/(?:@|channel/|c/)([a-zA-Z0-9_\-]{1,100})",
    }
    social = {}
    for platform, pattern in social_patterns.items():
        found = list(set(re.findall(pattern, text, re.IGNORECASE)))
        if found:
            social[platform] = found

    return {"emails": emails, "phones": phones, "social_profiles": social}

# ─────────────────────────────────────────────
# EXPORT HELPERS
# ─────────────────────────────────────────────
def to_csv_bytes(df: pd.DataFrame) -> bytes:
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    return buf.getvalue().encode()

def to_json_bytes(data) -> bytes:
    return json.dumps(data, indent=2, ensure_ascii=False).encode()

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:10px;padding:0.6rem 0 1.2rem 0;border-bottom:1px solid var(--border,#1a3828);margin-bottom:1rem;">
        {ICON_SPIDER}
        <div>
            <div style="font-family:'Space Mono',monospace;color:#00D4AA;font-size:1rem;font-weight:700;">WebHarvest</div>
            <div style="font-size:0.68rem;color:#5a7a65;font-family:'Space Mono',monospace;">PRO v2.0</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f'<div style="font-size:0.75rem;color:#00D4AA;font-family:\'Space Mono\',monospace;margin-bottom:0.5rem;display:flex;align-items:center;gap:6px;">{ICON_SETTINGS} REQUEST SETTINGS</div>', unsafe_allow_html=True)

    timeout = st.slider("Timeout (s)", 5, 60, 15)
    retries = st.slider("Retries", 1, 5, 2)
    delay = st.slider("Request Delay (s)", 0.0, 5.0, 1.0, 0.5)
    respect_robots = st.checkbox("Respect robots.txt", value=True)
    verify_ssl = st.checkbox("Verify SSL", value=True)

    st.markdown("---")
    st.markdown(f'<div style="font-size:0.75rem;color:#00D4AA;font-family:\'Space Mono\',monospace;margin-bottom:0.5rem;display:flex;align-items:center;gap:6px;">{ICON_CRAWL} CRAWLER SETTINGS</div>', unsafe_allow_html=True)

    max_pages = st.number_input("Max Pages (Crawler)", min_value=1, max_value=100, value=10)
    same_domain = st.checkbox("Same domain only", value=True)

    st.markdown("---")
    if st.button("Clear Logs + Results", type="secondary"):
        st.session_state.logs = []
        st.session_state.results = {}
        st.session_state.history = []
        st.rerun()

    st.markdown("---")
    # Activity log in sidebar
    st.markdown('<div style="font-size:0.72rem;color:#5a7a65;font-family:\'Space Mono\',monospace;margin-bottom:0.4rem;">ACTIVITY LOG</div>', unsafe_allow_html=True)
    log_html = ""
    for entry in st.session_state.logs[-30:][::-1]:
        cls = f"log-{entry['level']}"
        log_html += f'<div><span class="log-info">[{entry["ts"]}]</span> <span class="{cls}">{entry["msg"]}</span></div>'
    no_activity = '<span class="log-info">No activity yet.</span>'
    st.markdown(f'<div class="log-box" style="max-height:300px;">{log_html or no_activity}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown(f"""
<div class="wh-header">
    {ICON_SPIDER}
    <h1>WebHarvest Pro <span class="sub">Advanced Web Scraper</span></h1>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# URL INPUT
# ─────────────────────────────────────────────
col_url, col_btn = st.columns([5, 1])
with col_url:
    target_url = st.text_input(
        "Target URL",
        placeholder="https://example.com",
        label_visibility="collapsed",
    )
with col_btn:
    go = st.button("Scrape", use_container_width=True)

# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab_links, tab_tables, tab_images, tab_meta, tab_css, tab_article, tab_crawl, tab_structured = st.tabs([
    "Links", "Tables", "Images", "Meta / SEO", "CSS Selectors", "Article Text", "Site Crawler", "Emails & Social"
])

# ─────────────────────────────────────────────
# SCRAPING EXECUTION
# ─────────────────────────────────────────────
if go and target_url:
    if not target_url.startswith("http"):
        target_url = "https://" + target_url

    with st.spinner("Fetching page..."):
        if respect_robots and not check_robots(target_url):
            st.warning("robots.txt disallows scraping this URL. Proceeding may violate site policy.")
            log(f"robots.txt warning for {target_url}", "warn")

        resp = fetch_page(target_url, timeout=timeout, retries=retries, delay=delay)

    if resp:
        soup = get_soup(resp)

        # ── store results ──
        st.session_state.results["links"] = scrape_links(soup, target_url)
        st.session_state.results["tables"] = scrape_tables(soup)
        st.session_state.results["images"] = scrape_images(soup, target_url)
        st.session_state.results["meta"] = scrape_meta(soup, target_url)
        st.session_state.results["structured"] = extract_structured_data(soup)
        st.session_state.results["url"] = target_url
        st.session_state.results["resp"] = {
            "status": resp.status_code,
            "content_type": resp.headers.get("Content-Type", "—"),
            "server": resp.headers.get("Server", "—"),
            "size_kb": len(resp.content) // 1024,
            "encoding": resp.encoding or "—",
        }
        st.session_state.history.append({
            "url": target_url,
            "ts": datetime.now().strftime("%H:%M:%S"),
            "status": resp.status_code,
        })
        log(f"Scrape complete: {target_url}", "ok")
    else:
        st.error("Failed to fetch the page. Check the URL and your network.")

# ─────────────────────────────────────────────
# STAT BAR
# ─────────────────────────────────────────────
if st.session_state.results:
    r = st.session_state.results
    resp_info = r.get("resp", {})
    status = resp_info.get("status", "—")
    badge_cls = "badge-ok" if str(status).startswith("2") else "badge-err"

    st.markdown(f"""
    <div class="stat-row">
        <div class="stat-pill">
            <span class="val">{len(r.get('links', []))}</span>
            <span class="lbl">Links Found</span>
        </div>
        <div class="stat-pill">
            <span class="val">{len(r.get('tables', []))}</span>
            <span class="lbl">Tables</span>
        </div>
        <div class="stat-pill">
            <span class="val">{len(r.get('images', []))}</span>
            <span class="lbl">Images</span>
        </div>
        <div class="stat-pill">
            <span class="val">{resp_info.get('size_kb','—')} KB</span>
            <span class="lbl">Page Size</span>
        </div>
        <div class="stat-pill">
            <span class="val"><span class="badge {badge_cls}">{status}</span></span>
            <span class="lbl">HTTP Status</span>
        </div>
        <div class="stat-pill">
            <span class="val">{r.get('meta',{}).get('word_count','—')}</span>
            <span class="lbl">Words</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TAB: LINKS
# ─────────────────────────────────────────────
with tab_links:
    st.markdown(f'<div class="wh-card"><h4>{ICON_LINK} Extracted Links</h4>', unsafe_allow_html=True)

    if "links" in st.session_state.results:
        links = st.session_state.results["links"]
        if links:
            df_links = pd.DataFrame(links)

            col_f1, col_f2 = st.columns(2)
            with col_f1:
                link_filter = st.text_input("Filter by text or URL", key="lf", placeholder="Search...")
            with col_f2:
                ext_filter = st.selectbox("Type", ["All", "Internal", "External"], key="ef")

            filtered = df_links.copy()
            if link_filter:
                mask = (
                    filtered["text"].str.contains(link_filter, case=False, na=False) |
                    filtered["url"].str.contains(link_filter, case=False, na=False)
                )
                filtered = filtered[mask]
            if ext_filter == "External":
                filtered = filtered[filtered["external"] == True]
            elif ext_filter == "Internal":
                filtered = filtered[filtered["external"] == False]

            st.markdown(f'<div style="font-size:0.78rem;color:var(--muted,#5a7a65);margin-bottom:0.5rem;font-family:\'Space Mono\',monospace;">{len(filtered)} / {len(links)} links shown</div>', unsafe_allow_html=True)
            st.dataframe(filtered, use_container_width=True, hide_index=True)

            st.download_button(
                "Download CSV",
                data=to_csv_bytes(filtered),
                file_name="links.csv",
                mime="text/csv",
            )
        else:
            st.info("No links found on this page.")
    else:
        st.markdown('<div style="color:var(--muted,#5a7a65);font-size:0.85rem;">Enter a URL and click Scrape to extract links.</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TAB: TABLES
# ─────────────────────────────────────────────
with tab_tables:
    st.markdown(f'<div class="wh-card"><h4>{ICON_TABLE} Extracted Tables</h4>', unsafe_allow_html=True)

    if "tables" in st.session_state.results:
        tables = st.session_state.results["tables"]
        if tables:
            for i, df in enumerate(tables):
                with st.expander(f"Table {i+1}  —  {df.shape[0]} rows x {df.shape[1]} cols"):
                    st.dataframe(df, use_container_width=True)
                    st.download_button(
                        f"Download Table {i+1} CSV",
                        data=to_csv_bytes(df),
                        file_name=f"table_{i+1}.csv",
                        mime="text/csv",
                        key=f"dl_tbl_{i}",
                    )
        else:
            st.info("No HTML tables found on this page.")
    else:
        st.markdown('<div style="color:var(--muted,#5a7a65);font-size:0.85rem;">Enter a URL and click Scrape.</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TAB: IMAGES
# ─────────────────────────────────────────────
with tab_images:
    st.markdown(f'<div class="wh-card"><h4>{ICON_IMAGE} Extracted Images</h4>', unsafe_allow_html=True)

    if "images" in st.session_state.results:
        images = st.session_state.results["images"]
        if images:
            df_imgs = pd.DataFrame(images)
            img_search = st.text_input("Filter by src or alt text", key="img_f", placeholder="Search...")
            if img_search:
                df_imgs = df_imgs[
                    df_imgs["src"].str.contains(img_search, case=False, na=False) |
                    df_imgs["alt"].str.contains(img_search, case=False, na=False)
                ]

            st.dataframe(df_imgs, use_container_width=True, hide_index=True)

            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button("Download CSV", data=to_csv_bytes(df_imgs), file_name="images.csv", mime="text/csv")
            with col_d2:
                urls_only = "\n".join(df_imgs["src"].tolist())
                st.download_button("Download URL List", data=urls_only.encode(), file_name="image_urls.txt", mime="text/plain")
        else:
            st.info("No images found.")
    else:
        st.markdown('<div style="color:var(--muted,#5a7a65);font-size:0.85rem;">Enter a URL and click Scrape.</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TAB: META / SEO
# ─────────────────────────────────────────────
with tab_meta:
    st.markdown(f'<div class="wh-card"><h4>{ICON_META} Meta Data & SEO Analysis</h4>', unsafe_allow_html=True)

    if "meta" in st.session_state.results:
        meta = st.session_state.results["meta"]

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown("**Core SEO**")
            seo_items = {
                "Title": meta.get("title","—"),
                "Description": meta.get("description","—"),
                "Keywords": meta.get("keywords","—"),
                "Canonical": meta.get("canonical","—"),
                "Robots": meta.get("robots","—"),
                "Word Count": meta.get("word_count","—"),
            }
            for k, v in seo_items.items():
                st.markdown(f'<div class="result-item"><b style="color:#5a7a65;font-size:0.72rem;">{k}</b><br>{v}</div>', unsafe_allow_html=True)

        with col_m2:
            st.markdown("**Open Graph**")
            og = meta.get("_open_graph", {})
            if og:
                for k, v in og.items():
                    st.markdown(f'<div class="result-item"><b style="color:#5a7a65;font-size:0.72rem;">{k}</b><br>{v}</div>', unsafe_allow_html=True)
            else:
                st.info("No Open Graph tags found.")

        st.markdown("**H1 Tags**")
        for h in meta.get("h1", []):
            st.markdown(f'<div class="result-item">{h}</div>', unsafe_allow_html=True)

        st.markdown("**H2 Tags**")
        h2s = meta.get("h2", [])
        for h in h2s[:10]:
            st.markdown(f'<div class="result-item">{h}</div>', unsafe_allow_html=True)
        if len(h2s) > 10:
            st.caption(f"... and {len(h2s)-10} more")

        if meta.get("_schema_org"):
            with st.expander("Schema.org / JSON-LD Data"):
                st.json(meta["_schema_org"])

        st.download_button(
            "Download Meta JSON",
            data=to_json_bytes({k: v for k, v in meta.items() if not k.startswith("_")}),
            file_name="meta.json",
            mime="application/json",
        )
    else:
        st.markdown('<div style="color:var(--muted,#5a7a65);font-size:0.85rem;">Enter a URL and click Scrape.</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TAB: CSS SELECTORS
# ─────────────────────────────────────────────
with tab_css:
    st.markdown(f'<div class="wh-card"><h4>{ICON_SEARCH} CSS Selector Extractor</h4>', unsafe_allow_html=True)

    if "url" not in st.session_state.results:
        st.markdown('<div style="color:var(--muted,#5a7a65);font-size:0.85rem;">Scrape a page first, then use CSS selectors.</div>', unsafe_allow_html=True)
    else:
        selectors_input = st.text_area(
            "CSS Selectors (one per line)",
            placeholder="h1\n.product-title\n#price\narticle p\n[data-price]",
            height=120,
        )

        if st.button("Extract Elements", key="css_btn"):
            if selectors_input.strip():
                selectors = [s for s in selectors_input.strip().splitlines() if s.strip()]
                with st.spinner("Re-fetching & parsing..."):
                    resp = fetch_page(st.session_state.results["url"], timeout=timeout, retries=retries)
                if resp:
                    soup = get_soup(resp)
                    css_results = scrape_custom_css(soup, selectors)
                    st.session_state.results["css"] = css_results

        if "css" in st.session_state.results:
            css_results = st.session_state.results["css"]
            for sel, items in css_results.items():
                badge = f'<span class="badge badge-ok">{len(items)} match{"es" if len(items)!=1 else ""}</span>'
                if not items:
                    badge = '<span class="badge badge-warn">0 matches</span>'
                with st.expander(f"{sel}  {badge}", expanded=len(items) > 0):
                    if items:
                        for idx, item in enumerate(items[:50], 1):
                            st.markdown(f'<div class="result-item"><span style="color:var(--muted,#5a7a65);font-size:0.7rem;">#{idx}</span><br>{item}</div>', unsafe_allow_html=True)
                        if len(items) > 50:
                            st.caption(f"Showing 50 of {len(items)} matches.")
                    else:
                        st.caption("No elements matched this selector.")

            st.download_button(
                "Download Results JSON",
                data=to_json_bytes(css_results),
                file_name="css_results.json",
                mime="application/json",
            )

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TAB: ARTICLE TEXT
# ─────────────────────────────────────────────
with tab_article:
    st.markdown(f'<div class="wh-card"><h4>{ICON_META} Article / Main Content Extraction</h4>', unsafe_allow_html=True)

    art_url = st.text_input("Article URL (can differ from main URL)", key="art_url",
                            value=st.session_state.results.get("url", ""),
                            placeholder="https://example.com/article")

    if st.button("Extract Article Text", key="art_btn"):
        if art_url:
            with st.spinner("Extracting main content via trafilatura..."):
                article_text = scrape_article(art_url)
            st.session_state.results["article_text"] = article_text
            log(f"Article extraction: {len(article_text)} chars", "ok")

    if "article_text" in st.session_state.results:
        text = st.session_state.results["article_text"]
        words = len(text.split())
        st.markdown(f'<div style="font-family:\'Space Mono\',monospace;font-size:0.72rem;color:var(--muted,#5a7a65);margin-bottom:0.6rem;">{words} words extracted</div>', unsafe_allow_html=True)
        st.text_area("Extracted Content", value=text, height=350, key="art_out")
        st.download_button("Download as .txt", data=text.encode(), file_name="article.txt", mime="text/plain")

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TAB: SITE CRAWLER
# ─────────────────────────────────────────────
with tab_crawl:
    st.markdown(f'<div class="wh-card"><h4>{ICON_CRAWL} Multi-Page Site Crawler</h4>', unsafe_allow_html=True)

    crawl_url = st.text_input("Start URL", key="crawl_url",
                               value=st.session_state.results.get("url", ""),
                               placeholder="https://example.com")

    col_c1, col_c2, col_c3 = st.columns(3)
    with col_c1:
        c_max = st.number_input("Max Pages", 1, 100, int(max_pages), key="c_max")
    with col_c2:
        c_same = st.checkbox("Same domain only", value=same_domain, key="c_same")
    with col_c3:
        c_robots = st.checkbox("Respect robots.txt", value=respect_robots, key="c_robots")

    crawl_btn = st.button("Start Crawl", key="crawl_btn")

    crawl_status = st.empty()
    crawl_progress = st.empty()

    if crawl_btn and crawl_url:
        if not crawl_url.startswith("http"):
            crawl_url = "https://" + crawl_url

        log(f"Starting crawl: {crawl_url} (max {c_max} pages)", "ok")

        prog = crawl_progress.progress(0)
        status_ph = crawl_status.empty()

        with st.spinner("Crawling..."):
            crawl_results = crawl_site(
                crawl_url,
                max_pages=c_max,
                same_domain=c_same,
                delay=delay,
                respect_robots=c_robots,
                progress_bar=prog,
                status_text=status_ph,
            )

        st.session_state.results["crawl"] = crawl_results
        crawl_progress.empty()
        crawl_status.empty()
        log(f"Crawl complete: {len(crawl_results)} pages", "ok")

    if "crawl" in st.session_state.results:
        crawl_data = st.session_state.results["crawl"]
        df_crawl = pd.DataFrame(crawl_data)

        ok = sum(1 for p in crawl_data if str(p.get("status","")).startswith("2"))
        errs = len(crawl_data) - ok

        st.markdown(f"""
        <div class="stat-row">
            <div class="stat-pill"><span class="val">{len(crawl_data)}</span><span class="lbl">Pages Visited</span></div>
            <div class="stat-pill"><span class="val">{ok}</span><span class="lbl">Success</span></div>
            <div class="stat-pill"><span class="val">{errs}</span><span class="lbl">Errors</span></div>
        </div>
        """, unsafe_allow_html=True)

        st.dataframe(df_crawl, use_container_width=True, hide_index=True)
        st.download_button("Download Crawl CSV", data=to_csv_bytes(df_crawl), file_name="crawl_results.csv", mime="text/csv")

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TAB: EMAILS & SOCIAL
# ─────────────────────────────────────────────
with tab_structured:
    st.markdown(f'<div class="wh-card"><h4>{ICON_SEARCH} Emails, Phones & Social Profiles</h4>', unsafe_allow_html=True)

    if "structured" in st.session_state.results:
        s = st.session_state.results["structured"]

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            st.markdown("**Email Addresses**")
            emails = s.get("emails", [])
            if emails:
                for e in emails:
                    st.markdown(f'<div class="result-item"><a href="mailto:{e}">{e}</a></div>', unsafe_allow_html=True)
            else:
                st.caption("None found.")

            st.markdown("**Phone Numbers**")
            phones = s.get("phones", [])
            if phones:
                for p in phones[:15]:
                    st.markdown(f'<div class="result-item">{p.strip()}</div>', unsafe_allow_html=True)
            else:
                st.caption("None found.")

        with col_s2:
            st.markdown("**Social Profiles**")
            social = s.get("social_profiles", {})
            if social:
                for platform, handles in social.items():
                    st.markdown(f'<div class="result-item"><b style="color:var(--accent,#00D4AA);font-size:0.8rem;">{platform.upper()}</b><br>{"  /  ".join(handles[:5])}</div>', unsafe_allow_html=True)
            else:
                st.caption("No social profiles detected.")

        st.download_button(
            "Download JSON",
            data=to_json_bytes(s),
            file_name="contact_data.json",
            mime="application/json",
        )
    else:
        st.markdown('<div style="color:var(--muted,#5a7a65);font-size:0.85rem;">Enter a URL and click Scrape.</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SCRAPE HISTORY
# ─────────────────────────────────────────────
if st.session_state.history:
    st.markdown("---")
    st.markdown('<div style="font-family:\'Space Mono\',monospace;font-size:0.75rem;color:#5a7a65;margin-bottom:0.4rem;">SCRAPE HISTORY</div>', unsafe_allow_html=True)
    hist_html = ""
    for h in st.session_state.history[-10:][::-1]:
        badge_cls = "badge-ok" if str(h["status"]).startswith("2") else "badge-err"
        hist_html += f'<div class="result-item"><span class="badge {badge_cls}">{h["status"]}</span> <a href="{h["url"]}" target="_blank">{h["url"][:80]}</a> <span style="color:var(--muted,#5a7a65);font-size:0.7rem;float:right;">{h["ts"]}</span></div>'
    st.markdown(hist_html, unsafe_allow_html=True)
