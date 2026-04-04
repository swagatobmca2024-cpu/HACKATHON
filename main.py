"""
WebScraper Pro — Advanced Streamlit Web Scraper
================================================
Deploy-ready for Streamlit Cloud.
All icons are inline SVG — zero emoji dependency.
"""

from __future__ import annotations

import io
import json
import re
import time
from datetime import datetime
from typing import Optional
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE CONFIG  (must be first Streamlit call)
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="WebScraper Pro",
    page_icon="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>🕸</text></svg>",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "Get Help": "https://github.com",
        "Report a bug": None,
        "About": "WebScraper Pro — Advanced Streamlit Web Scraper",
    },
)

# ─────────────────────────────────────────────────────────────────────────────
#  INLINE SVG ICON LIBRARY
# ─────────────────────────────────────────────────────────────────────────────
def svg(key: str, size: int = 18, color: str = "currentColor") -> str:
    """Return an inline SVG string for the given icon key."""
    icons = {
        "spider": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M12 12v10"/><path d="M8 14l-4 3"/><path d="M16 14l4 3"/><path d="M9 12l-5 1"/><path d="M15 12l5 1"/><path d="M8.5 6.5L4 4"/><path d="M15.5 6.5L20 4"/></svg>""",
        "link": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>""",
        "image": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>""",
        "heading": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6 12h12"/><path d="M6 4v16"/><path d="M18 4v16"/></svg>""",
        "table": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18"/><path d="M3 15h18"/><path d="M9 3v18"/></svg>""",
        "target": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg>""",
        "mail": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="4" width="20" height="16" rx="2"/><polyline points="22,4 12,13 2,4"/></svg>""",
        "text": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polyline points="4 7 4 4 20 4 20 7"/><line x1="9" y1="20" x2="15" y2="20"/><line x1="12" y1="4" x2="12" y2="20"/></svg>""",
        "zap": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>""",
        "search": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>""",
        "download": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>""",
        "globe": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>""",
        "settings": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>""",
        "activity": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>""",
        "check": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>""",
        "alert": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><triangle points="10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>""",
        "copy": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>""",
        "filter": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>""",
        "layers": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></svg>""",
        "info": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>""",
        "cpu": f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="16" height="16" rx="2"/><rect x="9" y="9" width="6" height="6"/><line x1="9" y1="1" x2="9" y2="4"/><line x1="15" y1="1" x2="15" y2="4"/><line x1="9" y1="20" x2="9" y2="23"/><line x1="15" y1="20" x2="15" y2="23"/><line x1="20" y1="9" x2="23" y2="9"/><line x1="20" y1="14" x2="23" y2="14"/><line x1="1" y1="9" x2="4" y2="9"/><line x1="1" y1="14" x2="4" y2="14"/></svg>""",
    }
    return icons.get(key, f"""<svg width="{size}" height="{size}" viewBox="0 0 24 24"/>""")


def icon_html(key: str, size: int = 16, color: str = "#a78bfa") -> str:
    return f'<span style="display:inline-flex;align-items:center;vertical-align:middle;">{svg(key, size, color)}</span>'


# ─────────────────────────────────────────────────────────────────────────────
#  GLOBAL CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500;700&family=Sora:wght@300;400;500;600;700&display=swap');

/* ── tokens ───────────────────────────────────────── */
:root {
  --bg:       #080810;
  --s1:       #0f0f1a;
  --s2:       #14141f;
  --s3:       #1a1a2e;
  --border:   #1e1e35;
  --p:        #6d28d9;
  --p2:       #8b5cf6;
  --cyan:     #06b6d4;
  --amber:    #f59e0b;
  --green:    #10b981;
  --red:      #ef4444;
  --txt:      #e2e8f0;
  --muted:    #475569;
  --mono:     'JetBrains Mono', monospace;
  --sans:     'Sora', sans-serif;
}

/* ── base ─────────────────────────────────────────── */
html, body, [class*="css"] { font-family: var(--sans); background: var(--bg) !important; color: var(--txt); }
.stApp { background: var(--bg); }
.block-container { padding-top: 1.2rem !important; max-width: 1300px; }

/* ── scrollbars ───────────────────────────────────── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: var(--s1); }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: var(--p); }

/* ── hero ─────────────────────────────────────────── */
.hero {
  background: linear-gradient(135deg, var(--s1) 0%, #0d0d22 60%, #080d1a 100%);
  border: 1px solid var(--border);
  border-radius: 20px;
  padding: 2.2rem 2.8rem;
  margin-bottom: 1.6rem;
  position: relative;
  overflow: hidden;
}
.hero::before {
  content: '';
  position: absolute; inset: 0;
  background:
    radial-gradient(ellipse 60% 80% at 15% 50%, rgba(109,40,217,.14) 0%, transparent 70%),
    radial-gradient(ellipse 40% 60% at 80% 30%, rgba(6,182,212,.08) 0%, transparent 70%);
  pointer-events: none;
}
.hero-grid {
  position: absolute; inset: 0; opacity: .04;
  background-image:
    linear-gradient(var(--cyan) 1px, transparent 1px),
    linear-gradient(90deg, var(--cyan) 1px, transparent 1px);
  background-size: 40px 40px;
}
.hero-title {
  font-family: var(--mono); font-size: 2.1rem; font-weight: 700;
  background: linear-gradient(100deg, #8b5cf6 0%, #06b6d4 60%, #f59e0b 100%);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  background-clip: text; margin: 0; line-height: 1.2;
}
.hero-sub { color: var(--muted); font-size: .88rem; margin-top: .35rem; font-weight: 300; letter-spacing: .02em; }
.hero-badge {
  display: inline-flex; align-items: center; gap: 6px;
  background: rgba(109,40,217,.12); border: 1px solid rgba(109,40,217,.3);
  border-radius: 999px; padding: 4px 12px;
  font-family: var(--mono); font-size: .7rem; color: #a78bfa;
  margin-top: .8rem;
}

/* ── sidebar ──────────────────────────────────────── */
section[data-testid="stSidebar"] {
  background: var(--s1) !important;
  border-right: 1px solid var(--border);
}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div { color: var(--txt) !important; }
.sidebar-section-title {
  font-family: var(--mono); font-size: .7rem; letter-spacing: .12em;
  text-transform: uppercase; color: var(--muted); margin: 1rem 0 .5rem;
}
.sidebar-mode-item {
  display: flex; align-items: center; gap: 8px;
  padding: 7px 10px; border-radius: 8px;
  font-size: .82rem; color: var(--muted);
  margin-bottom: 2px; transition: background .15s, color .15s;
}
.sidebar-mode-item:hover { background: var(--s3); color: var(--txt); }

/* ── inputs ───────────────────────────────────────── */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea {
  background: var(--s2) !important;
  border: 1px solid var(--border) !important;
  border-radius: 10px !important;
  color: var(--txt) !important;
  font-family: var(--mono) !important;
  font-size: .85rem !important;
  padding: .55rem .9rem !important;
  transition: border-color .2s, box-shadow .2s;
}
.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
  border-color: var(--p2) !important;
  box-shadow: 0 0 0 3px rgba(139,92,246,.15) !important;
}
.stSelectbox > div > div {
  background: var(--s2) !important;
  border: 1px solid var(--border) !important;
  border-radius: 10px !important;
  color: var(--txt) !important;
}

/* ── buttons ──────────────────────────────────────── */
.stButton > button {
  background: linear-gradient(135deg, var(--p) 0%, #4c1d95 100%) !important;
  color: #fff !important; border: none !important;
  border-radius: 10px !important;
  font-family: var(--mono) !important;
  font-size: .82rem !important; font-weight: 600 !important;
  padding: .65rem 1.8rem !important; letter-spacing: .04em;
  transition: transform .15s, box-shadow .15s !important;
}
.stButton > button:hover {
  transform: translateY(-2px);
  box-shadow: 0 6px 24px rgba(109,40,217,.45) !important;
}
.stButton > button:active { transform: translateY(0); }
.stDownloadButton > button {
  background: var(--s3) !important;
  border: 1px solid var(--border) !important;
  color: var(--cyan) !important;
  border-radius: 8px !important;
  font-family: var(--mono) !important; font-size: .76rem !important;
}
.stDownloadButton > button:hover {
  border-color: var(--cyan) !important;
  box-shadow: 0 0 10px rgba(6,182,212,.2) !important;
}

/* ── tabs ─────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
  background: var(--s2) !important;
  border-radius: 12px !important; padding: 4px !important; gap: 3px !important;
  border: 1px solid var(--border);
}
.stTabs [data-baseweb="tab"] {
  background: transparent !important;
  color: var(--muted) !important;
  border-radius: 8px !important;
  font-family: var(--mono) !important; font-size: .75rem !important;
  padding: .4rem .9rem !important;
}
.stTabs [aria-selected="true"] {
  background: var(--p) !important; color: #fff !important;
}

/* ── expanders ────────────────────────────────────── */
.streamlit-expanderHeader {
  background: var(--s2) !important;
  border: 1px solid var(--border) !important;
  border-radius: 10px !important;
  font-family: var(--mono) !important; font-size: .8rem !important;
  color: var(--txt) !important;
}
.streamlit-expanderContent {
  background: var(--s1) !important;
  border: 1px solid var(--border) !important;
  border-top: none !important; border-radius: 0 0 10px 10px !important;
}

/* ── metric card ──────────────────────────────────── */
.metric-card {
  background: var(--s2);
  border: 1px solid var(--border);
  border-radius: 14px; padding: 1.3rem 1rem;
  text-align: center; position: relative; overflow: hidden;
  transition: border-color .2s, transform .2s;
}
.metric-card:hover { border-color: var(--p2); transform: translateY(-2px); }
.metric-card::before {
  content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
  background: linear-gradient(90deg, var(--p), var(--cyan));
}
.metric-val {
  font-family: var(--mono); font-size: 2rem; font-weight: 700;
  background: linear-gradient(135deg, #a78bfa, var(--cyan));
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  background-clip: text;
}
.metric-lbl {
  font-size: .72rem; color: var(--muted);
  text-transform: uppercase; letter-spacing: .1em; margin-top: 2px;
}
.metric-icon { margin-bottom: 6px; opacity: .6; }

/* ── result rows ──────────────────────────────────── */
.result-row {
  background: var(--s2);
  border: 1px solid var(--border);
  border-left: 3px solid var(--p);
  border-radius: 10px; padding: .9rem 1.1rem;
  margin: .4rem 0; font-size: .84rem; line-height: 1.5;
  transition: border-left-color .2s, background .2s;
}
.result-row:hover { border-left-color: var(--cyan); background: var(--s3); }
.result-row a { color: var(--cyan); text-decoration: none; }
.result-row a:hover { text-decoration: underline; }
.tag {
  display: inline-flex; align-items: center; gap: 4px;
  background: rgba(109,40,217,.15);
  border: 1px solid rgba(109,40,217,.3);
  color: #a78bfa; border-radius: 6px;
  padding: 2px 9px; font-family: var(--mono); font-size: .7rem;
  margin-right: 6px;
}
.tag-cyan { background: rgba(6,182,212,.1); border-color: rgba(6,182,212,.3); color: var(--cyan); }
.tag-amber { background: rgba(245,158,11,.1); border-color: rgba(245,158,11,.3); color: var(--amber); }
.tag-green { background: rgba(16,185,129,.1); border-color: rgba(16,185,129,.3); color: var(--green); }
.tag-red   { background: rgba(239,68,68,.1);  border-color: rgba(239,68,68,.3);  color: var(--red);   }

/* ── log terminal ─────────────────────────────────── */
.terminal {
  background: #03030a;
  border: 1px solid var(--border);
  border-radius: 12px; padding: 1rem 1.2rem;
  font-family: var(--mono); font-size: .76rem;
  color: var(--green); max-height: 180px; overflow-y: auto;
  line-height: 1.8; letter-spacing: .01em;
}
.terminal .ts { color: var(--muted); }
.terminal .ok { color: var(--green); }
.terminal .err { color: var(--red); }
.terminal .info { color: var(--cyan); }

/* ── page meta block ──────────────────────────────── */
.meta-block {
  background: var(--s2); border: 1px solid var(--border);
  border-radius: 14px; padding: 1.2rem 1.4rem; margin-bottom: 1rem;
}
.meta-key {
  font-family: var(--mono); font-size: .72rem;
  color: var(--muted); text-transform: uppercase; letter-spacing: .08em;
}
.meta-val { font-size: .88rem; color: var(--txt); margin-top: 2px; word-break: break-all; }

/* ── section header ───────────────────────────────── */
.sec-header {
  display: flex; align-items: center; gap: 10px;
  font-family: var(--mono); font-size: .9rem; font-weight: 600;
  color: #a78bfa; margin: 1.2rem 0 .7rem;
  padding-bottom: .5rem; border-bottom: 1px solid var(--border);
}

/* ── empty state ──────────────────────────────────── */
.empty-state {
  text-align: center; padding: 3.5rem 1rem;
  background: var(--s1); border: 1px dashed var(--border);
  border-radius: 16px; margin: 1rem 0;
}
.empty-title { font-family: var(--mono); font-size: 1.1rem; color: #a78bfa; margin: 1rem 0 .4rem; }
.empty-sub { color: var(--muted); font-size: .85rem; max-width: 420px; margin: 0 auto; line-height: 1.6; }

/* ── dataframe override ───────────────────────────── */
.stDataFrame { border-radius: 10px !important; overflow: hidden; }
.stDataFrame thead th { background: var(--s3) !important; font-family: var(--mono) !important; }

/* ── progress / spinner ───────────────────────────── */
.stProgress > div > div { background: var(--p) !important; }
.stSpinner > div { border-top-color: var(--p2) !important; }

/* ── divider ──────────────────────────────────────── */
hr { border-color: var(--border) !important; margin: 1.5rem 0 !important; }

/* ── welcome grid cards ───────────────────────────── */
.mode-card {
  background: var(--s2); border: 1px solid var(--border);
  border-radius: 14px; padding: 1.1rem 1.2rem; height: 120px;
  display: flex; flex-direction: column; justify-content: center;
  transition: border-color .2s, transform .2s;
}
.mode-card:hover { border-color: var(--p2); transform: translateY(-3px); }
.mode-card-icon { margin-bottom: 8px; }
.mode-card-name {
  font-family: var(--mono); font-size: .78rem; color: #a78bfa; font-weight: 600;
}
.mode-card-desc { font-size: .73rem; color: var(--muted); margin-top: 3px; line-height: 1.4; }

/* ── checkbox ─────────────────────────────────────── */
.stCheckbox > label { color: var(--muted) !important; font-size: .82rem !important; }

/* ── slider ───────────────────────────────────────── */
.stSlider > div > div > div { background: var(--p) !important; }

/* ── alerts ───────────────────────────────────────── */
.stAlert { border-radius: 10px !important; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
#  SESSION STATE INIT
# ─────────────────────────────────────────────────────────────────────────────
for key, default in {
    "scrape_log": [],
    "last_result": None,
    "last_url": "",
    "last_mode": "",
    "request_count": 0,
    "total_items": 0,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


# ─────────────────────────────────────────────────────────────────────────────
#  CONSTANTS & HELPERS
# ─────────────────────────────────────────────────────────────────────────────
REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "DNT": "1",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")


def log(msg: str, kind: str = "ok"):
    ts = datetime.now().strftime("%H:%M:%S")
    st.session_state.scrape_log.append(
        f'<span class="ts">[{ts}]</span> <span class="{kind}">{msg}</span>'
    )


def to_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def to_json(data) -> bytes:
    return json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")


def metric_card(col, icon_key: str, value, label: str, color: str = "#06b6d4"):
    col.markdown(
        f"""<div class="metric-card">
          <div class="metric-icon">{svg(icon_key, 22, color)}</div>
          <div class="metric-val">{value}</div>
          <div class="metric-lbl">{label}</div>
        </div>""",
        unsafe_allow_html=True,
    )


def section_header(icon_key: str, title: str):
    st.markdown(
        f"""<div class="sec-header">
          {svg(icon_key, 18, "#8b5cf6")}&nbsp;{title}
        </div>""",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
#  SCRAPING FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(ttl=120, show_spinner=False)
def fetch_page(url: str, timeout: int) -> tuple[Optional[str], int, str]:
    """Cached page fetch — returns (html, status_code, error)."""
    try:
        r = requests.get(url, headers=REQUEST_HEADERS, timeout=timeout)
        return r.text, r.status_code, ""
    except requests.exceptions.ConnectionError:
        return None, 0, "Connection refused or DNS failure."
    except requests.exceptions.Timeout:
        return None, 0, f"Timed out after {timeout}s."
    except requests.exceptions.InvalidURL:
        return None, 0, "Invalid URL format."
    except Exception as exc:
        return None, 0, str(exc)


def parse(html: str) -> BeautifulSoup:
    return BeautifulSoup(html, "lxml")


def extract_meta(soup: BeautifulSoup) -> dict:
    def _get(soup, **kwargs):
        tag = soup.find("meta", attrs=kwargs)
        return tag.get("content", "—") if tag else "—"

    title_tag = soup.find("title")
    return {
        "Title": title_tag.get_text(strip=True) if title_tag else "—",
        "Description": _get(soup, attrs={"name": "description"}),
        "Keywords": _get(soup, attrs={"name": "keywords"}),
        "Author": _get(soup, attrs={"name": "author"}),
        "Robots": _get(soup, attrs={"name": "robots"}),
        "OG Title": _get(soup, property="og:title"),
        "OG Description": _get(soup, property="og:description"),
        "OG Image": _get(soup, property="og:image"),
        "OG Type": _get(soup, property="og:type"),
        "Twitter Card": _get(soup, attrs={"name": "twitter:card"}),
        "Canonical": (lambda t: t["href"] if t else "—")(soup.find("link", rel="canonical")),
        "Charset": (lambda t: t.get("charset", "—") if t else "—")(soup.find("meta", charset=True)),
    }


def extract_links(soup: BeautifulSoup, base: str) -> pd.DataFrame:
    rows = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href or href.startswith(("javascript:", "#", "mailto:", "tel:")):
            continue
        full = urljoin(base, href)
        parsed = urlparse(full)
        rows.append({
            "Text": a.get_text(strip=True)[:120] or "(no text)",
            "URL": full,
            "Domain": parsed.netloc,
            "Path": parsed.path,
            "External": parsed.netloc != urlparse(base).netloc,
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def extract_images(soup: BeautifulSoup, base: str) -> pd.DataFrame:
    rows = []
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-src") or img.get("data-lazy-src") or ""
        src = src.strip()
        full = urljoin(base, src) if src else ""
        rows.append({
            "Alt": img.get("alt", "")[:100],
            "URL": full,
            "Width": img.get("width", ""),
            "Height": img.get("height", ""),
            "Loading": img.get("loading", ""),
            "Class": " ".join(img.get("class", [])),
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def extract_headings(soup: BeautifulSoup) -> pd.DataFrame:
    rows = []
    for tag in ["h1", "h2", "h3", "h4", "h5", "h6"]:
        for el in soup.find_all(tag):
            text = el.get_text(strip=True)
            if text:
                rows.append({"Level": tag.upper(), "Text": text, "ID": el.get("id", "")})
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def extract_tables(soup: BeautifulSoup) -> list[pd.DataFrame]:
    out = []
    for tbl in soup.find_all("table"):
        try:
            df = pd.read_html(io.StringIO(str(tbl)))[0]
            if not df.empty:
                out.append(df)
        except Exception:
            pass
    return out


def extract_css(soup: BeautifulSoup, selector: str) -> pd.DataFrame:
    try:
        els = soup.select(selector)
        rows = [{
            "Tag": el.name,
            "Text": el.get_text(strip=True)[:200],
            "Class": " ".join(el.get("class", [])),
            "ID": el.get("id", ""),
            "Href": el.get("href", ""),
            "Src": el.get("src", ""),
        } for el in els]
        return pd.DataFrame(rows) if rows else pd.DataFrame()
    except Exception as e:
        return pd.DataFrame({"Error": [str(e)]})


def extract_emails(soup: BeautifulSoup) -> list[str]:
    return sorted(set(EMAIL_RE.findall(soup.get_text())))


def extract_paragraphs(soup: BeautifulSoup) -> list[str]:
    return [p.get_text(strip=True) for p in soup.find_all("p") if len(p.get_text(strip=True)) > 20]


def extract_scripts_styles(soup: BeautifulSoup, base: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    scripts, styles = [], []
    for s in soup.find_all("script"):
        src = s.get("src", "")
        scripts.append({"Type": s.get("type", "text/javascript"), "Src": urljoin(base, src) if src else "(inline)", "Async": "async" in s.attrs, "Defer": "defer" in s.attrs})
    for l in soup.find_all("link", rel=lambda v: v and "stylesheet" in v):
        href = l.get("href", "")
        styles.append({"Href": urljoin(base, href) if href else "", "Media": l.get("media", "all")})
    return pd.DataFrame(scripts) if scripts else pd.DataFrame(), pd.DataFrame(styles) if styles else pd.DataFrame()


# ─────────────────────────────────────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:10px;padding:.5rem 0 .3rem;">
      {svg('spider', 28, '#8b5cf6')}
      <span style="font-family:var(--mono);font-size:1rem;font-weight:700;
        background:linear-gradient(90deg,#8b5cf6,#06b6d4);
        -webkit-background-clip:text;-webkit-text-fill-color:transparent;
        background-clip:text;">WebScraper Pro</span>
    </div>
    <div style="font-size:.72rem;color:#475569;margin-bottom:1rem;font-family:var(--mono);">
      v2.0 · BeautifulSoup + Requests
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f'<div class="sidebar-section-title">{icon_html("settings",12,"#475569")} Configuration</div>', unsafe_allow_html=True)

    timeout = st.slider("Request Timeout (s)", 5, 30, 12, key="timeout")
    delay   = st.slider("Polite Delay (s)", 0.0, 3.0, 0.5, step=0.5, key="delay")
    max_rows = st.slider("Max Rows to Display", 50, 500, 200, step=50, key="max_rows")
    show_raw = st.checkbox("Show raw HTML snippet", value=False, key="show_raw")

    st.markdown("---")
    st.markdown(f'<div class="sidebar-section-title">{icon_html("layers",12,"#475569")} Scrape Modes</div>', unsafe_allow_html=True)

    mode_map = {
        "Full Analysis":  ("layers",  "Everything at once"),
        "Links":          ("link",    "All anchor tags"),
        "Images":         ("image",   "img src + alt"),
        "Headings":       ("heading", "H1–H6 structure"),
        "Tables":         ("table",   "HTML tables → CSV"),
        "CSS Selector":   ("target",  "Custom element query"),
        "Emails":         ("mail",    "Regex email harvest"),
        "Paragraphs":     ("text",    "Body text extraction"),
        "Assets":         ("cpu",     "Scripts & stylesheets"),
    }
    for name, (icon, desc) in mode_map.items():
        st.markdown(
            f'<div class="sidebar-mode-item">{svg(icon,14,"#6d28d9")}'
            f'<span style="flex:1">{name}</span>'
            f'<span style="font-size:.65rem;color:#334155">{desc}</span></div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")
    if st.session_state.request_count:
        st.markdown(f"""
        <div style="background:var(--s2);border:1px solid var(--border);border-radius:10px;padding:.9rem 1rem;">
          <div style="font-family:var(--mono);font-size:.7rem;color:var(--muted);margin-bottom:.5rem;">SESSION STATS</div>
          <div style="display:flex;justify-content:space-between;font-size:.78rem;">
            <span style="color:#a78bfa">Requests</span>
            <span style="font-family:var(--mono);color:var(--txt)">{st.session_state.request_count}</span>
          </div>
          <div style="display:flex;justify-content:space-between;font-size:.78rem;margin-top:4px;">
            <span style="color:#a78bfa">Items scraped</span>
            <span style="font-family:var(--mono);color:var(--txt)">{st.session_state.total_items:,}</span>
          </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div style="font-size:.68rem;color:#1e293b;margin-top:1rem;line-height:1.7;">
      Always respect robots.txt &amp; ToS.<br>Use polite delays on production.
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
#  HERO HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="hero">
  <div class="hero-grid"></div>
  <div style="position:relative;z-index:1;display:flex;align-items:center;gap:1.2rem;">
    <div>{svg('spider', 48, '#8b5cf6')}</div>
    <div>
      <div class="hero-title">WebScraper Pro</div>
      <div class="hero-sub">Advanced web data extraction &mdash; BeautifulSoup · lxml · Requests · Pandas</div>
      <div class="hero-badge">
        {svg('zap',12,'#a78bfa')} Deploy-ready &nbsp;|&nbsp; {svg('globe',12,'#a78bfa')} Any public URL &nbsp;|&nbsp;
        {svg('download',12,'#a78bfa')} CSV &amp; JSON export
      </div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
#  INPUT ROW
# ─────────────────────────────────────────────────────────────────────────────
col_url, col_mode = st.columns([3, 2])
with col_url:
    url_input = st.text_input(
        f"{icon_html('globe',14,'#8b5cf6')} Target URL",
        placeholder="https://example.com",
        label_visibility="visible",
        key="url_input",
    )
with col_mode:
    mode = st.selectbox(
        f"{icon_html('layers',14,'#8b5cf6')} Scrape Mode",
        list(mode_map.keys()),
        key="mode_select",
    )

custom_selector = ""
if mode == "CSS Selector":
    custom_selector = st.text_input(
        f"{icon_html('target',14,'#8b5cf6')} CSS Selector",
        placeholder="e.g.  article h2,  div.price,  table.wikitable,  nav a",
        key="css_selector",
    )

scrape_btn = st.button(
    "SCRAPE NOW",
    use_container_width=True,
    key="scrape_btn",
)

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN SCRAPING BLOCK
# ─────────────────────────────────────────────────────────────────────────────
if scrape_btn:
    url = url_input.strip()
    if not url:
        st.error("Please enter a URL first.")
        st.stop()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    # reset log for this run
    st.session_state.scrape_log = []
    log(f"Initialising scrape → {url}", "info")

    prog = st.progress(0, text="Fetching page…")

    with st.spinner(""):
        html, status, err = fetch_page(url, timeout)
        time.sleep(delay)

    prog.progress(25, text="Parsing HTML…")

    if err or html is None:
        st.error(f"Fetch failed — {err}")
        log(f"ERROR: {err}", "err")
        st.stop()

    log(f"HTTP {status} · {len(html):,} bytes received", "ok")
    soup = parse(html)
    st.session_state.request_count += 1

    prog.progress(50, text="Extracting data…")

    # ── FULL ANALYSIS ────────────────────────────────────────────────────────
    if mode == "Full Analysis":
        meta       = extract_meta(soup)
        df_links   = extract_links(soup, url)
        df_images  = extract_images(soup, url)
        df_heads   = extract_headings(soup)
        emails     = extract_emails(soup)
        tables     = extract_tables(soup)
        paragraphs = extract_paragraphs(soup)

        prog.progress(90, text="Rendering…")
        log(f"Full analysis: {len(df_links)} links, {len(df_images)} images, {len(df_heads)} headings, {len(emails)} emails, {len(tables)} tables", "ok")

        # metrics
        cols = st.columns(6)
        metric_card(cols[0], "link",    len(df_links),    "Links",      "#06b6d4")
        metric_card(cols[1], "image",   len(df_images),   "Images",     "#8b5cf6")
        metric_card(cols[2], "heading", len(df_heads),    "Headings",   "#f59e0b")
        metric_card(cols[3], "mail",    len(emails),      "Emails",     "#10b981")
        metric_card(cols[4], "table",   len(tables),      "Tables",     "#ef4444")
        metric_card(cols[5], "text",    len(paragraphs),  "Paragraphs", "#a78bfa")

        # meta
        section_header("info", "Page Metadata")
        g1, g2 = st.columns(2)
        meta_items = list(meta.items())
        for i, (k, v) in enumerate(meta_items):
            col = g1 if i % 2 == 0 else g2
            col.markdown(
                f'<div class="meta-block"><div class="meta-key">{k}</div>'
                f'<div class="meta-val">{v or "—"}</div></div>',
                unsafe_allow_html=True,
            )

        # tabs
        t1, t2, t3, t4, t5, t6 = st.tabs(["Links", "Images", "Headings", "Emails", "Tables", "Text"])

        with t1:
            if not df_links.empty:
                st.dataframe(df_links.head(max_rows), use_container_width=True, height=350)
                c1, c2 = st.columns(2)
                c1.download_button(f"{icon_html('download',12,'#06b6d4')} CSV", to_csv(df_links), "links.csv", "text/csv")
                c2.download_button(f"{icon_html('download',12,'#06b6d4')} JSON", to_json(df_links.to_dict("records")), "links.json")
            else:
                st.info("No links found.")

        with t2:
            if not df_images.empty:
                st.dataframe(df_images.head(max_rows), use_container_width=True, height=300)
                preview_cols = st.columns(3)
                shown = 0
                for _, row in df_images.iterrows():
                    if shown >= 6: break
                    if row["URL"].startswith("http"):
                        try:
                            preview_cols[shown % 3].image(row["URL"], caption=row["Alt"][:40], use_container_width=True)
                            shown += 1
                        except Exception:
                            pass
                st.download_button(f"{icon_html('download',12,'#06b6d4')} CSV", to_csv(df_images), "images.csv", "text/csv")

        with t3:
            if not df_heads.empty:
                level_colors = {"H1":"#8b5cf6","H2":"#06b6d4","H3":"#f59e0b","H4":"#10b981","H5":"#ef4444","H6":"#a78bfa"}
                for _, row in df_heads.iterrows():
                    c = level_colors.get(row["Level"], "#94a3b8")
                    indent = "&nbsp;" * (int(row["Level"][1]) - 1) * 6
                    st.markdown(
                        f'<div class="result-row">{indent}'
                        f'<span class="tag" style="background:rgba(0,0,0,.2);border-color:{c};color:{c};">{row["Level"]}</span>'
                        f' {row["Text"]}</div>',
                        unsafe_allow_html=True,
                    )
                st.download_button(f"{icon_html('download',12,'#06b6d4')} CSV", to_csv(df_heads), "headings.csv", "text/csv")

        with t4:
            if emails:
                for em in emails:
                    st.markdown(f'<div class="result-row">{svg("mail",14,"#10b981")} &nbsp;<code style="color:#10b981">{em}</code></div>', unsafe_allow_html=True)
                df_em = pd.DataFrame({"email": emails})
                st.download_button(f"{icon_html('download',12,'#06b6d4')} CSV", to_csv(df_em), "emails.csv", "text/csv")
            else:
                st.info("No emails found.")

        with t5:
            if tables:
                for i, tbl in enumerate(tables):
                    with st.expander(f"Table {i+1} — {tbl.shape[0]} rows × {tbl.shape[1]} cols"):
                        st.dataframe(tbl, use_container_width=True)
                        st.download_button(f"{icon_html('download',12,'#06b6d4')} CSV", to_csv(tbl), f"table_{i+1}.csv", "text/csv", key=f"tbl_{i}")
            else:
                st.info("No HTML tables found.")

        with t6:
            kw = st.text_input("Filter paragraphs", placeholder="keyword…", key="full_para_filter")
            filtered_p = [p for p in paragraphs if kw.lower() in p.lower()] if kw else paragraphs
            for p in filtered_p[:60]:
                st.markdown(f'<div class="result-row">{p}</div>', unsafe_allow_html=True)

        st.session_state.total_items += len(df_links) + len(df_images) + len(df_heads)

    # ── LINKS ────────────────────────────────────────────────────────────────
    elif mode == "Links":
        df = extract_links(soup, url)
        prog.progress(90, text="Rendering…")
        log(f"Extracted {len(df)} links", "ok")

        metric_card(st.columns(3)[1], "link", len(df), "Links Found", "#06b6d4")
        section_header("filter", "Filter & Explore")

        col_f1, col_f2, col_f3 = st.columns(3)
        kw       = col_f1.text_input("URL keyword filter", placeholder="/blog or .pdf")
        ext_only = col_f2.checkbox("External links only")
        int_only = col_f3.checkbox("Internal links only")

        if not df.empty:
            if kw:        df = df[df["URL"].str.contains(kw, case=False, na=False)]
            if ext_only:  df = df[df["External"] == True]
            if int_only:  df = df[df["External"] == False]

            st.dataframe(df.head(max_rows), use_container_width=True, height=420)
            c1, c2 = st.columns(2)
            c1.download_button(f"{icon_html('download',12,'#06b6d4')} CSV",  to_csv(df),                        "links.csv",  "text/csv")
            c2.download_button(f"{icon_html('download',12,'#06b6d4')} JSON", to_json(df.to_dict("records")), "links.json")
        else:
            st.info("No links found on this page.")
        st.session_state.total_items += len(df)

    # ── IMAGES ───────────────────────────────────────────────────────────────
    elif mode == "Images":
        df = extract_images(soup, url)
        prog.progress(90, text="Rendering…")
        log(f"Extracted {len(df)} images", "ok")

        cols3 = st.columns(3)
        metric_card(cols3[0], "image", len(df), "Images", "#8b5cf6")
        metric_card(cols3[1], "image", df["Alt"].astype(bool).sum() if not df.empty else 0, "With Alt Text", "#10b981")
        metric_card(cols3[2], "image", (df["Loading"] == "lazy").sum() if not df.empty else 0, "Lazy Loaded", "#f59e0b")

        if not df.empty:
            kw = st.text_input(f"{icon_html('search',12,'#8b5cf6')} Filter images by alt text or URL")
            if kw:
                df = df[df["Alt"].str.contains(kw, case=False, na=False) | df["URL"].str.contains(kw, case=False, na=False)]

            section_header("image", "Image Preview (first 9)")
            preview_cols = st.columns(3)
            shown = 0
            for _, row in df.iterrows():
                if shown >= 9: break
                if row["URL"].startswith("http"):
                    try:
                        preview_cols[shown % 3].image(row["URL"], caption=row["Alt"][:50] or "(no alt)", use_container_width=True)
                        shown += 1
                    except Exception:
                        pass

            section_header("table", "Full Image Data")
            st.dataframe(df.head(max_rows), use_container_width=True, height=300)
            st.download_button(f"{icon_html('download',12,'#06b6d4')} CSV", to_csv(df), "images.csv", "text/csv")
        else:
            st.info("No images found.")
        st.session_state.total_items += len(df)

    # ── HEADINGS ─────────────────────────────────────────────────────────────
    elif mode == "Headings":
        df = extract_headings(soup)
        prog.progress(90, text="Rendering…")
        log(f"Extracted {len(df)} headings", "ok")

        if not df.empty:
            level_counts = df["Level"].value_counts().to_dict()
            cols = st.columns(min(len(level_counts), 6))
            level_colors = {"H1":"#8b5cf6","H2":"#06b6d4","H3":"#f59e0b","H4":"#10b981","H5":"#ef4444","H6":"#a78bfa"}
            for i, (lv, cnt) in enumerate(sorted(level_counts.items())):
                metric_card(cols[i], "heading", cnt, f"{lv} Tags", level_colors.get(lv, "#94a3b8"))

            section_header("heading", "Heading Hierarchy")
            for _, row in df.iterrows():
                depth = int(row["Level"][1])
                c = level_colors.get(row["Level"], "#94a3b8")
                indent = "&nbsp;" * (depth - 1) * 8
                st.markdown(
                    f'<div class="result-row">{indent}'
                    f'<span style="color:{c};font-family:var(--mono);font-size:.72rem;font-weight:700;">{row["Level"]}</span>'
                    f'&nbsp;&nbsp;{row["Text"]}'
                    + (f' <span class="tag tag-cyan" style="float:right">#{row["ID"]}</span>' if row["ID"] else "")
                    + '</div>',
                    unsafe_allow_html=True,
                )
            st.download_button(f"{icon_html('download',12,'#06b6d4')} CSV", to_csv(df), "headings.csv", "text/csv")
        else:
            st.info("No headings found.")
        st.session_state.total_items += len(df)

    # ── TABLES ───────────────────────────────────────────────────────────────
    elif mode == "Tables":
        tables = extract_tables(soup)
        prog.progress(90, text="Rendering…")
        log(f"Extracted {len(tables)} tables", "ok")

        metric_card(st.columns(3)[1], "table", len(tables), "Tables Found", "#ef4444")

        if tables:
            total_cells = sum(t.shape[0] * t.shape[1] for t in tables)
            st.markdown(f'<div class="result-row">{svg("info",14,"#a78bfa")} &nbsp; <b>{total_cells:,}</b> total cells across all tables</div>', unsafe_allow_html=True)
            for i, tbl in enumerate(tables):
                with st.expander(f"Table {i+1}  ·  {tbl.shape[0]} rows × {tbl.shape[1]} columns", expanded=i == 0):
                    st.dataframe(tbl.head(max_rows), use_container_width=True)
                    cc1, cc2 = st.columns(2)
                    cc1.download_button(f"{icon_html('download',12,'#06b6d4')} CSV",  to_csv(tbl),                        f"table_{i+1}.csv",  "text/csv",           key=f"csv_{i}")
                    cc2.download_button(f"{icon_html('download',12,'#06b6d4')} JSON", to_json(tbl.to_dict("records")), f"table_{i+1}.json",                          key=f"json_{i}")
        else:
            st.info("No HTML `<table>` elements found on this page.")
        st.session_state.total_items += sum(len(t) for t in tables)

    # ── CSS SELECTOR ─────────────────────────────────────────────────────────
    elif mode == "CSS Selector":
        if not custom_selector.strip():
            st.warning("Enter a CSS selector above to begin.")
        else:
            df = extract_css(soup, custom_selector)
            prog.progress(90, text="Rendering…")
            log(f"Selector '{custom_selector}' matched {len(df)} elements", "ok")

            cols2 = st.columns(2)
            metric_card(cols2[0], "target", len(df), "Elements Matched", "#8b5cf6")
            unique_tags = df["Tag"].nunique() if not df.empty and "Tag" in df.columns else 0
            metric_card(cols2[1], "layers", unique_tags, "Unique Tag Types", "#06b6d4")

            if not df.empty:
                section_header("target", f"Results for: {custom_selector}")
                st.dataframe(df.head(max_rows), use_container_width=True, height=400)
                cc1, cc2 = st.columns(2)
                cc1.download_button(f"{icon_html('download',12,'#06b6d4')} CSV",  to_csv(df),                        "selector.csv",  "text/csv")
                cc2.download_button(f"{icon_html('download',12,'#06b6d4')} JSON", to_json(df.to_dict("records")), "selector.json")
            else:
                st.info(f"No elements matched `{custom_selector}`")
            st.session_state.total_items += len(df)

    # ── EMAILS ───────────────────────────────────────────────────────────────
    elif mode == "Emails":
        emails = extract_emails(soup)
        prog.progress(90, text="Rendering…")
        log(f"Found {len(emails)} email addresses", "ok")

        metric_card(st.columns(3)[1], "mail", len(emails), "Emails Found", "#10b981")

        if emails:
            section_header("mail", "Discovered Emails")
            for em in emails:
                domain = em.split("@")[-1]
                st.markdown(
                    f'<div class="result-row">{svg("mail",14,"#10b981")}&nbsp;&nbsp;'
                    f'<code style="color:#10b981">{em}</code>'
                    f'&nbsp;<span class="tag tag-cyan">@{domain}</span></div>',
                    unsafe_allow_html=True,
                )
            df_em = pd.DataFrame({"email": emails, "domain": [e.split("@")[-1] for e in emails]})
            st.download_button(f"{icon_html('download',12,'#06b6d4')} CSV", to_csv(df_em), "emails.csv", "text/csv")
        else:
            st.info("No email addresses found.")
        st.session_state.total_items += len(emails)

    # ── PARAGRAPHS ───────────────────────────────────────────────────────────
    elif mode == "Paragraphs":
        paragraphs = extract_paragraphs(soup)
        prog.progress(90, text="Rendering…")
        log(f"Extracted {len(paragraphs)} paragraphs", "ok")

        word_count = sum(len(p.split()) for p in paragraphs)
        char_count = sum(len(p) for p in paragraphs)
        cols3 = st.columns(3)
        metric_card(cols3[0], "text",   len(paragraphs), "Paragraphs", "#a78bfa")
        metric_card(cols3[1], "search", word_count,      "Words",      "#06b6d4")
        metric_card(cols3[2], "copy",   char_count,      "Characters", "#f59e0b")

        section_header("search", "Search & Browse")
        kw = st.text_input("Keyword filter", placeholder="Search within paragraphs…")
        filtered = [p for p in paragraphs if kw.lower() in p.lower()] if kw else paragraphs

        if kw:
            st.markdown(f'<span class="tag tag-amber">{len(filtered)} / {len(paragraphs)} match</span>', unsafe_allow_html=True)

        for p in filtered[:max_rows]:
            highlighted = p
            if kw:
                highlighted = re.sub(
                    f"({re.escape(kw)})",
                    r'<mark style="background:#f59e0b33;color:#f59e0b;border-radius:3px;">\1</mark>',
                    highlighted, flags=re.IGNORECASE,
                )
            st.markdown(f'<div class="result-row">{highlighted}</div>', unsafe_allow_html=True)

        if paragraphs:
            st.download_button(f"{icon_html('download',12,'#06b6d4')} Download text", "\n\n".join(paragraphs).encode(), "paragraphs.txt", "text/plain")
        st.session_state.total_items += len(paragraphs)

    # ── ASSETS ───────────────────────────────────────────────────────────────
    elif mode == "Assets":
        df_scripts, df_styles = extract_scripts_styles(soup, url)
        prog.progress(90, text="Rendering…")
        log(f"Found {len(df_scripts)} scripts, {len(df_styles)} stylesheets", "ok")

        cols2 = st.columns(2)
        metric_card(cols2[0], "cpu",    len(df_scripts), "Scripts",      "#8b5cf6")
        metric_card(cols2[1], "layers", len(df_styles),  "Stylesheets",  "#06b6d4")

        section_header("cpu", "JavaScript Files")
        if not df_scripts.empty:
            st.dataframe(df_scripts, use_container_width=True, height=300)
            st.download_button(f"{icon_html('download',12,'#06b6d4')} CSV", to_csv(df_scripts), "scripts.csv", "text/csv")
        else:
            st.info("No script tags found.")

        section_header("layers", "Stylesheets")
        if not df_styles.empty:
            st.dataframe(df_styles, use_container_width=True, height=220)
            st.download_button(f"{icon_html('download',12,'#06b6d4')} CSS list", to_csv(df_styles), "styles.csv", "text/csv")
        else:
            st.info("No linked stylesheets found.")

    prog.progress(100, text="Done")

    # ── raw HTML snippet ─────────────────────────────────────────────────────
    if show_raw:
        section_header("copy", "Raw HTML (first 4000 chars)")
        st.code(html[:4000], language="html")

    # ── activity log ─────────────────────────────────────────────────────────
    log("Scrape complete.", "ok")
    section_header("activity", "Activity Log")
    log_html = "<br>".join([f"&gt; {l}" for l in st.session_state.scrape_log])
    st.markdown(f'<div class="terminal">{log_html}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
#  WELCOME / IDLE STATE
# ─────────────────────────────────────────────────────────────────────────────
else:
    st.markdown(f"""
    <div class="empty-state">
      {svg('spider', 52, '#4c1d95')}
      <div class="empty-title">Ready to extract</div>
      <div class="empty-sub">
        Paste a URL above, choose a scrape mode, and hit
        <strong style="color:#8b5cf6">SCRAPE NOW</strong>.<br>
        Supports links, images, headings, tables, emails, CSS selectors, paragraphs &amp; asset maps.
      </div>
    </div>
    """, unsafe_allow_html=True)

    section_header("layers", "Available Modes")
    grid_cols = st.columns(3)
    modes_desc = [
        ("layers",  "Full Analysis",  "Complete page breakdown: meta, links, images, headings, emails, tables all at once."),
        ("link",    "Links",          "All anchor tags with URL, text, domain and internal/external detection."),
        ("image",   "Images",         "Harvests img src + alt, lazy-load status, with inline previews."),
        ("heading", "Headings",       "H1–H6 tag hierarchy, colour-coded by level, with ID attributes."),
        ("table",   "Tables",         "Parses every HTML table into a DataFrame, downloadable as CSV."),
        ("target",  "CSS Selector",   "Query any element using standard CSS selectors — div.price, nav a, etc."),
        ("mail",    "Emails",         "Regex scan of full page text to harvest all email addresses."),
        ("text",    "Paragraphs",     "All <p> body text with word count, keyword filter and highlight."),
        ("cpu",     "Assets",         "Inventory all JavaScript files and linked CSS stylesheets."),
    ]
    for i, (icon, name, desc) in enumerate(modes_desc):
        grid_cols[i % 3].markdown(
            f'<div class="mode-card">'
            f'<div class="mode-card-icon">{svg(icon, 20, "#6d28d9")}</div>'
            f'<div class="mode-card-name">{name}</div>'
            f'<div class="mode-card-desc">{desc}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
