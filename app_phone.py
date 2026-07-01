# -- coding: utf-8 --
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
from scipy.optimize import linprog
import base64
import uuid
import html
import json
import re
import math
import urllib.request
from html.parser import HTMLParser
from datetime import datetime
from pathlib import Path

# =========================
# إعدادات عامة
# =========================
st.set_page_config(
    page_title="تكوين العلائق العلفية لتغذية الماشية",
    page_icon="🌿",
    layout="centered",
)

# ── Brand tokens ────────────────────────────────────────────────────────────
NAVY   = "#2C3149"
PURPLE = "#6B4CA0"
GOLD   = "#F5C842"
GREEN  = "#4A8C3F"
WHITE  = "#FFFFFF"
LIGHT_BG = "#F7F6FB"

NO_PRICE_TEXT = "لا يوجد معلومات عن السعر"
PRICE_INDICATOR_URL = "https://sanabilaljouf.com.sa/ar/price-indicator"
FILE_PATH  = "Livestock_NRC_v2 _final.xlsx"
WEB_LOGO_PATH  = "white_logo_nlfdp.png"
PRINT_LOGO_PATH = "NLFDP_full.png"

SPECIES_SHEETS = {
    "ضأن - النعاج (الأمهات)":    "Sheep_Ewes",
    "ضأن - الفحول":               "Sheep_Rams",
    "ضأن - حملان التسمين":        "Sheep_Fattening_Lambs",
    "ماعز - (الأمهات)":           "Goat_Female",
    "ماعز - الفحول":              "Goat_Rams",
    "ماعز - جديان التسمين":       "Goat_Fattining",
    "إبل - النوق (الأمهات)":      "Camel_Female",
    "إبل - الفحول":               "Camel_Male",
    "إبل - حواشي تسمين":          "Camel_Fattining",
}
FEED_SHEET = "Feed_Bank"

ANIMAL_TYPES = {
    "ضأن": "🐑",
    "ماعز": "🐐",
    "إبل": "🐪",
}

# ── Feed classification (V5): plain vs. has additives ─────────────────────
# Same membership-list pattern as ROUGHAGE_FEEDS / CONCENTRATE_FEEDS below.
# Any feed name containing one of these brand substrings is treated as a
# commercial / fortified ("has additives") product. Everything else is
# "plain" (raw single-ingredient feed). Edit this list as your feed bank
# grows or if a brand should be reclassified.
ADDITIVE_BRAND_KEYWORDS = [
    "أراسكو",
    "بريمير",
    "زاد",
    "المراح الذهبي",
    "المتحدة",
]

DEFAULT_PACKAGE_KG = 50.0
BAG_ROUND_UP_THRESHOLD = 0.20

# =========================
# Global CSS
# =========================
def inject_css():
    st.markdown(
        f"""
        <style>
        /* ── Reset & base ─────────────────────────────────────────── */
        html, body, [class*="css"] {{
            font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
            direction: rtl;
        }}
        .main .block-container {{
            padding-top: 2rem;
            padding-bottom: 4rem;
            max-width: 1280px;
        }}
        div[data-testid="stExpander"] details summary {{
            display: flex;
            align-items: center;
            gap: 0.9rem;
        }}
        div[data-testid="stExpander"] details summary span[data-testid="stExpanderToggleIcon"] {{
            margin-inline: 0.55rem;
            flex-shrink: 0;
        }}
        div[data-testid="stExpander"] details summary p {{
            margin: 0;
        }}

        /* ── App header ────────────────────────────────────────────── */
        .app-hero {{
            background: {NAVY};
            border-radius: 18px;
            padding: 28px 32px 22px;
            margin-bottom: 28px;
            text-align: center;
            position: relative;
            overflow: hidden;
        }}
        .app-hero::before {{
            content: '';
            position: absolute;
            top: -40px; left: -40px;
            width: 180px; height: 180px;
            border-radius: 50%;
            background: {PURPLE};
            opacity: 0.18;
        }}
        .app-hero::after {{
            content: '';
            position: absolute;
            bottom: -50px; right: -30px;
            width: 200px; height: 200px;
            border-radius: 50%;
            background: {GOLD};
            opacity: 0.10;
        }}
        .app-hero img {{
            width: 200px;
            max-width: 80%;
            border-radius: 10px;
            margin-bottom: 14px;
        }}
        .app-hero-title {{
            color: {WHITE};
            font-size: 1.65rem;
            font-weight: 800;
            margin: 0 0 6px;
            position: relative;
            z-index: 1;
        }}
        .app-hero-subtitle {{
            color: rgba(255,255,255,0.60);
            font-size: 0.90rem;
            position: relative;
            z-index: 1;
        }}
        .gold-dot {{
            display: inline-block;
            width: 8px; height: 8px;
            border-radius: 50%;
            background: {GOLD};
            margin: 0 6px;
            vertical-align: middle;
        }}

        /* ── Section cards ─────────────────────────────────────────── */
        .section-card {{
            background: {WHITE};
            border: 1.5px solid #E8E4F0;
            border-radius: 16px;
            padding: 22px 24px 18px;
            margin-bottom: 20px;
            box-shadow: 0 2px 12px rgba(44,49,73,0.06);
        }}
        .section-title {{
            font-size: 1.05rem;
            font-weight: 800;
            color: {NAVY};
            margin-bottom: 16px;
            padding-bottom: 10px;
            border-bottom: 2px solid {PURPLE};
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .section-title-icon {{
            background: {PURPLE};
            color: {WHITE};
            width: 28px; height: 28px;
            border-radius: 8px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 0.85rem;
            flex-shrink: 0;
        }}

        /* ── Metric pills (top summary) ────────────────────────────── */
        .metric-strip {{
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
            margin-bottom: 20px;
        }}
        .metric-pill {{
            background: {NAVY};
            color: {WHITE};
            font-size: 0.78rem;
            font-weight: 700;
            padding: 6px 14px;
            border-radius: 999px;
        }}
        .metric-pill span {{
            color: {GOLD};
            margin-right: 4px;
        }}

        /* ── Feed chips (legacy, still used for selected-feed pills) ─── */
        .chip-badge {{
            display: inline-block;
            padding: 5px 14px;
            border-radius: 999px;
            margin: 4px 3px;
            font-size: 0.82rem;
            font-weight: 700;
            white-space: nowrap;
        }}
        .chip-roughage {{
            background: #E8F7ED;
            color: #166534;
            border: 1px solid #B7E4C4;
        }}
        .chip-concentrate {{
            background: #EAF2FF;
            color: #1D4ED8;
            border: 1px solid #BBD3FF;
        }}
        .chip-category-label {{
            display: inline-block;
            padding: 5px 14px;
            border-radius: 999px;
            margin: 6px 0 10px;
            font-size: 0.9rem;
            font-weight: 800;
        }}
        .selected-feeds-box {{
            border: 1px solid #E8E4F0;
            border-radius: 10px;
            padding: 12px 14px;
            background: {LIGHT_BG};
            margin-top: 10px;
            line-height: 2.4;
        }}

        /* ── V5: feed selection + pricing, merged section ────────────── */
        .feed-type-legend {{
            display: flex;
            gap: 16px;
            margin-bottom: 14px;
            font-size: 0.78rem;
            color: #555;
        }}
        .legend-dot {{
            display: inline-block;
            width: 9px; height: 9px;
            border-radius: 50%;
            margin-left: 6px;
            vertical-align: -1px;
        }}
        .legend-dot.plain    {{ background: #166534; }}
        .legend-dot.additive {{ background: #1D4ED8; }}

        .feed-chip-row {{
            display: flex;
            align-items: stretch;
            border: 1.5px solid #E0E0E0;
            border-radius: 10px;
            overflow: hidden;
            margin-bottom: 6px;
            background: {WHITE};
        }}
        .feed-chip-row.plain    {{ border-color: #B7E4C4; }}
        .feed-chip-row.additive {{ border-color: #BBD3FF; }}
        .feed-chip-row.feed-selected.plain {{
            border-color: #166534;
            box-shadow: 0 0 0 1px #166534 inset;
        }}
        .feed-chip-row.feed-selected.additive {{
            border-color: #1D4ED8;
            box-shadow: 0 0 0 1px #1D4ED8 inset;
        }}
        .chip-price-tag {{
            flex-shrink: 0;
            padding: 6px 10px;
            font-size: 0.72rem;
            font-weight: 700;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            min-width: 86px;
            text-align: center;
            border-left: 1px solid rgba(0,0,0,0.06);
        }}
        .chip-price-tag.plain    {{ background: #E8F7ED; color: #166534; }}
        .chip-price-tag.additive {{ background: #EAF2FF; color: #1D4ED8; }}
        .chip-price-tag.no-price {{ background: #FFF8E6; color: #92400E; font-style: italic; }}
        .chip-price-unit {{ font-size: 0.6rem; opacity: 0.75; font-weight: 500; }}

        .price-panel-row {{
            border-bottom: 0.5px solid #eee;
            padding: 10px 4px;
        }}
        .price-panel-row:last-child {{ border-bottom: none; }}
        .price-panel-row.highlight {{
            background: #FAF6FF;
            border-radius: 10px;
            padding: 10px 12px;
            border: 1px solid #D9CCF0;
        }}
        .price-panel-feed-name {{
            font-size: 0.85rem;
            font-weight: 800;
        }}
        .price-panel-feed-name.plain    {{ color: #166534; }}
        .price-panel-feed-name.additive {{ color: #1D4ED8; }}
        .missing-price-flag {{
            font-size: 0.66rem;
            color: #92400E;
            background: #FFF8E6;
            padding: 2px 9px;
            border-radius: 999px;
            margin-right: 8px;
        }}
        .price-panel-hint {{
            font-size: 0.70rem;
            color: #999;
            margin: 2px 0 0;
        }}

        .feed-price-pill {{
            display: inline-flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 8px;
            padding: 6px 13px;
            border-radius: 999px;
            margin: 2px 0 8px;
            font-size: 0.76rem;
            font-weight: 800;
            border: 1px solid #D9CCF0;
            background: #F7F3FC;
            color: {NAVY};
        }}
        .feed-price-pill.normal {{
            background: #E8F7ED;
            color: #166534;
            border-color: #B7E4C4;
        }}
        .feed-price-pill.concentrate {{
            background: #EAF2FF;
            color: #1D4ED8;
            border-color: #BBD3FF;
        }}
        .feed-price-pill.no-price {{
            background: #FFF8E6;
            color: #92400E;
            border-color: #F5C842;
        }}
        .feed-price-pill-value {{
            color: {PURPLE};
            font-size: 0.70rem;
            font-weight: 700;
        }}
        .feed-price-pill.no-price .feed-price-pill-value {{
            color: #92400E;
        }}
        .price-review-box {{
            margin-top: 16px;
            padding: 14px;
            border: 1.5px solid #D9CCF0;
            border-radius: 12px;
            background: #FAF8FD;
        }}
        .price-review-title {{
            color: {NAVY};
            font-size: 0.82rem;
            font-weight: 800;
            margin-bottom: 8px;
        }}
        .price-review-pills {{
            display: flex;
            flex-wrap: wrap;
            gap: 7px;
        }}
        .price-review-pill {{
            display: inline-block;
            padding: 5px 11px;
            border-radius: 999px;
            background: #F3EEF9;
            color: {PURPLE};
            border: 1px solid #D9CCF0;
            font-size: 0.70rem;
            font-weight: 700;
        }}
        .price-review-pill.missing {{
            background: #FFF8E6;
            color: #92400E;
            border-color: #F5C842;
        }}

        .feed-category-title {{
            display: inline-block;
            padding: 5px 13px;
            border-radius: 999px;
            margin: 10px 0 8px;
            font-size: 0.78rem;
            font-weight: 800;
        }}
        .feed-category-title.normal {{
            background: #E8F7ED;
            color: #166534;
            border: 1px solid #B7E4C4;
        }}
        .feed-category-title.concentrate {{
            background: #EAF2FF;
            color: #1D4ED8;
            border: 1px solid #BBD3FF;
        }}
        div[class*="st-key-feed_choice_normal_"] button {{
            border-color: #B7E4C4 !important;
            color: #166534 !important;
            background: #FFFFFF !important;
        }}
        div[class*="st-key-feed_choice_selected_normal_"] button {{
            border-color: #7CCB96 !important;
            color: #166534 !important;
            background: #E8F7ED !important;
            box-shadow: 0 0 0 1px #7CCB96 inset !important;
        }}
        div[class*="st-key-feed_choice_concentrate_"] button {{
            border-color: #BBD3FF !important;
            color: #1D4ED8 !important;
            background: #FFFFFF !important;
        }}
        div[class*="st-key-feed_choice_selected_concentrate_"] button {{
            border-color: #7DAAFF !important;
            color: #1D4ED8 !important;
            background: #EAF2FF !important;
            box-shadow: 0 0 0 1px #7DAAFF inset !important;
        }}
        div[class*="st-key-price_summary_normal_"] button {{
            min-height: 32px !important;
            padding: 4px 9px !important;
            border-color: #B7E4C4 !important;
            background: #E8F7ED !important;
            color: #166534 !important;
            font-size: 0.68rem !important;
        }}
        div[class*="st-key-price_summary_concentrate_"] button {{
            min-height: 32px !important;
            padding: 4px 9px !important;
            border-color: #BBD3FF !important;
            background: #EAF2FF !important;
            color: #1D4ED8 !important;
            font-size: 0.68rem !important;
        }}
        div[class*="st-key-price_summary_"] button p {{
            white-space: pre-line !important;
            line-height: 1.35 !important;
            margin: 0 !important;
        }}
        .compact-price-summary-title {{
            color: {NAVY};
            font-size: 0.78rem;
            font-weight: 800;
            margin: 12px 0 6px;
        }}

        /* ── V5 checkout-style feed pricing ───────────────────────── */
        .checkout-feed-heading {{
            display: inline-block;
            padding: 5px 13px;
            border-radius: 999px;
            margin: 6px 0 8px;
            font-size: 0.78rem;
            font-weight: 800;
        }}
        .checkout-feed-heading.plain {{
            background: #E8F7ED;
            color: #166534;
            border: 1px solid #B7E4C4;
        }}
        .checkout-feed-heading.additive {{
            background: #EAF2FF;
            color: #1D4ED8;
            border: 1px solid #BBD3FF;
        }}
        .checkout-feed-heading.raw {{
            background: #FFF7E6;
            color: #92400E;
            border: 1px solid #F7D58A;
        }}
        .feed-pill-card {{
            border: 1px solid #E5E7EB;
            border-radius: 16px;
            overflow: hidden;
            background: #FFFFFF;
            margin-bottom: 8px;
            min-height: 118px;
            box-shadow: 0 2px 10px rgba(44,49,73,0.05);
            position: relative;
        }}
        .feed-pill-card.selected {{
            box-shadow: 0 0 0 2px {PURPLE} inset, 0 5px 18px rgba(107,76,160,0.14);
        }}
        .feed-pill-header {{
            min-height: 58px;
            padding: 9px 12px 8px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
            gap: 2px;
            position: relative;
        }}
        .feed-pill-card.plain .feed-pill-header {{ background: #E8F7ED; color: #166534; }}
        .feed-pill-card.raw .feed-pill-header {{ background: #FFF7E6; color: #92400E; }}
        .feed-pill-card.additive .feed-pill-header {{ background: #EAF2FF; color: #1D4ED8; }}
        .feed-pill-card.other .feed-pill-header {{ background: #F9FAFB; color: #374151; }}
        .feed-pill-name {{ font-size: 0.82rem; font-weight: 900; line-height: 1.25; }}
        .feed-pill-price {{ font-size: 0.66rem; font-weight: 800; opacity: 0.86; line-height: 1.25; }}
        .feed-pill-check {{
            position: absolute;
            top: 6px;
            inset-inline-start: 8px;
            width: 19px;
            height: 19px;
            border-radius: 999px;
            background: {PURPLE};
            color: #FFFFFF;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 0.68rem;
            font-weight: 900;
        }}
        .feed-pill-macros {{
            padding: 8px 8px 9px;
            display: flex;
            justify-content: center;
            align-items: center;
            flex-wrap: wrap;
            gap: 5px;
        }}
        .feed-macro-tag {{
            display: inline-block;
            border-radius: 999px;
            padding: 3px 8px;
            font-size: 0.58rem;
            line-height: 1.2;
            font-weight: 900;
            white-space: nowrap;
        }}
        .feed-macro-protein {{ background: #EEEDFE; color: #3C3489; }}
        .feed-macro-fiber {{ background: #FAECE7; color: #712B13; }}
        .feed-macro-energy {{ background: #FBEAF0; color: #72243E; }}
        div[class*="st-key-checkout_feed_plain_"],
        div[class*="st-key-checkout_feed_selected_plain_"],
        div[class*="st-key-checkout_feed_raw_"],
        div[class*="st-key-checkout_feed_selected_raw_"],
        div[class*="st-key-checkout_feed_additive_"],
        div[class*="st-key-checkout_feed_selected_additive_"] {{
            margin-top: -126px;
            height: 126px;
            position: relative;
            z-index: 3;
        }}
        div[class*="st-key-checkout_feed_plain_"] button,
        div[class*="st-key-checkout_feed_selected_plain_"] button,
        div[class*="st-key-checkout_feed_raw_"] button,
        div[class*="st-key-checkout_feed_selected_raw_"] button,
        div[class*="st-key-checkout_feed_additive_"] button,
        div[class*="st-key-checkout_feed_selected_additive_"] button {{
            height: 118px !important;
            min-height: 118px !important;
            width: 100% !important;
            opacity: 0 !important;
            border: 0 !important;
            background: transparent !important;
            color: transparent !important;
            padding: 0 !important;
            cursor: pointer !important;
        }}
        .price-panel-section {{
            padding: 10px 0;
        }}
        .price-panel-heading {{
            color: {NAVY};
            font-size: 0.90rem;
            font-weight: 800;
            margin-bottom: 2px;
        }}
        .price-panel-caption {{
            color: #888;
            font-size: 0.68rem;
            line-height: 1.45;
            margin-bottom: 10px;
        }}
        .price-panel-available-box {{
            border: 1px solid #E5E7EB;
            border-radius: 12px;
            background: #F9FAFB;
            padding: 10px 12px;
            margin-bottom: 12px;
        }}
        .price-panel-available-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 10px;
            padding: 6px 0;
            border-bottom: 1px dashed #E5E7EB;
            font-size: 0.76rem;
        }}
        .price-panel-available-row:last-child {{
            border-bottom: none;
            padding-bottom: 0;
        }}
        .price-panel-available-name {{
            font-weight: 800;
            color: #2C3149;
        }}
        .price-panel-available-value {{
            font-weight: 700;
            color: #166534;
            white-space: nowrap;
        }}
        .price-panel-feed-label {{
            font-size: 0.72rem;
            font-weight: 800;
            margin: 7px 0 3px;
        }}
        .price-panel-missing-row {{
            color: #92400E;
        }}
        .price-panel-priced-row {{
            color: {PURPLE};
            align-self: center;
        }}
        div[data-testid="stHorizontalBlock"] div[data-testid="stNumberInput"] label p {{
            font-size: 0.66rem !important;
        }}

        /* ── Recommendation cards ──────────────────────────────────── */
        .rec-grid-header {{
            font-size: 1.0rem;
            font-weight: 800;
            color: {NAVY};
            margin: 24px 0 12px;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        .rec-grid-header::after {{
            content: '';
            flex: 1;
            height: 1.5px;
            background: linear-gradient(to left, transparent, {PURPLE}44);
        }}

        /* ── Status badge (market price) ───────────────────────────── */
        .status-ok-badge {{
            display: inline-block;
            background: #E8F7ED;
            color: #166534;
            border: 1px solid #B7E4C4;
            border-radius: 999px;
            font-size: 0.72rem;
            font-weight: 700;
            padding: 3px 10px;
            margin-bottom: 10px;
        }}

        /* ── Streamlit widget overrides ────────────────────────────── */
        div[role="radiogroup"] {{
            gap: 0.35rem;
        }}
        div[role="radiogroup"] > label {{
            min-height: 30px;
            padding: 4px 8px;
            border: 1.5px solid #D1D5DB;
            border-radius: 999px;
            background: {WHITE};
            font-size: 0.66rem;
            line-height: 1.1;
            transition: all 0.15s;
        }}
        div[role="radiogroup"] > label p {{
            font-size: 0.66rem !important;
        }}
        div[role="radiogroup"] > label:has(input:checked) {{
            background: #F3E8FF;
            border-color: {PURPLE};
            color: #5B21B6;
            font-weight: 700;
        }}
        div[role="radiogroup"] > label input[type="radio"] {{
            display: block;
            width: 0.82rem;
            height: 0.82rem;
            margin: 0 0.35rem 0 0;
            accent-color: {PURPLE};
            flex-shrink: 0;
        }}
        div[data-testid="stCheckbox"] input[type="checkbox"] {{
            accent-color: #2563EB !important;
            width: 0.95rem;
            height: 0.95rem;
        }}
        div[data-testid="stCheckbox"] label {{
            gap: 0.30rem;
        }}
        div[data-testid="stCheckbox"] label p {{
            margin-right: 0.18rem !important;
        }}
        div[data-testid="stExpander"] details summary p {{
            margin-right: 0.22rem !important;
        }}

        .stButton > button {{
            border-radius: 999px !important;
            border: 1.5px solid #D1D5DB !important;
            font-size: 0.82rem !important;
            padding: 5px 12px !important;
            transition: all 0.15s !important;
        }}
        .stButton > button:hover {{
            border-color: {PURPLE} !important;
            color: {PURPLE} !important;
            background: #F3E8FF !important;
        }}

        /* primary CTA */
        .cta-btn > button {{
            background: {NAVY} !important;
            color: {WHITE} !important;
            border: none !important;
            border-radius: 12px !important;
            font-size: 1.0rem !important;
            font-weight: 800 !important;
            padding: 12px 0 !important;
            letter-spacing: 0.02em;
            transition: background 0.2s !important;
        }}
        .cta-btn > button:hover {{
            background: {PURPLE} !important;
        }}

        /* Footer note */
        .app-footer-note {{
            text-align: center;
            font-size: 0.78rem;
            color: #AAA;
            margin-top: 32px;
            padding-top: 16px;
            border-top: 1px solid #EEE;
        }}

        @media (max-width: 760px) {{
            .main .block-container {{
                padding-top: 0.8rem;
                padding-bottom: 2rem;
                padding-left: 0.75rem;
                padding-right: 0.75rem;
                max-width: 100%;
            }}
            .main .block-container,
            .section-card,
            .selected-feeds-box,
            .price-panel-available-box,
            .price-panel-section,
            .rec-card-wide,
            .rec-grid-header {{
                direction: rtl;
                text-align: right;
            }}
            .app-hero {{
                padding: 18px 16px 16px;
                margin-bottom: 18px;
                border-radius: 14px;
            }}
            .app-hero img {{
                width: 148px;
                max-width: 70%;
                margin-bottom: 10px;
            }}
            .app-hero-title {{
                font-size: 1.15rem;
                line-height: 1.35;
            }}
            .app-hero-subtitle {{
                font-size: 0.76rem;
                line-height: 1.55;
            }}
            .section-card {{
                padding: 16px 14px 14px;
                border-radius: 14px;
                margin-bottom: 14px;
            }}
            .section-title {{
                font-size: 0.92rem;
                margin-bottom: 12px;
                padding-bottom: 8px;
            }}
            .section-title-icon {{
                width: 24px;
                height: 24px;
                font-size: 0.75rem;
            }}
            .metric-strip {{
                gap: 8px;
                margin-bottom: 14px;
                justify-content: flex-start;
            }}
            .metric-pill {{
                font-size: 0.70rem;
                padding: 6px 10px;
            }}
            .selected-feeds-box {{
                line-height: 2;
            }}
            .chip-badge {{
                direction: rtl;
            }}
            .checkout-feed-heading {{
                display: table;
                font-size: 0.72rem;
                padding: 4px 10px;
                margin-right: 0;
                margin-left: auto;
                text-align: right;
            }}
            .feed-pill-card {{
                min-height: 104px;
                border-radius: 14px;
            }}
            .feed-pill-header {{
                min-height: 54px;
                padding: 8px 10px 7px;
            }}
            .feed-pill-name {{
                font-size: 0.76rem;
            }}
            .feed-pill-price {{
                font-size: 0.62rem;
            }}
            .feed-pill-macros {{
                padding: 7px 7px 8px;
                gap: 4px;
            }}
            .feed-macro-tag {{
                font-size: 0.54rem;
                padding: 3px 7px;
            }}
            div[class*="st-key-checkout_feed_plain_"],
            div[class*="st-key-checkout_feed_selected_plain_"],
            div[class*="st-key-checkout_feed_raw_"],
            div[class*="st-key-checkout_feed_selected_raw_"],
            div[class*="st-key-checkout_feed_additive_"],
            div[class*="st-key-checkout_feed_selected_additive_"] {{
                margin-top: -112px;
                height: 112px;
            }}
            div[class*="st-key-checkout_feed_plain_"] button,
            div[class*="st-key-checkout_feed_selected_plain_"] button,
            div[class*="st-key-checkout_feed_raw_"] button,
            div[class*="st-key-checkout_feed_selected_raw_"] button,
            div[class*="st-key-checkout_feed_additive_"] button,
            div[class*="st-key-checkout_feed_selected_additive_"] button {{
                height: 104px !important;
                min-height: 104px !important;
            }}
            .price-panel-heading {{
                font-size: 0.84rem;
            }}
            .price-panel-caption,
            .price-panel-feed-label,
            .price-panel-available-row {{
                font-size: 0.68rem;
            }}
            .price-panel-available-row {{
                align-items: flex-start;
                flex-direction: column;
                gap: 3px;
            }}
            .price-panel-available-value {{
                white-space: normal;
            }}
            .rec-grid-header {{
                font-size: 0.88rem;
                margin: 18px 0 10px;
                gap: 8px;
            }}
            .rec-card-top,
            .rec-gauges-col,
            .rec-feed-col,
            .rec-card-footer-wide {{
                padding-left: 14px;
                padding-right: 14px;
            }}
            .rec-price-block {{
                gap: 10px;
            }}
            .rec-price-value {{
                font-size: 0.92rem;
            }}
            .rec-col-title {{
                text-align: right;
            }}
            .gauge-row {{
                margin-bottom: 28px;
            }}
            .gauge-label-row {{
                display: block;
                margin-bottom: 12px;
                min-height: 72px;
            }}
            .gauge-name,
            .gauge-value {{
                display: block;
                width: 100%;
                white-space: normal;
            }}
            .gauge-name {{
                font-size: 0.92rem;
                line-height: 1.35;
                margin-bottom: 6px;
            }}
            .gauge-name small {{
                display: block;
                margin-right: 0;
                margin-top: 2px;
                font-size: 0.72rem;
            }}
            .gauge-value {{
                font-size: 0.88rem;
                line-height: 1.35;
            }}
            .gauge-marker {{
                top: 86px;
            }}
            .gauge-marker-line {{
                top: 94px;
                height: 8px;
            }}
            .gauge-tooltip {{
                display: none;
            }}
            .gauge-scale-labels {{
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                gap: 4px;
                margin-top: 8px;
                font-size: 0.44rem;
                direction: ltr;
            }}
            .gauge-scale-labels span {{
                width: auto !important;
                white-space: normal;
                line-height: 1.2;
                gap: 2px;
            }}
            .gauge-scale-labels small {{
                font-size: 0.40rem;
            }}
            .rec-feed-table-wide {{
                min-width: 0;
                font-size: 0.60rem;
            }}
            .rec-feed-table-wide th {{
                font-size: 0.50rem;
            }}
            .rec-feed-table-wide td,
            .rec-feed-table-wide th {{
                padding-left: 2px;
                padding-right: 2px;
                white-space: normal;
                word-break: break-word;
            }}
            .rec-unit-badge {{
                white-space: normal;
                line-height: 1.3;
            }}
            .rec-footer-label,
            .rec-footer-value {{
                font-size: 0.68rem;
            }}
            div[data-testid="stHorizontalBlock"] {{
                flex-direction: column !important;
                gap: 0.7rem !important;
            }}
            div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {{
                width: 100% !important;
                min-width: 0 !important;
                flex: 1 1 100% !important;
            }}
            div[data-testid="stHorizontalBlock"] div[data-testid="stNumberInput"],
            div[data-testid="stHorizontalBlock"] div[data-testid="stSelectbox"],
            div[data-testid="stHorizontalBlock"] div[data-testid="stMultiSelect"] {{
                width: 100% !important;
            }}
            .stButton > button {{
                min-height: 42px;
                font-size: 0.88rem !important;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

inject_css()

# =========================
# مساعدات UI – الشعار والهيدر
# =========================
def load_logo_b64(path: str):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return None
web_logo_b64 = load_logo_b64(WEB_LOGO_PATH)
print_logo_b64 = load_logo_b64(PRINT_LOGO_PATH)
if web_logo_b64:
    logo_tag = f'<img src="data:image/png;base64,{web_logo_b64}" />'
else:
    logo_tag = ""

st.markdown(
    f"""
    <div class="app-hero">
        {logo_tag}
        <div class="app-hero-title">حاسبة الاعلاف لقياس تكلفة العليقة و كميات القطيع </div>
        <div class="app-hero-subtitle">
            ضأن <span class="gold-dot"></span>
            ماعز <span class="gold-dot"></span>
            إبل — تحسين تكلفة العليقة مع حساب كميات القطيع
        </div>
    </div>
    <div class="app-footer-note"> تعتمد التوصيات اولا على تحقيق الاحتياجات الغذائية، و قد يختلف السعر المعروض عن السعر الفعلي </div>
    """,
    unsafe_allow_html=True,
)

# ── helper: open / close a section card ────────────────────────────────────
def section_open(icon: str, title: str):
    st.markdown(
        f"""<div class="section-card">
            <div class="section-title">
                <span class="section-title-icon">{icon}</span> {title}
            </div>""",
        unsafe_allow_html=True,
    )

def section_close():
    st.markdown("</div>", unsafe_allow_html=True)


# =========================
# دوال البيانات  (لا تغيير في المنطق)
# =========================
@st.cache_data
def load_all(path):
    xls = pd.ExcelFile(path)
    sheets = {}
    for name in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=name, header=None)
        if len(df) > 2:
            df.columns = df.iloc[1].astype(str)
            df = df.iloc[2:].reset_index(drop=True)
        sheets[name] = df
    return sheets


def find_col(df, keywords):
    for col in df.columns:
        c = str(col).strip()
        for kw in keywords:
            if kw in c:
                return col
    return None


def to_num(x):
    try:
        v = float(str(x).replace(",", "."))
        if np.isfinite(v):
            return v
    except Exception:
        pass
    return float("nan")


def format_display_num(value, digits=2, use_commas=False):
    try:
        v = float(value)
        if not np.isfinite(v):
            return str(value)
        factor = 10 ** digits
        truncated = np.trunc(v * factor) / factor
        spec = f",.{digits}f" if use_commas else f".{digits}f"
        return format(truncated, spec)
    except Exception:
        return str(value)


def normalize_feed_name(value):
    text = str(value or "").strip().lower()
    for old, new in {"أ": "ا","إ": "ا","آ": "ا","ى": "ي","ة": "ه","ـ": ""}.items():
        text = text.replace(old, new)
    text = re.sub(r"[^\w\s%]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def extract_number(value):
    match = re.search(r"\d+(?:[.,]\d+)?", str(value or ""))
    if not match:
        return float("nan")
    return float(match.group(0).replace(",", "."))


class PriceTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.tables = []
        self._table_depth = 0
        self._rows = []
        self._row = None
        self._cell = None

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self._table_depth += 1
            if self._table_depth == 1:
                self._rows = []
        elif self._table_depth and tag == "tr":
            self._row = []
        elif self._table_depth and tag in {"td", "th"}:
            self._cell = []

    def handle_data(self, data):
        if self._cell is not None:
            data = data.strip()
            if data:
                self._cell.append(data)

    def handle_endtag(self, tag):
        if self._table_depth and tag in {"td", "th"} and self._cell is not None:
            if self._row is not None:
                self._row.append(" ".join(self._cell).strip())
            self._cell = None
        elif self._table_depth and tag == "tr" and self._row is not None:
            if any(self._row):
                self._rows.append(self._row)
            self._row = None
        elif tag == "table" and self._table_depth:
            self._table_depth -= 1
            if self._table_depth == 0 and self._rows:
                self.tables.append(self._rows)
                self._rows = []


def parse_price_tables(html_text):
    parser = PriceTableParser()
    parser.feed(html_text)
    return parser.tables


@st.cache_data(ttl=60 * 60)
def load_market_prices(url):
    try:
        request = urllib.request.Request(
            url, headers={"User-Agent": "Mozilla/5.0 (compatible; feed-calculator/1.0)"}
        )
        with urllib.request.urlopen(request, timeout=20) as response:
            charset = response.headers.get_content_charset() or "utf-8"
            html_text = response.read().decode(charset, errors="replace")
    except Exception as e:
        return {}, f"تعذر تحميل أسعار السوق من الرابط: {e}"

    price_rows = None
    for rows in parse_price_tables(html_text):
        header_index = None
        for i, row in enumerate(rows):
            joined = " ".join(row)
            if "المنتج" in joined and "السعر" in joined:
                header_index = i
                break
        if header_index is not None:
            price_rows = rows[header_index:]
            break

    if not price_rows:
        return {}, "لم يتم العثور على جدول أسعار يحتوي على المنتج والسعر."

    headers    = price_rows[0]
    product_idx = next((i for i, h in enumerate(headers) if "المنتج" in h), None)
    price_idx   = next((i for i, h in enumerate(headers) if "السعر"  in h), None)
    weight_idx  = next((i for i, h in enumerate(headers) if "الوزن"  in h), None)
    brand_idx   = next((i for i, h in enumerate(headers) if "العلامة" in h), None)

    if product_idx is None or price_idx is None:
        return {}, "لم يتم العثور على أعمدة المنتج والسعر في جدول الأسعار."

    prices = {}
    for row in price_rows[1:]:
        if len(row) <= max(product_idx, price_idx):
            continue
        product = str(row[product_idx]).strip()
        if not product or product.lower() == "nan":
            continue
        package_price  = extract_number(row[price_idx])
        package_weight = extract_number(row[weight_idx]) if weight_idx is not None and len(row) > weight_idx else float("nan")
        if np.isnan(package_price):
            continue
        price_per_kg = package_price / package_weight if package_weight and not np.isnan(package_weight) else package_price
        brand = str(row[brand_idx]).strip() if brand_idx is not None and len(row) > brand_idx else ""
        prices[normalize_feed_name(product)] = {
            "product":        product,
            "brand":          "" if brand.lower() == "nan" else brand,
            "package_price":  float(package_price),
            "package_weight": float(package_weight) if not np.isnan(package_weight) else np.nan,
            "price_per_kg":   float(price_per_kg),
        }

    return prices, None


def find_market_price(feed_name, market_prices):
    feed_key = normalize_feed_name(feed_name)
    if feed_key in market_prices:
        return market_prices[feed_key]
    candidates = []
    feed_tokens = set(feed_key.split())
    for product_key, payload in market_prices.items():
        if not product_key:
            continue
        product_tokens = set(product_key.split())
        shared = feed_tokens & product_tokens
        if feed_key in product_key or product_key in feed_key:
            score = 100 + len(shared)
        elif shared:
            score = len(shared) / max(len(feed_tokens), len(product_tokens))
        else:
            score = 0
        if score:
            candidates.append((score, payload))
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]


def get_req_columns(req_df):
    REQ_STAGE = find_col(req_df, ["الفترة", "الحالة", "الإنتاجية", "stage"])
    REQ_CP    = find_col(req_df, ["بروتين"])
    REQ_EN    = find_col(req_df, ["طاقة"])
    REQ_FI    = find_col(req_df, ["ألياف"])
    REQ_WT    = find_col(req_df, ["وزن"])
    REQ_DM    = find_col(req_df, ["المادة الجافة", "جافة مأكولة", "Dry", "DM"])
    if REQ_DM is None:
        try:
            if req_df.shape[1] > 62:
                REQ_DM = req_df.columns[62]
            elif req_df.shape[1] > 61:
                REQ_DM = req_df.columns[61]
        except Exception:
            REQ_DM = None
    return REQ_STAGE, REQ_CP, REQ_EN, REQ_FI, REQ_WT, REQ_DM


def get_feed_columns(feeds_df):
    FEED_NAME = find_col(feeds_df, ["نوع العلف", "اسم العلف", "العلف"])
    FEED_CP   = find_col(feeds_df, ["بروتين"])
    FEED_EN   = find_col(feeds_df, ["طاقة"])
    FEED_FI   = find_col(feeds_df, ["ألياف"])
    return FEED_NAME, FEED_CP, FEED_EN, FEED_FI


def build_stage_options(req_df, REQ_STAGE):
    return (req_df[REQ_STAGE].dropna().astype(str).str.strip()
            .replace("", np.nan).dropna().unique().tolist())


def build_weight_options(stage_rows, REQ_WT):
    if REQ_WT is None:
        return []
    return (stage_rows[REQ_WT].dropna().astype(str).str.strip()
            .replace("", np.nan).dropna().unique().tolist())


def solve_ration(feeds_df, FEED_NAME, FEED_CP, FEED_EN, FEED_FI,
                 selected_feeds, feed_prices, feed_mins, feed_maxs,
                 cp_req, en_req, fi_req):
    sub = feeds_df[feeds_df[FEED_NAME].astype(str).str.strip().isin(selected_feeds)].copy()
    if sub.empty:
        return None, "❌ لم يتم العثور على الأعلاف المختارة في بنك الأعلاف."

    cp_arr = pd.to_numeric(sub[FEED_CP], errors="coerce").to_numpy()
    en_arr = pd.to_numeric(sub[FEED_EN], errors="coerce").to_numpy()
    fi_arr = pd.to_numeric(sub[FEED_FI], errors="coerce").to_numpy()
    names  = sub[FEED_NAME].astype(str).str.strip().tolist()
    p_arr  = np.array([feed_prices.get(n, np.nan) for n in names], dtype=float)

    mask = ~(np.isnan(cp_arr) | np.isnan(en_arr) | np.isnan(fi_arr) | np.isnan(p_arr))
    cp_arr, en_arr, fi_arr, p_arr = cp_arr[mask], en_arr[mask], fi_arr[mask], p_arr[mask]
    names = [n for i, n in enumerate(names) if mask[i]]
    if len(names) == 0:
        return None, "❌ لا توجد بيانات رقمية كافية بعد التنظيف."

    lb, ub = [], []
    for n in names:
        mn = max(0.0, min(1.0, float(feed_mins.get(n, 0.0)) / 100.0))
        mx = max(0.0, min(1.0, float(feed_maxs.get(n, 100.0)) / 100.0))
        if mx < mn: mx = mn
        lb.append(mn); ub.append(mx)

    if sum(lb) > 1.0 + 1e-9:
        return None, "❌ مجموع الحدود الدنيا Min% أكبر من 100%."
    if sum(ub) < 1.0 - 1e-9:
        return None, "❌ مجموع الحدود القصوى Max% أقل من 100%."

    A_ub = np.vstack([-cp_arr, -en_arr, -fi_arr])
    b_ub = np.array([-cp_req, -en_req, -fi_req], dtype=float)
    A_eq = np.ones((1, len(names)))
    b_eq = np.array([1.0])
    bounds = [(lb[i], ub[i]) for i in range(len(names))]

    res = linprog(p_arr, A_ub=A_ub, b_ub=b_ub, A_eq=A_eq, b_eq=b_eq,
                  bounds=bounds, method="highs")
    if not res.success:
        en_s = en_arr / max(1000.0, float(np.nanmax(en_arr)))
        en_rs = en_req / max(1000.0, float(np.nanmax(en_arr)))
        res = linprog(p_arr,
                      A_ub=np.vstack([-cp_arr, -en_s, -fi_arr]),
                      b_ub=np.array([-cp_req, -en_rs, -fi_req]),
                      A_eq=A_eq, b_eq=b_eq, bounds=bounds, method="highs")
    if not res.success:
        return None, "❌ لم يتم العثور على توليفة تحقق الشروط الدنيا. جرّب تعديل الحدود أو إضافة أعلاف."

    x = res.x
    return {
        "names":           names,
        "x":               x,
        "prices":          p_arr,
        "cp_ach":          float(np.sum(cp_arr * x)),
        "en_ach":          float(np.sum(en_arr * x)),
        "fi_ach":          float(np.sum(fi_arr * x)),
        "cost_per_kg_mix": float(np.sum(p_arr * x)),
    }, None


def coverage_judge(pct: float):
    if pct < 95:   return "🔴 أقل من المطلوب"
    if pct <= 105: return "✅ مناسب"
    if pct < 121: return "مرتفع قليلاً ⚠️"
    return "❗ مرتفع جداً"


def make_analysis_table(cp_req, en_req, fi_req, cp_ach, en_ach, fi_ach):
    df = pd.DataFrame({
        "العنصر الغذائي": ["البروتين %", "الطاقة (kcal/kg)", "الألياف %"],
        "الاحتياج":        [float(cp_req), float(en_req), float(fi_req)],
        "المتحقق من الخلطة": [float(cp_ach), float(en_ach), float(fi_ach)],
    })
    df["نسبة التغطية %"] = df["المتحقق من الخلطة"] / df["الاحتياج"] * 100
    df["التقييم"]        = df["نسبة التغطية %"].apply(coverage_judge)
    return df


# ── V5: feed classification helper ──────────────────────────────────────────
def feed_has_additives(feed_name: str) -> bool:
    """True if the feed name matches a known commercial/additive brand."""
    name = str(feed_name)
    return any(kw in name for kw in ADDITIVE_BRAND_KEYWORDS)


# =========================
# Session State
# =========================
for key, default in [
    ("herd_groups",          []),
    ("selected_animal_types", []),
    ("selected_feed_chips",  []),
    ("editable_feed_prices", {}),       # {feed: price_per_bag (SAR)}
    ("editable_feed_bag_sizes", {}),    # {feed: bag size (kg)}
    ("price_source_choice",  {}),       # {feed: "market" | "custom"}
    ("feed_highlight",       None),     # feed name to visually highlight in panel
    ("last_group_payloads",  []),
    ("last_recommendation_days", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


def blank_group(animal_type=None):
    return {
        "id": str(uuid.uuid4())[:8],
        "animal_type": animal_type,
        "species": None,
        "stage": None,
        "weight": None,
        "n": 1,
    }

def add_group(group):    st.session_state.herd_groups.append(group)
def remove_group(gid):   st.session_state.herd_groups = [g for g in st.session_state.herd_groups if g["id"] != gid]

if not st.session_state.herd_groups:
    st.session_state.herd_groups = [blank_group()]

# =========================
# تحميل البيانات
# =========================
with st.spinner("جاري تحميل بيانات الأعلاف والاحتياجات…"):
    try:
        sheets = load_all(FILE_PATH)
    except Exception as e:
        st.error(f"خطأ في تحميل الملف: {e}")
        st.stop()

if FEED_SHEET not in sheets:
    st.error("شيت الأعلاف (Feed_Bank) غير موجود في الملف.")
    st.stop()

feeds_df = sheets[FEED_SHEET].copy()
FEED_NAME, FEED_CP, FEED_EN, FEED_FI = get_feed_columns(feeds_df)

if not all([FEED_NAME, FEED_CP, FEED_EN, FEED_FI]):
    st.error("❌ أعمدة مفقودة في بنك الأعلاف (اسم العلف / بروتين / طاقة / ألياف)")
    st.stop()

feed_names = (feeds_df[FEED_NAME].dropna().astype(str).str.strip()
              .replace("", np.nan).dropna().unique().tolist())

market_prices, market_price_error = load_market_prices(PRICE_INDICATOR_URL)

def get_feed_market_record(feed_name):
    return find_market_price(feed_name, market_prices) if market_prices else None

def build_auto_feed_prices(feeds):
    prices, records, missing = {}, {}, []
    for feed in feeds:
        record = get_feed_market_record(feed)
        if record and record.get("price_per_kg", 0) > 0:
            prices[feed] = float(record["price_per_kg"])
            records[feed] = record
        else:
            missing.append(feed)
    return prices, records, missing

def get_requirement_row(species, stage, weight):
    sheet_name = SPECIES_SHEETS[species]
    req_df = sheets[sheet_name].copy()
    REQ_STAGE, REQ_CP, REQ_EN, REQ_FI, REQ_WT, REQ_DM = get_req_columns(req_df)
    if not all([REQ_STAGE, REQ_CP, REQ_EN, REQ_FI]):
        return None, "❌ أعمدة مفقودة في جدول الاحتياجات."
    rows = req_df[req_df[REQ_STAGE].astype(str).str.strip() == str(stage).strip()].copy()
    if REQ_WT is not None and weight is not None:
        rows = rows[rows[REQ_WT].astype(str).str.strip() == str(weight).strip()].copy()
    if rows.empty:
        return None, "❌ لم يتم العثور على سطر احتياجات لهذه الحالة/الوزن."
    r = rows.iloc[0]
    cp_req = to_num(r[REQ_CP]); en_req = to_num(r[REQ_EN]); fi_req = to_num(r[REQ_FI])
    dm_req = to_num(r[REQ_DM]) if REQ_DM is not None else np.nan
    if np.isnan(dm_req) or dm_req <= 0: dm_req = 1.0
    if any(np.isnan([cp_req, en_req, fi_req])) or en_req <= 0:
        return None, "❌ قيم الاحتياجات غير صالحة في الإكسل."
    return {"cp_req": cp_req, "en_req": en_req, "fi_req": fi_req, "dm_req": dm_req, "row": r}, None


def recommendation_quality_score(analysis_df):
    coverage = analysis_df["نسبة التغطية %"].to_numpy(dtype=float)
    green_count = int(((coverage >= 95) & (coverage <= 105)).sum())
    yellow_count = int(((coverage > 105) & (coverage < 121)).sum())
    low_count = int((coverage < 95).sum())
    high_count = int((coverage >= 121).sum())
    low_gap = float(np.sum(np.maximum(95.0 - coverage, 0.0) / 95.0))
    high_gap = float(np.sum(np.maximum(coverage - 121.0, 0.0) / 121.0))
    distance_from_ideal = float(np.sum(np.abs(coverage - 100.0) / 100.0))
    return {
        "green_count": green_count,
        "yellow_count": yellow_count,
        "red_count": low_count + high_count,
        "low_count": low_count,
        "high_count": high_count,
        "out_of_range_gap": low_gap + high_gap,
        "distance_from_ideal": distance_from_ideal,
    }


def recommendation_sort_key(sol):
    score = sol.get("quality_score", {})
    return (
        -score.get("green_count", 0),
        -score.get("yellow_count", 0),
        score.get("red_count", 0),
        score.get("out_of_range_gap", 0.0),
        score.get("distance_from_ideal", 0.0),
        sol["cost_per_kg_mix"],
        sol.get("feed_count", 0),
    )


def solve_suitable_ration(feeds_df, feed_names_subset, feed_prices, cp_req, en_req, fi_req):
    sub = feeds_df[feeds_df[FEED_NAME].astype(str).str.strip().isin(feed_names_subset)].copy()
    if sub.empty: return None
    cp_arr = pd.to_numeric(sub[FEED_CP], errors="coerce").to_numpy(dtype=float)
    en_arr = pd.to_numeric(sub[FEED_EN], errors="coerce").to_numpy(dtype=float)
    fi_arr = pd.to_numeric(sub[FEED_FI], errors="coerce").to_numpy(dtype=float)
    names  = sub[FEED_NAME].astype(str).str.strip().tolist()
    p_arr  = np.array([feed_prices.get(n, np.nan) for n in names], dtype=float)
    mask   = ~(np.isnan(cp_arr)|np.isnan(en_arr)|np.isnan(fi_arr)|np.isnan(p_arr))
    cp_arr,en_arr,fi_arr,p_arr = cp_arr[mask],en_arr[mask],fi_arr[mask],p_arr[mask]
    names  = [n for i,n in enumerate(names) if mask[i]]
    if len(names) < 2: return None

    req = np.array([cp_req, en_req, fi_req], dtype=float)
    if np.any(req <= 0) or np.any(np.isnan(req)):
        return None

    nutr_ratio = np.vstack([cp_arr / cp_req, en_arr / en_req, fi_arr / fi_req])
    n = len(names)
    n_vars = n + 6

    # Optimize nutrient fit only. Price is kept for display and final sorting, not for pass/fail.
    c = np.concatenate([np.zeros(n), np.ones(6)])
    a_ub = []
    b_ub = []
    for i in range(3):
        row = np.zeros(n_vars)
        row[:n] = -nutr_ratio[i]
        a_ub.append(row)
        b_ub.append(-0.95)

        row = np.zeros(n_vars)
        row[:n] = nutr_ratio[i]
        a_ub.append(row)
        b_ub.append(1.21)

        row = np.zeros(n_vars)
        row[:n] = nutr_ratio[i]
        row[n + i] = -1.0
        a_ub.append(row)
        b_ub.append(1.0)

        row = np.zeros(n_vars)
        row[:n] = -nutr_ratio[i]
        row[n + 3 + i] = -1.0
        a_ub.append(row)
        b_ub.append(-1.0)

    a_eq = np.zeros((1, n_vars))
    a_eq[0, :n] = 1.0
    res = linprog(
        c,
        A_ub=np.array(a_ub),
        b_ub=np.array(b_ub),
        A_eq=a_eq,
        b_eq=np.array([1.0]),
        bounds=[(0.0, 1.0) for _ in range(n)] + [(0.0, None) for _ in range(6)],
        method="highs",
    )
    if not res.success: return None

    x    = res.x[:n]
    keep = x > 1e-5
    names  = [name for i, name in enumerate(names) if keep[i]]
    x, p_arr = x[keep], p_arr[keep]
    cp_arr,en_arr,fi_arr = cp_arr[keep],en_arr[keep],fi_arr[keep]
    if len(names) == 0: return None
    x = x / np.sum(x)

    cp_ach = float(np.sum(cp_arr*x)); en_ach = float(np.sum(en_arr*x)); fi_ach = float(np.sum(fi_arr*x))
    analysis_df = make_analysis_table(cp_req,en_req,fi_req,cp_ach,en_ach,fi_ach)
    quality_score = recommendation_quality_score(analysis_df)
    if quality_score["green_count"] < 1: return None
    if quality_score["red_count"] == 0:
        status = "✅ مناسب" if quality_score["yellow_count"] == 0 else "⚠️ مقبول: عناصر مرتفعة قليلاً"
    else:
        status = "⚠️ توصية مقبولة لوجود عنصر غذائي مناسب واحد على الأقل"
    return {
        "names": names, "x": x, "prices": p_arr,
        "cp_ach": cp_ach, "en_ach": en_ach, "fi_ach": fi_ach,
        "cost_per_kg_mix": float(np.sum(p_arr*x)),
        "analysis_df": analysis_df, "status_summary": status,
        "is_fallback": False, "issue_messages": [], "fit_penalty": quality_score["out_of_range_gap"],
        "quality_score": quality_score,
    }

def ration_issue_messages(analysis_df):
    messages = []
    slightly_high = 0
    for _, row in analysis_df.iterrows():
        nutrient = str(row["العنصر الغذائي"]).replace(" %", "").replace(" (kcal/kg)", "")
        pct = float(row["نسبة التغطية %"])
        achieved = row["المتحقق من الخلطة"]
        requirement = row["الاحتياج"]
        if pct < 95:
            messages.append(
                f"{nutrient}: التغطية {format_display_num(pct)}% فقط، أقل من الحد المقبول 95%. "
                f"المتحقق {format_display_num(achieved)} مقابل احتياج {format_display_num(requirement)}."
            )
        elif pct >= 121:
            messages.append(
                f"{nutrient}: التغطية {format_display_num(pct)}%، أعلى من الحد المقبول 121%. "
                f"المتحقق {format_display_num(achieved)} مقابل احتياج {format_display_num(requirement)}."
            )
        elif 95 <= pct <= 105:
            slightly_high += 1
    if slightly_high < 1:
        messages.append(
            "لا يوجد عنصر غذائي واحد ضمن الحد المناسب 95–105%."
        )
    return messages


def solve_best_effort_ration(feeds_df, feed_names_subset, feed_prices, cp_req, en_req, fi_req):
    sub = feeds_df[feeds_df[FEED_NAME].astype(str).str.strip().isin(feed_names_subset)].copy()
    if sub.empty: return None
    cp_arr = pd.to_numeric(sub[FEED_CP], errors="coerce").to_numpy(dtype=float)
    en_arr = pd.to_numeric(sub[FEED_EN], errors="coerce").to_numpy(dtype=float)
    fi_arr = pd.to_numeric(sub[FEED_FI], errors="coerce").to_numpy(dtype=float)
    names  = sub[FEED_NAME].astype(str).str.strip().tolist()
    p_arr  = np.array([feed_prices.get(n, np.nan) for n in names], dtype=float)
    mask   = ~(np.isnan(cp_arr)|np.isnan(en_arr)|np.isnan(fi_arr)|np.isnan(p_arr))
    cp_arr,en_arr,fi_arr,p_arr = cp_arr[mask],en_arr[mask],fi_arr[mask],p_arr[mask]
    names  = [n for i,n in enumerate(names) if mask[i]]
    if len(names) < 2: return None

    req = np.array([cp_req, en_req, fi_req], dtype=float)
    if np.any(req <= 0) or np.any(np.isnan(req)):
        return None

    nutr_ratio = np.vstack([cp_arr / cp_req, en_arr / en_req, fi_arr / fi_req])
    n = len(names)
    n_vars = n + 6
    c = np.concatenate([np.zeros(n), np.ones(6)])

    a_ub = []
    b_ub = []
    for i in range(3):
        row = np.zeros(n_vars)
        row[:n] = -nutr_ratio[i]
        row[n + i] = -1.0
        a_ub.append(row)
        b_ub.append(-0.95)

        row = np.zeros(n_vars)
        row[:n] = nutr_ratio[i]
        row[n + 3 + i] = -1.0
        a_ub.append(row)
        b_ub.append(1.21)

    a_eq = np.zeros((1, n_vars))
    a_eq[0, :n] = 1.0
    bounds = [(0.0, 1.0) for _ in range(n)] + [(0.0, None) for _ in range(6)]
    res = linprog(
        c,
        A_ub=np.array(a_ub),
        b_ub=np.array(b_ub),
        A_eq=a_eq,
        b_eq=np.array([1.0]),
        bounds=bounds,
        method="highs",
    )
    if not res.success:
        return None

    x = res.x[:n]
    keep = x > 1e-5
    names = [name for i, name in enumerate(names) if keep[i]]
    x, p_arr = x[keep], p_arr[keep]
    cp_arr,en_arr,fi_arr = cp_arr[keep],en_arr[keep],fi_arr[keep]
    if len(names) == 0:
        return None

    x = x / np.sum(x)
    cp_ach = float(np.sum(cp_arr*x)); en_ach = float(np.sum(en_arr*x)); fi_ach = float(np.sum(fi_arr*x))
    analysis_df = make_analysis_table(cp_req,en_req,fi_req,cp_ach,en_ach,fi_ach)
    coverage = analysis_df["نسبة التغطية %"].to_numpy(dtype=float)
    low_gap = np.maximum(95.0 - coverage, 0.0) / 95.0
    high_gap = np.maximum(coverage - 121.0, 0.0) / 121.0
    quality_score = recommendation_quality_score(analysis_df)
    fit_penalty = float(np.sum(low_gap + high_gap) + (0.25 if quality_score["green_count"] < 1 else 0.0))
    messages = ration_issue_messages(analysis_df)
    if not messages:
        messages = ["هذه الخلطة قريبة من الاحتياج، لكنها لم تمر من فلتر الملاءمة الصارم."]

    return {
        "names": names, "x": x, "prices": p_arr,
        "cp_ach": cp_ach, "en_ach": en_ach, "fi_ach": fi_ach,
        "cost_per_kg_mix": float(np.sum(p_arr*x)),
        "analysis_df": analysis_df,
        "status_summary": "⚠️ أقرب خلطة ممكنة وليست مناسبة بالكامل",
        "is_fallback": True,
        "issue_messages": messages,
        "fit_penalty": fit_penalty,
        "quality_score": quality_score,
    }


def generate_recommendations(cp_req, en_req, fi_req, candidate_feeds, feed_prices, max_results=30):
    from itertools import combinations
    results, seen = [], set()
    candidate_feeds = sorted(candidate_feeds)
    max_combo = min(5, len(candidate_feeds))
    min_combo = 1 if len(candidate_feeds)==1 else 2
    for size in range(min_combo, max_combo+1):
        for combo in combinations(candidate_feeds, size):
            sol = solve_suitable_ration(feeds_df, combo, feed_prices, cp_req, en_req, fi_req)
            if not sol: continue
            key = tuple(sorted(sol["names"]))
            if key in seen: continue
            seen.add(key)
            sol["feed_count"] = len(sol["names"])
            results.append(sol)
    if results:
        results.sort(key=recommendation_sort_key)
        return results[:max_results]

    fallback_results, fallback_seen = [], set()
    for size in range(min_combo, max_combo+1):
        for combo in combinations(candidate_feeds, size):
            sol = solve_best_effort_ration(feeds_df, combo, feed_prices, cp_req, en_req, fi_req)
            if not sol: continue
            key = tuple(sorted(sol["names"]))
            if key in fallback_seen: continue
            fallback_seen.add(key)
            sol["feed_count"] = len(sol["names"])
            fallback_results.append(sol)
    promoted_results = []
    warning_results = []
    for sol in fallback_results:
        if sol.get("quality_score", {}).get("green_count", 0) >= 1:
            sol["is_fallback"] = False
            sol["issue_messages"] = []
            sol["status_summary"] = "⚠️ توصية مقبولة لوجود عنصر غذائي مناسب واحد على الأقل"
            promoted_results.append(sol)
        else:
            warning_results.append(sol)

    if promoted_results:
        promoted_results.sort(key=recommendation_sort_key)
        return promoted_results[:max_results]

    warning_results.sort(key=recommendation_sort_key)
    return warning_results[:max(3, min(max_results, 6))]


ROUGHAGE_FEEDS = ["أتبان", "اتبان", "البرسيم", "برسيم", "الرودس", "رودس", "بلوبونيك"]
RAW_FEEDS = ["نخالة", "مخالة", "مخاله", "شعير مفرود", "ذرة"]
CONCENTRATE_FEEDS = [
    "وفير","جزل","حملان 18% أراسكو","حملان بلس 24% أراسكو",
    "المراح الذهبي","حملان 18% المتحدة","حملان 21% المتحدة","علف ماشية بريمير",
    "المرعى بريمير","حملان 18% بريمير","حملان 21% بريمير","بادي حملان زاد",
    "زاد تسمين","زاد تربية",
]

def feed_category(feed):
    if feed in ROUGHAGE_FEEDS:    return "roughage"
    if feed in RAW_FEEDS:         return "raw"
    if feed in CONCENTRATE_FEEDS: return "concentrate"
    return "other"


def animal_type_for_species(species):
    for animal_type in ANIMAL_TYPES:
        if str(species or "").startswith(animal_type):
            return animal_type
    return None


def sync_herd_groups_with_animals(selected_animals):
    selected_set = set(selected_animals)
    st.session_state.herd_groups = [
        group for group in st.session_state.herd_groups
        if not group.get("species")
        or animal_type_for_species(group["species"]) in selected_set
    ]
    if not st.session_state.herd_groups:
        st.session_state.herd_groups = [blank_group()]


def render_animal_selector():
    section_open("🐾", "ما الحيوانات الموجودة لديك؟")
    selected = [
        animal for animal in ANIMAL_TYPES
        if animal in st.session_state.selected_animal_types
    ]

    columns = st.columns(len(ANIMAL_TYPES))
    for index, (animal, icon) in enumerate(ANIMAL_TYPES.items()):
        is_selected = animal in selected
        label = f"✓ {icon} {animal}" if is_selected else f"{icon} {animal}"
        if columns[index].button(
            label,
            key=f"animal_type_{animal}",
            use_container_width=True,
        ):
            if is_selected:
                selected = [item for item in selected if item != animal]
            else:
                selected.append(animal)
            st.session_state.selected_animal_types = selected
            sync_herd_groups_with_animals(selected)
            st.rerun()

    if selected:
        selected_labels = "".join(
            f"<span class='chip-badge chip-roughage'>{ANIMAL_TYPES[item]} {item}</span>"
            for item in selected
        )
        st.markdown(
            f"<div class='selected-feeds-box'><b>الحيوانات المختارة:</b><br>{selected_labels}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.info("اختر نوع حيوان واحداً على الأقل للمتابعة.")

    section_close()
    return selected


# =============================================================
# V5 — قسم اختيار الأعلاف وتحديد الأسعار (مدمج في قسم واحد)
# =============================================================
#
# يجمع هذا القسم بين اختيار الأعلاف وتحديد/تعديل أسعارها في مكان واحد:
#   - كل علف يظهر كشريحة (chip) ملوّنة بحسب نوعه (عادي / مضاف إليه)
#     مع معاينة السعر مباشرة على الشريحة.
#   - الضغط على اسم العلف داخل الشريحة: تبديل التحديد (مثل السابق).
#   - الضغط على جزء السعر في الشريحة: يحدد العلف (إن لم يكن محدداً)
#     وينقل المستخدم/يبرز سطره في لوحة المراجعة أسفل الشبكة.
#   - الأعلاف التي لا تحتوي على سعر سوقي: يُطلب من المستخدم إدخال
#     سعر الكيس وحجمه بالكيلوجرام معاً (لأن هذه الأعلاف لا تُباع
#     بالكيلو الواحد).
#
def get_feed_package_kg(feed_name, price_records):
    """حجم الكيس بالكيلوجرام لعلف معيّن، أو القيمة الافتراضية إن لم تتوفر."""
    record = price_records.get(feed_name) if price_records else None
    if record:
        package_weight = record.get("package_weight", np.nan)
        try:
            package_weight = float(package_weight)
        except (TypeError, ValueError):
            package_weight = np.nan
        if np.isfinite(package_weight) and package_weight > 0:
            return package_weight
    return DEFAULT_PACKAGE_KG


def get_effective_feed_package_kg(feed_name, price_records, entered_bag_sizes=None):
    """Use only the user-entered bag size for bag calculations."""
    if entered_bag_sizes and feed_name in entered_bag_sizes:
        try:
            package_weight = float(entered_bag_sizes[feed_name])
        except (TypeError, ValueError):
            package_weight = np.nan
        if np.isfinite(package_weight) and package_weight > 0:
            return float(package_weight)
    return np.nan


def calculate_bags_to_buy(total_kg, package_kg, round_up_threshold=BAG_ROUND_UP_THRESHOLD):
    """Round bag counts using a configurable fractional remainder threshold."""
    if not package_kg or package_kg <= 0:
        return 0, 0.0, -float(total_kg)

    total_kg = float(total_kg)
    package_kg = float(package_kg)
    full_bags = math.floor(total_kg / package_kg)
    remainder = total_kg - (full_bags * package_kg)
    threshold_kg = package_kg * float(round_up_threshold)

    bags_to_buy = full_bags + 1 if remainder >= threshold_kg and remainder > 0 else full_bags
    purchased_kg = bags_to_buy * package_kg
    delta_kg = purchased_kg - total_kg
    return bags_to_buy, purchased_kg, delta_kg


def missing_feed_input_names(feed_names, entered_prices, entered_bag_sizes):
    missing = []
    for name in feed_names:
        price_ok = name in entered_prices and np.isfinite(float(entered_prices[name])) and float(entered_prices[name]) > 0
        try:
            bag_size = float((entered_bag_sizes or {}).get(name, np.nan))
        except (TypeError, ValueError):
            bag_size = np.nan
        bag_ok = np.isfinite(bag_size) and bag_size > 0
        if not price_ok or not bag_ok:
            missing.append(name)
    return missing


def build_missing_price_note(feed_names, entered_prices, entered_bag_sizes):
    missing = missing_feed_input_names(feed_names, entered_prices, entered_bag_sizes)
    if not missing:
        return ""
    return "تعتمد الأسعار وعدد الأكياس المعروضة على البيانات المدخلة. لم تُحتسب الأعلاف التي لم يُدخل سعرها أو حجم عبوتها"


def build_missing_price_subnote():
    return "إذا لم يتوفر السعر، يمكنك إدخال سعر تقديري."


def render_feed_selection_and_pricing(feed_names, price_records):
    """
    Render feed selection and price review side by side.
    Returns {feed_name: price_per_kg} for downstream optimization.
    """
    section_open("🌾", "اختيار الأعلاف وتحديد الأسعار")

    def positive_float(value, default=0.0):
        try:
            value = float(value)
        except (TypeError, ValueError):
            return float(default)
        return value if np.isfinite(value) and value > 0 else float(default)

    def manual_values(feed):
        bag_size = st.session_state.get(
            f"bag_{feed}",
            st.session_state.editable_feed_bag_sizes.get(feed, 0.0),
        )
        bag_price = st.session_state.get(
            f"price_{feed}",
            st.session_state.editable_feed_prices.get(feed, 0.0),
        )
        return positive_float(bag_size), positive_float(bag_price)

    def market_values(feed):
        record = price_records.get(feed)
        if not record:
            return 0.0, 0.0
        bag_size = get_feed_package_kg(feed, price_records)
        price_per_kg = positive_float(record.get("price_per_kg"))
        return bag_size, price_per_kg * bag_size

    selected = [
        feed for feed in st.session_state.selected_feed_chips
        if feed in feed_names
    ]
    st.session_state.selected_feed_chips = selected

    feed_nutrient_lookup = {}
    nutrient_cols = [FEED_CP, FEED_EN, FEED_FI]
    if all(nutrient_cols):
        for _, row in feeds_df.iterrows():
            feed_name = str(row.get(FEED_NAME, "")).strip()
            if not feed_name:
                continue
            feed_nutrient_lookup[feed_name] = {
                "protein": to_num(row.get(FEED_CP)),
                "energy": to_num(row.get(FEED_EN)),
                "fiber": to_num(row.get(FEED_FI)),
            }

    def nutrient_text(feed, key):
        value = feed_nutrient_lookup.get(feed, {}).get(key, np.nan)
        return "غير متوفر" if np.isnan(value) else format_display_num(value)

    feed_column, price_column = st.columns([2, 1])

    with feed_column:
        feed_groups = [
            (
                "أعلاف خشنة",
                [feed for feed in feed_names if feed_category(feed) == "roughage"],
                "plain",
            ),
            (
                "اعلاف خام",
                [feed for feed in feed_names if feed_category(feed) == "raw"],
                "raw",
            ),
            (
                "أعلاف مركزة",
                [feed for feed in feed_names if feed_category(feed) == "concentrate"],
                "additive",
            ),
        ]

        for group_title, group_feeds, group_class in feed_groups:
            if not group_feeds:
                continue
            st.markdown(
                f"<span class='checkout-feed-heading {group_class}'>"
                f"{group_title}</span>",
                unsafe_allow_html=True,
            )
            grid_columns = st.columns(3)

            for index, feed in enumerate(group_feeds):
                is_selected = feed in selected
                record = price_records.get(feed)

                if record:
                    bag_size, bag_price = market_values(feed)
                    price_text = (
                        f"{format_display_num(bag_price)} ر.س / {format_display_num(bag_size)} كجم"
                        if bag_size > 0 and bag_price > 0
                        else "لا يوجد سعر محلي"
                    )
                else:
                    price_text = "لا يوجد سعر محلي"

                selected_prefix = "selected_" if is_selected else ""
                button_key = (
                    f"checkout_feed_{selected_prefix}{group_class}_{index}"
                )
                button_label = " "
                selected_class = " selected" if is_selected else ""
                check_html = "<span class='feed-pill-check'>✓</span>" if is_selected else ""
                pill_html = (
                    f"<div class='feed-pill-card {group_class}{selected_class}'>"
                    f"<div class='feed-pill-header'>{check_html}"
                    f"<div class='feed-pill-name'>{html.escape(str(feed))}</div>"
                    f"<div class='feed-pill-price'>{html.escape(price_text)}</div>"
                    "</div>"
                    "<div class='feed-pill-macros'>"
                    f"<span class='feed-macro-tag feed-macro-protein'>بروتين {nutrient_text(feed, 'protein')}%</span>"
                    f"<span class='feed-macro-tag feed-macro-fiber'>ألياف {nutrient_text(feed, 'fiber')}%</span>"
                    f"<span class='feed-macro-tag feed-macro-energy'>طاقة {nutrient_text(feed, 'energy')}</span>"
                    "</div></div>"
                )

                with grid_columns[index % 3]:
                    st.markdown(pill_html, unsafe_allow_html=True)
                    if st.button(
                        button_label,
                        key=button_key,
                        use_container_width=True,
                    ):
                        if is_selected:
                            st.session_state.selected_feed_chips = [
                                item for item in selected if item != feed
                            ]
                        else:
                            st.session_state.selected_feed_chips = selected + [feed]
                        st.rerun()

    selected = [
        feed for feed in st.session_state.selected_feed_chips
        if feed in feed_names
    ]
    st.session_state.selected_feed_chips = selected

    manually_priced = set()
    for feed in selected:
        bag_size, bag_price = manual_values(feed)
        if bag_size > 0 and bag_price > 0:
            manually_priced.add(feed)

    missing_price_feeds = [
        feed for feed in selected
        if feed not in price_records and feed not in manually_priced
    ]
    priced_feeds = [
        feed for feed in selected
        if feed in price_records or feed in manually_priced
    ]

    confirmed_prices = {}

    with price_column:
        with st.container(border=True):
            if not selected:
                st.info("اختر نوعين أو أكثر من الأعلاف لعرض الأسعار.")
            else:
                if missing_price_feeds:
                    st.markdown(
                        "<div class='price-panel-heading'>"
                        "ادخل سعر الأعلاف</div>",
                        unsafe_allow_html=True,
                    )
                    st.markdown(
                        "<div class='price-panel-caption'>"
                        "أدخل حجم الكيس وسعره. إذا لم تعرف السعر أدخل سعرًا متوقعًا.</div>",
                        unsafe_allow_html=True,
                    )

                    for feed in missing_price_feeds:
                        safe_feed = html.escape(str(feed))
                        st.markdown(
                            f"<div class='price-panel-feed-label "
                            f"price-panel-missing-row'>{safe_feed}</div>",
                            unsafe_allow_html=True,
                        )

                        current_bag, current_price = manual_values(feed)
                        bag_col, price_col = st.columns(2)
                        with bag_col:
                            bag_size = st.number_input(
                                "حجم الكيس (كج)",
                                min_value=0.0,
                                step=1.0,
                                value=float(current_bag),
                                key=f"bag_{feed}",
                            )
                        with price_col:
                            bag_price = st.number_input(
                                "السعر (ر.س)",
                                min_value=0.0,
                                step=0.5,
                                value=float(current_price),
                                key=f"price_{feed}",
                            )

                        st.session_state.editable_feed_bag_sizes[feed] = bag_size
                        st.session_state.editable_feed_prices[feed] = bag_price
                        st.session_state.price_source_choice[
                            f"price_source_{feed}"
                        ] = "custom"

                        if bag_size > 0 and bag_price > 0:
                            confirmed_prices[feed] = bag_price / bag_size

                if missing_price_feeds and priced_feeds:
                    st.divider()

                if priced_feeds:
                    st.markdown(
                        "<div class='price-panel-heading'>"
                        "تعديل الأسعار المتوفرة</div>",
                        unsafe_allow_html=True,
                    )
                    st.markdown(
                        "<div class='price-panel-caption'>"
                        "أدخل الحجم والسعر المعتمدين للحساب. أسعار سنابل الجوف المعروضة مرجعية فقط، وإذا لم تعرف السعر أدخل سعرًا متوقعًا.</div>",
                        unsafe_allow_html=True,
                    )

                    for feed in priced_feeds:
                        safe_feed = html.escape(str(feed))
                        record = price_records.get(feed)
                        source_key = f"price_source_{feed}"
                        manual_bag, manual_price = manual_values(feed)

                        if record:
                            market_bag, market_price = market_values(feed)
                            source = st.session_state.price_source_choice.get(
                                source_key,
                                "market",
                            )
                            if (
                                source == "custom"
                                and manual_bag > 0
                                and manual_price > 0
                            ):
                                default_bag = manual_bag
                                default_price = manual_price
                            else:
                                default_bag = market_bag
                                default_price = market_price
                        else:
                            market_bag = market_price = 0.0
                            default_bag = manual_bag
                            default_price = manual_price

                        label_col, bag_col, price_col = st.columns(
                            [1.35, 1, 1]
                        )
                        with label_col:
                            reference_html = ""
                            if record and market_bag > 0 and market_price > 0:
                                reference_html = (
                                    f"<div class='price-panel-caption'>مرجع السوق: {format_display_num(market_price)} ر.س / {format_display_num(market_bag)} كجم</div>"
                                )
                            st.markdown(
                                f"<div class='price-panel-feed-label "
                                f"price-panel-priced-row'>{safe_feed}</div>{reference_html}",
                                unsafe_allow_html=True,
                            )
                        with bag_col:
                            bag_size = st.number_input(
                                "حجم الكيس (كج)",
                                min_value=0.0,
                                step=1.0,
                                value=float(default_bag),
                                key=f"bag_{feed}",
                            )
                        with price_col:
                            bag_price = st.number_input(
                                "السعر (ر.س)",
                                min_value=0.0,
                                step=0.5,
                                value=float(default_price),
                                key=f"price_{feed}",
                            )

                        if record:
                            uses_market_price = (
                                np.isclose(bag_size, market_bag)
                                and np.isclose(bag_price, market_price)
                            )
                            st.session_state.price_source_choice[source_key] = (
                                "market" if uses_market_price else "custom"
                            )
                        else:
                            st.session_state.price_source_choice[source_key] = "custom"

                        st.session_state.editable_feed_bag_sizes[feed] = bag_size
                        st.session_state.editable_feed_prices[feed] = bag_price

                        if bag_size > 0 and bag_price > 0:
                            confirmed_prices[feed] = bag_price / bag_size

                current_price_rows = []
                for feed in selected:
                    bag_size, bag_price = manual_values(feed)
                    if bag_size > 0 and bag_price > 0:
                        summary_text = (
                            f"{format_display_num(bag_price)} ر.س / "
                            f"{format_display_num(bag_size)} كجم"
                        )
                    else:
                        summary_text = "0.00 ر.س / 0.00 كجم"
                    current_price_rows.append(
                        "<div class='price-panel-available-row'>"
                        f"<span class='price-panel-available-name'>{html.escape(str(feed))}</span>"
                        f"<span class='price-panel-available-value'>{html.escape(summary_text)}</span>"
                        "</div>"
                    )

                if current_price_rows:
                    st.markdown(
                        "<div class='price-panel-available-box'>"
                        "<div class='price-panel-heading'>اسعار الاعلاف</div>"
                        "<div class='price-panel-caption'>القيم التالية تعكس الأسعار والأحجام الحالية للأعلاف المختارة بعد التعديل أو الإدخال.</div>"
                        f"{''.join(current_price_rows)}"
                        "</div>",
                        unsafe_allow_html=True,
                    )

    section_close()
    return confirmed_prices

def herd_group_title(index):
    titles = {
        1: "القطيع الاول",
        2: "القطيع الثاني",
        3: "القطيع الثالث",
        4: "القطيع الرابع",
        5: "القطيع الخامس",
        6: "القطيع السادس",
        7: "القطيع السابع",
        8: "القطيع الثامن",
        9: "القطيع التاسع",
        10: "القطيع العاشر",
    }
    title = titles.get(index, f"القطيع {index}")
    return f"\u2003{title}\u2003"


def render_group_editor(selected_animal_types):
    section_open("🐑", "تفاصيل القطيع")

    ordered_animal_types = [
        animal_type for animal_type in ANIMAL_TYPES
        if animal_type in selected_animal_types
    ]
    animal_display_names = {
        "ضأن": "الضأن",
        "ماعز": "الماعز",
        "إبل": "الإبل",
    }

    for group in st.session_state.herd_groups:
        if group.get("species"):
            group["animal_type"] = animal_type_for_species(group["species"])
        elif not group.get("animal_type") and len(ordered_animal_types) == 1:
            group["animal_type"] = ordered_animal_types[0]

    st.session_state.herd_groups = [
        group for group in st.session_state.herd_groups
        if not group.get("animal_type")
        or group["animal_type"] in ordered_animal_types
    ]

    for animal_type in ordered_animal_types:
        animal_groups = [
            group for group in st.session_state.herd_groups
            if group.get("animal_type") == animal_type
            or animal_type_for_species(group.get("species")) == animal_type
        ]
        if not animal_groups:
            add_group(blank_group(animal_type))
            animal_groups = [st.session_state.herd_groups[-1]]

        allowed_species = [
            species for species in SPECIES_SHEETS
            if animal_type_for_species(species) == animal_type
        ]

        for idx, g in enumerate(animal_groups, start=1):
            title = herd_group_title(idx).strip()
            animal_label = animal_display_names.get(animal_type, animal_type)
            with st.expander(f"\u2003معلومات {title} - {animal_label}\u2003", expanded=True):
                species_options = ["اختر فئة الحيوان"] + allowed_species

                field_cols = st.columns(4)
                with field_cols[0]:
                    species_index = species_options.index(g["species"]) if g.get("species") in species_options else 0
                    selected_species = st.selectbox(
                        "فئة الحيوان",
                        species_options,
                        index=species_index,
                        key=f"g_species_{g['id']}",
                    )
                    g["species"] = None if selected_species == "اختر فئة الحيوان" else selected_species
                    g["animal_type"] = animal_type

                stage_options  = ["اختر الفترة / الحالة الإنتاجية"]
                weight_options = ["اختر وزن الحيوان"]
                stage_disabled = weight_disabled = True

                if g["species"]:
                    req_df = sheets[SPECIES_SHEETS[g["species"]]].copy()
                    REQ_STAGE,_,_,_,_,_ = get_req_columns(req_df)
                    stage_options += build_stage_options(req_df, REQ_STAGE)
                    stage_disabled = False

                with field_cols[1]:
                    stage_index = stage_options.index(g["stage"]) if g.get("stage") in stage_options else 0
                    selected_stage = st.selectbox(
                        "الفترة / الحالة الإنتاجية",
                        stage_options,
                        index=stage_index,
                        key=f"g_stage_{g['id']}",
                        disabled=stage_disabled,
                    )
                    g["stage"] = None if selected_stage == "اختر الفترة / الحالة الإنتاجية" else selected_stage

                if g["species"] and g["stage"]:
                    req_df = sheets[SPECIES_SHEETS[g["species"]]].copy()
                    REQ_STAGE,_,_,_,REQ_WT,_ = get_req_columns(req_df)
                    stage_rows = req_df[req_df[REQ_STAGE].astype(str).str.strip() == g["stage"]].copy()
                    weights = build_weight_options(stage_rows, REQ_WT) if REQ_WT is not None else []
                    if weights:
                        weight_options += weights
                        weight_disabled = False
                    else:
                        g["weight"] = None

                with field_cols[2]:
                    weight_index = weight_options.index(g["weight"]) if g.get("weight") in weight_options else 0
                    selected_weight = st.selectbox(
                        "متوسط الأوزان (كجم)",
                        weight_options,
                        index=weight_index,
                        key=f"g_weight_{g['id']}",
                        disabled=weight_disabled,
                    )
                    g["weight"] = None if selected_weight == "اختر وزن الحيوان" else selected_weight

                with field_cols[3]:
                    g["n"] = st.number_input(
                        "عدد الحيوانات",
                        min_value=1,
                        step=1,
                        value=int(g.get("n", 1)),
                        key=f"g_n_{g['id']}",
                    )

                if len(st.session_state.herd_groups) > 1:
                    if st.button("🗑  حذف هذا القطيع", key=f"del_{g['id']}"):
                        remove_group(g["id"])
                        st.rerun()

        if st.button(f"➕  إضافة قطيع {animal_type}", key=f"add_animal_group_{animal_type}", use_container_width=True):
            add_group(blank_group(animal_type))
            st.rerun()

    herd_days = st.number_input("عدد الأيام", min_value=1, step=1, value=30, key="herd_days")
    section_close()
    return int(herd_days)

def build_optimizer_prices(selected_feeds, entered_prices, market_prices):
    prices = {}
    known_prices = []
    for feed in selected_feeds:
        if feed in entered_prices:
            price = float(entered_prices[feed])
            prices[feed] = price
            known_prices.append(price)
        elif feed in market_prices:
            price = float(market_prices[feed])
            prices[feed] = price
            known_prices.append(price)

    fallback_price = float(np.nanmedian(known_prices)) if known_prices else 1.0
    for feed in selected_feeds:
        if feed not in prices:
            prices[feed] = fallback_price
    return prices


def group_is_complete(group):
    if not group.get("species") or not group.get("stage"): return False
    req_df = sheets[SPECIES_SHEETS[group["species"]]].copy()
    _,_,_,_,REQ_WT,_ = get_req_columns(req_df)
    if REQ_WT is None: return True
    rows = req_df[req_df[get_req_columns(req_df)[0]].astype(str).str.strip() == str(group["stage"]).strip()].copy()
    return not build_weight_options(rows, REQ_WT) or group.get("weight") is not None


def macro_badge_class(status):
    if "✅" in str(status):   return "macro-ok"
    if "⚠" in str(status):   return "macro-warn"
    return "macro-bad"


GAUGE_CARD_CSS = """
.rec-card-wide {
  border: 2px solid #E0DCF0;
  border-radius: 16px;
  overflow: hidden;
  background: #fff;
  margin-bottom: 18px;
  box-shadow: 0 4px 18px rgba(44,49,73,0.06);
}
.rec-card-wide.best-card {
  border: 2.5px solid #6B4CA0;
  box-shadow: 0 6px 24px rgba(107,76,160,0.16);
}
.rec-best-strip {
  background: #6B4CA0;
  color: #fff;
  font-size: 0.78rem;
  font-weight: 800;
  text-align: center;
  padding: 6px 10px;
}
.rec-card-top {
  background: #2C3149;
  padding: 14px 22px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: wrap;
  gap: 12px;
}
.rec-applies-text {
  color: #fff;
  font-size: 0.82rem;
  font-weight: 700;
}
.rec-applies-text span {
  color: #9BB8D4;
  font-weight: 500;
  font-size: 0.70rem;
  display: block;
  margin-top: 3px;
}
.rec-price-block {
  display: flex;
  gap: 18px;
  align-items: baseline;
  flex-wrap: wrap;
}
.rec-price-item { text-align: right; }
.rec-price-label {
  font-size: 0.56rem;
  color: rgba(255,255,255,0.55);
  margin-bottom: 2px;
}
.rec-price-value {
  font-size: 1.05rem;
  font-weight: 800;
  color: #F5C842;
}
.rec-price-value small {
  font-size: 0.62rem;
  color: rgba(255,255,255,0.55);
  font-weight: 500;
}
.rec-card-main {
  display: block;
}
.rec-gauges-col {
  padding: 18px 22px 16px;
}
.rec-feed-col {
  padding: 16px 22px 18px;
  background: #FAF9FD;
  border-top: 1px solid #F0EEF7;
  overflow-x: auto;
}
.rec-col-title {
  font-size: 0.68rem;
  color: #888;
  font-weight: 700;
  margin-bottom: 12px;
}
.gauge-row {
  margin: 0 0 18px;
  position: relative;
}
.gauge-row:last-child { margin-bottom: 0; }
.gauge-label-row {
  display: flex;
  justify-content: space-between;
  align-items: baseline;
  gap: 10px;
  margin-bottom: 5px;
}
.gauge-name {
  font-size: 0.76rem;
  font-weight: 700;
  color: #2C3149;
}
.gauge-name small {
  font-size: 0.62rem;
  color: #999;
  font-weight: 500;
  margin-right: 3px;
}
.gauge-value {
  font-size: 0.80rem;
  font-weight: 800;
  color: #2C3149;
  white-space: nowrap;
}
.gauge-strip {
  display: flex;
  direction: ltr;
  height: 11px;
  border-radius: 6px;
  overflow: hidden;
  position: relative;
}
.gauge-zone { height: 100%; }
.gauge-zone-red { background: #E85C4A; }
.gauge-zone-yellow { background: #F0B429; }
.gauge-zone-green { background: #4A8C3F; }
.gauge-marker {
  position: absolute;
  top: 21px;
  width: 0;
  height: 0;
  border-left: 6px solid transparent;
  border-right: 6px solid transparent;
  border-top: 8px solid #2C3149;
  transform: translateX(-50%);
  cursor: pointer;
}
.gauge-marker-line {
  position: absolute;
  top: 29px;
  width: 2px;
  height: 6px;
  background: #2C3149;
  transform: translateX(-50%);
}
.gauge-tooltip {
  position: absolute;
  top: -25px;
  background: #2C3149;
  color: #fff;
  font-size: 0.66rem;
  padding: 3px 8px;
  border-radius: 5px;
  white-space: nowrap;
  transform: translateX(-50%);
  font-weight: 700;
  opacity: 0;
  transition: opacity 0.15s;
  pointer-events: none;
  z-index: 5;
}
.gauge-row:hover .gauge-tooltip { opacity: 1; }
.gauge-scale-labels {
  display: flex;
  direction: ltr;
  margin-top: 5px;
  font-size: 0.54rem;
  color: #888;
  text-align: center;
}
.gauge-scale-labels span {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 1px;
  line-height: 1.15;
  white-space: nowrap;
}
.gauge-scale-labels b {
  color: #666;
  font-weight: 800;
}
.gauge-scale-labels small {
  color: #aaa;
  font-size: 0.50rem;
}
.rec-feed-table-wide {
  width: 100%;
  min-width: 420px;
  border-collapse: collapse;
  font-size: 0.68rem;
}
.rec-feed-table-wide th {
  font-size: 0.56rem;
  font-weight: 700;
  color: #888;
  text-align: right;
  padding: 0 4px 4px;
  border-bottom: 1.5px solid #e5e2ec;
}
.rec-feed-table-wide td {
  padding: 4px 4px;
  text-align: right;
  border-bottom: 0.5px solid #E9E5F1;
  color: #333;
}
.rec-feed-table-wide td:first-child {
  color: #555;
  font-weight: 600;
}
.rec-feed-table-wide tr:last-child td { border-bottom: none; }
.rec-unit-badge {
  display: inline-block;
  background: #EAF2FF;
  color: #1D4ED8;
  font-size: 0.56rem;
  font-weight: 700;
  padding: 1px 5px;
  border-radius: 999px;
  white-space: nowrap;
}
.rec-card-footer-wide {
  background: #F3EEF9;
  border-top: 1px solid #E0D6F0;
  padding: 10px 22px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 12px;
  font-size: 0.76rem;
}
.rec-footer-label { color: #8D7BA3; }
.rec-footer-value {
  font-weight: 800;
  color: #6B4CA0;
  white-space: nowrap;
}
.rec-missing-price-alert {
  background: #FFF8E6;
  border-bottom: 1px solid #F5C842;
  padding: 6px 16px;
  font-size: 0.68rem;
  line-height: 1.4;
  color: #7a5500;
  margin: 0;
}
.rec-missing-price-alert-sub {
  font-size: 0.60rem;
  color: #8B6A1B;
  margin-top: 2px;
}
@media (max-width: 700px) {
  .rec-card-top,
  .rec-card-footer-wide {
    align-items: flex-start;
    flex-direction: column;
  }
  .rec-gauges-col,
  .rec-feed-col {
    padding: 15px;
  }
}
"""

st.markdown(f"<style>{GAUGE_CARD_CSS}</style>", unsafe_allow_html=True)

GAUGE_ZONES = {
    "default": {
        "low_red_end": 95,
        "green_start": 95,
        "green_end": 105,
        "high_red_start": 121,
    },
}


def get_gauge_zones(macro_name):
    return GAUGE_ZONES.get(macro_name, GAUGE_ZONES["default"])


def render_gauge_bar(macro_name, requirement_value, achieved_value,
                     coverage_pct, judgment_text, unit_label=""):
    zones = get_gauge_zones(macro_name)
    low_red_end = zones["low_red_end"]
    green_start = zones["green_start"]
    green_end = zones["green_end"]
    high_red_start = zones["high_red_start"]

    bar_ceiling = max(high_red_start * 1.5, coverage_pct * 1.1, 150)
    w_red_low = (low_red_end / bar_ceiling) * 100
    w_green = ((green_end - green_start) / bar_ceiling) * 100
    w_yellow = ((high_red_start - green_end) / bar_ceiling) * 100
    w_red_high = max(100 - w_red_low - w_green - w_yellow, 0)
    marker_pos = min(100, max(0, (coverage_pct / bar_ceiling) * 100))

    tooltip = html.escape(
        f"{format_display_num(achieved_value)}{unit_label} — {judgment_text} "
        f"({format_display_num(coverage_pct)}% من الاحتياج)"
    )
    macro = html.escape(str(macro_name))
    return (
        "<div class='gauge-row'>"
        "<div class='gauge-label-row'>"
        f"<span class='gauge-name'>{macro} "
        f"<small>(الاحتياج: {format_display_num(requirement_value)}{unit_label})</small></span>"
        f"<span class='gauge-value'>{format_display_num(achieved_value)}{unit_label}</span>"
        "</div>"
        "<div class='gauge-strip'>"
        f"<div class='gauge-zone gauge-zone-red' style='width:{w_red_low:.2f}%'></div>"
        f"<div class='gauge-zone gauge-zone-green' style='width:{w_green:.2f}%'></div>"
        f"<div class='gauge-zone gauge-zone-yellow' style='width:{w_yellow:.2f}%'></div>"
        f"<div class='gauge-zone gauge-zone-red' style='width:{w_red_high:.2f}%'></div>"
        "</div>"
        f"<div class='gauge-marker' style='left:{marker_pos:.2f}%'></div>"
        f"<div class='gauge-marker-line' style='left:{marker_pos:.2f}%'></div>"
        f"<div class='gauge-tooltip' style='left:{marker_pos:.2f}%'>{tooltip}</div>"
        "<div class='gauge-scale-labels'>"
        f"<span style='width:{w_red_low:.2f}%'><b>&lt;{low_red_end:g}%</b><small>أقل من الحد</small></span>"
        f"<span style='width:{w_green:.2f}%'><b>{green_start:g}–{green_end:g}%</b><small>مناسب</small></span>"
        f"<span style='width:{w_yellow:.2f}%'><b>{green_end:g}–{high_red_start:g}%</b><small>مرتفع قليلاً</small></span>"
        f"<span style='width:{w_red_high:.2f}%'><b>&gt;{high_red_start:g}%</b><small>أعلى من الحد</small></span>"
        "</div></div>"
    )


def render_recommendation_card_wide(card, price_records, entered_prices, entered_bag_sizes, is_best=False):
    sol = card["sol"]
    dm_req = card["dm_req"]
    n_days = card["n_days"]
    total_animals = card["total_animals"]
    group_labels = card["group_labels"]

    missing_names = missing_feed_input_names(sol["names"], entered_prices, entered_bag_sizes)
    has_missing_price = bool(missing_names)
    estimated_cost_per_kg = 0.0
    for name, x_val in zip(sol["names"], sol["x"]):
        if name in entered_prices:
            estimated_cost_per_kg += float(entered_prices[name]) * float(x_val)
    head_period_cost = estimated_cost_per_kg * dm_req * n_days
    period_cost = head_period_cost * total_animals
    period_label = f"{format_display_num(period_cost, use_commas=True)} <small>ر.س</small>"
    head_period_label = f"{format_display_num(head_period_cost, use_commas=True)} <small>ر.س</small>"

    card_class = "rec-card-wide best-card" if is_best else "rec-card-wide"
    feed_mix_title = html.escape(" + ".join(sol["names"]))
    best_strip = (
        "<div class='rec-best-strip'>⭐ الأفضل قيمةً</div>"
        if is_best else ""
    )

    gauges_html = ""
    unit_lookup = {
        "البروتين %": "%",
        "الطاقة (kcal/kg)": " kcal/كجم",
        "الألياف %": "%",
    }
    for _, row in sol["analysis_df"].iterrows():
        raw_name = str(row["العنصر الغذائي"])
        clean_name = raw_name.replace(" %", "").replace(" (kcal/kg)", "")
        gauges_html += render_gauge_bar(
            clean_name,
            row["الاحتياج"],
            row["المتحقق من الخلطة"],
            row["نسبة التغطية %"],
            row["التقييم"],
            unit_lookup.get(raw_name, ""),
        )

    feed_rows = ""
    for name, x_val in zip(sol["names"], sol["x"]):
        pct_val = x_val * 100
        kg_day_head = x_val * dm_req
        kg_total_period = kg_day_head * total_animals * n_days
        package_kg = get_effective_feed_package_kg(name, price_records, entered_bag_sizes)
        if np.isfinite(package_kg) and package_kg > 0:
            bags_total = kg_total_period / package_kg
            bag_text = f"{format_display_num(bags_total)} كيس ({format_display_num(package_kg)} كجم)"
        else:
            bag_text = "-"
        feed_rows += (
            "<tr>"
            f"<td>{html.escape(str(name))}</td>"
            f"<td>{format_display_num(pct_val)}%</td>"
            f"<td>{format_display_num(kg_day_head)}</td>"
            f"<td><span class='rec-unit-badge'>{html.escape(bag_text)}</span></td>"
            "</tr>"
        )

    alert_html = ""
    if has_missing_price:
        note_text = html.escape(build_missing_price_note(sol["names"], entered_prices, entered_bag_sizes))
        subnote_text = html.escape(build_missing_price_subnote())
        alert_html = (
            "<div class='rec-missing-price-alert'>"
            f"<div>{note_text}</div>"
            f"<div class='rec-missing-price-alert-sub'>{subnote_text}</div>"
            "</div>"
        )

    total_herd_kg_day = dm_req * total_animals
    card_html = (
        f"<div class='{card_class}'>"
        f"{best_strip}"
        "<div class='rec-card-top'>"
        "<div class='rec-applies-text'>خلطة الأعلاف"
        f"<span>{feed_mix_title}</span></div>"
        "<div class='rec-price-block'>"
        f"<div class='rec-price-item'><div class='rec-price-label'>تكلفة تقديرية للرأس للفترة ({n_days} يوم)</div>"
        f"<div class='rec-price-value'>{head_period_label}</div></div>"
        f"<div class='rec-price-item'><div class='rec-price-label'>تكلفة تقديرية لإجمالي الفترة ({n_days} يوم)</div>"
        f"<div class='rec-price-value'>{period_label}</div></div>"
        "</div></div>"
        f"{alert_html}"
        "<div class='rec-card-main'>"
        "<div class='rec-gauges-col'><div class='rec-col-title'>مؤشرات التغطية الغذائية</div>"
        f"{gauges_html}</div>"
        "<div class='rec-feed-col'><div class='rec-col-title'>مكونات الخلطة</div>"
        "<table class='rec-feed-table-wide'><thead><tr>"
        "<th>العلف</th><th>النسبة</th><th>كجم/يوم/رأس</th>"
        "<th>عدد الاكياس تقريبا لكل الفترة</th></tr></thead>"
        f"<tbody>{feed_rows}</tbody></table></div>"
        "</div>"
        "<div class='rec-card-footer-wide'>"
        f"<span class='rec-footer-label'>إجمالي يومي للقطيع ({total_animals} رؤوس)</span>"
        f"<span class='rec-footer-value'>{format_display_num(total_herd_kg_day)} كجم/يوم</span>"
        "</div></div>"
    )
    st.markdown(card_html, unsafe_allow_html=True)
def _print_page_header(title, subtitle=""):
    logo_html = f"<img class='print-logo' src='data:image/png;base64,{print_logo_b64}' />" if print_logo_b64 else ""
    subtitle_html = f"<div class='print-page-subtitle'>{html.escape(subtitle)}</div>" if subtitle else ""
    return (
        "<div class='print-page-header'>"
        f"{logo_html}"
        "</div>"
        "<div class='print-title-block'>"
        f"<h1>{html.escape(title)}</h1>"
        f"{subtitle_html}"
        "</div>"
    )


def _format_bag_quantity_text(total_kg, package_kg):
    if package_kg is None or not np.isfinite(package_kg) or package_kg <= 0:
        return "-", "-"
    raw_bags = total_kg / package_kg if package_kg else 0.0
    rounded_bags, _, delta_kg = calculate_bags_to_buy(total_kg, package_kg)
    raw_text = f"{format_display_num(raw_bags)}"
    rounded_text = f"{raw_text} ← {rounded_bags} كيس"
    if delta_kg > 0:
        status_text = f"فائض {format_display_num(delta_kg)} كجم"
    elif delta_kg < 0:
        status_text = f"عجز {format_display_num(abs(delta_kg))} كجم"
    else:
        status_text = "مطابق"
    return rounded_text, status_text


def _print_group_info_table(group):
    rows = [
        ("الفئة", group.get("species") or "غير محدد"),
        ("الحالة / المرحلة", group.get("stage") or "غير محدد"),
        ("الوزن", group.get("weight") or "غير محدد"),
        ("عدد الرؤوس", int(group.get("n", 1))),
    ]
    body = "".join(
        f"<tr><th>{html.escape(str(label))}</th><td>{html.escape(str(value))}</td></tr>"
        for label, value in rows
    )
    return (
        "<div class='print-table-card'>"
        "<h2>بيانات القطيع</h2>"
        f"<table class='print-info-table'><tbody>{body}</tbody></table>"
        "</div>"
    )


def _print_cost_table(sol, dm_req, total_animals, n_days, entered_prices, entered_bag_sizes):
    missing_names = missing_feed_input_names(sol["names"], entered_prices, entered_bag_sizes)
    estimated_cost_per_kg = 0.0
    for name, x_val in zip(sol["names"], sol["x"]):
        if name in entered_prices:
            estimated_cost_per_kg += float(entered_prices[name]) * float(x_val)
    per_head_value = estimated_cost_per_kg * dm_req * n_days
    total_value = per_head_value * total_animals
    total_cost = f"{format_display_num(total_value, use_commas=True)} ر.س"
    per_head_cost = f"{format_display_num(per_head_value, use_commas=True)} ر.س"
    note_text = build_missing_price_note(sol["names"], entered_prices, entered_bag_sizes) if missing_names else ""

    rows = [
        ("عدد الأيام", n_days),
        ("التكلفة التقريبيه", total_cost),
        ("التكلفة لكل راس (تقريبا)", per_head_cost),
    ]
    if note_text:
        rows.append(("ملاحظة", note_text))
    body = "".join(
        f"<tr><th>{html.escape(str(label))}</th><td>{html.escape(str(value))}</td></tr>"
        for label, value in rows
    )
    return (
        "<div class='print-table-card'>"
        "<h2>المدة والتكلفة</h2>"
        f"<table class='print-info-table'><tbody>{body}</tbody></table>"
        "</div>"
    )


def _print_feed_mix_table(sol, dm_req, total_animals, n_days, price_records, entered_bag_sizes, entered_prices):
    rows = []
    for name, x_val in zip(sol["names"], sol["x"]):
        kg_period = x_val * dm_req * total_animals * n_days
        package_kg = get_effective_feed_package_kg(name, price_records, entered_bag_sizes)
        bags_text, status_text = _format_bag_quantity_text(kg_period, package_kg)
        if np.isfinite(package_kg) and package_kg > 0:
            rounded_bags, _, _ = calculate_bags_to_buy(kg_period, package_kg)
        else:
            rounded_bags = 0
        price_per_kg = entered_prices.get(name)
        if price_per_kg is None or not (np.isfinite(package_kg) and package_kg > 0):
            total_price_text = NO_PRICE_TEXT
        else:
            total_price = rounded_bags * package_kg * float(price_per_kg)
            total_price_text = f"{format_display_num(total_price, use_commas=True)} ر.س"
        rows.append(
            "<tr>"
            f"<td>{html.escape(str(name))}</td>"
            f"<td>{html.escape(bags_text)}</td>"
            f"<td>{html.escape(total_price_text)}</td>"
            f"<td>{html.escape(status_text)}</td>"
            "</tr>"
        )

    return (
        "<div class='print-table-card'>"
        "<h2>تفاصيل الأعلاف</h2>"
        "<table class='print-merged-table'>"
        "<thead><tr>"
        "<th>العلف</th><th>عدد الأكياس</th><th>السعر التقريبي</th><th>الفائض / العجز</th>"
        "</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody></table>"
        "</div>"
    )


def _print_farm_summary_page(plans, price_records):
    feed_totals = {}
    feed_groups = {}
    feed_bag_sizes = {}
    feed_prices = {}
    farm_total_cost = 0.0
    has_missing_price = False
    pill_classes = ["pill-a", "pill-b", "pill-c", "pill-d", "pill-e"]

    for plan in plans:
        sol = plan["sol"]
        dm_req = plan["dm_req"]
        total_animals = plan["total_animals"]
        n_days = plan["n_days"]
        entered_bag_sizes = plan.get("entered_bag_sizes") or {}
        entered_prices = plan.get("entered_prices") or {}
        group = plan.get("group") or {}
        group_label = f"{group.get('species') or 'غير محدد'} | {group.get('stage') or 'غير محدد'}"

        for name, x_val in zip(sol["names"], sol["x"]):
            total_kg = x_val * dm_req * total_animals * n_days
            feed_totals[name] = feed_totals.get(name, 0.0) + total_kg
            feed_groups.setdefault(name, [])
            if group_label not in feed_groups[name]:
                feed_groups[name].append(group_label)
            if name not in feed_bag_sizes and name in entered_bag_sizes:
                try:
                    bag_size = float(entered_bag_sizes[name])
                except (TypeError, ValueError):
                    bag_size = np.nan
                if np.isfinite(bag_size) and bag_size > 0:
                    feed_bag_sizes[name] = bag_size
            if name not in feed_prices and name in entered_prices:
                try:
                    feed_prices[name] = float(entered_prices[name])
                except (TypeError, ValueError):
                    pass

    unique_group_labels = sorted({label for labels in feed_groups.values() for label in labels})
    pill_class_by_group = {
        label: pill_classes[index % len(pill_classes)]
        for index, label in enumerate(unique_group_labels)
    }

    rows = []
    for name in sorted(feed_totals):
        total_kg = feed_totals[name]
        package_kg = get_effective_feed_package_kg(name, price_records, feed_bag_sizes)
        bags_text, _ = _format_bag_quantity_text(total_kg, package_kg)
        if np.isfinite(package_kg) and package_kg > 0:
            rounded_bags, _, _ = calculate_bags_to_buy(total_kg, package_kg)
        else:
            rounded_bags = 0
        price_per_kg = feed_prices.get(name)
        if price_per_kg is None or not (np.isfinite(package_kg) and package_kg > 0):
            has_missing_price = True
            total_price_text = NO_PRICE_TEXT
        else:
            feed_total_cost = rounded_bags * package_kg * price_per_kg
            farm_total_cost += feed_total_cost
            total_price_text = f"{format_display_num(feed_total_cost, use_commas=True)} ر.س"

        group_pills = "".join(
            f"<span class='print-group-pill {pill_class_by_group.get(label, pill_classes[0])}'>{html.escape(label)}</span>"
            for label in feed_groups.get(name, [])
        )
        rows.append(
            "<tr>"
            f"<td>{html.escape(str(name))}</td>"
            f"<td>{html.escape(bags_text)}</td>"
            f"<td><div class='print-pill-wrap'>{group_pills}</div></td>"
            f"<td>{html.escape(total_price_text)}</td>"
            "</tr>"
        )

    total_cost_text = f"{format_display_num(farm_total_cost, use_commas=True)} ر.س"
    if has_missing_price:
        missing_summary = []
        for name in sorted(feed_totals):
            if name not in feed_prices or not (np.isfinite(float(feed_bag_sizes.get(name, np.nan))) and float(feed_bag_sizes.get(name, np.nan)) > 0):
                missing_summary.append(str(name))
        total_cost_text += " (لا يشمل الأعلاف التالية لعدم إدخال السعر والحجم: " + "، ".join(missing_summary) + ")"

    farm_days = int(plans[0].get("n_days", 0)) if plans else 0
    return (
        "<section class='print-page'>"
        f"{_print_page_header('ملخص المزرعة', 'ملخص موحد للأعلاف المختارة لجميع القطعان')}"
        "<div class='print-summary-total'>"
        f"<span>إجمالي السعر التقريبي لخطط التغذية للمزرعة ({farm_days} يوم): <b>{html.escape(total_cost_text)}</b></span>"
        "</div>"
        "<div class='print-table-card'>"
        "<h2>ملخص الأعلاف</h2>"
        "<table class='print-merged-table'>"
        "<thead><tr>"
        "<th>العلف</th><th>عدد الأكياس (خام ← مقرب)</th><th>ينطبق على</th><th>السعر التقريبي</th>"
        "</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody></table>"
        "</div>"
        "</section>"
    )


def _print_plan_section(plan, index, price_records):
    sol = plan["sol"]
    group = plan.get("group") or {}
    title = f"القطيع #{index}: {group.get('species') or 'غير محدد'} | {group.get('stage') or 'غير محدد'}"
    subtitle = datetime.now().strftime('%Y-%m-%d %H:%M')
    return (
        "<section class='print-page'>"
        f"{_print_page_header(title, subtitle)}"
        f"{_print_group_info_table(group)}"
        f"{_print_cost_table(sol, plan['dm_req'], plan['total_animals'], plan['n_days'], plan['entered_prices'], plan.get('entered_bag_sizes') or {})}"
        f"{_print_feed_mix_table(sol, plan['dm_req'], plan['total_animals'], plan['n_days'], price_records, plan.get('entered_bag_sizes') or {}, plan['entered_prices'])}"
        "</section>"
    )


def build_print_report_html(plans, price_records):
    if not plans:
        return ""

    sections = [_print_farm_summary_page(plans, price_records)]
    sections.extend(
        _print_plan_section(plan, index, price_records)
        for index, plan in enumerate(plans, start=1)
    )

    return (
        "<!doctype html><html lang='ar' dir='rtl'><head><meta charset='utf-8'>"
        "<title>تقرير توصيات الاعلاف </title>"
        "<style>"
        "@page{size:A4 portrait;margin:8mm;}"
        "*{box-sizing:border-box;}"
        "body{margin:0;font-family:Arial,Tahoma,sans-serif;color:#111;direction:rtl;background:#FFF;}"
        ".print-page{width:100%;min-height:100%;padding:0;margin:0;page-break-after:always;}"
        ".print-logo{max-height:18mm;max-width:48mm;object-fit:contain;margin-top:-3.2mm;}"
        ".print-page-header{display:flex;justify-content:flex-end;align-items:flex-start;margin:9mm 0 4mm;}"
        ".print-logo{max-height:18mm;max-width:48mm;object-fit:contain;margin-top:-3.2mm;}"
        ".print-title-block{border-bottom:2px solid #6B4CA0;margin:0 0 3mm;padding-top:4.5mm;padding-bottom:2.2mm;}"
        ".print-title-block h1{font-size:13pt;color:#2C3149;margin:0;font-weight:900;}"
        ".print-page-subtitle{font-size:7.5pt;color:#666;margin-top:2.2mm;}"
        ".print-summary-total{margin:0 0 4mm;padding:2.2mm 2.6mm;border:1px solid #D8D8D8;background:#F8F6FC;font-size:8.5pt;font-weight:700;color:#2C3149;}"
        ".print-table-card{margin:0 0 4mm;}"
        ".print-table-card h2{font-size:9pt;color:#6B4CA0;margin:0 0 1.5mm;font-weight:900;}"
        ".print-pill-wrap{display:flex;flex-wrap:wrap;gap:1.5mm;}"
        ".print-group-pill{display:inline-block;padding:0.8mm 2.2mm;border-radius:999px;font-size:7pt;font-weight:800;white-space:nowrap;}"
        ".print-group-pill.pill-a{background:#E8F7ED;color:#166534;border:1px solid #B7E4C4;}"
        ".print-group-pill.pill-b{background:#EAF2FF;color:#1D4ED8;border:1px solid #BBD3FF;}"
        ".print-group-pill.pill-c{background:#FFF4E5;color:#9A3412;border:1px solid #FED7AA;}"
        ".print-group-pill.pill-d{background:#F3E8FF;color:#6B21A8;border:1px solid #D8B4FE;}"
        ".print-group-pill.pill-e{background:#ECFEFF;color:#155E75;border:1px solid #A5F3FC;}"
        ".print-info-table,.print-merged-table{width:100%;border-collapse:collapse;font-size:8pt;}"
        ".print-info-table th,.print-info-table td,.print-merged-table th,.print-merged-table td{border:1px solid #D8D8D8;padding:1.8mm;text-align:right;vertical-align:middle;}"
        ".print-info-table th,.print-merged-table th{background:#F1F1F1;font-weight:900;width:34%;}"
        ".print-merged-table thead th{background:#EDE7F6;color:#2C3149;width:auto;}"
        "</style></head><body>"
        f"{''.join(sections)}"
        "</body></html>"
    )


def render_print_plan_selector(selected_plans, price_records):
    if not selected_plans:
        st.info("حدد خطة واحدة على الأقل من المربعات بجانب البطاقات للطباعة.")
        return

    report_html = build_print_report_html(selected_plans, price_records)
    report_json = json.dumps(report_html)
    components.html(
        f"""
        <button id="print-selected-plans" style="
          width:100%;border:1px solid #6B4CA0;border-radius:10px;
          background:#6B4CA0;color:white;padding:0.65rem 1rem;
          font-weight:800;cursor:pointer;margin:4px 0 10px;
          font-family:inherit;font-size:15px;">
          🖨️ طباعة الخطط المحددة
        </button>
        <script>
          const reportHtml = {report_json};
          document.getElementById('print-selected-plans').addEventListener('click', () => {{
            const printWindow = window.open('', '_blank', 'width=1200,height=800');
            if (!printWindow) {{
              alert('يرجى السماح بالنوافذ المنبثقة للطباعة.');
              return;
            }}
            printWindow.document.open();
            printWindow.document.write(reportHtml);
            printWindow.document.close();
            printWindow.focus();
            setTimeout(() => {{ printWindow.print(); }}, 300);
          }});
        </script>
        """,
        height=58,
    )


def render_unsuitable_feed_warning(group, sol):
    species_tag = html.escape(str(group.get("species") or "غير محدد"))
    stage_tag = html.escape(str(group.get("stage") or "غير محدد"))
    rows_html = ""

    for _, row in sol["analysis_df"].iterrows():
        nutrient = str(row["العنصر الغذائي"]).replace(" %", "").replace(" (kcal/kg)", "")
        pct = float(row["نسبة التغطية %"])
        achieved = row["المتحقق من الخلطة"]
        requirement = row["الاحتياج"]
        bar_width = min(100, max(0, (pct / 150) * 100))

        if pct < 95:
            note_tag = "↓ أقل من الحد المقبول"
            fill = "#6B7280"
            tag_bg = "#F3F4F6"
            tag_border = "#D1D5DB"
            tag_color = "#374151"
            detail_color = "#4B5563"
        elif pct >= 121:
            note_tag = "↑ أعلى من الحد المقبول"
            fill = "#DC2626"
            tag_bg = "#FEE2E2"
            tag_border = "#FCA5A5"
            tag_color = "#991B1B"
            detail_color = "#991B1B"
        elif pct > 105:
            note_tag = "↑ مرتفع قليلاً"
            fill = "#DC2626"
            tag_bg = "#FEE2E2"
            tag_border = "#FCA5A5"
            tag_color = "#991B1B"
            detail_color = "#991B1B"
        else:
            note_tag = "✓ ضمن الحد المناسب"
            fill = "#16A34A"
            tag_bg = "#ECFDF5"
            tag_border = "#86EFAC"
            tag_color = "#166534"
            detail_color = "#166534"

        rows_html += (
            "<div style='margin-top:8px;'>"
            "<div style='display:flex;justify-content:space-between;gap:8px;align-items:center;'>"
            f"<b style='font-size:0.72rem;'>{html.escape(nutrient)}</b>"
            f"<span style='display:inline-block;background:{tag_bg};border:1px solid {tag_border};"
            f"color:{tag_color};border-radius:999px;padding:1px 7px;font-size:0.60rem;font-weight:900;'>{format_display_num(pct)}% {note_tag}</span>"
            "</div>"
            "<div style='height:6px;background:#FEF3C7;border-radius:999px;overflow:hidden;margin-top:3px;'>"
            f"<div style='height:100%;width:{bar_width:.1f}%;background:{fill};border-radius:999px;'></div>"
            "</div>"
            f"<div style='display:flex;justify-content:space-between;color:{detail_color};font-size:0.66rem;margin-top:3px;'>"
            f"<span>المتحقق: {format_display_num(achieved)}</span>"
            f"<span>الاحتياج: {format_display_num(requirement)}</span>"
            "</div>"
            "</div>"
        )

    warning_html = (
        "<div style='background:#FFFBEB;border:1.5px solid #F5C842;"
        "border-radius:10px;padding:10px 12px;margin:10px 0;color:#7A5500;direction:rtl;'>"
        "<div style='font-weight:900;font-size:0.82rem;margin-bottom:7px;'>"
        "لم تحقق الأعلاف المختارة القيم الغذائية المناسبة للقطيع:"
        "</div>"
        "<div style='display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px;'>"
        f"<span style='display:inline-block;background:#FFFFFF;border:1px solid #F5C842;"
        f"border-radius:999px;padding:2px 8px;font-size:0.68rem;font-weight:800;'>{species_tag}</span>"
        f"<span style='display:inline-block;background:#FFFFFF;border:1px solid #F5C842;"
        f"border-radius:999px;padding:2px 8px;font-size:0.68rem;font-weight:800;'>{stage_tag}</span>"
        "</div>"
        "<div style='font-weight:900;font-size:0.76rem;margin-bottom:4px;'>السبب:</div>"
        f"{rows_html}"
        "</div>"
    )
    st.markdown(warning_html, unsafe_allow_html=True)


def show_grouped_recommendations(group_payloads, n_days, price_records):
    if not group_payloads:
        st.error("أضف قطيع مكتمل واحد على الأقل.")
        return

    cards_by_group = []
    selected_print_plans = []
    print_plan_counter = 0

    for payload_index, payload in enumerate(group_payloads, start=1):
        group = payload["group"]
        recs  = payload["recommendations"]
        group_key = group.get("id") or f"group_{payload_index}"
        label = (f"{group['species']} | {group['stage']} | "
                 f"وزن: {group.get('weight') or 'غير محدد'} | عدد: {int(group.get('n',1))}")
        if not recs:
            st.error(f"لا توجد توصية ممكنة للقطيع: {label}. تحقق من بيانات الأعلاف والأسعار المدخلة.")
            continue
        if recs and recs[0].get("is_fallback"):
            print_plan_counter += 1
            warning_plan = {
                "id": f"warning_{group_key}",
                "kind": "warning",
                "group": group.copy(),
                "sol": recs[0],
                "dm_req": payload["dm_req"],
                "n_days": n_days,
                "total_animals": int(group.get("n", 1)),
                "entered_prices": payload["entered_prices"],
                "entered_bag_sizes": payload["entered_bag_sizes"],
            }
            selector_col, card_col = st.columns([0.12, 0.88])
            with selector_col:
                if st.checkbox(" طباعة", value=False, key=f"print_select_{warning_plan['id']}"):
                    selected_print_plans.append(warning_plan)
            with card_col:
                render_unsuitable_feed_warning(group, recs[0])
            continue
        cards = [
            {"rank": rank, "sol": sol, "dm_req": payload["dm_req"],
             "n_days": n_days, "total_animals": int(group.get("n",1)),
             "group_labels": [label], "entered_prices": payload["entered_prices"],
             "entered_bag_sizes": payload["entered_bag_sizes"],
             "group": group.copy()}
            for rank, sol in enumerate(recs, start=1)
        ]
        cards_by_group.append((label, group_key, cards))

    if not cards_by_group and not selected_print_plans:
        return

    for group_label, group_key, cards in cards_by_group:
        st.markdown(
            f"<div class='rec-grid-header'>🐾 {group_label}</div>",
            unsafe_allow_html=True,
        )
        if any(card["sol"].get("is_fallback") for card in cards):
            ordered = sorted(
                cards,
                key=lambda c: (
                    c["sol"].get("fit_penalty", 0),
                    c["sol"]["cost_per_kg_mix"],
                    c["rank"],
                ),
            )
        else:
            ordered = sorted(cards, key=lambda c: recommendation_sort_key(c["sol"]))
        top3    = ordered[:3]

        for i, card in enumerate(top3):
            print_plan_counter += 1
            plan = {
                "id": f"recommendation_{group_key}_{card['rank']}",
                "kind": "recommendation",
                "group": card["group"],
                "sol": card["sol"],
                "dm_req": card["dm_req"],
                "n_days": card["n_days"],
                "total_animals": card["total_animals"],
                "entered_prices": card["entered_prices"],
                "entered_bag_sizes": card["entered_bag_sizes"],
            }
            selector_col, card_col = st.columns([0.12, 0.88])
            with selector_col:
                if st.checkbox(" طباعة", value=(i == 0), key=f"print_select_{plan['id']}"):
                    selected_print_plans.append(plan)
            with card_col:
                render_recommendation_card_wide(card, price_records,
                                                card["entered_prices"], card["entered_bag_sizes"], is_best=(i == 0))

        if len(ordered) > 3:
            with st.expander(" عرض المزيد من الخيارات", expanded=False):
                for i, card in enumerate(ordered[3:]):
                    print_plan_counter += 1
                    plan = {
                        "id": f"recommendation_more_{group_key}_{card['rank']}",
                        "kind": "recommendation",
                        "group": card["group"],
                        "sol": card["sol"],
                        "dm_req": card["dm_req"],
                        "n_days": card["n_days"],
                        "total_animals": card["total_animals"],
                        "entered_prices": card["entered_prices"],
                        "entered_bag_sizes": card["entered_bag_sizes"],
                    }
                    selector_col, card_col = st.columns([0.12, 0.88])
                    with selector_col:
                        if st.checkbox(" طباعة", value=False, key=f"print_select_{plan['id']}"):
                            selected_print_plans.append(plan)
                    with card_col:
                        render_recommendation_card_wide(card, price_records,
                                                        card["entered_prices"], card["entered_bag_sizes"], is_best=False)

    render_print_plan_selector(selected_print_plans, price_records)


# =========================
# Market price status banner
# =========================
if market_price_error:
    st.warning(market_price_error)
elif market_prices:
    st.markdown(
        f"<span class='status-ok-badge'>"
        f"✦ تم تحميل {len(market_prices)} سعر تلقائياً من مؤشر سنابل الجوف"
        f"</span>",
        unsafe_allow_html=True,
    )

# =========================
# Build feed price lists
# =========================
feed_prices, price_records, missing_price_feeds = build_auto_feed_prices(feed_names)
available_feeds = list(feed_names)
if not available_feeds:
    st.error("لا توجد أعلاف في بنك الأعلاف. لا يمكن توليد توصيات.")
    st.stop()

# =========================
# Main UI flow - V5
# =========================
selected_animal_types = render_animal_selector()
if not selected_animal_types:
    st.stop()

herd_days = render_group_editor(selected_animal_types)

editable_feed_prices = render_feed_selection_and_pricing(feed_names, price_records)
selected_candidate_feeds = st.session_state.selected_feed_chips

if len(selected_candidate_feeds) < 2:
    st.info("اختر نوعين على الأقل من الأعلاف للانتقال إلى تفاصيل القطيع.")
    st.stop()

st.markdown("<div class='cta-btn'>", unsafe_allow_html=True)
generate = st.button("✦  توليد التوصيات", key="generate_recommendations", use_container_width=True)
st.markdown("</div>", unsafe_allow_html=True)

if generate:
    if len(selected_candidate_feeds) < 2:
        st.error("اختر نوعين على الأقل من الأعلاف لتوليد خلطة متعددة الأعلاف.")
        st.stop()

    optimizer_prices = build_optimizer_prices(
        selected_candidate_feeds, editable_feed_prices, feed_prices
    )
    valid_groups = [g for g in st.session_state.herd_groups if group_is_complete(g)]
    if not valid_groups:
        st.error("أكمل بيانات فئة حيوان واحدة على الأقل، بما في ذلك الوزن إذا ظهر في النموذج.")
        st.stop()

    group_payloads = []
    for g in valid_groups:
        req, err = get_requirement_row(g["species"], g["stage"], g.get("weight"))
        if err:
            st.error(f"{g.get('species') or ''}: {err}")
            continue
        recs = generate_recommendations(
            req["cp_req"], req["en_req"], req["fi_req"],
            selected_candidate_feeds, optimizer_prices,
        )
        group_payloads.append({
            "group":          g.copy(),
            "dm_req":         req["dm_req"],
            "recommendations": recs,
            "entered_prices": optimizer_prices.copy(),
            "entered_bag_sizes": st.session_state.editable_feed_bag_sizes.copy(),
        })

    st.session_state.last_group_payloads = group_payloads
    st.session_state.last_recommendation_days = int(herd_days)

if st.session_state.last_group_payloads:
    show_grouped_recommendations(
        st.session_state.last_group_payloads,
        int(st.session_state.last_recommendation_days or herd_days),
        price_records,
    )

