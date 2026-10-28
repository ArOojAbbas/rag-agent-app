import io
import os
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st
from docx import Document as DocxDocument
from docx.table import Table
from docx.text.paragraph import Paragraph
from langchain_community.document_loaders import PyPDFLoader
from langchain_community.vectorstores import FAISS
from langchain_core.documents import Document
from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.runnables import RunnableLambda, RunnablePassthrough
from langchain_openai import ChatOpenAI, OpenAIEmbeddings
from langchain_text_splitters import RecursiveCharacterTextSplitter
from openai import (
    APIConnectionError,
    AuthenticationError,
    OpenAI,
    PermissionDeniedError,
    RateLimitError,
)

# ----------------------------------------------------------------
# Page setup
# ----------------------------------------------------------------
st.set_page_config(page_title="RAG Agent", page_icon=":material/description:", layout="wide", initial_sidebar_state="expanded")


def html(s: str):
    """Render HTML safely: strip indentation so markdown never treats it as a code block."""
    st.markdown(" ".join(l.strip() for l in s.splitlines() if l.strip()), unsafe_allow_html=True)


LOGO_SVG = """
<svg width="{s}" height="{s}" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
<defs>
<linearGradient id="bg{s}" x1="0" y1="0" x2="1" y2="1">
<stop offset="0" stop-color="#6366F1"/><stop offset=".55" stop-color="#8B5CF6"/><stop offset="1" stop-color="#06B6D4"/>
</linearGradient>
<linearGradient id="dc{s}" x1="0" y1="0" x2="0" y2="1">
<stop offset="0" stop-color="#FFFFFF"/><stop offset="1" stop-color="#E0E7FF"/>
</linearGradient>
</defs>
<rect width="64" height="64" rx="18" fill="url(#bg{s})"/>
<rect x="2" y="2" width="60" height="30" rx="16" fill="#fff" opacity=".13"/>
<path d="M19 11h18l11 11v28a3 3 0 0 1-3 3H19a3 3 0 0 1-3-3V14a3 3 0 0 1 3-3z" fill="url(#dc{s})"/>
<path d="M37 11v8a3 3 0 0 0 3 3h8z" fill="#A5B4FC"/>
<line x1="25" y1="31" x2="37" y2="28" stroke="#6366F1" stroke-width="1.7" stroke-linecap="round"/>
<line x1="37" y1="28" x2="31" y2="42" stroke="#6366F1" stroke-width="1.7" stroke-linecap="round"/>
<line x1="25" y1="31" x2="31" y2="42" stroke="#6366F1" stroke-width="1.7" stroke-linecap="round"/>
<circle class="node n1" cx="25" cy="31" r="3.2" fill="#6366F1"/>
<circle class="node n2" cx="37" cy="28" r="3.2" fill="#06B6D4"/>
<circle class="node n3" cx="31" cy="42" r="3.2" fill="#8B5CF6"/>
<circle cx="48" cy="49" r="10" fill="#0B1020" stroke="#fff" stroke-opacity=".35" stroke-width="1.2"/>
<path class="spark" d="M48 42l2 5 5 2-5 2-2 5-2-5-5-2 5-2z" fill="#67E8F9"/>
</svg>
"""


def logo(size: int) -> str:
    return f'<span class="logo">{LOGO_SVG.format(s=size)}</span>'


st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Playfair+Display:wght@600;700&display=swap');
html, body, [class*="css"], .stApp, .stApp p, .stApp label, .stApp input, .stApp textarea, .stApp button,
.stApp li, .stApp a, .stApp div, .stApp h1, .stApp h2, .stApp h3, .stApp h4 { font-family: 'DM Sans', sans-serif !important; }
.stApp .brand-name, .stApp .hero h1, .stApp .login-head .name, .stApp .headline { font-family: 'Playfair Display', serif !important; }

/* ---------- Base: soft light ---------- */
.stApp { background: linear-gradient(135deg,#FBF8F3 0%,#F5F0FA 55%,#EEF5F4 100%); color:#1F1B2E; }
.stApp::before {
    content:""; position:fixed; inset:0; z-index:0; pointer-events:none;
    background:
      radial-gradient(520px 420px at 12% 12%, rgba(196,181,253,.45), transparent 70%),
      radial-gradient(460px 400px at 88% 18%, rgba(253,186,160,.38), transparent 70%),
      radial-gradient(520px 420px at 70% 92%, rgba(167,226,206,.40), transparent 70%);
}
header[data-testid="stHeader"] { background: transparent; }
#MainMenu, footer, [data-testid="stAppDeployButton"], [data-testid="stMainMenu"], [data-testid="stDecoration"] { visibility: hidden; }
[data-testid="stToolbar"] { background: transparent; }

/* ---------- Sidebar show / hide buttons: always visible and styled ---------- */
[data-testid="stExpandSidebarButton"], [data-testid="stSidebarCollapsedControl"], [data-testid="collapsedControl"] {
    visibility: visible !important; opacity: 1 !important; z-index: 999991;
}
[data-testid="stExpandSidebarButton"], [data-testid="stSidebarCollapsedControl"] button, [data-testid="collapsedControl"] button {
    background:#FFFFFF !important; border:1px solid #E1DBEE !important; border-radius:12px !important;
    box-shadow:0 6px 18px rgba(80,60,140,.14);
}
[data-testid="stSidebarCollapseButton"], [data-testid="stSidebarCollapseButton"] button { opacity:1 !important; visibility:visible !important; }
[data-testid="stExpandSidebarButton"] *, [data-testid="stSidebarCollapsedControl"] *, [data-testid="collapsedControl"] *, [data-testid="stSidebarCollapseButton"] * { color:#1F1B2E !important; }
.block-container { position:relative; z-index:1; padding-top: 2.6rem; max-width: 1040px; }

/* ---------- Text colors ---------- */
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5 { color:#1F1B2E; }
.stApp p, .stApp li, [data-testid="stMarkdownContainer"] { color:#1F1B2E; }
.stApp label, .stApp [data-testid="stWidgetLabel"] p, .stApp [data-testid="stWidgetLabel"] span { color:#4B4760 !important; font-weight:500; }
[data-testid="stCaptionContainer"], .stApp small { color:#7A7690 !important; }

/* ---------- Animations ---------- */
@keyframes fadeUp { from { opacity:0; transform: translateY(16px); } to { opacity:1; transform:none; } }
@keyframes drift { to { transform: translate3d(2.5%,3%,0) scale(1.06); } }
.stApp::before { inset:-10%; animation: drift 22s ease-in-out infinite alternate; }

/* ---------- Login: left showcase ---------- */
.headline { font-size:2.8rem; line-height:1.08; font-weight:700; color:#1F1B2E; margin:0 0 1rem; letter-spacing:-.01em; }
.headline em { font-style:italic; background:linear-gradient(90deg,#6D5AE6,#C084FC,#F59E7B); -webkit-background-clip:text; -webkit-text-fill-color:transparent; padding-right:.12em; }
.lead { color:#6B6880; font-size:1.02rem; line-height:1.7; max-width:470px; margin:0 0 1.8rem; }
.mock { background:#FFFFFF; border:1px solid #E7E3EC; border-radius:20px; padding:18px 18px 16px; max-width:460px; box-shadow:0 18px 44px rgba(80,60,140,.12); }
.mock .dots { display:flex; gap:6px; margin-bottom:14px; }
.mock .dots i { width:9px; height:9px; border-radius:50%; background:#E7E3EC; display:block; }
.bubble-u { margin-left:auto; width:fit-content; max-width:82%; background:#1F1B2E; color:#FFFFFF; padding:10px 15px; border-radius:16px 16px 4px 16px; font-size:.88rem; }
.bubble-a { margin-top:10px; max-width:92%; background:#F6F3FB; color:#1F1B2E; padding:12px 15px; border-radius:16px 16px 16px 4px; font-size:.88rem; line-height:1.55; }
.cite { display:inline-flex; align-items:center; gap:6px; margin:10px 6px 0 0; font-size:.72rem; padding:4px 10px; border-radius:999px; background:#FFFFFF; border:1px solid #E1DBEE; color:#4B4760; }
.cite b { font-size:.58rem; letter-spacing:.08em; background:#F0ECF7; padding:1px 5px; border-radius:4px; }
.feats { display:flex; gap:22px; margin-top:1.9rem; flex-wrap:wrap; }
.feat { display:flex; align-items:center; gap:8px; font-size:.85rem; color:#4B4760; font-weight:500; }
.feat svg { color:#6D5AE6; flex-shrink:0; }
.err { display:flex; align-items:flex-start; gap:10px; background:#FDEEEE; border:1px solid #F4C9C9; border-radius:12px; padding:11px 14px; margin:14px 0 4px; }
.err svg { color:#C53030; flex-shrink:0; margin-top:2px; }
.err b { display:block; color:#9B1C1C; font-size:.9rem; font-weight:700; }
.err span { color:#B45454; font-size:.8rem; }

/* ---------- Logo ---------- */
.logo { display:inline-block; line-height:0; filter: drop-shadow(0 14px 28px rgba(99,102,241,.32)); }

/* ---------- Sidebar ---------- */
[data-testid="stSidebar"] { background:rgba(255,255,255,.72); backdrop-filter: blur(10px); border-right:1px solid #E7E3EC; }
[data-testid="stSidebarUserContent"] { padding-top: 1.2rem; }
.brand { display:flex; align-items:center; gap:12px; margin-bottom:4px; }
.brand-name { font-family:'Playfair Display',serif; font-size:1.4rem; font-weight:700; color:#1F1B2E; }
.brand-sub { font-size:.76rem; color:#7A7690; margin-bottom:16px; }
.side-title { font-size:.68rem; letter-spacing:.14em; text-transform:uppercase; color:#9A96AE; font-weight:700; margin:20px 0 6px; }
.status { display:flex; align-items:center; gap:8px; font-size:.8rem; color:#6B6880; margin:14px 0 2px; }
.status i { width:8px; height:8px; border-radius:50%; background:#34D399; display:inline-block; }

/* ---------- Hero ---------- */
.hero { text-align:center; padding:.4rem 0 .2rem; }
.hero h1 { font-family:'Playfair Display',serif; font-size:3.2rem; font-weight:700; letter-spacing:-.01em; margin:1rem 0 .5rem; line-height:1.1; color:#1F1B2E; }
.hero p { color:#6B6880 !important; font-size:1.05rem; line-height:1.65; margin:0 auto; max-width:620px; }

/* ---------- Cards ---------- */
.cards { display:grid; grid-template-columns:repeat(3,1fr); gap:16px; margin:1.6rem 0; }
.card { background:rgba(255,255,255,.8); border:1px solid #E7E3EC; border-radius:16px; padding:22px; transition:.2s; box-shadow:0 6px 24px rgba(80,60,140,.05); }
.card:hover { border-color:#CFC8DD; box-shadow:0 12px 32px rgba(80,60,140,.10); }
.card .num { font-size:.78rem; font-weight:700; letter-spacing:.12em; color:#9A96AE; }
.card h4 { margin:.7rem 0 .35rem; font-size:1.02rem; font-weight:600; color:#1F1B2E; }
.card span { color:#6B6880; font-size:.88rem; line-height:1.55; }

/* ---------- Login page ---------- */
.login-head { display:flex; align-items:center; justify-content:center; gap:14px; margin-bottom:.6rem; }
.login-head .name { font-family:'Playfair Display',serif; font-size:2.15rem; font-weight:700; color:#1F1B2E; line-height:1; letter-spacing:-.01em; }
.login-title { text-align:center; font-family:'Playfair Display',serif; font-size:1.55rem; font-weight:700; color:#1F1B2E; margin:0 0 .35rem; }
.login-sub { text-align:center; color:#6B6880; font-size:.93rem; line-height:1.55; margin:0 0 1.5rem; }
.note { display:flex; justify-content:center; align-items:center; gap:8px; color:#7A7690; font-size:.76rem; white-space:nowrap; margin-top:1.4rem; }
.note svg { flex-shrink:0; display:block; }
.links { text-align:center; font-size:.82rem; color:#7A7690; margin-top:.6rem; }
.links a { color:#1F1B2E; font-weight:600; text-decoration:none; border-bottom:1px solid #CFC8DD; }
.works { display:flex; justify-content:center; align-items:center; gap:8px; margin-top:2rem; flex-wrap:wrap; }
.works span { font-size:.72rem; color:#9A96AE; letter-spacing:.12em; text-transform:uppercase; margin-right:4px; }
.works b { font-size:.74rem; font-weight:600; color:#4B4760; padding:5px 12px; border-radius:999px; background:rgba(255,255,255,.75); border:1px solid #E7E3EC; }
[data-testid="stForm"] {
    background:rgba(255,255,255,.82); backdrop-filter: blur(14px); border:1px solid rgba(255,255,255,.95);
    border-radius:24px; padding:34px 34px 26px; animation: fadeUp .7s ease both; box-shadow:0 30px 80px rgba(80,60,140,.14), 0 2px 6px rgba(80,60,140,.06);
}

/* ---------- Inputs ---------- */
div[data-baseweb="input"], div[data-baseweb="base-input"] { background:#FFFFFF !important; border-radius:12px !important; border:1px solid #DDD7E8 !important; }
div[data-baseweb="input"]:focus-within { border-color:#1F1B2E !important; box-shadow:0 0 0 3px rgba(31,27,46,.08) !important; }
.stTextInput input { color:#1F1B2E !important; background:transparent !important; height:30px; }
.stTextInput input::placeholder { color:#A8A4BA !important; }
div[data-baseweb="select"] > div { background:#FFFFFF !important; color:#1F1B2E !important; border:1px solid #DDD7E8 !important; border-radius:12px !important; }
div[data-baseweb="select"] * { color:#1F1B2E !important; }
div[data-baseweb="popover"] ul, div[data-baseweb="popover"] li { background:#FFFFFF !important; color:#1F1B2E !important; }
div[data-baseweb="popover"] li:hover { background:#F3EFF9 !important; }
[data-testid="stSlider"] [data-baseweb="slider"] { filter: grayscale(1) brightness(.3); }
[data-testid="stSliderThumbValue"], [data-testid="stTickBarMin"], [data-testid="stTickBarMax"] { color:#4B4760 !important; }

/* ---------- Buttons ---------- */
.stButton > button, [data-testid="stFileUploader"] button {
    background:#FFFFFF !important; color:#1F1B2E !important; border:1px solid #DDD7E8 !important;
    border-radius:12px !important; font-weight:500 !important; padding:.5rem 1rem !important; transition:.15s; width:100%;
}
.stButton > button:hover, [data-testid="stFileUploader"] button:hover { border-color:#1F1B2E !important; background:#FAF8FD !important; }
[data-testid="stFileUploader"] button { width:auto; }
[data-testid="stFormSubmitButton"] > button {
    width:auto; min-width:210px; height:48px; padding:0 2.4rem !important; background:#1F1B2E !important; color:#FFFFFF !important; border:0 !important;
    border-radius:12px !important; font-weight:600 !important; font-size:.98rem; transition:.2s;
}
[data-testid="stFormSubmitButton"] > button:hover { background:#342F4D !important; box-shadow:0 10px 26px rgba(31,27,46,.25); transform:translateY(-1px); }
.stButton > button p, [data-testid="stFileUploader"] button span { color:#1F1B2E !important; }
[data-testid="stFormSubmitButton"] > button p { color:#FFFFFF !important; }

[data-testid="stFormSubmitButton"] { width:100% !important; display:flex !important; justify-content:center !important; margin-top:.4rem; }
[data-testid="stFormSubmitButton"] > button { margin:0 auto !important; }
[data-testid="stElementContainer"]:has([data-testid="stFormSubmitButton"]), .stElementContainer:has([data-testid="stFormSubmitButton"]) { display:flex; justify-content:center; width:100% !important; }

/* ---------- Uploader ---------- */
[data-testid="stFileUploader"] section { background:rgba(255,255,255,.7); border:1.5px dashed #CFC8DD; border-radius:16px; padding:1.5rem; transition:.2s; }
[data-testid="stFileUploader"] section:hover { border-color:#1F1B2E; background:rgba(255,255,255,.95); }
[data-testid="stFileUploaderDropzoneInstructions"] * { color:#7A7690 !important; }
[data-testid="stFileUploaderFile"] * { color:#1F1B2E !important; }

/* ---------- Alerts ---------- */
[data-testid="stAlert"] { background:rgba(255,255,255,.9) !important; border:1px solid #E7E3EC; border-radius:12px; }
[data-testid="stAlert"] * { color:#1F1B2E !important; }

/* ---------- Chat ---------- */
[data-testid="stChatMessage"] { background:rgba(255,255,255,.85); border:1px solid #E7E3EC; border-radius:16px; padding:14px 18px; margin-bottom:10px; }
[data-testid="stChatMessage"] * { color:#1F1B2E; }
[data-testid="stBottom"], [data-testid="stBottom"] > div { background: transparent !important; }
[data-testid="stChatInput"] { background:#FFFFFF !important; border:1px solid #D9D3E5 !important; border-radius:16px !important; box-shadow:0 8px 28px rgba(80,60,140,.08); }
[data-testid="stChatInput"]:focus-within { border-color:#1F1B2E !important; }
[data-testid="stChatInput"] textarea { color:#1F1B2E !important; background:transparent !important; }
[data-testid="stChatInput"] textarea::placeholder { color:#A8A4BA !important; }
[data-testid="stExpander"] { background:rgba(255,255,255,.8); border:1px solid #E7E3EC; border-radius:12px; }
[data-testid="stExpander"] * { color:#1F1B2E; }
hr { border-color:#E7E3EC !important; }

/* ---------- File chips + stats ---------- */
.chip { display:inline-flex; align-items:center; gap:8px; padding:6px 12px; margin:3px 6px 3px 0; border-radius:10px; font-size:.82rem; color:#1F1B2E; background:#FFFFFF; border:1px solid #E7E3EC; }
.chip i { font-style:normal; font-size:.66rem; font-weight:700; letter-spacing:.08em; padding:2px 6px; border-radius:5px; background:#F0ECF7; color:#4B4760; }
.stat { background:rgba(255,255,255,.85); border:1px solid #E7E3EC; border-radius:14px; padding:14px; text-align:center; color:#7A7690; font-size:.8rem; }
.stat b { display:block; font-size:1.6rem; font-weight:700; color:#1F1B2E; margin-bottom:2px; }
@media (max-width: 800px) { .cards { grid-template-columns:1fr; } .hero h1 { font-size:2.3rem; } }
</style>
""",
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------
# API key verification (real check against OpenAI)
# ----------------------------------------------------------------
def verify_api_key(key: str):
    """Return (is_valid, title, hint). Makes a lightweight authenticated call to OpenAI."""
    if not key.startswith("sk-") or len(key) < 20:
        return False, "Invalid API key", "Please check the key and try again."
    try:
        OpenAI(api_key=key, timeout=20).models.list()
        return True, "", ""
    except AuthenticationError:
        return False, "Invalid API key", "OpenAI didn't accept this key. Please check it and try again."
    except PermissionDeniedError:
        return True, "", ""  # restricted key: it authenticated, it just can't list models
    except RateLimitError:
        return True, "", ""  # key is valid; quota issues will be reported when you chat
    except APIConnectionError:
        return False, "Connection problem", "Couldn't reach OpenAI. Check your internet and try again."
    except Exception as e:
        return False, "Couldn't verify the key", str(e)[:120]


# ----------------------------------------------------------------
# Login screen (API key is asked every time the app is opened)
# ----------------------------------------------------------------
if "api_key" not in st.session_state:
    st.session_state.api_key = ""

if not st.session_state.api_key:
    # Login page only: wider layout, vertically centered columns
    st.markdown(
        "<style>.block-container{max-width:1180px !important; padding-top:5vh !important;}"
        '[data-testid="stHorizontalBlock"]{align-items:center !important; gap:3rem !important; flex-wrap:nowrap !important;}'
        '[data-testid="stColumn"], [data-testid="column"]{min-width:0 !important;}'
        '@media (max-width:900px){[data-testid="stHorizontalBlock"]{flex-wrap:wrap !important;}}'
        # Key box is a normal text box shown as dots, so Chrome's "Suggest strong password" popup never appears
        '[data-testid="stForm"] input{-webkit-text-security:disc !important; text-security:disc !important;}</style>',
        unsafe_allow_html=True,
    )
    left, right = st.columns([1.15, 1])

    with left:
        html("""
        <div class="headline">Ask your documents <em>anything.</em></div>
        <div class="lead">Upload PDF, Word and Excel files and get answers grounded in your own content, with the exact page or sheet cited.</div>
        <div class="mock">
          <div class="dots"><i></i><i></i><i></i></div>
          <div class="bubble-u">Which region grew fastest in Q3?</div>
          <div class="bubble-a">APAC grew fastest at 18%, ahead of EMEA at 11% and the Americas at 7%.
            <div><span class="cite"><b>XLSX</b>sales_2026.xlsx · Q3</span><span class="cite"><b>PDF</b>report.pdf · p. 12</span></div>
          </div>
        </div>
        <div class="feats">
          <div class="feat"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12.5l3 3 5-6"/></svg>Cited answers</div>
          <div class="feat"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/></svg>PDF, Word &amp; Excel</div>
          <div class="feat"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l8 3v6c0 4.5-3.2 7.8-8 9-4.8-1.2-8-4.5-8-9V6z"/></svg>Private by design</div>
        </div>
        """)

    with right:
        with st.form("login"):
            html(f"""
            <div class="login-head">{logo(52)}<div class="name">RAG Agent</div></div>
            <div class="login-sub">Enter your OpenAI API key to continue.</div>
            """)
            try:
                key_in = st.text_input("OpenAI API key", placeholder="sk-...", autocomplete="off")
            except TypeError:  # older Streamlit without the autocomplete argument
                key_in = st.text_input("OpenAI API key", placeholder="sk-...")
            err = st.empty()
            go = st.form_submit_button("Continue")
            html("""
            <div class="note">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg>
              <span>Your key is used only for this session and is never stored.</span>
            </div>
            """)

        def show_err(title, hint):
            err.markdown(
                '<div class="err"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
                'stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/>'
                f'<path d="M12 8v5M12 16.5v.01"/></svg><div><b>{title}</b><span>{hint}</span></div></div>',
                unsafe_allow_html=True,
            )

        if go:
            k = key_in.strip()
            if not k:
                show_err("API key required", "Please paste your OpenAI API key.")
            else:
                with st.spinner("Verifying your API key..."):
                    ok, title, hint = verify_api_key(k)
                if ok:
                    st.session_state.api_key = k
                    st.rerun()
                else:
                    show_err(title, hint)
        html('<div class="links">Don\'t have a key? <a href="https://platform.openai.com/api-keys" target="_blank">Create one on OpenAI</a></div>')
    st.stop()

api_key = st.session_state.api_key
os.environ["OPENAI_API_KEY"] = api_key

# ----------------------------------------------------------------
# Styles used only after login (login page above is not affected)
# ----------------------------------------------------------------
st.markdown(
    """
<style>
/* Room at the bottom so the chat box never covers the content */
.block-container { padding-bottom: 9rem !important; }
[data-testid="stBottom"] { background: linear-gradient(to top, #F5F0FA 58%, rgba(245,240,250,0)) !important; }
[data-testid="stBottom"] > div { background: transparent !important; }

/* Sidebar a little wider */
[data-testid="stSidebar"][aria-expanded="true"] { min-width: 316px !important; max-width: 316px !important; }

/* Question box: clearly visible */
[data-testid="stChatInput"] {
    background:#FFFFFF !important; border:1.5px solid #8E87A8 !important; border-radius:18px !important;
    box-shadow:0 12px 32px rgba(31,27,46,.16), 0 2px 6px rgba(31,27,46,.08) !important;
}
[data-testid="stChatInput"]:focus-within { border-color:#1F1B2E !important; box-shadow:0 14px 36px rgba(31,27,46,.22) !important; }
[data-testid="stChatInput"] div[data-baseweb] { background:#FFFFFF !important; border:0 !important; box-shadow:none !important; border-radius:16px !important; }
[data-testid="stChatInput"] textarea { color:#1F1B2E !important; font-size:1rem !important; font-weight:500 !important; }
[data-testid="stChatInput"] textarea::placeholder { color:#5E5A74 !important; opacity:1 !important; font-weight:400; }
[data-testid="stChatInputSubmitButton"], [data-testid="stChatInput"] button {
    background:#1F1B2E !important; color:#FFFFFF !important; border-radius:12px !important; border:0 !important;
}
[data-testid="stChatInputSubmitButton"] svg, [data-testid="stChatInput"] button svg { fill:#FFFFFF !important; color:#FFFFFF !important; }
[data-testid="stChatInputSubmitButton"]:hover, [data-testid="stChatInput"] button:hover { background:#342F4D !important; }

/* Chat bubbles: your question in soft violet, answer in white */
[data-testid="stChatMessage"] { box-shadow:0 4px 16px rgba(80,60,140,.06); }
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]),
[data-testid="stChatMessage"]:has([data-testid="chatAvatarIcon-user"]) { background:#F1ECFD; border-color:#DCD3F5; }

/* Hero: logo and name on one line */
.hero-row { display:flex; align-items:center; justify-content:center; gap:16px; margin:.2rem 0 .7rem; }
.hero .hero-row h1 { margin:0 !important; font-size:2.9rem; }

/* Stat cards with a short explanation */
.stat { padding:16px 14px; }
.stat .t { font-weight:600; color:#1F1B2E; font-size:.9rem; }
.stat .d { font-size:.75rem; color:#7A7690; margin-top:5px; line-height:1.45; }

/* Sidebar: chat history */
.side-empty { font-size:.82rem; color:#8A86A0; margin:2px 0 4px; }
[class*="st-key-hrow_"] { margin-bottom:-.35rem; }
[class*="st-key-hrow_"] [data-testid="stHorizontalBlock"] { gap:.3rem !important; flex-wrap:nowrap !important; align-items:center !important; }
[class*="st-key-hrow_"] [data-testid="stColumn"], [class*="st-key-hrow_"] [data-testid="column"] { min-width:0 !important; }
[class*="st-key-hrow_"] .stButton > button { min-height:38px; padding:0 .5rem !important; }
[class*="st-key-open_"] .stButton > button { justify-content:flex-start; text-align:left; padding:0 .8rem !important; }
[class*="st-key-open_"] .stButton > button > div { width:100%; justify-content:flex-start !important; }
[class*="st-key-open_"] .stButton > button p { white-space:nowrap; overflow:hidden; text-overflow:ellipsis; font-size:.88rem; }
.stButton > button[data-testid="stBaseButton-primary"], .stButton > button[kind="primary"] {
    background:#EFEAFB !important; border:1px solid #C9BFEA !important; color:#1F1B2E !important;
}
[class*="st-key-pin_"] .stButton > button, [class*="st-key-del_"] .stButton > button { padding:0 !important; width:100% !important; }
[class*="st-key-pin_"] .stButton > button[data-testid="stBaseButton-primary"], [class*="st-key-pin_"] .stButton > button[kind="primary"] {
    background:#1F1B2E !important; border-color:#1F1B2E !important;
}
[class*="st-key-pin_"] .stButton > button[data-testid="stBaseButton-primary"] *, [class*="st-key-pin_"] .stButton > button[kind="primary"] * { color:#FFFFFF !important; }
[class*="st-key-del_"] .stButton > button:hover { border-color:#C53030 !important; background:#FDEEEE !important; }
[class*="st-key-del_"] .stButton > button:hover * { color:#C53030 !important; }
</style>
""",
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------
# Chat history state (kept for this browser session)
# ----------------------------------------------------------------
for _k, _v in (("chats", []), ("next_id", 1), ("active", None), ("uploader_id", 0)):
    st.session_state.setdefault(_k, _v)


def new_chat(reset_uploader: bool = True):
    """Open a fresh chat. Every chat keeps its own documents (vs / chunks / files / stats / sig)."""
    cid = st.session_state.next_id
    st.session_state.next_id += 1
    st.session_state.chats = [c for c in st.session_state.chats if c["messages"]]  # drop unused empty chats
    st.session_state.chats.append(
        {"id": cid, "title": "", "messages": [], "pinned": False,
         "vs": None, "chunks": None, "files": None, "stats": None, "sig": None}
    )
    st.session_state.active = cid
    if reset_uploader:
        st.session_state.uploader_id += 1  # empties the upload box


def get_active():
    for c in st.session_state.chats:
        if c["id"] == st.session_state.active:
            return c
    new_chat()
    return st.session_state.chats[-1]


def chat_title(q: str) -> str:
    t = " ".join(q.split())
    for ch in "*_`#[]<>":
        t = t.replace(ch, "")
    return t if len(t) <= 28 else t[:27].rstrip() + "…"


def render_history():
    chats = [c for c in st.session_state.chats if c["messages"] or c["id"] == st.session_state.active]
    if not chats:
        html('<div class="side-title">Chats</div><div class="side-empty">Your conversations will appear here.</div>')
        return
    pinned = [c for c in reversed(chats) if c["pinned"]]
    recent = [c for c in reversed(chats) if not c["pinned"]]
    for label, group in (("Pinned", pinned), ("Recent", recent)):
        if not group:
            continue
        html(f'<div class="side-title">{label}</div>')
        for c in group:
            cid = c["id"]
            with st.container(key=f"hrow_{cid}"):
                a, b, d = st.columns([6, 1.5, 1.5], gap="small")
                if a.button(
                    c["title"] or "New chat",
                    key=f"open_{cid}",
                    type="primary" if cid == st.session_state.active else "secondary",
                ):
                    if cid != st.session_state.active:
                        st.session_state.chats = [x for x in st.session_state.chats if x["messages"] or x["id"] == cid]
                        st.session_state.active = cid
                        st.session_state.uploader_id += 1  # upload box shows nothing from the other chat
                    st.rerun()
                if b.button(
                    ":material/push_pin:",
                    key=f"pin_{cid}",
                    type="primary" if c["pinned"] else "secondary",
                    help="Unpin chat" if c["pinned"] else "Pin chat",
                ):
                    c["pinned"] = not c["pinned"]
                    st.rerun()
                if d.button(":material/delete:", key=f"del_{cid}", help="Delete chat"):
                    st.session_state.chats = [x for x in st.session_state.chats if x["id"] != cid]
                    if st.session_state.active == cid:
                        st.session_state.active = None
                    st.rerun()


chat = get_active()

# ----------------------------------------------------------------
# Sidebar: brand + actions + history + settings
# ----------------------------------------------------------------
with st.sidebar:
    html(f"""
    <div class="brand">{logo(40)}<div class="brand-name">RAG Agent</div></div>
    <div class="brand-sub">Document Q&amp;A for PDF, Word and Excel</div>
    """)
    if st.button(":material/add:  New chat"):
        if chat["messages"]:
            new_chat()
        else:  # already an empty chat: just clear its documents
            for _f in ("vs", "chunks", "files", "stats", "sig"):
                chat[_f] = None
            st.session_state.uploader_id += 1
        st.rerun()

    render_history()

    html('<div class="side-title">Settings</div>')
    with st.expander("Model and search options"):
        model_name = st.selectbox("Model", ["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini"])
        chunk_size = st.slider("Chunk size", 500, 2000, 1000, step=100, help="Size of each text section the documents are split into.")
        Chunk_overlap = st.slider("Chunk overlap", 0, 500, 100, step=50, help="Characters shared between neighbouring sections.")
        top_k = st.slider("Sources per answer", 2, 10, 6, help="How many document sections are retrieved for each question.")

    if st.button(":material/key:  Change API key"):
        for k in ["api_key", "chats", "active", "next_id", "uploader_id"]:
            st.session_state.pop(k, None)
        st.rerun()
    html('<div class="status"><i></i> API key verified</div>')

# ----------------------------------------------------------------
# Hero + upload
# ----------------------------------------------------------------
html(f"""
<div class="hero">
  <div class="hero-row">{logo(64)}<h1>RAG Agent</h1></div>
  <p>Upload your documents and ask questions in plain language. Every answer is based only on your files and cites its source.</p>
</div>
""")
st.write("")

uploaded_files = st.file_uploader(
    "Upload your documents",
    type=["pdf", "docx", "xlsx", "xlsm", "csv"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_id}",
    help="Supported formats: PDF, Word (.docx), Excel (.xlsx, .xlsm) and CSV. You can upload several files at once.",
)


# ----------------------------------------------------------------
# Loaders (PDF / Word / Excel / CSV)
# ----------------------------------------------------------------
def load_pdf(name: str, data: bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(data)
        tmp_path = tmp.name
    try:
        docs = PyPDFLoader(tmp_path).load()
    finally:
        os.unlink(tmp_path)
    for d in docs:
        d.metadata = {"file": name, "page": d.metadata.get("page")}
    return docs


def load_docx_plain(name: str, data: bytes):
    """Fallback reader: pulls the text straight from word/document.xml.
    Works even when an image inside the .docx is corrupt (e.g. 'Bad CRC-32')."""
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    parts = []
    for p in root.iter(w + "p"):
        text = "".join(t.text or "" for t in p.iter(w + "t")).strip()
        if text:
            parts.append(text)
    return [Document(page_content="\n".join(parts), metadata={"file": name})]


def load_docx(name: str, data: bytes):
    try:
        doc = DocxDocument(io.BytesIO(data))
    except Exception:
        return load_docx_plain(name, data)
    parts = []
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            text = Paragraph(child, doc).text.strip()
            if text:
                parts.append(text)
        elif child.tag.endswith("}tbl"):
            for row in Table(child, doc).rows:
                cells = [c.text.strip() for c in row.cells]
                if any(cells):
                    parts.append(" | ".join(cells))
    return [Document(page_content="\n".join(parts), metadata={"file": name})]


def load_excel(name: str, data: bytes, rows_per_block: int = 25):
    try:
        if name.lower().endswith(".csv"):
            sheets = {"CSV": pd.read_csv(io.BytesIO(data), dtype=str)}
        else:
            sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, dtype=str)
    except ImportError:
        raise ImportError("Excel support needs the 'openpyxl' package. Run: pip install openpyxl")

    docs = []
    for sheet, df in sheets.items():
        df = df.fillna("")
        lines = []
        for i, row in df.iterrows():
            line = " | ".join(f"{c}: {v}" for c, v in row.items() if str(v).strip())
            if line:
                lines.append((i + 2, line))  # +2 = Excel row number (header is row 1)
        for s in range(0, len(lines), rows_per_block):
            block = lines[s : s + rows_per_block]
            text = f"Sheet: {sheet}\n" + "\n".join(f"Row {n}: {t}" for n, t in block)
            docs.append(
                Document(
                    page_content=text,
                    metadata={"file": name, "sheet": sheet, "rows": f"{block[0][0]}-{block[-1][0]}"},
                )
            )
    return docs


def load_file(name: str, data: bytes):
    low = name.lower()
    if low.endswith(".pdf"):
        return load_pdf(name, data)
    if low.endswith(".docx"):
        return load_docx(name, data)
    return load_excel(name, data)


def build_vectorstore(files, size: int, overlap: int):
    """Load every file, split into chunks, embed, store in FAISS"""
    docs = []
    for name, data in files:
        docs.extend(load_file(name, data))

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=size,
        chunk_overlap=overlap,
        separators=["\n\n", "\n", ". ", " ", ""],
    )
    chunks = [c for c in splitter.split_documents(docs) if c.page_content.strip()]
    if not chunks:
        raise ValueError("No readable text was found. Scanned PDFs need OCR before they can be searched.")

    embeddings = OpenAIEmbeddings(model="text-embedding-3-small")
    return FAISS.from_documents(chunks, embeddings), len(chunks), chunks


def source_label(d) -> str:
    m = d.metadata
    name = m.get("file", "document")
    if isinstance(m.get("page"), int):
        return f"{name} (p. {m['page'] + 1})"
    if "sheet" in m:
        return f"{name} (sheet '{m['sheet']}', rows {m['rows']})"
    return name


TAGS = {"pdf": "PDF", "docx": "DOCX", "xlsx": "XLSX", "xlsm": "XLSM", "csv": "CSV"}

if uploaded_files:
    # Re-index only when new files (or new settings) arrive
    file_sig = (tuple((f.name, f.size) for f in uploaded_files), chunk_size, Chunk_overlap)
    if chat["sig"] != file_sig:
        try:
            with st.spinner("Indexing your documents..."):
                files = [(f.name, f.getvalue()) for f in uploaded_files]
                vs, n_chunks, all_chunks = build_vectorstore(files, chunk_size, Chunk_overlap)
            # This chat already has questions -> the new documents go into a fresh chat
            if chat["messages"]:
                new_chat(reset_uploader=False)  # keep the files in the upload box
                chat = get_active()
            chat["vs"] = vs
            chat["chunks"] = all_chunks
            chat["sig"] = file_sig
            chat["files"] = [f.name for f in uploaded_files]
            chat["stats"] = (len(uploaded_files), n_chunks)
            st.success(
                f"Indexed {len(uploaded_files)} document(s) into {n_chunks} searchable sections. You can start asking questions.",
                icon=":material/check_circle:",
            )
        except AuthenticationError:
            st.error("OpenAI rejected your API key. Use \"Change API key\" in the sidebar to enter a valid one.")
        except Exception as e:
            st.error(f"Couldn't index the documents: {e}")

if chat["files"]:
    chips = "".join(
        f'<span class="chip"><i>{TAGS.get(n.rsplit(".", 1)[-1].lower(), "FILE")}</i>{n}</span>'
        for n in chat["files"]
    )
    html(chips)
    st.write("")
    s1, s2, s3 = st.columns(3)
    s1.markdown(
        f'<div class="stat"><b>{chat["stats"][0]}</b><div class="t">Documents</div>'
        '<div class="d">The files you uploaded.</div></div>',
        unsafe_allow_html=True,
    )
    s2.markdown(
        f'<div class="stat"><b>{chat["stats"][1]}</b><div class="t">Sections indexed</div>'
        '<div class="d">Small pieces your files were split into, so the right part can be found quickly.</div></div>',
        unsafe_allow_html=True,
    )
    s3.markdown(
        f'<div class="stat"><b>{top_k}</b><div class="t">Sources per answer</div>'
        '<div class="d">How many sections the AI reads to answer each question. Change it in Settings.</div></div>',
        unsafe_allow_html=True,
    )
    st.write("")

# ----------------------------------------------------------------
# RAG chain
# ----------------------------------------------------------------
RAG_PROMPT = ChatPromptTemplate.from_messages([
    ("system",
     "You are a helpful assistant that answers questions strictly using the "
     "provided context from the user's documents (PDF, Word and Excel files).\n"
     "Rules:\n"
     "1. Answer ONLY from the context below.\n"
     "2. If the answer is not in the context, say: "
     "\"I couldn't find that in the documents.\"\n"
     "3. Cite your sources using the labels shown in the context, "
     "e.g. (report.pdf, p. 3) or (sales.xlsx, sheet 'Q1').\n"
     "4. Headings such as '2. SCOPE' or 'ABSTRACT' mark sections: if the question "
     "names a section or topic, summarize the text under the matching heading.\n"
     "5. For general questions (what is in this file, summarize it), give an overview "
     "of whatever the context contains.\n\n"
     "Context:\n{context}"),
    ("human", "{question}"),
])


def format_docs(docs) -> str:
    """Join retrieved chunks, tagging each with its file / page / sheet."""
    return "\n\n".join(f"[{source_label(d)}]\n{d.page_content}" for d in docs)


STOPWORDS = {
    "the", "and", "for", "are", "what", "whats", "which", "who", "how", "why", "when", "where", "tell", "about",
    "me", "this", "that", "with", "from", "into", "does", "did", "can", "you", "your", "please", "give", "show",
    "explain", "describe", "file", "document", "documents", "list", "have", "has", "was", "were",
    "any", "all", "its", "their", "there", "than", "then", "also", "not", "but", "his", "her", "our", "out",
}


def keyword_chunks(question: str, chunks, n: int = 3):
    """Chunks that literally contain the question's keywords (headings count extra)."""
    words = {w for w in re.findall(r"[A-Za-z0-9\u0600-\u06FF]{3,}", question.lower()) if w not in STOPWORDS}
    if not words:
        return []
    scored = []
    for c in chunks:
        text = c.page_content.lower()
        score = sum(text.count(w) for w in words)
        if score:
            lines = [re.sub(r"^[\d\.\)\s]+", "", ln.strip()) for ln in text.splitlines()]
            if any(ln in words for ln in lines):
                score += 5  # a heading that matches the keyword
            scored.append((score, c))
    scored.sort(key=lambda x: -x[0])
    return [c for _, c in scored[:n]]


def retrieve(question: str, vectorstore: FAISS, chunks, k: int):
    """Hybrid search: keyword matches first, then the usual meaning-based (vector) matches."""
    docs = keyword_chunks(question, chunks)
    seen = {d.page_content for d in docs}
    for d in vectorstore.similarity_search(question, k=k):
        if d.page_content not in seen:
            docs.append(d)
            seen.add(d.page_content)
    return docs


def get_chain(vectorstore: FAISS, model: str, k: int, chunks):
    retriever = RunnableLambda(lambda q: retrieve(q, vectorstore, chunks, k))
    llm = ChatOpenAI(model=model, temperature=0)
    return (
        {"context": retriever | format_docs, "question": RunnablePassthrough()}
        | RAG_PROMPT
        | llm
        | StrOutputParser()
    ), retriever


# ----------------------------------------------------------------
# Chat interface
# ----------------------------------------------------------------
USER_AVATAR = ":material/person:"
BOT_AVATAR = ":material/auto_awesome:"


def show_sources(sources):
    with st.expander("Sources (retrieved sections)"):
        for label, text in sources:
            st.markdown(f"**{label}**")
            st.text(text)
            st.divider()


if chat["vs"] is not None:
    messages = chat["messages"]

    # Replay history of the open chat
    for msg in messages:
        with st.chat_message(msg["role"], avatar=USER_AVATAR if msg["role"] == "user" else BOT_AVATAR):
            st.markdown(msg["content"])
            if msg.get("sources"):
                show_sources(msg["sources"])

    question = st.chat_input("Ask a question about your documents...")
    if question:
        messages.append({"role": "user", "content": question})
        if not chat["title"]:
            chat["title"] = chat_title(question)
        with st.chat_message("user", avatar=USER_AVATAR):
            st.markdown(question)

        chain, retriever = get_chain(chat["vs"], model_name, top_k, chat["chunks"])

        sources = []
        with st.chat_message("assistant", avatar=BOT_AVATAR):
            try:
                with st.spinner("Searching your documents..."):
                    answer = chain.invoke(question)
                st.markdown(answer)

                # Show the retrieved chunks for transparency
                sources = [(source_label(d), d.page_content[:500]) for d in retriever.invoke(question)]
                show_sources(sources)
            except AuthenticationError:
                answer = "Your API key was rejected by OpenAI. Use \"Change API key\" in the sidebar to enter a valid one."
                st.error(answer)
            except Exception as e:
                answer = f"Something went wrong while answering: {e}"
                st.error(answer)

        messages.append({"role": "assistant", "content": answer, "sources": sources})
        if len(messages) == 2:
            st.rerun()  # first exchange: refresh the sidebar so the new chat appears in the history
else:
    html("""
    <div class="cards">
      <div class="card"><div class="num">01</div><h4>Upload</h4><span>Add PDF, Word, Excel or CSV files. You can upload several at once.</span></div>
      <div class="card"><div class="num">02</div><h4>Index</h4><span>Your documents are split into sections and indexed automatically.</span></div>
      <div class="card"><div class="num">03</div><h4>Ask</h4><span>Get answers grounded in your files, with the file, page or sheet cited.</span></div>
    </div>
    """)
    st.info("Upload one or more documents above to begin.", icon=":material/upload_file:")