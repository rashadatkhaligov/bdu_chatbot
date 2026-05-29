"""
TARİX AI Köməkçi — BDU Tarix Fakültəsi
Streamlit versiyası · Context Cache ilə token qənaəti

İşə salmaq:
    streamlit run tarix.py
"""

import os
import re
import time
from pathlib import Path

import streamlit as st
import openpyxl
from google import genai
from google.genai import types

# ─────────────────────────────────────────────
# KONFIQURASIYA
# ─────────────────────────────────────────────
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL   = "gemini-2.5-flash"
MAX_MSG_LEN    = 800
MAX_HIST_TURNS = 8
CACHE_TTL_SEC  = 21600         # keş 6 saat saxlanılır
KNOWLEDGE_FILE = Path(__file__).parent / "BDU Tarix Chatbot.xlsx"

# ─────────────────────────────────────────────
# SƏHIFƏ AYARLARI
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="TARİX AI · BDU",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&family=Noto+Serif:wght@600;700&display=swap');

/* ── Açıq rejim (default) ── */
:root {
    --navy:       #0d2137;
    --navy-mid:   #163352;
    --blue:       #1e4d8c;
    --blue-light: #2a6abf;
    --gold:       #c8962a;
    --gold-light: #e8b84b;
    --bg:         #f7f4ee;
    --surface:    #ffffff;
    --text:       #1a1e2e;
    --text-soft:  #4a5568;
    --border:     #e2e8f0;
    --shadow:     0 4px 24px rgba(13,33,55,0.10);
    --input-bg:   #ffffff;
    --font-main:  'Noto Sans', sans-serif;
    --font-serif: 'Noto Serif', serif;
}

/* ── Tünd rejim ── */
@media (prefers-color-scheme: dark) {
    :root {
        --bg:        #0f1923;
        --surface:   #1a2635;
        --text:      #e8edf5;
        --text-soft: #94a3b8;
        --border:    #2a3a4e;
        --shadow:    0 4px 24px rgba(0,0,0,0.35);
        --input-bg:  #1a2635;
    }
}

html, body, [class*="css"], * { font-family: var(--font-main) !important; }
.stApp { background: var(--bg) !important; }

/* ── Sidebar — həmişə tünd ── */
[data-testid="stSidebar"] { background: var(--navy) !important; border-right: none !important; }
[data-testid="stSidebar"] * { font-family: var(--font-main) !important; color: rgba(255,255,255,0.88) !important; }
[data-testid="stSidebar"] .stButton button {
    background: rgba(200,150,42,0.15) !important;
    border: 1px solid rgba(200,150,42,0.4) !important;
    color: var(--gold-light) !important;
    border-radius: 8px !important; font-weight: 500 !important;
    transition: all .2s !important;
}
[data-testid="stSidebar"] .stButton button:hover {
    background: rgba(200,150,42,0.28) !important;
    border-color: var(--gold-light) !important;
}
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.1) !important; }

.sidebar-brand { padding: 8px 0 20px; border-bottom: 1px solid rgba(255,255,255,0.08); margin-bottom: 20px; }
.sidebar-brand h3 { font-family: var(--font-serif) !important; font-size: 1.15rem !important; font-weight: 700 !important; color: #fff !important; margin: 0 !important; }
.sidebar-brand span { font-size: 0.72rem; color: var(--gold-light) !important; letter-spacing: 0.07em; text-transform: uppercase; }

.status-line { font-size: 0.75rem; color: rgba(255,255,255,0.5) !important; display: flex; align-items: center; margin-bottom: 14px; }
.status-dot { display: inline-block; width: 7px; height: 7px; border-radius: 50%; margin-right: 7px; flex-shrink: 0; }
.status-dot.online  { background: #22c55e; animation: pulse 2s infinite; }
.status-dot.offline { background: #ef4444; }
@keyframes pulse { 0%,100%{ opacity:1; transform:scale(1); } 50%{ opacity:.6; transform:scale(1.35); } }

.stat-card { background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.1); border-radius: 10px; padding: 11px 14px; margin-bottom: 8px; }
.stat-card .s-label { font-size: 0.68rem; color: rgba(255,255,255,0.38) !important; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 3px; }
.stat-card .s-value { font-size: 1.05rem; font-weight: 600; color: var(--gold-light) !important; }

.lang-chips { display: flex; gap: 6px; flex-wrap: wrap; margin-top: 6px; }
.lang-chip { background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; padding: 4px 10px; font-size: 0.74rem; color: rgba(255,255,255,0.7) !important; }

/* ── Başlıq — həmişə tünd ── */
.page-header {
    background: linear-gradient(135deg, var(--navy) 0%, var(--navy-mid) 60%, var(--blue) 100%);
    border-radius: 18px; padding: 30px 34px; margin-bottom: 6px;
    display: flex; align-items: center; gap: 22px;
    position: relative; overflow: hidden; box-shadow: var(--shadow);
}
.page-header::before { content: ""; position: absolute; top: -40px; right: -40px; width: 220px; height: 220px; border-radius: 50%; background: radial-gradient(circle, rgba(200,150,42,0.18) 0%, transparent 70%); }
.page-header::after  { content: ""; position: absolute; bottom: -60px; left: 30%; width: 300px; height: 300px; border-radius: 50%; background: radial-gradient(circle, rgba(42,106,191,0.15) 0%, transparent 70%); }
.header-icon { width: 62px; height: 62px; border-radius: 14px; background: linear-gradient(135deg, var(--gold), var(--gold-light)); display: flex; align-items: center; justify-content: center; font-size: 1.9rem; box-shadow: 0 8px 24px rgba(200,150,42,0.35); flex-shrink: 0; position: relative; z-index: 1; }
.header-text { position: relative; z-index: 1; }
.header-text h1 { font-family: var(--font-serif) !important; color: #fff !important; font-size: 1.7rem !important; font-weight: 700 !important; margin: 0 0 4px !important; line-height: 1.2 !important; }
.header-text p { color: rgba(255,255,255,0.58) !important; font-size: 0.8rem !important; margin: 0 !important; letter-spacing: 0.04em !important; text-transform: uppercase !important; }
.header-badges { display: flex; gap: 8px; margin-top: 10px; flex-wrap: wrap; }
.badge { background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.18); border-radius: 20px; padding: 3px 11px; font-size: 0.71rem; color: rgba(255,255,255,0.75) !important; }

.accent-bar { height: 3px; background: linear-gradient(90deg, var(--gold), var(--gold-light), transparent); border-radius: 2px; margin-bottom: 22px; }

/* ── Chat mesajları — fona uyğun ── */
[data-testid="stChatMessage"] { background: transparent !important; border: none !important; padding: 4px 0 !important; }
[data-testid="stChatMessage"] p,
[data-testid="stChatMessage"] li,
[data-testid="stChatMessage"] span,
[data-testid="stChatMessage"] div,
[data-testid="stChatMessageContent"] { font-family: var(--font-main) !important; color: var(--text) !important; }
.stMarkdown p, .stMarkdown li, .stMarkdown span, .stMarkdown strong { color: var(--text) !important; }

/* ── Giriş sahəsi — fona uyğun ── */
[data-testid="stChatInput"] {
    background: var(--input-bg) !important;
    border-radius: 14px !important;
    border: 2px solid var(--border) !important;
    box-shadow: 0 2px 12px rgba(13,33,55,0.07) !important;
    transition: border-color .2s !important;
}
[data-testid="stChatInput"]:focus-within { border-color: var(--blue-light) !important; box-shadow: 0 2px 16px rgba(30,77,140,0.12) !important; }
[data-testid="stChatInput"] textarea { font-family: var(--font-main) !important; font-size: 0.9rem !important; color: var(--text) !important; background: transparent !important; }
[data-testid="stChatInput"] button { background: linear-gradient(135deg, var(--navy-mid), var(--blue)) !important; border-radius: 10px !important; border: none !important; }

.stSpinner > div { border-top-color: var(--gold) !important; }
::-webkit-scrollbar { width: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 1.5rem !important; max-width: 960px !important; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# MƏLUMAT BAZASINI YÜKLƏ (bir dəfə)
# ─────────────────────────────────────────────
@st.cache_resource(show_spinner="📚 Məlumat bazası yüklənir…")
def load_knowledge() -> str:
    if not KNOWLEDGE_FILE.exists():
        return "BDU Tarix fakültəsi haqqında məlumat bazası tapılmadı."
    try:
        wb   = openpyxl.load_workbook(str(KNOWLEDGE_FILE), read_only=True)
        skip = {"Sual","Cavab","Soyad, ad, ata adı","№",
                "Toplanmış Yekun Bal","Struktur","Qiymətləndirmə"}
        parts = ["BDU TARİX FAKÜLTƏSİ — MƏLUMAT BAZASI\n"]

        for sn in wb.sheetnames:
            ws   = wb[sn]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            parts.append(f"\n{'='*50}\nBÖLMƏ: {sn}\n{'='*50}")

            # ── İmtahan cədvəli ──────────────────────────
            if sn.lower() == "imtahanlar":
                parts.append("İmtahan Cədvəli (Tarix fakültəsi):")
                parts.append("Tarix | Fənn | Kurs | Qrup | Tələbə sayı | Otaq | İmtahan növü")
                for row in rows[1:]:
                    if not any(c for c in row if c is not None):
                        continue
                    tarix   = str(row[0]).split(" ")[0] if row[0] else ""
                    # datetime formatını düzəlt
                    if hasattr(row[0], 'strftime'):
                        tarix = row[0].strftime("%d.%m.%Y")
                    fenn    = str(row[1]).strip() if row[1] else ""
                    kurs    = str(row[2]).strip() if row[2] else ""
                    qrup    = str(row[3]).strip() if row[3] else ""
                    say     = str(row[4]).strip() if row[4] else ""
                    otaq    = str(row[6]).strip() if row[6] else ""
                    nov     = str(row[7]).strip() if row[7] else ""
                    if fenn:
                        parts.append(f"{tarix} | {fenn} | {kurs} | {qrup} | {say} nəfər | {otaq} | {nov}")
                continue

            # ── Əməkdaş cədvəli ─────────────────────────
            if "əməkdaş" in sn.lower():
                parts.append("Fakültənin əməkdaşları:")
                for row in rows[1:]:
                    vals = [str(c).strip() for c in row if c is not None and str(c).strip()]
                    if vals:
                        parts.append(" | ".join(vals))
                continue

            if "Cədvəl" in sn:
                for row in rows:
                    vals = [str(c).strip() for c in row if c is not None and str(c).strip()]
                    if vals:
                        parts.append(" | ".join(vals))
                continue

            for row in rows[1:]:
                if not any(c for c in row if c is not None):
                    continue
                cells = [str(c).strip() if c is not None else "" for c in row]
                if len(cells) >= 3 and cells[1] and cells[2]:
                    s, c2 = cells[1], cells[2]
                    if s not in skip and c2 not in skip:
                        parts.append(f"\nSual: {s}\nCavab: {c2}")

        return "\n".join(parts)
    except Exception as e:
        return f"Məlumat bazası oxunarkən xəta: {e}"


# ─────────────────────────────────────────────
# CONTEXT CACHE — məlumat bazasını bir dəfə
# Gemini-yə yükləyirik, hər sorğuda yalnız sual gedir
# ─────────────────────────────────────────────
@st.cache_resource(show_spinner="⚡ Keş yaradılır…")
def create_cache(knowledge: str, system: str) -> str | None:
    """
    Məlumat bazasını Gemini Context Cache-ə yükləyir.
    Cache adını qaytarır. Xəta olarsa None qaytarır.
    Minimum 32.768 token tələb olunur.
    """
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        cache  = client.caches.create(
            model=GEMINI_MODEL,
            config=types.CreateCachedContentConfig(
                system_instruction=system,
                contents=[
                    types.Content(
                        role="user",
                        parts=[types.Part(text=knowledge)]
                    ),
                    types.Content(
                        role="model",
                        parts=[types.Part(text="Məlumat bazası qəbul edildi. Suallarınızı gözləyirəm.")]
                    ),
                ],
                ttl=f"{CACHE_TTL_SEC}s",
            )
        )
        return cache.name
    except Exception as e:
        # Cache minimum token tələbi yerinə yetirilmədikdə fallback-ə keçir
        st.session_state["cache_error"] = str(e)
        return None


# ─────────────────────────────────────────────
# DİL AŞKARLAMA
# ─────────────────────────────────────────────
def detect_language(text: str) -> str:
    if re.search(r"[\u0400-\u04FF]", text):
        return "ru"
    if re.search(r"[əıöüğşçƏIÖÜĞŞÇ]", text):
        return "az"
    az_words = {"salam","necəsən","necəsiniz","bəli","xeyr","hə","yox","nə",
                "bu","bir","var","kim","harada","nece","niyə","nədir","hansı",
                "neçə","təşəkkür","sagol","sağol","əla","yaxşı","pis","cənab",
                "xanım","dekan","kafedra","fakulte","imtahan","telebe","tələbə"}
    words = set(text.lower().split())
    if words & az_words:
        return "az"
    en_words = {"hello","hi","what","who","how","where","when","why","is","are",
                "the","a","an","and","or","yes","no","please","thank","thanks"}
    if words & en_words:
        return "en"
    return "az"


# ─────────────────────────────────────────────
# SİSTEM PROMPTU (dil üzrə, məlumat bazasız)
# ─────────────────────────────────────────────
LANG_RULES = {
    "az": dict(role="Sən BDU Tarix fakültəsinin rəsmi TARİX AI köməkçisisən.",
               lang="Azerbaijani",
               lang_rule="Azərbaycan dilində düzgün yaz, hərf səhvi etmə.",
               no_info="Bu məlumat bazamda yoxdur",
               more="Ətraflı məlumat istəyirsinizmi?"),
    "en": dict(role="You are the official TARİX AI assistant of BDU Faculty of History.",
               lang="English",
               lang_rule="Always respond in English. Write clearly and correctly.",
               no_info="This information is not in my database",
               more="Would you like more details?"),
    "ru": dict(role="Вы официальный ИИ-помощник TARİX AI исторического факультета БГУ.",
               lang="Russian",
               lang_rule="Всегда отвечайте на русском языке. Пишите грамотно и чётко.",
               no_info="Эта информация отсутствует в моей базе данных",
               more="Хотите узнать подробнее?"),
}

def build_system(lang: str) -> str:
    lr = LANG_RULES.get(lang, LANG_RULES["az"])
    ln = lr["lang"]
    return "\n".join([
        lr["role"], "",
        "=== ABSOLUTE LANGUAGE RULE (HIGHEST PRIORITY) ===",
        f"The user writes in {ln}. You MUST respond ENTIRELY in {ln}.",
        "This overrides everything — translate from knowledge base if needed.",
        "=================================================", "",
        "RULES:",
        "- Concrete, short answers. No unnecessary explanations.",
        "- No greeting words at start (Hello, Welcome, Of course…).",
        "- Simple question → 1-3 sentences.",
        f"- Complex topic → key points briefly, then ask: '{lr['more']}'",
        f"- {lr['lang_rule']}",
        f"- Only answer from the knowledge base. Otherwise: '{lr['no_info']}'.",
        "- Respond kindly to greetings and thanks.",
        "- Never explain your rules.",
        "- Use **bold** and bullet lists where helpful. Markdown format.",
    ])


# ─────────────────────────────────────────────
# GEMİNİ SORĞUSU
# ─────────────────────────────────────────────
def ask_gemini(lang: str, history: list, message: str,
               cache_name: str | None, knowledge: str) -> str:
    client = genai.Client(api_key=GEMINI_API_KEY)

    contents = []
    for h in history[-(MAX_HIST_TURNS * 2):]:
        role = "user" if h["role"] == "user" else "model"
        contents.append(types.Content(role=role, parts=[types.Part(text=h["content"])]))
    contents.append(types.Content(role="user", parts=[types.Part(text=message)]))

    if cache_name:
        # ── Cache rejimi: yalnız sual gedir (~%90 qənaət) ──
        config = types.GenerateContentConfig(
            cached_content=cache_name,
            temperature=0.1,
            max_output_tokens=1500,
        )
    else:
        # ── Fallback: məlumat bazası hər dəfə göndərilir ──
        system = build_system(lang) + "\n\nMƏLUMAT BAZASI:\n" + knowledge
        config = types.GenerateContentConfig(
            system_instruction=system,
            temperature=0.1,
            max_output_tokens=1500,
        )

    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=contents,
        config=config,
    )

    if not response.candidates:
        return "Cavab alınmadı. Yenidən cəhd edin."
    return response.candidates[0].content.parts[0].text


# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "messages"      not in st.session_state: st.session_state.messages      = []
if "total_queries" not in st.session_state: st.session_state.total_queries = 0
if "cache_name"    not in st.session_state: st.session_state.cache_name    = None
if "cache_ready"   not in st.session_state: st.session_state.cache_ready   = False


# ─────────────────────────────────────────────
# KENAR PANEL
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <h3>🏛️ TARİX AI</h3>
        <span>BDU · Tarix Fakültəsi</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="status-line">
        <span class="status-dot online"></span>Sistem aktiv
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    kb_exists   = KNOWLEDGE_FILE.exists()
    cache_label = "⚡ Aktiv" if st.session_state.cache_ready else "○ Gözləyir"
    st.markdown(f"""
    <div class="stat-card">
        <div class="s-label">Məlumat Bazası</div>
        <div class="s-value">{'✓ Yüklənib' if kb_exists else '✗ Tapılmadı'}</div>
    </div>
    <div class="stat-card">
        <div class="s-label">Context Cache</div>
        <div class="s-value">{cache_label}</div>
    </div>
    <div class="stat-card">
        <div class="s-label">Ümumi Sorğular</div>
        <div class="s-value">{st.session_state.total_queries}</div>
    </div>
    <div class="stat-card">
        <div class="s-label">Söhbətdəki Mesajlar</div>
        <div class="s-value">{len(st.session_state.messages)}</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    st.markdown("""
    <div style='font-size:0.7rem;color:rgba(255,255,255,0.38);text-transform:uppercase;
                letter-spacing:0.07em;margin-bottom:8px;'>Dil Dəstəyi</div>
    <div class="lang-chips">
        <span class="lang-chip">🇦🇿 Azərbaycan</span>
        <span class="lang-chip">🇬🇧 İngilis</span>
        <span class="lang-chip">🇷🇺 Rus</span>
    </div>
    <div style='font-size:0.7rem;color:rgba(255,255,255,0.28);margin-top:7px;'>
        Dil avtomatik aşkarlanır
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🗑 Sil", use_container_width=True):
            st.session_state.messages = []
            st.rerun()
    with col2:
        if st.button("↺ Yenilə", use_container_width=True):
            st.cache_resource.clear()
            st.session_state.cache_name  = None
            st.session_state.cache_ready = False
            st.rerun()

    st.markdown("""
    <div style='font-size:0.67rem;color:rgba(255,255,255,0.18);text-align:center;
                margin-top:28px;line-height:1.7;'>
        TARİX AI v3.0<br>Gemini 2.5 Flash · Context Cache
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# ANA SƏHİFƏ
# ─────────────────────────────────────────────
st.markdown("""
<div class="page-header">
    <div class="header-icon">🏛️</div>
    <div class="header-text">
        <h1>TARİX AI Köməkçi</h1>
        <p>Bakı Dövlət Universiteti · Tarix Fakültəsi · Rəsmi Rəqəmsal Köməkçi</p>
        <div class="header-badges">
            <span class="badge">📚 Fakültə Məlumatları</span>
            <span class="badge">👨‍🏫 Əməkdaşlar</span>
            <span class="badge">📋 Tədris Prosesi</span>
            <span class="badge">⚡ Context Cache</span>
        </div>
    </div>
</div>
<div class="accent-bar"></div>
""", unsafe_allow_html=True)

if not KNOWLEDGE_FILE.exists():
    st.warning(f"⚠️ `{KNOWLEDGE_FILE.name}` faylını bu qovluğa əlavə edin.")

# ── Cache-i bir dəfə yarat ───────────────────
if not st.session_state.cache_ready and GEMINI_API_KEY:
    knowledge = load_knowledge()
    system    = build_system("az")  # cache dil-neytral saxlanılır
    name      = create_cache(knowledge, system)
    st.session_state.cache_name  = name
    st.session_state.cache_ready = True

knowledge = load_knowledge()

# ── Mesaj tarixçəsi ──────────────────────────
if not st.session_state.messages:
    with st.chat_message("assistant", avatar="🏛️"):
        st.markdown(
            "Salam! Mən **BDU Tarix fakültəsinin** rəqəmsal köməkçisiyəm.\n\n"
            "Aşağıdakı mövzularda kömək edə bilərəm:\n"
            "- 🏛️ Fakültə tarixi və strukturu\n"
            "- 👨‍🏫 Kafedra müdirləri və əməkdaşlar\n"
            "- 📚 İxtisaslar və magistratura proqramları\n"
            "- 📋 Tədris qaydaları, imtahanlar, təqaüd\n"
            "- 🎓 Tələbə hüquq və vəzifələri\n\n"
            "Sualınızı Azərbaycan, İngilis və ya Rus dilində yaza bilərsiniz."
        )
else:
    for msg in st.session_state.messages:
        avatar = "🧑" if msg["role"] == "user" else "🏛️"
        with st.chat_message(msg["role"], avatar=avatar):
            st.markdown(msg["content"])

# ── Giriş sahəsi ────────────────────────────
if prompt := st.chat_input("Sualınızı yazın…"):
    prompt = prompt[:MAX_MSG_LEN]

    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user", avatar="🧑"):
        st.markdown(prompt)

    with st.chat_message("assistant", avatar="🏛️"):
        with st.spinner(""):
            try:
                lang   = detect_language(prompt)
                answer = ask_gemini(
                    lang,
                    st.session_state.messages[:-1],
                    prompt,
                    st.session_state.cache_name,
                    knowledge,
                )
                st.session_state.total_queries += 1
            except Exception as e:
                answer = f"⚠️ Xəta baş verdi: `{e}`"
        st.markdown(answer)

    st.session_state.messages.append({"role": "assistant", "content": answer})
    st.rerun()
