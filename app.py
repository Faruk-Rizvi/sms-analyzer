"""
SMS Forensic Analyzer — Streamlit App
Bangladesh Law Enforcement | v12
"""
import streamlit as st
import tempfile, os, sys, traceback

# ── Page config ─────────────────────────────────────────────
st.set_page_config(
    page_title="SMS Forensic Analyzer",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Hide default Streamlit chrome ────────────────────────────
st.markdown("""
<style>
#MainMenu, footer, header {visibility: hidden;}
.block-container {padding-top: 1.5rem; padding-bottom: 1rem;}
</style>
""", unsafe_allow_html=True)

# ── Load analyzer module ─────────────────────────────────────
@st.cache_resource(show_spinner=False)
def load_analyzer():
    """sms_analyzer.py import করো — error হলে clearly দেখাও।"""
    try:
        import sms_analyzer as sa
        return sa
    except Exception as e:
        return e

sa = load_analyzer()

# ── Header ───────────────────────────────────────────────────
st.markdown("""
<h2 style='margin:0 0 4px 0; color:#1e293b; font-family: Georgia, serif;'>
  🔍 SMS Forensic Analyzer
</h2>
<p style='color:#64748b; margin:0 0 20px 0; font-size:14px;'>
  Bangladesh Law Enforcement — CDR Analysis Tool
</p>
""", unsafe_allow_html=True)

# ── Import error ─────────────────────────────────────────────
if isinstance(sa, Exception):
    st.error("⚠️ sms_analyzer.py লোড করতে সমস্যা হয়েছে")
    st.code(traceback.format_exc(), language="python")
    st.info("নিশ্চিত করুন `sms_analyzer.py` repository-তে আছে এবং `requirements.txt`-এ `openpyxl` আছে।")
    st.stop()

# ── Upload ───────────────────────────────────────────────────
col1, col2 = st.columns([3, 1])
with col1:
    uploaded = st.file_uploader(
        "CDR ফাইল আপলোড করুন (.xlsx)",
        type=["xlsx"],
        help="Grameenphone, Robi, Banglalink, NIP — সব format সাপোর্টেড।",
    )
with col2:
    target = st.text_input(
        "Target নম্বর (optional)",
        placeholder="01711385207",
        help="শুধু এই নম্বর সংক্রান্ত SMS দেখাবে।",
    )

# ── Process ──────────────────────────────────────────────────
if uploaded:
    # temp file-এ save করো
    suffix = os.path.splitext(uploaded.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_in:
        tmp_in.write(uploaded.read())
        tmp_xlsx = tmp_in.name

    tmp_html = tmp_xlsx.replace(suffix, ".html")

    with st.spinner("বিশ্লেষণ চলছে…"):
        try:
            result = sa.process_file(
                tmp_xlsx,
                tmp_html,
                target_number=target.strip() if target.strip() else None,
            )
        except Exception as e:
            st.error(f"প্রসেসিং ব্যর্থ: {e}")
            st.code(traceback.format_exc(), language="python")
            os.unlink(tmp_xlsx)
            st.stop()

    # ── Stats bar ────────────────────────────────────────────
    s = result
    st.markdown(f"""
    <div style='display:flex;gap:12px;flex-wrap:wrap;margin:12px 0;'>
      <div style='background:#f1f5f9;border-radius:8px;padding:8px 16px;font-size:13px;'>
        📊 <b>{s.get("total_sms",0):,}</b> মোট
      </div>
      <div style='background:#f1f5f9;border-radius:8px;padding:8px 16px;font-size:13px;'>
        ✅ <b>{s.get("total_shown",0):,}</b> shown
      </div>
      <div style='background:#fee2e2;border-radius:8px;padding:8px 16px;font-size:13px;color:#b91c1c;'>
        🔴 <b>{s.get("high",0)}</b> High (১লক্ষ+)
      </div>
      <div style='background:#ffedd5;border-radius:8px;padding:8px 16px;font-size:13px;color:#c2410c;'>
        🟠 <b>{s.get("med",0)}</b> Medium
      </div>
      <div style='background:#fef9c3;border-radius:8px;padding:8px 16px;font-size:13px;color:#a16207;'>
        🟡 <b>{s.get("small",0)}</b> Small
      </div>
      <div style='background:#f1f5f9;border-radius:8px;padding:8px 16px;font-size:13px;color:#64748b;'>
        🚫 <b>{s.get("filtered",0):,}</b> filtered
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Download & embed ─────────────────────────────────────
    with open(tmp_html, "rb") as f:
        html_bytes = f.read()

    dl_col, _ = st.columns([1, 3])
    with dl_col:
        st.download_button(
            label="📥 রিপোর্ট ডাউনলোড (.html)",
            data=html_bytes,
            file_name=f"report_{uploaded.name.replace('.xlsx','')}.html",
            mime="text/html",
            use_container_width=True,
        )

    # ── Embed HTML report in iframe ───────────────────────────
    html_str  = html_bytes.decode("utf-8", errors="replace")
    # Streamlit-এ iframe-এ embed করো
    st.components.v1.html(html_str, height=800, scrolling=True)

    # cleanup
    try:
        os.unlink(tmp_xlsx)
        os.unlink(tmp_html)
    except Exception:
        pass

else:
    # Welcome state
    st.markdown("""
    <div style='background:#f8fafc;border:2px dashed #cbd5e1;border-radius:12px;
                padding:40px;text-align:center;color:#94a3b8;margin-top:20px;'>
        <div style='font-size:48px;margin-bottom:12px;'>📂</div>
        <div style='font-size:16px;font-weight:600;color:#64748b;'>
            উপরে CDR ফাইল (.xlsx) আপলোড করুন
        </div>
        <div style='font-size:13px;margin-top:8px;'>
            GP · Robi · Banglalink · NIP সব format সাপোর্টেড
        </div>
    </div>
    """, unsafe_allow_html=True)
