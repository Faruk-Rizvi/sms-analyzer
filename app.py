import streamlit as st
import tempfile, os, re, base64
from sms_analyzer import process_file

st.set_page_config(
    page_title="SMS Intelligence Dashboard",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

*,*::before,*::after{box-sizing:border-box}
html,body,.stApp{background:#f4f5f7!important;font-family:'Outfit',sans-serif}
.block-container{padding:0!important;max-width:100%!important}
.stApp>header{display:none}
section[data-testid="stSidebar"]{display:none}
footer{display:none}

/* ── TOPBAR ── */
.topbar{background:#fff;border-bottom:1px solid #e8eaed;padding:0 48px;height:64px;
  display:flex;align-items:center;justify-content:space-between;
  position:sticky;top:0;z-index:1000;box-shadow:0 1px 8px rgba(0,0,0,.06)}
.tb-brand{display:flex;align-items:center;gap:12px}
.tb-logo{width:40px;height:40px;background:linear-gradient(135deg,#c0392b,#96281b);
  border-radius:10px;display:flex;align-items:center;justify-content:center;
  font-size:20px;box-shadow:0 4px 12px rgba(192,57,43,.3)}
.tb-title{font-size:17px;font-weight:700;color:#1a1a2e;letter-spacing:-.3px}
.tb-sub{font-size:11px;color:#9ca3af;margin-top:1px}
.tb-nav{display:flex;gap:28px}
.tb-nav a{font-size:13px;font-weight:500;color:#525866;text-decoration:none}
.tb-nav a:hover{color:#c0392b}

/* ── HERO ── */
.hero{background:linear-gradient(135deg,#1a0505 0%,#7b0d0d 42%,#c0392b 70%,#a93226 100%);
  padding:52px 48px 44px;position:relative;overflow:hidden}
.hero::before{content:'';position:absolute;inset:0;
  background:url("data:image/svg+xml,%3Csvg width='40' height='40' viewBox='0 0 40 40' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='%23ffffff' fill-opacity='0.03'%3E%3Ccircle cx='20' cy='20' r='1.5'/%3E%3C/g%3E%3C/svg%3E")}
.hero-content{position:relative;z-index:1}
.hero-badge{display:inline-flex;align-items:center;gap:6px;
  background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);
  color:#fff;font-size:11px;font-weight:600;letter-spacing:1.5px;
  text-transform:uppercase;padding:4px 14px;border-radius:20px;margin-bottom:18px}
.hero-title{font-size:40px;font-weight:800;color:#fff;line-height:1.1;
  letter-spacing:-1px;margin-bottom:12px}
.hero-sub{font-size:16px;color:rgba(255,255,255,.75);max-width:520px;
  line-height:1.6;margin-bottom:28px}
.hero-features{display:flex;gap:24px;flex-wrap:wrap}
.hf{display:flex;align-items:center;gap:8px;color:rgba(255,255,255,.85);font-size:13px;font-weight:500}
.hf-icon{width:30px;height:30px;background:rgba(255,255,255,.15);border-radius:7px;
  display:flex;align-items:center;justify-content:center;font-size:14px}

/* ── PANELS ── */
.main-wrap{padding:36px 48px}
.panel{background:#fff;border:1px solid #e8eaed;border-radius:14px;
  padding:26px;box-shadow:0 2px 12px rgba(0,0,0,.05);margin-bottom:20px}
.ph{display:flex;align-items:center;gap:12px;margin-bottom:20px}
.ph-icon{width:40px;height:40px;border-radius:10px;
  display:flex;align-items:center;justify-content:center;
  font-size:18px;flex-shrink:0;box-shadow:0 4px 10px rgba(0,0,0,.15)}
.ph-icon.red{background:linear-gradient(135deg,#c0392b,#96281b)}
.ph-icon.blue{background:linear-gradient(135deg,#2563eb,#1d4ed8)}
.ph-title{font-size:17px;font-weight:700;color:#1a1a2e;letter-spacing:-.3px}
.ph-sub{font-size:12px;color:#9ca3af;margin-top:2px}

/* ── DETECT PANEL ── */
.detect-panel{background:#fff;border:1px solid #e8eaed;border-radius:14px;padding:22px}
.detect-title{font-size:16px;font-weight:700;color:#1a1a2e;margin-bottom:14px}
.d-item{display:flex;align-items:center;justify-content:space-between;
  padding:10px 0;border-bottom:1px solid #f3f4f6;font-size:13px;color:#374151;font-weight:500}
.d-item:last-child{border-bottom:none}
.d-left{display:flex;align-items:center;gap:10px}
.d-dot{width:28px;height:28px;border-radius:7px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}

/* ── UPLOAD META ── */
.upload-meta{display:flex;justify-content:space-between;font-size:11px;color:#9ca3af;padding:6px 2px 0}

/* ── ENGINE BADGE ── */
.eng{background:linear-gradient(135deg,#fef9f0,#fef3c7);border:1px solid #fcd34d;
  border-radius:12px;padding:14px 18px;display:flex;align-items:center;
  gap:14px;margin:20px 0}
.eng-title{font-size:13px;font-weight:600;color:#78350f;margin-bottom:2px}
.eng-sub{font-size:12px;color:#92400e}
.secure-badge{margin-left:auto;text-align:center;flex-shrink:0;
  background:#f0fdf4;border:1px solid #86efac;border-radius:8px;padding:8px 14px}

/* ── PREVIEW CARDS ── */
.prev-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.prev-card{border-radius:12px;padding:16px;border:1px solid transparent}
.prev-card.r{background:#fef2f2;border-color:#fecaca}
.prev-card.o{background:#fff7ed;border-color:#fed7aa}
.prev-card.a{background:#fffbeb;border-color:#fde68a}
.prev-card.b{background:#eff6ff;border-color:#bfdbfe}
.pc-icon{font-size:20px;margin-bottom:8px;display:block}
.pc-lbl{font-size:10px;font-weight:600;color:#6b7280;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}
.pc-val{font-size:20px;font-weight:800;font-family:'JetBrains Mono',monospace;letter-spacing:-1px}
.pc-desc{font-size:11px;color:#9ca3af;margin-top:3px}
.r .pc-val{color:#b91c1c}.o .pc-val{color:#c2410c}
.a .pc-val{color:#b45309}.b .pc-val{color:#1d4ed8}

/* ── CAPABILITIES ── */
.cap-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.cap-item{display:flex;align-items:center;gap:10px;padding:9px 12px;
  background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;font-size:12px;font-weight:500;color:#374151}

/* ── RESULT BANNER ── */
.result-banner{background:linear-gradient(135deg,#064e3b,#065f46);border-radius:14px;
  padding:22px 28px;display:flex;align-items:center;justify-content:space-between;
  margin-bottom:24px;box-shadow:0 4px 20px rgba(6,78,59,.2)}
.rb-left h3{font-size:19px;font-weight:700;color:#fff;margin-bottom:4px}
.rb-left p{font-size:13px;color:rgba(255,255,255,.7)}
.rb-stats{display:flex;gap:20px}
.rs{text-align:center}
.rs-num{font-size:24px;font-weight:800;font-family:'JetBrains Mono',monospace;color:#fff;line-height:1}
.rs-lbl{font-size:9px;color:rgba(255,255,255,.55);text-transform:uppercase;letter-spacing:.5px;margin-top:2px}
.rs.r .rs-num{color:#fca5a5}.rs.o .rs-num{color:#fdba74}.rs.a .rs-num{color:#fde047}

/* ── BOTTOM META ── */
.bm{display:grid;grid-template-columns:repeat(4,1fr);
  background:#fff;border:1px solid #e8eaed;border-radius:14px;margin-top:28px;overflow:hidden}
.bm-item{padding:18px 22px;border-right:1px solid #e8eaed}
.bm-item:last-child{border-right:none}
.bm-title{font-size:13px;font-weight:700;color:#374151;margin-bottom:3px;display:flex;align-items:center;gap:6px}
.bm-sub{font-size:11px;color:#9ca3af}

/* ── STREAMLIT OVERRIDES ── */
[data-testid="stFileUploaderDropzone"]{
  border:2.5px dashed #d1d5db!important;border-radius:12px!important;
  background:#fafbfc!important;padding:28px!important;transition:all .25s!important}
[data-testid="stFileUploaderDropzone"]:hover{border-color:#c0392b!important;background:#fef9f9!important}
.stTextInput>div>div>input{
  border:1.5px solid #d1d5db!important;border-radius:10px!important;
  font-family:'JetBrains Mono',monospace!important;font-size:13px!important;
  padding:10px 14px!important;background:#fafbfc!important;color:#1a1a2e!important}
.stTextInput>div>div>input:focus{border-color:#c0392b!important;background:#fff!important;
  box-shadow:0 0 0 3px rgba(192,57,43,.1)!important}
.stButton>button{
  background:linear-gradient(135deg,#c0392b,#96281b)!important;color:#fff!important;
  border:none!important;border-radius:10px!important;font-family:'Outfit',sans-serif!important;
  font-size:15px!important;font-weight:700!important;padding:12px 32px!important;
  width:100%!important;box-shadow:0 4px 14px rgba(192,57,43,.35)!important;
  letter-spacing:.3px!important;transition:all .2s!important}
.stButton>button:hover{transform:translateY(-1px)!important}
.stDownloadButton>button{
  background:linear-gradient(135deg,#c0392b,#96281b)!important;color:#fff!important;
  border:none!important;border-radius:10px!important;font-family:'Outfit',sans-serif!important;
  font-size:14px!important;font-weight:700!important;padding:12px 32px!important;
  width:100%!important;box-shadow:0 4px 14px rgba(192,57,43,.35)!important;
  letter-spacing:.3px!important}
.stDownloadButton>button:hover{transform:translateY(-1px)!important}
label[data-baseweb="form-control-label"]{display:none!important}
</style>
""", unsafe_allow_html=True)

# ── TOPBAR ──────────────────────────────────────────────────
st.markdown("""
<div class="topbar">
  <div class="tb-brand">
    <div class="tb-logo">🔍</div>
    <div>
      <div class="tb-title">SMS Intelligence Dashboard</div>
      <div class="tb-sub">Smart SMS Analysis • Forensic Intelligence</div>
    </div>
  </div>
  <div class="tb-nav">
    <a href="#">ⓘ How It Works</a>
    <a href="#">🛡 Security</a>
    <a href="#">❓ Help</a>
  </div>
</div>
""", unsafe_allow_html=True)

# ── HERO ────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-content">
    <div class="hero-badge">⚡ Forensic • Secure • Intelligent</div>
    <div class="hero-title">SMS Intelligence Dashboard</div>
    <div class="hero-sub">Upload SMS records from Excel files and generate categorized forensic intelligence reports instantly.</div>
    <div class="hero-features">
      <div class="hf"><div class="hf-icon">⚙️</div>Smart Categorization</div>
      <div class="hf"><div class="hf-icon">🏦</div>Financial Transaction Detection</div>
      <div class="hf"><div class="hf-icon">💬</div>Personal Conversation Analysis</div>
      <div class="hf"><div class="hf-icon">📄</div>HTML Forensic Reports</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── MAIN CONTENT ─────────────────────────────────────────────
st.markdown('<div class="main-wrap">', unsafe_allow_html=True)

col_left, col_right = st.columns([2, 1], gap="large")

with col_left:
    # Upload
    st.markdown("""
    <div class="panel">
      <div class="ph">
        <div class="ph-icon red">📤</div>
        <div>
          <div class="ph-title">1. Upload SMS Excel File</div>
          <div class="ph-sub">Upload your SMS dataset in XLSX / CSV format for analysis.</div>
        </div>
      </div>
    </div>""", unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "SMS File", type=["xlsx","xls","csv"],
        help="Supported: .xlsx, .xls, .csv — Max 200MB",
        label_visibility="collapsed"
    )
    st.markdown("""
    <div class="upload-meta">
      <span>✅ Supported: .XLSX / .XLS / .CSV</span>
      <span>ℹ️ Max file size: 200MB</span>
    </div>""", unsafe_allow_html=True)

    # Target number
    st.markdown("""
    <div style="height:18px"></div>
    <div class="panel">
      <div class="ph">
        <div class="ph-icon blue">📱</div>
        <div>
          <div class="ph-title">2. Target Number <span style="font-size:11px;font-weight:500;color:#9ca3af">(Optional)</span></div>
          <div class="ph-sub">Filter and analyze SMS for a specific mobile number.</div>
        </div>
      </div>
    </div>""", unsafe_allow_html=True)

    target_num = st.text_input(
        "Target Number",
        placeholder="e.g. 8801711XXXXXXX or 01711XXXXXXX",
        label_visibility="collapsed"
    )
    st.markdown("""
    <div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;
      padding:8px 12px;font-size:12px;color:#6b7280;margin-top:8px">
      ℹ️ Leave blank to analyze all numbers in the uploaded file.
      When specified, the report will focus on SMS associated with this number.
    </div>""", unsafe_allow_html=True)

    # Engine badge
    st.markdown("""
    <div class="eng">
      <div style="font-size:22px">🔬</div>
      <div>
        <div class="eng-title">Intelligent Analysis Engine</div>
        <div class="eng-sub">Automatically categorizes SMS, detects financial activities, merges
          multi-part messages, identifies SMSMT/SMSMO direction, and generates a comprehensive
          forensic HTML report.</div>
      </div>
      <div class="secure-badge">
        <div style="font-size:18px">🔒</div>
        <div style="font-size:10px;font-weight:700;color:#15803d;letter-spacing:.5px;margin-top:2px">
          100% SECURE<br>LOCAL PROCESSING</div>
      </div>
    </div>""", unsafe_allow_html=True)

    analyze_btn = st.button("🔍  Run Forensic Analysis", use_container_width=True)

with col_right:
    st.markdown("""
    <div class="detect-panel">
      <div class="detect-title">What This Tool Detects</div>
      <div class="d-item"><div class="d-left"><div class="d-dot" style="background:#fef2f2">🔴</div>Financial Transactions (High / Medium / Small)</div>›</div>
      <div class="d-item"><div class="d-left"><div class="d-dot" style="background:#eff6ff">🏦</div>Banking &amp; Mobile Financial Services</div>›</div>
      <div class="d-item"><div class="d-left"><div class="d-dot" style="background:#dcfce7">📶</div>SMSMT (Incoming) vs SMSMO (Outgoing)</div>›</div>
      <div class="d-item"><div class="d-left"><div class="d-dot" style="background:#f5f3ff">💬</div>Personal Conversations</div>›</div>
      <div class="d-item"><div class="d-left"><div class="d-dot" style="background:#eff6ff">✈️</div>Travel, Visa &amp; Immigration SMS</div>›</div>
      <div class="d-item"><div class="d-left"><div class="d-dot" style="background:#f0fdf4">🛒</div>Shopping, Billing &amp; Courier</div>›</div>
      <div class="d-item"><div class="d-left"><div class="d-dot" style="background:#fff7ed">🔗</div>Multi-part (Concatenated) SMS</div>›</div>
    </div>""", unsafe_allow_html=True)

# ── PREVIEW + CAPABILITIES ───────────────────────────────────
st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
col_p, col_c = st.columns([3, 2], gap="large")

with col_p:
    st.markdown("""
    <div style="font-size:17px;font-weight:700;color:#1a1a2e;margin-bottom:4px">Report Preview</div>
    <div style="font-size:12px;color:#9ca3af;margin-bottom:16px">What you'll get in the HTML forensic report.</div>
    <div class="prev-grid">
      <div class="prev-card r"><span class="pc-icon">📈</span>
        <div class="pc-lbl">High Value</div><div class="pc-val">৳ 1,00,000+</div>
        <div class="pc-desc">Critical financial alerts</div></div>
      <div class="prev-card o"><span class="pc-icon">📊</span>
        <div class="pc-lbl">Medium Value</div><div class="pc-val">৳ 50K–1L</div>
        <div class="pc-desc">Important financial alerts</div></div>
      <div class="prev-card a"><span class="pc-icon">💰</span>
        <div class="pc-lbl">Small Value</div><div class="pc-val">৳ 10K–50K</div>
        <div class="pc-desc">Regular financial activities</div></div>
      <div class="prev-card b"><span class="pc-icon">💬</span>
        <div class="pc-lbl">Personal</div><div class="pc-val" style="font-size:15px;padding-top:3px">Merged &<br>Categorized</div>
        <div class="pc-desc">messages</div></div>
    </div>
    <div style="display:flex;gap:20px;margin-top:14px;font-size:11px;color:#9ca3af;
      padding-top:14px;border-top:1px solid #f3f4f6">
      <span>📅 Monthly Grouping</span><span>🔍 Advanced Filtering</span>
      <span>🔎 Keyword Search</span><span>📊 Bank/MFS Summary</span>
    </div>""", unsafe_allow_html=True)

with col_c:
    st.markdown("""
    <div style="font-size:17px;font-weight:700;color:#1a1a2e;margin-bottom:14px">Key Capabilities</div>
    <div class="cap-grid">
      <div class="cap-item"><span>⚙️</span>Automatic SMS Categorization</div>
      <div class="cap-item"><span>💬</span>Personal Conversation Analysis</div>
      <div class="cap-item"><span>🏦</span>Financial Transaction Detection</div>
      <div class="cap-item"><span>🔗</span>Multi-part SMS Merging</div>
      <div class="cap-item"><span>📶</span>SMSMT / SMSMO Direction Tag</div>
      <div class="cap-item"><span>📄</span>Searchable &amp; Filterable Report</div>
      <div class="cap-item"><span>📅</span>Monthly Wise Organization</div>
      <div class="cap-item"><span>⬇️</span>In-report HTML Download Button</div>
    </div>""", unsafe_allow_html=True)

# ── BOTTOM META ──────────────────────────────────────────────
st.markdown("""
<div class="bm">
  <div class="bm-item"><div class="bm-title">🔒 Secure Processing</div>
    <div class="bm-sub">Data is processed locally on your device only.</div></div>
  <div class="bm-item"><div class="bm-title">🚫 No Data Stored</div>
    <div class="bm-sub">We do not store or transmit your data anywhere.</div></div>
  <div class="bm-item"><div class="bm-title">⚡ Fast &amp; Reliable</div>
    <div class="bm-sub">Get forensic results in seconds.</div></div>
  <div class="bm-item"><div class="bm-title" style="color:#c0392b">👤 Developed By</div>
    <div class="bm-sub" style="font-weight:600;color:#374151">Md. Omar Faruk Mazumder</div></div>
</div>
""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ── SESSION STATE init ───────────────────────────────────────
if "result"      not in st.session_state: st.session_state.result      = None
if "html_bytes"  not in st.session_state: st.session_state.html_bytes  = None
if "docx_bytes"  not in st.session_state: st.session_state.docx_bytes  = None
if "dl_name"     not in st.session_state: st.session_state.dl_name     = None
if "file_name"   not in st.session_state: st.session_state.file_name   = None

# ── ANALYSIS ─────────────────────────────────────────────────
if analyze_btn:
    if not uploaded:
        st.error("⚠️ Please upload an Excel or CSV file before running analysis.")
    else:
        with st.spinner("🔬 Analyzing SMS records — please wait..."):
            suffix = os.path.splitext(uploaded.name)[1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded.read()); tmp_in = tmp.name
            tmp_html = tempfile.mktemp(suffix=".html")
            tmp_docx = tempfile.mktemp(suffix=".docx")
            try:
                result = process_file(tmp_in, tmp_html,
                                      docx_out=tmp_docx,
                                      target_number=target_num.strip() if target_num else "")
            except Exception as e:
                st.error(f"❌ Analysis failed: {e}")
                try: os.unlink(tmp_in)
                except: pass
                st.stop()
            try: os.unlink(tmp_in)
            except: pass

        base_name = os.path.splitext(uploaded.name)[0]
        tn_sfx    = f"_{result.get('target_number','')}" if result.get("target_number") else ""
        dl_name   = f"{base_name}{tn_sfx}_intelligence_report.html"

        html_bytes = None
        if os.path.exists(tmp_html):
            with open(tmp_html,"rb") as f: html_bytes = f.read()
            try: os.unlink(tmp_html)
            except: pass

        docx_bytes = None
        if os.path.exists(tmp_docx):
            with open(tmp_docx,"rb") as f: docx_bytes = f.read()
            try: os.unlink(tmp_docx)
            except: pass

        # session_state-এ save — download button click করলেও হারাবে না
        st.session_state.result     = result
        st.session_state.html_bytes = html_bytes
        st.session_state.docx_bytes = docx_bytes
        st.session_state.dl_name    = dl_name
        st.session_state.file_name  = uploaded.name

# ── RESULT (session_state থেকে — সবসময় দেখাবে) ──────────────
if st.session_state.result:
    result     = st.session_state.result
    html_bytes = st.session_state.html_bytes
    docx_bytes = st.session_state.docx_bytes
    dl_name    = st.session_state.dl_name
    fname      = st.session_state.file_name

    tn_display = result.get("target_number","") or "All Numbers"

    # Result Banner
    st.markdown('<div class="main-wrap" style="padding-top:0">', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="result-banner">
      <div class="rb-left">
        <h3>✅ Analysis Complete</h3>
        <p>Target: <strong style="color:#fff">{tn_display}</strong>
           &nbsp;·&nbsp; File: <strong style="color:#fff">{fname}</strong></p>
      </div>
      <div class="rb-stats">
        <div class="rs r"><div class="rs-num">{result.get("high",0)}</div><div class="rs-lbl">High (100K+)</div></div>
        <div class="rs o"><div class="rs-num">{result.get("med",0)}</div><div class="rs-lbl">Medium (50K)</div></div>
        <div class="rs a"><div class="rs-num">{result.get("small",0)}</div><div class="rs-lbl">Small (10K)</div></div>
        <div class="rs"><div class="rs-num">{result.get("total_shown",0)}</div><div class="rs-lbl">Total Shown</div></div>
        <div class="rs"><div class="rs-num">{result.get("total_sms",0)}</div><div class="rs-lbl">Total SMS</div></div>
      </div>
    </div>""", unsafe_allow_html=True)

    # Download buttons — session_state-এ আছে তাই reload-এও থাকবে
    if html_bytes:
        if docx_bytes:
            c1, c2, c3 = st.columns([1, 1, 1])
            with c1:
                st.download_button(
                    label="⬇️  HTML Report",
                    data=html_bytes,
                    file_name=dl_name,
                    mime="text/html",
                    use_container_width=True,
                    key="dl_html",
                )
            with c2:
                st.download_button(
                    label="📄  Word (.docx) Report",
                    data=docx_bytes,
                    file_name=dl_name.replace(".html", ".docx"),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_docx",
                )
        else:
            c1, c2, c3 = st.columns([1,2,1])
            with c2:
                st.download_button(
                    label="⬇️   Download HTML Intelligence Report",
                    data=html_bytes,
                    file_name=dl_name,
                    mime="text/html",
                    use_container_width=True,
                    key="dl_html_only",
                )

    st.markdown("""
    <div style="text-align:center;font-size:12px;color:#9ca3af;margin-top:10px">
      ℹ️ The HTML report also contains a built-in Download button for offline saving.
    </div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)
