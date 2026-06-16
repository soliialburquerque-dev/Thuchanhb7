import os
import json
import io
import streamlit as st
import streamlit.components.v1 as components
import xlsxwriter
from scorer import calculate_score as local_calculate_score, run_scoring, JSON_PATH, CSV_PATH, SHEET_CSV_URL

# ─── Page Config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AI Lead Scoring Console",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Theme State ─────────────────────────────────────────────────────────────
if "theme" not in st.session_state:
    st.session_state.theme = "dark"

def toggle_theme():
    st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"

IS_DARK = st.session_state.theme == "dark"

# ─── CSS Variables ───────────────────────────────────────────────────────────
BG          = "#090d16"     if IS_DARK else "#f8fafc"
BG_SUBTLE   = "#111827"     if IS_DARK else "#f1f5f9"
CARD        = "rgba(19,29,48,0.85)"  if IS_DARK else "rgba(255,255,255,0.95)"
BORDER      = "rgba(255,255,255,0.09)" if IS_DARK else "rgba(0,0,0,0.09)"
TEXT        = "#f8fafc"     if IS_DARK else "#0f172a"
TEXT_MUTED  = "#94a3b8"     if IS_DARK else "#475569"
TEXT_DIM    = "#64748b"     if IS_DARK else "#94a3b8"
TH_BG       = "rgba(255,255,255,0.03)" if IS_DARK else "rgba(0,0,0,0.03)"
ROW_HOVER   = "rgba(0,242,254,0.05)"  if IS_DARK else "rgba(0,120,200,0.06)"

st.markdown(f"""
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root {{
    --bg: {BG};
    --bg-subtle: {BG_SUBTLE};
    --card: {CARD};
    --border: {BORDER};
    --text: {TEXT};
    --text-muted: {TEXT_MUTED};
    --text-dim: {TEXT_DIM};
    --accent: #00f2fe;
    --vip-color: #10b981;  --vip-bg: rgba(16,185,129,0.12); --vip-border: rgba(16,185,129,0.3);
    --normal-color: #f59e0b; --normal-bg: rgba(245,158,11,0.12); --normal-border: rgba(245,158,11,0.3);
    --spam-color: #ef4444;  --spam-bg: rgba(239,68,68,0.12); --spam-border: rgba(239,68,68,0.3);
}}
html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"], .main, .block-container {{
    background-color: var(--bg) !important;
    color: var(--text) !important;
    font-family: 'DM Sans', 'Segoe UI', sans-serif !important;
}}
.block-container {{ padding: 2rem 3rem 3rem !important; max-width: 1440px !important; }}
header[data-testid="stHeader"], #MainMenu, footer, [data-testid="stToolbar"],
[data-testid="stDecoration"], [data-testid="stStatusWidget"], .stDeployButton {{ display: none !important; }}

/* Subheader override */
h2, h3 {{ color: var(--text) !important; }}

/* Sidebar */
[data-testid="stSidebar"] {{ background: var(--bg-subtle) !important; border-right: 1px solid var(--border) !important; }}
[data-testid="stSidebar"] * {{ color: var(--text) !important; }}
[data-testid="stSidebar"] button {{ background: var(--card) !important; border: 1px solid var(--border) !important; color: var(--text) !important; }}

/* Inputs */
[data-testid="stTextInput"] input, [data-testid="stSelectbox"] select,
.stSelectbox > div > div {{ background: var(--card) !important; color: var(--text) !important; border-color: var(--border) !important; }}
.stTextInput label, .stSelectbox label, .stNumberInput label {{ color: var(--text-muted) !important; }}

/* Buttons */
.stButton > button, .stFormSubmitButton > button {{
    background: linear-gradient(135deg, #00f2fe, #4facfe) !important;
    color: #0f172a !important; font-weight: 700 !important;
    border: none !important; border-radius: 8px !important;
    transition: opacity .2s !important;
}}
.stButton > button:hover, .stFormSubmitButton > button:hover {{ opacity: 0.85 !important; }}

/* Slider */
[data-testid="stSlider"] [data-testid="stMarkdownContainer"] p {{ color: var(--text) !important; }}

/* Selectbox in form */
[data-testid="stForm"] [data-testid="stSelectbox"] > div > div {{ background: var(--card) !important; color: var(--text) !important; }}

/* Checkbox */
[data-testid="stCheckbox"] label {{ color: var(--text) !important; }}

/* Text area */
.stTextArea textarea {{ background: var(--card) !important; color: var(--text) !important; border-color: var(--border) !important; }}

/* KPI Cards */
.kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(190px,1fr)); gap: 1.2rem; margin-bottom: 2rem; }}
.kpi-card {{
    background: {CARD}; backdrop-filter: blur(12px);
    border: 1px solid {BORDER}; border-radius: 14px;
    padding: 1.3rem 1.5rem; box-shadow: 0 4px 24px rgba(0,0,0,0.18);
    transition: transform .2s, box-shadow .2s;
}}
.kpi-card:hover {{ transform: translateY(-2px); box-shadow: 0 8px 32px rgba(0,0,0,0.28); }}
.kpi-label {{ font-size: .72rem; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: {TEXT_MUTED}; }}
.kpi-value {{ font-size: 2rem; font-weight: 800; margin-top: .2rem; color: {TEXT}; }}

/* Badge styles */
.badge {{ display:inline-block; padding:3px 10px; border-radius:9999px; font-size:.68rem; font-weight:700; letter-spacing:.04em; text-transform:uppercase; }}
.bvip   {{ color:#10b981; background:rgba(16,185,129,.15); border:1px solid rgba(16,185,129,.35); }}
.bnorm  {{ color:#f59e0b; background:rgba(245,158,11,.15); border:1px solid rgba(245,158,11,.35); }}
.bspam  {{ color:#ef4444; background:rgba(239,68,68,.15);  border:1px solid rgba(239,68,68,.35); }}
.brev   {{ color:#a855f7; background:rgba(168,85,247,.15); border:1px solid rgba(168,85,247,.3); }}
.bunrev {{ color:{TEXT_DIM}; background:rgba(255,255,255,.05); border:1px solid {BORDER}; }}

/* Detail box */
.detail-box {{
    background: {"rgba(0,0,0,0.25)" if IS_DARK else "rgba(0,0,0,0.05)"};
    border: 1px solid {BORDER}; border-radius: 8px;
    padding: .9rem 1rem; font-size: .87rem; line-height: 1.6; margin-bottom: 1rem;
    color: {TEXT};
}}
.reason-tag {{
    display: inline-block; font-size: .7rem; padding: 3px 8px;
    border-radius: 5px; margin: 3px 5px 3px 0;
}}
.rpos {{ color: #a7f3d0; background: rgba(16,185,129,.1);  border: 1px solid rgba(16,185,129,.25); }}
.rneg {{ color: #fecaca; background: rgba(239,68,68,.1);   border: 1px solid rgba(239,68,68,.25); }}
.rnet {{ color: #fde68a; background: rgba(245,158,11,.1);  border: 1px solid rgba(245,158,11,.25); }}

/* Brand header */
.brand-header {{ display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid {BORDER}; padding-bottom:1rem; margin-bottom:2rem; }}
.brand-name {{ font-size:1.6rem; font-weight:800; color:{TEXT}; }}
.brand-sub  {{ font-size:.78rem; color:{TEXT_MUTED}; margin-top:.1rem; }}
</style>
""", unsafe_allow_html=True)


# ─── Data Helpers ─────────────────────────────────────────────────────────────
def load_leads():
    if not os.path.exists(JSON_PATH):
        run_scoring()
    with open(JSON_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def save_leads(data):
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def generate_excel(leads_list):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output)
    ws = wb.add_worksheet("Leads Scored")
    hfmt = wb.add_format({'bold':True,'font_name':'Segoe UI','font_size':11,'font_color':'white','bg_color':'#1F4E78','border':1,'align':'center','valign':'vcenter'})
    cfmt = wb.add_format({'font_name':'Segoe UI','font_size':10,'border':1,'valign':'vcenter'})
    mfmt = wb.add_format({'font_name':'Segoe UI','font_size':10,'border':1,'align':'center','valign':'vcenter'})
    vfmt = wb.add_format({'font_name':'Segoe UI','font_size':10,'bold':True,'font_color':'#375623','bg_color':'#E2EFDA','border':1,'align':'center','valign':'vcenter'})
    nfmt = wb.add_format({'font_name':'Segoe UI','font_size':10,'bold':True,'font_color':'#7F6000','bg_color':'#FFF2CC','border':1,'align':'center','valign':'vcenter'})
    sfmt = wb.add_format({'font_name':'Segoe UI','font_size':10,'bold':True,'font_color':'#C00000','bg_color':'#FCE4D6','border':1,'align':'center','valign':'vcenter'})
    ws.set_column('A:A',8); ws.set_column('B:B',22); ws.set_column('C:C',15)
    ws.set_column('D:D',65); ws.set_column('E:E',12); ws.set_column('F:F',15)
    ws.set_column('G:G',14); ws.set_column('H:H',28); ws.set_column('I:I',55)
    ws.set_row(0, 32)
    headers = ["ID","Khách hàng","Số điện thoại","Nhu cầu mô tả","Điểm số","Trạng thái","Đã duyệt","Ghi chú","Lý do hệ thống"]
    for ci, h in enumerate(headers):
        ws.write(0, ci, h, hfmt)
    for ri, lead in enumerate(leads_list, 1):
        ws.set_row(ri, 22)
        ws.write(ri, 0, lead["id"], mfmt)
        ws.write(ri, 1, lead["ten_khach"], cfmt)
        ws.write(ri, 2, lead["sdt"], mfmt)
        ws.write(ri, 3, lead["nhu_cau_mo_ta"], cfmt)
        ws.write(ri, 4, lead["score"], mfmt)
        st_fmt = vfmt if lead["status"]=="VIP" else (sfmt if lead["status"]=="SPAM" else nfmt)
        ws.write(ri, 5, lead["status"], st_fmt)
        ws.write(ri, 6, "Đã duyệt" if lead.get("reviewed") else "Chưa duyệt", mfmt)
        ws.write(ri, 7, lead.get("notes",""), cfmt)
        ws.write(ri, 8, lead.get("reason",""), cfmt)
    wb.close()
    return output.getvalue()


# ─── Brand Header ─────────────────────────────────────────────────────────────
hl, hr = st.columns([10, 2])
with hl:
    st.markdown(f"""
    <div class="brand-header">
        <div>
            <div class="brand-name">🏢 AI Lead Scoring Console</div>
            <div class="brand-sub">Real Estate · Automated Lead Processing Dashboard</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
with hr:
    lbl = "☀️ Light Mode" if IS_DARK else "🌙 Dark Mode"
    st.button(lbl, on_click=toggle_theme, use_container_width=True)

# ─── Load Data ────────────────────────────────────────────────────────────────
leads_list = load_leads()
total      = len(leads_list)
vip_cnt    = sum(1 for l in leads_list if l["status"] == "VIP")
norm_cnt   = sum(1 for l in leads_list if l["status"] == "NORMAL")
spam_cnt   = sum(1 for l in leads_list if l["status"] == "SPAM")
rev_cnt    = sum(1 for l in leads_list if l.get("reviewed"))

# ─── KPI Cards ───────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card">
    <div class="kpi-label">Tổng khách hàng</div>
    <div class="kpi-value">{total}</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label" style="color:var(--vip-color);">Siêu tiềm năng (VIP)</div>
    <div class="kpi-value" style="color:var(--vip-color);">{vip_cnt}</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label" style="color:var(--normal-color);">Khách tiềm năng</div>
    <div class="kpi-value" style="color:var(--normal-color);">{norm_cnt}</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label" style="color:var(--spam-color);">Spam / Không tiềm năng</div>
    <div class="kpi-value" style="color:var(--spam-color);">{spam_cnt}</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">Đã kiểm duyệt</div>
    <div class="kpi-value">{rev_cnt} <span style="font-size:1rem;color:var(--text-muted);">/ {total}</span></div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─── Main Layout ──────────────────────────────────────────────────────────────
col_left, col_right = st.columns([7, 3], gap="large")

with col_left:
    st.subheader("📋 Danh sách khách hàng")

    tf1, tf2, tf3 = st.columns([4, 3, 3])
    with tf1:
        search_q = st.text_input("🔍 Tìm kiếm tên / SĐT / nhu cầu...", placeholder="Nhập từ khóa...")
    with tf2:
        status_sel = st.selectbox("Lọc trạng thái:", ["Tất cả", "VIP (≥ 80đ)", "NORMAL (30-79đ)", "SPAM (< 30đ)"])
    with tf3:
        review_sel = st.selectbox("Lọc kiểm duyệt:", ["Tất cả", "Đã duyệt", "Chưa duyệt"])

    # Filtering
    filtered = leads_list
    if search_q:
        sq = search_q.lower()
        filtered = [l for l in filtered if sq in l["ten_khach"].lower() or sq in l["sdt"].lower() or sq in l["nhu_cau_mo_ta"].lower()]
    if status_sel != "Tất cả":
        tgt = "VIP" if "VIP" in status_sel else ("SPAM" if "SPAM" in status_sel else "NORMAL")
        filtered = [l for l in filtered if l["status"] == tgt]
    if review_sel != "Tất cả":
        want_rev = review_sel == "Đã duyệt"
        filtered = [l for l in filtered if bool(l.get("reviewed")) == want_rev]

    # Pagination
    PER_PAGE = 15
    total_f  = len(filtered)
    n_pages  = max(1, (total_f + PER_PAGE - 1) // PER_PAGE)
    _, pc, _ = st.columns([4, 4, 4])
    with pc:
        page_no = st.number_input("Trang", min_value=1, max_value=n_pages, value=1, step=1)

    start = (page_no - 1) * PER_PAGE
    end   = min(start + PER_PAGE, total_f)
    page_leads = filtered[start:end]

    # Build HTML table with inline styles (no external CSS class dependency)
    rows_html = ""
    for lead in page_leads:
        s = lead["score"]
        score_color = "#10b981" if s >= 80 else ("#ef4444" if s < 30 else "#f59e0b")

        if lead["status"] == "VIP":
            badge_st = '<span style="display:inline-block;padding:3px 10px;border-radius:9999px;font-size:.68rem;font-weight:700;color:#10b981;background:rgba(16,185,129,.15);border:1px solid rgba(16,185,129,.35);">VIP</span>'
        elif lead["status"] == "SPAM":
            badge_st = '<span style="display:inline-block;padding:3px 10px;border-radius:9999px;font-size:.68rem;font-weight:700;color:#ef4444;background:rgba(239,68,68,.15);border:1px solid rgba(239,68,68,.35);">SPAM</span>'
        else:
            badge_st = '<span style="display:inline-block;padding:3px 10px;border-radius:9999px;font-size:.68rem;font-weight:700;color:#f59e0b;background:rgba(245,158,11,.15);border:1px solid rgba(245,158,11,.35);">NORMAL</span>'

        if lead.get("reviewed"):
            badge_rv = '<span style="display:inline-block;padding:3px 10px;border-radius:9999px;font-size:.68rem;font-weight:700;color:#a855f7;background:rgba(168,85,247,.15);border:1px solid rgba(168,85,247,.3);">✓ Đã duyệt</span>'
        else:
            badge_rv = '<span style="display:inline-block;padding:3px 10px;border-radius:9999px;font-size:.68rem;font-weight:700;color:#64748b;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.09);">Chưa duyệt</span>'

        desc = str(lead["nhu_cau_mo_ta"]).replace("<","&lt;").replace(">","&gt;")
        name = str(lead["ten_khach"]).replace("<","&lt;").replace(">","&gt;")

        rows_html += f"""
        <tr style="border-bottom:1px solid rgba(255,255,255,0.07);">
            <td style="padding:.75rem .9rem;font-family:monospace;text-align:center;font-weight:700;color:#94a3b8;">{lead['id']}</td>
            <td style="padding:.75rem .9rem;font-weight:600;">{name}</td>
            <td style="padding:.75rem .9rem;font-family:monospace;text-align:center;color:#94a3b8;">{lead['sdt']}</td>
            <td style="padding:.75rem .9rem;max-width:360px;"><div style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:360px;color:#94a3b8;">{desc}</div></td>
            <td style="padding:.75rem .9rem;text-align:center;font-family:monospace;font-weight:800;font-size:1rem;color:{score_color};">{s}đ</td>
            <td style="padding:.75rem .9rem;text-align:center;">{badge_st}</td>
            <td style="padding:.75rem .9rem;text-align:center;">{badge_rv}</td>
        </tr>"""

    if not rows_html:
        rows_html = '<tr><td colspan="7" style="padding:48px;text-align:center;color:#64748b;">Không tìm thấy khách hàng nào.</td></tr>'

    table_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
  body {{ margin:0; background:transparent; font-family:'DM Sans',sans-serif; color:#f8fafc; }}
  table {{ width:100%; border-collapse:collapse; background:rgba(19,29,48,0.85); border:1px solid rgba(255,255,255,0.09); border-radius:12px; overflow:hidden; font-size:.85rem; }}
  thead tr {{ background:rgba(255,255,255,0.04); }}
  th {{ padding:.8rem .9rem; text-align:left; color:#64748b; font-size:.72rem; font-weight:700; text-transform:uppercase; letter-spacing:.06em; border-bottom:1px solid rgba(255,255,255,0.09); white-space:nowrap; }}
  tbody tr {{ transition:background .15s; }}
  tbody tr:hover {{ background:rgba(0,242,254,0.06); }}
  .footer {{ font-size:.72rem; color:#64748b; text-align:right; margin-top:.6rem; padding-right:.4rem; }}
</style>
</head>
<body>
<table>
  <thead>
    <tr>
      <th style="width:56px;text-align:center;">ID</th>
      <th style="width:160px;">Khách hàng</th>
      <th style="width:120px;text-align:center;">Số điện thoại</th>
      <th>Nhu cầu chi tiết</th>
      <th style="width:72px;text-align:center;">Điểm</th>
      <th style="width:100px;text-align:center;">Trạng thái</th>
      <th style="width:110px;text-align:center;">Duyệt</th>
    </tr>
  </thead>
  <tbody>
    {rows_html}
  </tbody>
</table>
<div class="footer">Hiển thị {start+1 if total_f > 0 else 0} – {end} trong số {total_f} leads được lọc</div>
</body>
</html>"""

    # Use st.components to render table properly
    components.html(table_html, height=min(600, 80 + len(page_leads) * 50), scrolling=False)


with col_right:
    st.subheader("✏️ Duyệt kết quả")

    all_ids = [l["id"] for l in filtered]
    if not all_ids:
        st.info("Không có leads nào trong danh sách được lọc.")
    else:
        if "sel_id" not in st.session_state or st.session_state.sel_id not in all_ids:
            st.session_state.sel_id = all_ids[0]

        sel_id = st.selectbox("Chọn ID khách hàng:", all_ids, index=all_ids.index(st.session_state.sel_id))
        st.session_state.sel_id = sel_id

        lead_obj = next(l for l in leads_list if l["id"] == sel_id)

        name_safe = str(lead_obj["ten_khach"]).replace("<","&lt;").replace(">","&gt;")
        desc_safe = str(lead_obj["nhu_cau_mo_ta"]).replace("<","&lt;").replace(">","&gt;")

        s = lead_obj["score"]
        score_color = "#10b981" if s >= 80 else ("#ef4444" if s < 30 else "#f59e0b")

        st.markdown(f"""
        <div style="background:{CARD};border:1px solid {BORDER};border-radius:12px;padding:1.25rem;box-shadow:0 4px 16px rgba(0,0,0,0.18);margin-bottom:.8rem;">
            <div style="font-size:.68rem;color:{TEXT_DIM};font-weight:700;text-transform:uppercase;letter-spacing:.06em;margin-bottom:.4rem;">Chi tiết khách hàng</div>
            <div style="font-size:1.2rem;font-weight:800;color:{TEXT};">{name_safe}</div>
            <div style="font-family:monospace;font-size:.82rem;color:{TEXT_MUTED};margin-bottom:1rem;">SĐT: {lead_obj['sdt']} &nbsp;|&nbsp; ID: #{lead_obj['id']} &nbsp;|&nbsp; <span style="color:{score_color};font-weight:800;">{s}đ</span></div>
            <div style="font-size:.68rem;color:{TEXT_DIM};font-weight:700;text-transform:uppercase;letter-spacing:.06em;margin-bottom:.4rem;">Nhu cầu mô tả:</div>
            <div class="detail-box">{desc_safe}</div>
        </div>
        """, unsafe_allow_html=True)

        # Reasons
        reasons_html = ""
        if lead_obj.get("reason"):
            for r in lead_obj["reason"].split(";"):
                r = r.strip()
                if not r: continue
                cls = "rpos" if r.startswith("+") else ("rneg" if r.startswith("-") else "rnet")
                color_map = {"rpos":"#a7f3d0","rneg":"#fecaca","rnet":"#fde68a"}
                bg_map    = {"rpos":"rgba(16,185,129,.1)","rneg":"rgba(239,68,68,.1)","rnet":"rgba(245,158,11,.1)"}
                bd_map    = {"rpos":"rgba(16,185,129,.25)","rneg":"rgba(239,68,68,.25)","rnet":"rgba(245,158,11,.25)"}
                r_safe = str(r).replace("<","&lt;").replace(">","&gt;")
                reasons_html += f'<span style="display:inline-block;font-size:.7rem;padding:3px 8px;border-radius:5px;margin:3px 5px 3px 0;color:{color_map[cls]};background:{bg_map[cls]};border:1px solid {bd_map[cls]};">{r_safe}</span>'

        if reasons_html:
            st.markdown(f"""
            <div style="margin-bottom:1.2rem;">
                <div style="font-size:.68rem;color:{TEXT_DIM};font-weight:700;text-transform:uppercase;letter-spacing:.06em;margin-bottom:.5rem;">Lý do chấm điểm tự động:</div>
                <div>{reasons_html}</div>
            </div>
            """, unsafe_allow_html=True)

        # Review Form
        with st.form("review_form"):
            f_score = st.slider("Điều chỉnh điểm (0–100):", 0, 100, int(lead_obj["score"]))
            def_idx = 0 if lead_obj["status"] == "VIP" else (2 if lead_obj["status"] == "SPAM" else 1)
            f_status = st.selectbox("Phân loại:", ["VIP", "NORMAL", "SPAM"], index=def_idx)
            f_notes = st.text_area("Ghi chú kiểm duyệt:", value=lead_obj.get("notes", ""), height=90)
            f_reviewed = st.checkbox("✅ Đánh dấu đã duyệt", value=bool(lead_obj.get("reviewed")))
            submitted = st.form_submit_button("💾 Lưu kết quả duyệt", use_container_width=True)
            if submitted:
                lead_obj["score"]    = f_score
                lead_obj["status"]   = f_status
                lead_obj["notes"]    = f_notes
                lead_obj["reviewed"] = f_reviewed
                if not lead_obj.get("reason","").startswith("[Manual]"):
                    lead_obj["reason"] = f"[Manual] {lead_obj.get('reason','')}"
                save_leads(leads_list)
                st.success(f"✅ Đã lưu kết quả duyệt cho KH #{sel_id}!")
                st.rerun()


# ─── Sidebar ──────────────────────────────────────────────────────────────────
st.sidebar.header("⚙️ Cài đặt & Xuất dữ liệu")

excel_bytes = generate_excel(leads_list)
st.sidebar.download_button(
    label="📥 Tải Báo Cáo Excel",
    data=excel_bytes,
    file_name="Leads_Tiem_Nang_Da_Cham_Diem.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.sidebar.markdown("---")
st.sidebar.subheader("🔄 Chấm điểm lại")
recalc_key = st.sidebar.text_input("Gemini API Key (Tùy chọn):", type="password",
    help="Để trống để dùng thuật toán Keyword Regex cục bộ")
recalc_overwrite = st.sidebar.checkbox("Ghi đè leads đã duyệt thủ công")

if st.sidebar.button("🚀 Chấm điểm lại toàn bộ", use_container_width=True):
    with st.spinner("Đang chấm điểm lại 500 khách hàng..."):
        try:
            ok = run_scoring(api_key=recalc_key if recalc_key else None, overwrite=recalc_overwrite)
            if ok:
                st.sidebar.success("✅ Hoàn tất chấm điểm lại!")
                st.rerun()
            else:
                st.sidebar.error("❌ Có lỗi xảy ra.")
        except Exception as e:
            st.sidebar.error(f"Lỗi: {e}")

st.sidebar.markdown("---")
st.sidebar.markdown(f"""
<div style="font-size:.75rem;color:{TEXT_MUTED};line-height:1.8;">
    <b>Bộ tiêu chí chấm điểm:</b><br>
    🟢 VIP: ≥ 80 điểm<br>
    🟡 NORMAL: 30–79 điểm<br>
    🔴 SPAM: &lt; 30 điểm<br><br>
    <b>Nguồn dữ liệu:</b><br>
    Google Sheet (500 leads)<br>
    Chấm điểm: Rule-Based AI Engine
</div>
""", unsafe_allow_html=True)
