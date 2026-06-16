import os
import json
import io
import streamlit as st
import xlsxwriter
from scorer import calculate_score as local_calculate_score, run_scoring, JSON_PATH, CSV_PATH, SHEET_CSV_URL

# ─── Page Config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AI Lead Scoring Console",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Theme ───────────────────────────────────────────────────────────────────
if "theme" not in st.session_state:
    st.session_state.theme = "dark"

def toggle_theme():
    st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"

IS_DARK = st.session_state.theme == "dark"

# ─── Global CSS via st.html (works in Streamlit 1.36+) ────────────────────────
BG         = "#090d16"    if IS_DARK else "#f5f7fa"
CARD       = "#131d30"    if IS_DARK else "#ffffff"
BORDER     = "rgba(255,255,255,0.09)" if IS_DARK else "rgba(0,0,0,0.09)"
TEXT       = "#f1f5f9"    if IS_DARK else "#0f172a"
MUTED      = "#94a3b8"    if IS_DARK else "#475569"
DIM        = "#64748b"    if IS_DARK else "#94a3b8"

st.html(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');

section[data-testid="stMain"] > div,
[data-testid="stAppViewContainer"],
[data-testid="stApp"] {{
    background-color: {BG} !important;
    font-family: 'DM Sans', sans-serif !important;
}}
.block-container {{
    padding: 1.5rem 2.5rem 3rem !important;
    max-width: 1500px !important;
    background-color: {BG} !important;
}}
header[data-testid="stHeader"] {{ display: none !important; }}
footer {{ display: none !important; }}
#MainMenu {{ display: none !important; }}
[data-testid="stToolbar"] {{ display: none !important; }}
[data-testid="stDecoration"] {{ display: none !important; }}

/* Text colors */
p, span, label, div, h1, h2, h3, h4 {{
    color: {TEXT} !important;
    font-family: 'DM Sans', sans-serif !important;
}}

/* Sidebar */
[data-testid="stSidebar"] {{
    background-color: {CARD} !important;
    border-right: 1px solid {BORDER} !important;
}}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] div {{
    color: {TEXT} !important;
}}

/* Buttons */
.stButton > button {{
    background: linear-gradient(135deg, #00f2fe 0%, #4facfe 100%) !important;
    color: #0a0f1e !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 8px !important;
    transition: opacity 0.2s ease !important;
    font-family: 'DM Sans', sans-serif !important;
}}
.stButton > button:hover {{ opacity: 0.88 !important; }}

/* Form submit button */
[data-testid="stFormSubmitButton"] > button {{
    background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
    color: white !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 8px !important;
}}

/* Inputs */
[data-testid="stTextInput"] input {{
    background-color: {CARD} !important;
    color: {TEXT} !important;
    border-color: {BORDER} !important;
    border-radius: 8px !important;
}}
.stSelectbox [data-testid="stSelectbox"] > div,
[data-testid="stSelectbox"] > div > div {{
    background-color: {CARD} !important;
    color: {TEXT} !important;
}}
.stTextArea textarea {{
    background-color: {CARD} !important;
    color: {TEXT} !important;
    border-color: {BORDER} !important;
}}
[data-testid="stNumberInput"] input {{
    background-color: {CARD} !important;
    color: {TEXT} !important;
    border-color: {BORDER} !important;
}}

/* Download button */
[data-testid="stDownloadButton"] > button {{
    background: {CARD} !important;
    color: {TEXT} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 8px !important;
}}
</style>
""")


# ─── Helpers ─────────────────────────────────────────────────────────────────
def load_leads():
    if not os.path.exists(JSON_PATH):
        run_scoring()
    with open(JSON_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def save_leads(data):
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def score_color(s):
    return "#10b981" if s >= 80 else ("#ef4444" if s < 30 else "#f59e0b")

def badge_status(status):
    cfg = {
        "VIP":    ("#10b981","rgba(16,185,129,.18)","rgba(16,185,129,.4)"),
        "NORMAL": ("#f59e0b","rgba(245,158,11,.18)","rgba(245,158,11,.4)"),
        "SPAM":   ("#ef4444","rgba(239,68,68,.18)", "rgba(239,68,68,.4)"),
    }
    c,bg,bd = cfg.get(status, ("#94a3b8","rgba(148,163,184,.1)","rgba(148,163,184,.3)"))
    return f'<span style="display:inline-block;padding:3px 11px;border-radius:9999px;font-size:.68rem;font-weight:700;letter-spacing:.04em;color:{c};background:{bg};border:1px solid {bd};">{status}</span>'

def badge_review(reviewed):
    if reviewed:
        return '<span style="display:inline-block;padding:3px 11px;border-radius:9999px;font-size:.68rem;font-weight:700;color:#a855f7;background:rgba(168,85,247,.18);border:1px solid rgba(168,85,247,.4);">✓ Đã duyệt</span>'
    return '<span style="display:inline-block;padding:3px 11px;border-radius:9999px;font-size:.68rem;font-weight:700;color:#64748b;background:rgba(100,116,139,.12);border:1px solid rgba(100,116,139,.25);">Chưa duyệt</span>'

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
    for col,w in [('A:A',8),('B:B',22),('C:C',15),('D:D',65),('E:E',12),('F:F',15),('G:G',14),('H:H',28),('I:I',55)]:
        ws.set_column(col, w)
    ws.set_row(0, 32)
    headers = ["ID","Khách hàng","Số điện thoại","Nhu cầu mô tả","Điểm số","Trạng thái","Đã duyệt","Ghi chú","Lý do hệ thống"]
    for ci, h in enumerate(headers): ws.write(0, ci, h, hfmt)
    for ri, lead in enumerate(leads_list, 1):
        ws.set_row(ri, 22)
        ws.write(ri, 0, lead["id"], mfmt)
        ws.write(ri, 1, lead["ten_khach"], cfmt)
        ws.write(ri, 2, lead["sdt"], mfmt)
        ws.write(ri, 3, lead["nhu_cau_mo_ta"], cfmt)
        ws.write(ri, 4, lead["score"], mfmt)
        stfmt = vfmt if lead["status"]=="VIP" else (sfmt if lead["status"]=="SPAM" else nfmt)
        ws.write(ri, 5, lead["status"], stfmt)
        ws.write(ri, 6, "Đã duyệt" if lead.get("reviewed") else "Chưa duyệt", mfmt)
        ws.write(ri, 7, lead.get("notes",""), cfmt)
        ws.write(ri, 8, lead.get("reason",""), cfmt)
    wb.close()
    return output.getvalue()


# ─── Load Data ────────────────────────────────────────────────────────────────
leads_list = load_leads()
total   = len(leads_list)
vip_cnt = sum(1 for l in leads_list if l["status"] == "VIP")
nor_cnt = sum(1 for l in leads_list if l["status"] == "NORMAL")
spm_cnt = sum(1 for l in leads_list if l["status"] == "SPAM")
rev_cnt = sum(1 for l in leads_list if l.get("reviewed"))


# ─── Header ──────────────────────────────────────────────────────────────────
hl, hr = st.columns([10, 2])
with hl:
    st.html(f"""
    <div style="display:flex;align-items:center;gap:12px;padding-bottom:1rem;border-bottom:1px solid {BORDER};margin-bottom:1.5rem;">
        <span style="font-size:2rem;">🏢</span>
        <div>
            <div style="font-size:1.55rem;font-weight:800;color:{TEXT};font-family:'DM Sans',sans-serif;line-height:1.2;">AI Lead Scoring Console</div>
            <div style="font-size:.78rem;color:{MUTED};font-family:'DM Sans',sans-serif;margin-top:2px;">Real Estate · Automated Lead Processing Dashboard</div>
        </div>
    </div>
    """)
with hr:
    lbl = "☀️ Light" if IS_DARK else "🌙 Dark"
    st.button(lbl, on_click=toggle_theme, use_container_width=True)


# ─── KPI Cards (via components.html for full CSS control) ────────────────────
kpi_bg   = "#131d30" if IS_DARK else "#ffffff"
kpi_sh   = "0 4px 24px rgba(0,0,0,0.25)" if IS_DARK else "0 2px 12px rgba(0,0,0,0.08)"
kpi_html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@600;800&display=swap" rel="stylesheet">
<style>
  body {{ margin:0; background:transparent; font-family:'DM Sans',sans-serif; }}
  .grid {{ display:grid; grid-template-columns:repeat(5,1fr); gap:16px; }}
  .card {{ background:{kpi_bg}; border:1px solid {BORDER}; border-radius:14px; padding:18px 20px;
           box-shadow:{kpi_sh}; transition:transform .18s; }}
  .card:hover {{ transform:translateY(-3px); }}
  .lbl {{ font-size:.68rem; font-weight:700; text-transform:uppercase; letter-spacing:.07em; color:{MUTED}; margin-bottom:6px; }}
  .val {{ font-size:2.1rem; font-weight:800; line-height:1; }}
</style></head><body>
<div class="grid">
  <div class="card"><div class="lbl">Tổng khách hàng</div><div class="val" style="color:{TEXT};">{total}</div></div>
  <div class="card"><div class="lbl" style="color:#10b981;">Siêu tiềm năng (VIP)</div><div class="val" style="color:#10b981;">{vip_cnt}</div></div>
  <div class="card"><div class="lbl" style="color:#f59e0b;">Khách tiềm năng</div><div class="val" style="color:#f59e0b;">{nor_cnt}</div></div>
  <div class="card"><div class="lbl" style="color:#ef4444;">Spam / Không tiềm năng</div><div class="val" style="color:#ef4444;">{spm_cnt}</div></div>
  <div class="card"><div class="lbl" style="color:{MUTED};">Đã kiểm duyệt</div>
    <div class="val" style="color:{TEXT};">{rev_cnt}<span style="font-size:1rem;color:{MUTED};font-weight:600;"> / {total}</span></div>
  </div>
</div></body></html>"""
st.iframe(kpi_html, height=120)

st.write("")  # spacer

# ─── Main Layout ──────────────────────────────────────────────────────────────
col_l, col_r = st.columns([7, 3], gap="large")

with col_l:
    st.subheader("📋 Danh sách khách hàng")
    f1, f2, f3 = st.columns([4, 3, 3])
    with f1:
        search_q   = st.text_input("🔍 Tìm kiếm tên / SĐT / nhu cầu...", placeholder="Nhập từ khóa...")
    with f2:
        status_sel = st.selectbox("Lọc trạng thái:", ["Tất cả", "VIP (≥80đ)", "NORMAL (30-79đ)", "SPAM (<30đ)"])
    with f3:
        review_sel = st.selectbox("Lọc kiểm duyệt:", ["Tất cả", "Đã duyệt", "Chưa duyệt"])

    # Filter
    filtered = leads_list
    if search_q:
        sq = search_q.lower()
        filtered = [l for l in filtered if sq in l["ten_khach"].lower() or sq in l["sdt"].lower() or sq in l["nhu_cau_mo_ta"].lower()]
    if "VIP" in status_sel:
        filtered = [l for l in filtered if l["status"] == "VIP"]
    elif "NORMAL" in status_sel:
        filtered = [l for l in filtered if l["status"] == "NORMAL"]
    elif "SPAM" in status_sel:
        filtered = [l for l in filtered if l["status"] == "SPAM"]
    if review_sel == "Đã duyệt":
        filtered = [l for l in filtered if l.get("reviewed")]
    elif review_sel == "Chưa duyệt":
        filtered = [l for l in filtered if not l.get("reviewed")]

    # Pagination
    PER = 15
    n_pages = max(1, (len(filtered) + PER - 1) // PER)
    _, pc, _ = st.columns([4, 4, 4])
    with pc:
        page_no = st.number_input("Trang", min_value=1, max_value=n_pages, value=1, step=1, label_visibility="collapsed")

    page_leads = filtered[(page_no-1)*PER : page_no*PER]

    # Build table HTML (self-contained)
    rows = ""
    for lead in page_leads:
        s  = lead["score"]
        sc = score_color(s)
        nm = str(lead["ten_khach"]).replace("<","&lt;").replace(">","&gt;")
        ds = str(lead["nhu_cau_mo_ta"]).replace("<","&lt;").replace(">","&gt;")
        rows += f"""
        <tr class="row">
          <td style="text-align:center;font-family:monospace;font-weight:700;color:#64748b;width:56px;">{lead['id']}</td>
          <td style="font-weight:600;width:155px;">{nm}</td>
          <td style="text-align:center;font-family:monospace;color:#94a3b8;width:125px;">{lead['sdt']}</td>
          <td style="max-width:340px;"><div style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:340px;color:#94a3b8;">{ds}</div></td>
          <td style="text-align:center;font-family:monospace;font-weight:800;font-size:1.05rem;color:{sc};width:72px;">{s}đ</td>
          <td style="text-align:center;width:105px;">{badge_status(lead['status'])}</td>
          <td style="text-align:center;width:115px;">{badge_review(bool(lead.get('reviewed')))}</td>
        </tr>"""

    if not rows:
        rows = '<tr><td colspan="7" style="text-align:center;padding:56px 0;color:#475569;">Không tìm thấy khách hàng nào phù hợp.</td></tr>'

    tbl_bg   = "#0d1525" if IS_DARK else "#ffffff"
    tbl_bd   = "rgba(255,255,255,0.08)" if IS_DARK else "rgba(0,0,0,0.08)"
    th_bg    = "rgba(255,255,255,0.03)" if IS_DARK else "rgba(0,0,0,0.03)"
    row_hvr  = "rgba(0,242,254,0.05)" if IS_DARK else "rgba(0,100,200,0.05)"
    pg_lbl   = f"Hiển thị {(page_no-1)*PER+1 if filtered else 0}–{min(page_no*PER, len(filtered))} trong {len(filtered)} leads"

    tbl_html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
  body  {{ margin:0; background:transparent; font-family:'DM Sans',sans-serif; font-size:.85rem; color:{TEXT}; }}
  table {{ width:100%; border-collapse:collapse; background:{tbl_bg}; border:1px solid {tbl_bd}; border-radius:12px; overflow:hidden; }}
  thead {{ background:{th_bg}; }}
  th    {{ padding:.75rem 1rem; text-align:left; color:#64748b; font-size:.68rem; font-weight:700;
           text-transform:uppercase; letter-spacing:.07em; border-bottom:1px solid {tbl_bd}; white-space:nowrap; }}
  td    {{ padding:.72rem 1rem; border-bottom:1px solid {tbl_bd}; vertical-align:middle; color:{TEXT}; }}
  .row:last-child td {{ border-bottom:none; }}
  .row:hover {{ background:{row_hvr}; }}
  .foot {{ font-size:.7rem; color:#64748b; text-align:right; margin-top:.5rem; padding-right:2px; }}
</style></head><body>
<table>
  <thead><tr>
    <th style="text-align:center;">ID</th>
    <th>Khách hàng</th>
    <th style="text-align:center;">Số điện thoại</th>
    <th>Nhu cầu chi tiết</th>
    <th style="text-align:center;">Điểm</th>
    <th style="text-align:center;">Trạng thái</th>
    <th style="text-align:center;">Duyệt</th>
  </tr></thead>
  <tbody>{rows}</tbody>
</table>
<div class="foot">{pg_lbl}</div>
</body></html>"""

    n_rows = max(1, len(page_leads))
    st.iframe(tbl_html, height=90 + n_rows * 52)


# ─── Right Panel: Review ──────────────────────────────────────────────────────
with col_r:
    st.subheader("✏️ Duyệt kết quả")

    all_ids = [l["id"] for l in filtered]
    if not all_ids:
        st.info("Không có leads nào trong danh sách.")
    else:
        if "sel_id" not in st.session_state or st.session_state.sel_id not in all_ids:
            st.session_state.sel_id = all_ids[0]
        sel_id = st.selectbox("Chọn ID khách hàng:", all_ids, index=all_ids.index(st.session_state.sel_id))
        st.session_state.sel_id = sel_id

        lo = next(l for l in leads_list if l["id"] == sel_id)
        sc = score_color(lo["score"])
        nm = str(lo["ten_khach"]).replace("<","&lt;").replace(">","&gt;")
        ds = str(lo["nhu_cau_mo_ta"]).replace("<","&lt;").replace(">","&gt;")

        detail_bg = "#0d1525" if IS_DARK else "#f8fafc"
        st.iframe(f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
  body {{ margin:0; background:transparent; font-family:'DM Sans',sans-serif; color:{TEXT}; }}
  .card {{ background:{detail_bg}; border:1px solid {BORDER}; border-radius:12px; padding:16px 18px; margin-bottom:12px; }}
  .lbl  {{ font-size:.66rem; font-weight:700; text-transform:uppercase; letter-spacing:.07em; color:{DIM}; margin-bottom:6px; }}
  .name {{ font-size:1.2rem; font-weight:800; color:{TEXT}; margin-bottom:4px; }}
  .meta {{ font-family:monospace; font-size:.8rem; color:{MUTED}; margin-bottom:12px; }}
  .desc {{ background:{"rgba(0,0,0,0.25)" if IS_DARK else "rgba(0,0,0,0.04)"}; border:1px solid {BORDER};
           border-radius:8px; padding:10px 12px; font-size:.85rem; line-height:1.65; color:{TEXT}; }}
</style></head><body>
<div class="card">
  <div class="lbl">Chi tiết khách hàng</div>
  <div class="name">{nm}</div>
  <div class="meta">SĐT: {lo['sdt']} &nbsp;|&nbsp; ID: #{lo['id']} &nbsp;|&nbsp;
    <span style="color:{sc};font-weight:800;">{lo['score']}đ &nbsp;{badge_status(lo['status'])}</span></div>
  <div class="lbl">Nhu cầu mô tả</div>
  <div class="desc">{ds}</div>
</div>
</body></html>""", height=260)

        # Reason tags
        if lo.get("reason"):
            tags_html = ""
            for r in lo["reason"].split(";"):
                r = r.strip()
                if not r: continue
                if r.startswith("+"): c,bg,bd = "#a7f3d0","rgba(16,185,129,.12)","rgba(16,185,129,.3)"
                elif r.startswith("-"): c,bg,bd = "#fecaca","rgba(239,68,68,.12)","rgba(239,68,68,.3)"
                else: c,bg,bd = "#fde68a","rgba(245,158,11,.12)","rgba(245,158,11,.3)"
                rs = r.replace("<","&lt;").replace(">","&gt;")
                tags_html += f'<span style="display:inline-block;font-size:.7rem;padding:3px 8px;border-radius:5px;margin:3px 4px 3px 0;color:{c};background:{bg};border:1px solid {bd};">{rs}</span>'
            st.iframe(f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600&display=swap" rel="stylesheet">
<style>body{{margin:0;background:transparent;font-family:'DM Sans',sans-serif;}}
.lbl{{font-size:.66rem;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:{DIM};margin-bottom:8px;}}
</style></head><body>
<div class="lbl">Lý do chấm điểm tự động</div>
<div>{tags_html}</div>
</body></html>""", height=max(60, 30 + len(lo["reason"].split(";")) * 28))

        # Review Form
        with st.form("review_form"):
            f_score = st.slider("Điều chỉnh điểm (0–100):", 0, 100, int(lo["score"]))
            def_idx = 0 if lo["status"] == "VIP" else (2 if lo["status"] == "SPAM" else 1)
            f_status   = st.selectbox("Phân loại:", ["VIP", "NORMAL", "SPAM"], index=def_idx)
            f_notes    = st.text_area("Ghi chú kiểm duyệt:", value=lo.get("notes",""), height=80)
            f_reviewed = st.checkbox("✅ Đánh dấu đã duyệt", value=bool(lo.get("reviewed")))
            if st.form_submit_button("💾 Lưu kết quả", use_container_width=True):
                lo["score"]    = f_score
                lo["status"]   = f_status
                lo["notes"]    = f_notes
                lo["reviewed"] = f_reviewed
                if not lo.get("reason","").startswith("[Manual]"):
                    lo["reason"] = f"[Manual] {lo.get('reason','')}"
                save_leads(leads_list)
                st.success(f"✅ Đã lưu kết quả cho KH #{sel_id}!")
                st.rerun()


# ─── Sidebar ──────────────────────────────────────────────────────────────────
st.sidebar.header("⚙️ Cài đặt & Xuất dữ liệu")
excel_bytes = generate_excel(leads_list)
st.sidebar.download_button(
    "📥 Tải Báo Cáo Excel",
    data=excel_bytes,
    file_name="Leads_Tiem_Nang_Da_Cham_Diem.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
st.sidebar.markdown("---")
st.sidebar.subheader("🔄 Chấm điểm lại")
recalc_key = st.sidebar.text_input("Gemini API Key (Tùy chọn):", type="password",
    help="Để trống → dùng Rule-Based Keyword Engine")
overwrite = st.sidebar.checkbox("Ghi đè leads đã duyệt thủ công")
if st.sidebar.button("🚀 Chấm điểm lại toàn bộ", use_container_width=True):
    with st.spinner("Đang xử lý 500 khách hàng..."):
        try:
            ok = run_scoring(api_key=recalc_key or None, overwrite=overwrite)
            if ok:
                st.sidebar.success("✅ Hoàn tất!")
                st.rerun()
            else:
                st.sidebar.error("❌ Có lỗi xảy ra.")
        except Exception as e:
            st.sidebar.error(f"Lỗi: {e}")

st.sidebar.markdown("---")
st.sidebar.markdown(f"""
🟢 **VIP**: ≥ 80 điểm  
🟡 **NORMAL**: 30–79 điểm  
🔴 **SPAM**: < 30 điểm  

📊 **Nguồn:** Google Sheet (500 leads)  
🤖 **Engine:** Rule-Based AI Keyword Scorer
""")
