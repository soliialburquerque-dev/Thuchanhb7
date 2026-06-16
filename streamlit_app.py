import os
import json
import csv
import io
import pandas as pd
import streamlit as st
import xlsxwriter
from scorer import calculate_score as local_calculate_score, run_scoring, JSON_PATH, CSV_PATH, SHEET_CSV_URL

# Page Config
st.set_page_config(
    page_title="AI Lead Scoring Console",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Theme Toggle State
if "theme" not in st.session_state:
    st.session_state.theme = "dark"

def toggle_theme():
    st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"

IS_DARK = st.session_state.theme == "dark"

# Dynamic CSS styling based on dark/light mode
css_theme = f"""
<style>
:root {{
    --bg: {"#090d16" if IS_DARK else "#f8fafc"};
    --bg-subtle: {"#111827" if IS_DARK else "#f1f5f9"};
    --card: {"rgba(19, 29, 48, 0.6)" if IS_DARK else "rgba(255, 255, 255, 0.8)"};
    --border: {"rgba(255, 255, 255, 0.08)" if IS_DARK else "rgba(0, 0, 0, 0.08)"};
    --text: {"#f8fafc" if IS_DARK else "#0f172a"};
    --text-muted: {"#94a3b8" if IS_DARK else "#475569"};
    --text-dim: {"#64748b" if IS_DARK else "#94a3b8"};
    --accent: #00f2fe;
    
    --vip-color: #10b981;
    --vip-bg: rgba(16, 185, 129, 0.12);
    --vip-border: rgba(16, 185, 129, 0.3);
    
    --normal-color: #f59e0b;
    --normal-bg: rgba(245, 158, 11, 0.12);
    --normal-border: rgba(245, 158, 11, 0.3);
    
    --spam-color: #ef4444;
    --spam-bg: rgba(239, 68, 68, 0.12);
    --spam-border: rgba(239, 68, 68, 0.3);
}}

/* Global Styling */
html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"], .main, .block-container, section[data-testid="stMain"] {{
    background-color: var(--bg) !important;
    color: var(--text) !important;
    font-family: 'DM Sans', 'Segoe UI', sans-serif !important;
}}

.block-container {{
    padding: 2rem 3rem 3rem !important;
    max-width: 1400px !important;
}}

/* Hide Streamlit components */
header[data-testid="stHeader"], #MainMenu, footer, [data-testid="stToolbar"],
[data-testid="stDecoration"], [data-testid="stStatusWidget"], .stDeployButton {{
    display: none !important;
}}

/* Custom Cards */
.brand-header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 1px solid var(--border);
    padding-bottom: 1rem;
    margin-bottom: 2rem;
}}
.brand-name {{
    font-size: 1.5rem;
    font-weight: 700;
    color: var(--text);
}}
.brand-sub {{
    font-size: 0.8rem;
    color: var(--text-muted);
}}

/* KPI Metric Cards */
.kpi-row {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 1.25rem;
    margin-bottom: 2rem;
}}
.kpi-card {{
    background: var(--card);
    backdrop-filter: blur(12px);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.25rem;
    box-shadow: 0 4px 20px rgba(0,0,0,0.15);
}}
.kpi-label {{
    font-size: 0.78rem;
    color: var(--text-muted);
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}}
.kpi-value {{
    font-size: 1.8rem;
    font-weight: 700;
    color: var(--text);
    margin-top: 0.25rem;
}}

/* Custom HTML table */
.data-table {{
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    font-size: 0.85rem;
    border: 1px solid var(--border);
    border-radius: 10px;
    overflow: hidden;
    background: var(--card);
}}
.data-table th {{
    text-align: left;
    padding: 0.9rem 1rem;
    color: var(--text-muted);
    font-weight: 600;
    font-size: 0.75rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    border-bottom: 1px solid var(--border);
    background: rgba(255, 255, 255, 0.02);
}}
.data-table td {{
    padding: 0.85rem 1rem;
    color: var(--text);
    border-bottom: 1px solid var(--border);
    vertical-align: middle;
}}
.data-table tr:last-child td {{
    border-bottom: none;
}}

/* Badges */
.badge {{
    display: inline-block;
    padding: 3px 9px;
    border-radius: 9999px;
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.03em;
    text-transform: uppercase;
}}
.badge-vip {{
    color: var(--vip-color);
    background: var(--vip-bg);
    border: 1px solid var(--vip-border);
}}
.badge-normal {{
    color: var(--normal-color);
    background: var(--normal-bg);
    border: 1px solid var(--normal-border);
}}
.badge-spam {{
    color: var(--spam-color);
    background: var(--spam-bg);
    border: 1px solid var(--spam-border);
}}
.badge-reviewed {{
    color: #a855f7;
    background: rgba(168, 85, 247, 0.12);
    border: 1px solid rgba(168, 85, 247, 0.3);
}}
.badge-unreviewed {{
    color: var(--text-dim);
    background: rgba(255, 255, 255, 0.04);
    border: 1px solid var(--border);
}}

/* Description details */
.detail-box {{
    background: rgba(0, 0, 0, 0.15);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 1rem;
    font-size: 0.88rem;
    line-height: 1.5;
    margin-bottom: 1rem;
}}
.reason-tag {{
    display: inline-block;
    font-size: 0.72rem;
    padding: 3px 8px;
    border-radius: 4px;
    margin-right: 6px;
    margin-bottom: 6px;
    background: rgba(255,255,255,0.05);
    border: 1px solid var(--border);
}}
.reason-positive {{
    color: #a7f3d0;
    background: rgba(16, 185, 129, 0.08);
    border-color: rgba(16, 185, 129, 0.2);
}}
.reason-negative {{
    color: #fecaca;
    background: rgba(239, 68, 68, 0.08);
    border-color: rgba(239, 68, 68, 0.2);
}}
.reason-neutral {{
    color: #fde68a;
    background: rgba(245, 158, 11, 0.08);
    border-color: rgba(245, 158, 11, 0.2);
}}
</style>
"""
st.markdown(css_theme, unsafe_allow_html=True)

# Helper function to load data
def load_leads_data():
    if not os.path.exists(JSON_PATH):
        run_scoring()
    
    with open(JSON_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

# Helper function to save data
def save_leads_data(data):
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# Helper to generate Excel binary
def generate_excel_bytes(leads_list):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("Leads Scored")
    
    header_format = workbook.add_format({
        'bold': True, 'font_name': 'Segoe UI', 'font_size': 11,
        'font_color': 'white', 'bg_color': '#1F4E78', 'border': 1,
        'align': 'center', 'valign': 'vcenter'
    })
    
    cell_format = workbook.add_format({'font_name': 'Segoe UI', 'font_size': 10, 'border': 1, 'valign': 'vcenter'})
    center_format = workbook.add_format({'font_name': 'Segoe UI', 'font_size': 10, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    
    vip_format = workbook.add_format({'font_name': 'Segoe UI', 'font_size': 10, 'bold': True, 'font_color': '#375623', 'bg_color': '#E2EFDA', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    normal_format = workbook.add_format({'font_name': 'Segoe UI', 'font_size': 10, 'bold': True, 'font_color': '#7F6000', 'bg_color': '#FFF2CC', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    spam_format = workbook.add_format({'font_name': 'Segoe UI', 'font_size': 10, 'bold': True, 'font_color': '#C00000', 'bg_color': '#FCE4D6', 'border': 1, 'align': 'center', 'valign': 'vcenter'})

    worksheet.set_column('A:A', 8)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 65)
    worksheet.set_column('E:E', 12)
    worksheet.set_column('F:F', 15)
    worksheet.set_column('G:G', 12)
    worksheet.set_column('H:H', 25)
    worksheet.set_column('I:I', 50)
    
    worksheet.set_row(0, 30)
    
    headers = ["ID", "Khách hàng", "Số điện thoại", "Nhu cầu mô tả", "Điểm số", "Trạng thái", "Đã duyệt", "Ghi chú thủ công", "Lý do hệ thống"]
    for col_idx, header in enumerate(headers):
        worksheet.write(0, col_idx, header, header_format)
        
    for row_idx, lead in enumerate(leads_list, start=1):
        worksheet.set_row(row_idx, 22)
        worksheet.write(row_idx, 0, lead["id"], center_format)
        worksheet.write(row_idx, 1, lead["ten_khach"], cell_format)
        worksheet.write(row_idx, 2, lead["sdt"], center_format)
        worksheet.write(row_idx, 3, lead["nhu_cau_mo_ta"], cell_format)
        worksheet.write(row_idx, 4, lead["score"], center_format)
        
        status = lead["status"]
        if status == "VIP":
            worksheet.write(row_idx, 5, status, vip_format)
        elif status == "SPAM":
            worksheet.write(row_idx, 5, status, spam_format)
        else:
            worksheet.write(row_idx, 5, status, normal_format)
            
        worksheet.write(row_idx, 6, "Đã duyệt" if lead.get("reviewed", False) else "Chưa duyệt", center_format)
        worksheet.write(row_idx, 7, lead.get("notes", ""), cell_format)
        worksheet.write(row_idx, 8, lead.get("reason", ""), cell_format)
        
    workbook.close()
    return output.getvalue()

# MAIN INTERFACE
# Brand Header Row
hl, hr = st.columns([10, 2])
with hl:
    st.markdown(f"""
    <div class="brand-header">
        <div>
            <span class="brand-name">🏢 AI Lead Scoring Console</span>
            <div class="brand-sub">Real Estate Leads Automated Processing Dashboard</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
with hr:
    theme_lbl = "☀️ Light Mode" if IS_DARK else "🌙 Dark Mode"
    st.button(theme_lbl, on_click=toggle_theme, use_container_width=True)

# Load Leads
leads_list = load_leads_data()

# Compute Stats
total_leads = len(leads_list)
vip_count = sum(1 for l in leads_list if l["status"] == "VIP")
normal_count = sum(1 for l in leads_list if l["status"] == "NORMAL")
spam_count = sum(1 for l in leads_list if l["status"] == "SPAM")
reviewed_count = sum(1 for l in leads_list if l.get("reviewed", False))

# Stats grid markup
st.markdown(f"""
<div class="kpi-row">
    <div class="kpi-card">
        <div class="kpi-label">Tổng khách hàng</div>
        <div class="kpi-value">{total_leads}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label" style="color:var(--vip-color);">Siêu tiềm năng (VIP)</div>
        <div class="kpi-value" style="color:var(--vip-color);">{vip_count}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label" style="color:var(--normal-color);">Khách tiềm năng</div>
        <div class="kpi-value" style="color:var(--normal-color);">{normal_count}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label" style="color:var(--spam-color);">Spam / Không tiềm năng</div>
        <div class="kpi-value" style="color:var(--spam-color);">{spam_count}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">Đã kiểm duyệt</div>
        <div class="kpi-value">{reviewed_count} / {total_leads}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# Main Grid Layout (Left: Table/Filters, Right: Sidebar review drawer)
col_left, col_right = st.columns([8, 4])

with col_left:
    st.subheader("Danh sách khách hàng")
    
    # Filter Toolbar
    tf1, tf2, tf3 = st.columns([5, 3, 4])
    with tf1:
        search_query = st.text_input("🔍 Tìm kiếm theo tên, số điện thoại, nhu cầu...", placeholder="Nhập từ khóa...")
    with tf2:
        status_sel = st.selectbox("Lọc Trạng thái:", ["Tất cả", "VIP (≥ 80đ)", "NORMAL (30-79đ)", "SPAM (< 30đ)"])
    with tf3:
        review_sel = st.selectbox("Lọc Kiểm duyệt:", ["Tất cả", "Đã duyệt (Human-in-the-loop)", "Chưa duyệt"])
        
    # Apply Filtering
    filtered_leads = leads_list
    if search_query:
        search_query = search_query.lower()
        filtered_leads = [
            l for l in filtered_leads 
            if search_query in l["ten_khach"].lower() 
            or search_query in l["sdt"].lower() 
            or search_query in l["nhu_cau_mo_ta"].lower()
        ]
        
    if status_sel != "Tất cả":
        target = "VIP" if "VIP" in status_sel else ("NORMAL" if "NORMAL" in status_sel else "SPAM")
        filtered_leads = [l for l in filtered_leads if l["status"] == target]
        
    if review_sel != "Tất cả":
        target_rev = True if "Đã duyệt" in review_sel else False
        filtered_leads = [l for l in filtered_leads if l.get("reviewed", False) == target_rev]

    # Pagination Setup
    items_per_page = 15
    total_filtered = len(filtered_leads)
    total_pages = max(1, (total_filtered + items_per_page - 1) // items_per_page)
    
    p1, p2, p3 = st.columns([4, 4, 4])
    with p2:
        page_num = st.number_input("Trang", min_value=1, max_value=total_pages, value=1, step=1)
        
    start_idx = (page_num - 1) * items_per_page
    end_idx = min(start_idx + items_per_page, total_filtered)
    page_leads = filtered_leads[start_idx:end_idx]
    
    # HTML Table Compilation
    table_rows = ""
    for lead in page_leads:
        # Badges
        status_cls = "badge-vip" if lead["status"] == "VIP" else ("badge-spam" if lead["status"] == "SPAM" else "badge-normal")
        status_badge = f'<span class="badge {status_cls}">{lead["status"]}</span>'
        
        review_cls = "badge-reviewed" if lead.get("reviewed", False) else "badge-unreviewed"
        review_lbl = "Đã duyệt" if lead.get("reviewed", False) else "Chưa duyệt"
        review_badge = f'<span class="badge {review_cls}">{review_lbl}</span>'
        
        table_rows += f"""
        <tr>
            <td style="font-family: monospace; text-align: center; font-weight:600;">{lead["id"]}</td>
            <td style="font-weight:600;">{lead["ten_khach"]}</td>
            <td style="font-family: monospace; text-align: center;">{lead["sdt"]}</td>
            <td><div style="max-width: 380px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: var(--text-muted);">{lead["nhu_cau_mo_ta"]}</div></td>
            <td style="font-family: monospace; text-align: center; font-weight: 700; color: {"var(--vip-color)" if lead["score"] >= 80 else ("var(--spam-color)" if lead["score"] < 30 else "var(--normal-color)")}">{lead["score"]}đ</td>
            <td style="text-align: center;">{status_badge}</td>
            <td style="text-align: center;">{review_badge}</td>
        </tr>
        """
        
    if not table_rows:
        table_rows = '<tr><td colspan="7" style="text-align:center; padding: 40px; color: var(--text-dim);">Không tìm thấy khách hàng nào khớp bộ lọc.</td></tr>'
        
    st.markdown(f"""
    <table class="data-table">
        <thead>
            <tr>
                <th width="60" style="text-align: center;">ID</th>
                <th width="150">Khách hàng</th>
                <th width="120" style="text-align: center;">Số điện thoại</th>
                <th>Nhu cầu chi tiết</th>
                <th width="80" style="text-align: center;">Điểm</th>
                <th width="100" style="text-align: center;">Trạng thái</th>
                <th width="110" style="text-align: center;">Duyệt</th>
            </tr>
        </thead>
        <tbody>
            {table_rows}
        </tbody>
    </table>
    <div style="font-size:0.75rem; color:var(--text-dim); margin-top: 0.5rem; text-align:right;">
        Hiển thị {start_idx + 1 if total_filtered > 0 else 0} - {end_idx} trong số {total_filtered} leads được lọc.
    </div>
    """, unsafe_allow_html=True)


with col_right:
    st.subheader("Duyệt kết quả (Human-in-the-loop)")
    
    # Choose Lead ID to review
    all_lead_ids = [l["id"] for l in filtered_leads]
    if not all_lead_ids:
        st.info("Không có leads nào để duyệt trong danh sách được lọc.")
    else:
        # Track selected lead in session state to enable smooth updates
        if "selected_lead_id" not in st.session_state or st.session_state.selected_lead_id not in all_lead_ids:
            st.session_state.selected_lead_id = all_lead_ids[0]
            
        selected_id = st.selectbox("Chọn ID khách hàng cần kiểm duyệt:", all_lead_ids, index=all_lead_ids.index(st.session_state.selected_lead_id))
        st.session_state.selected_lead_id = selected_id
        
        # Load lead object
        lead_obj = next(l for l in leads_list if l["id"] == selected_id)
        
        st.markdown(f"""
        <div style="background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 1.25rem; box-shadow: 0 4px 15px rgba(0,0,0,0.15);">
            <div style="font-size:0.75rem; color: var(--text-dim); font-weight:600; text-transform:uppercase; letter-spacing:0.04em;">Chi tiết khách hàng</div>
            <div style="font-size: 1.25rem; font-weight:700; margin-top:0.25rem;">{lead_obj["ten_khach"]}</div>
            <div style="font-family: monospace; font-size:0.85rem; color:var(--text-muted); margin-bottom:1rem;">SĐT: {lead_obj["sdt"]} | ID: #{lead_obj["id"]}</div>
            
            <div style="font-size:0.75rem; color: var(--text-dim); font-weight:600; text-transform:uppercase; letter-spacing:0.04em; margin-bottom: 0.25rem;">Nhu cầu mô tả:</div>
            <div class="detail-box">{lead_obj["nhu_cau_mo_ta"]}</div>
        </div>
        """, unsafe_allow_html=True)
        
        # Display reasons tags
        reasons_tags_html = ""
        if lead_obj.get("reason"):
            for r in lead_obj["reason"].split(";"):
                r = r.strip()
                if not r: continue
                cls = "reason-positive" if r.startswith("+") else ("reason-negative" if r.startswith("-") else "reason-neutral")
                reasons_tags_html += f'<span class="reason-tag {cls}">{r}</span>'
                
        st.markdown(f"""
        <div style="margin-top: 1rem; margin-bottom: 1.25rem;">
            <div style="font-size:0.75rem; color: var(--text-dim); font-weight:600; text-transform:uppercase; letter-spacing:0.04em; margin-bottom: 0.5rem;">Lý do chấm điểm tự động:</div>
            <div class="reason-tags">{reasons_tags_html or '<span class="reason-tag">Không có phân tích chi tiết</span>'}</div>
        </div>
        """, unsafe_allow_html=True)
        
        # Review Form Fields
        with st.form("review_form", clear_on_submit=False):
            f_score = st.slider("Điều chỉnh Điểm số (0-100):", min_value=0, max_value=100, value=int(lead_obj["score"]))
            
            # Status mapping
            default_status_idx = 0 if lead_obj["status"] == "VIP" else (2 if lead_obj["status"] == "SPAM" else 1)
            f_status = st.selectbox("Phân loại trạng thái:", ["VIP", "NORMAL", "SPAM"], index=default_status_idx)
            
            f_notes = st.text_area("Ghi chú/Nhận xét kiểm duyệt:", value=lead_obj.get("notes", ""))
            f_reviewed = st.checkbox("Đánh dấu đã duyệt kết quả này", value=lead_obj.get("reviewed", False))
            
            submit_btn = st.form_submit_button("Lưu kết quả duyệt", use_container_width=True)
            
            if submit_btn:
                # Update lead in memory
                lead_obj["score"] = f_score
                lead_obj["status"] = f_status
                lead_obj["notes"] = f_notes
                lead_obj["reviewed"] = f_reviewed
                if not lead_obj.get("reason", "").startswith("[Manual]"):
                    lead_obj["reason"] = f"[Manual] {lead_obj.get('reason', '')}"
                
                # Save changes
                save_leads_data(leads_list)
                st.success(f"Đã lưu kết quả duyệt cho khách hàng #{selected_id}!")
                st.rerun()

# Sidebar Settings / Action toolbar
st.sidebar.header("Cài đặt & Xuất dữ liệu")

# Export to Excel
excel_data = generate_excel_bytes(leads_list)
st.sidebar.download_button(
    label="📥 Tải Báo Cáo Excel",
    data=excel_data,
    file_name="Danh_sach_khach_hang_tiem_nang_da_duyet.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.sidebar.markdown("---")

# API Configuration and recalculation
st.sidebar.subheader("Chấm điểm lại tự động")
recalc_key = st.sidebar.text_input("Gemini API Key (Tùy chọn):", type="password", help="Để trống để sử dụng phân tích Regex từ khóa cục bộ")
recalc_overwrite = st.sidebar.checkbox("Ghi đè leads đã duyệt thủ công")

if st.sidebar.button("🚀 Bắt đầu chấm điểm lại", use_container_width=True):
    with st.spinner("Đang chạy chấm điểm lại dữ liệu khách hàng..."):
        try:
            success = run_scoring(api_key=recalc_key if recalc_key else None, overwrite=recalc_overwrite)
            if success:
                st.sidebar.success("Tiến trình chấm điểm lại đã hoàn tất thành công!")
                st.rerun()
            else:
                st.sidebar.error("Có lỗi xảy ra khi chạy tiến trình.")
        except Exception as e:
            st.sidebar.error(f"Lỗi: {e}")
