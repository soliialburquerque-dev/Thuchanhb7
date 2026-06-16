import os
import csv
import re
import requests
import xlsxwriter

# Configuration
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/16tCAf_qqtgYZxoumYQKMEOdBhKE0wg5A/export?format=csv&gid=1542775777"
OUTPUT_EXCEL_PATH = "leads_scored_final.xlsx"

def download_lead_data(url):
    """
    Downloads customer lead data in CSV format from the Google Sheets URL.
    """
    print(f"[*] Downloading lead data from Google Sheets: {url}")
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
    response = requests.get(url, headers=headers, timeout=20)
    response.raise_for_status()
    print(f"[+] Download successful! File size: {len(response.content)} bytes")
    
    # Save CSV locally as backup/audit log
    csv_filename = "data_source.csv"
    with open(csv_filename, "wb") as f:
        f.write(response.content)
    
    return csv_filename


def calculate_score(description):
    """
    Core local scoring ruleset engine based on tieu_chi_cham_diem.txt.
    Base score: 50.
    Adds 50 for VIP indicators, subtracts 50 for Spam/invalid indicators.
    Caps final score between 0 and 100.
    """
    if not description:
        return 50, "Không có nội dung mô tả", "NORMAL"
        
    text_lower = description.lower()
    score = 50
    reasons = []
    
    # VIP INDICATORS (+50)
    vip_matched = False
    
    # 1. Budget check (>= 20 billion VND or finance keywords)
    budget_keywords = ["tài chính mạnh", "không thành vấn đề", "tài chính cực mạnh", "ngân sách lớn"]
    has_large_budget = any(kw in text_lower for kw in budget_keywords)
    # Check for number sequences followed by 'tỷ' or 'tỉ'
    budget_numbers = re.findall(r'(\d+)\s*(?:tỷ|tỉ)', text_lower)
    for num in budget_numbers:
        if int(num) >= 20:
            has_large_budget = True
            break
            
    if has_large_budget:
        reasons.append("+50: Ngân sách lớn (>= 20 tỷ hoặc tài chính mạnh)")
        vip_matched = True

    # 2. Premium property types
    luxury_properties = ["biệt thự đơn lập", "penthouse", "shophouse mặt đường lớn", "quỹ đất công nghiệp", "sàn văn phòng diện tích lớn", "shophouse", "sản văn phòng", "đất công nghiệp"]
    has_luxury = any(kw in text_lower for kw in luxury_properties)
    if has_luxury:
        reasons.append("+50: Loại hình cao cấp (Biệt thự, Penthouse, Shophouse, Đất công nghiệp/Văn phòng)")
        vip_matched = True

    # 3. Prime locations
    prime_locations = ["quận 1", "ven sông", "vinhomes ocean park", "phú mỹ hưng", "q1", "q.1"]
    has_prime_location = any(kw in text_lower for kw in prime_locations)
    if has_prime_location:
        reasons.append("+50: Vị trí đắc địa (Quận 1, Ven sông, Vinhomes Ocean Park, Phú Mỹ Hưng)")
        vip_matched = True

    # 4. VIP profiles
    vip_profiles = ["chủ doanh nghiệp", "nhà đầu tư chuyên nghiệp", "mua sỉ", "mua số lượng lớn", "gom sỉ"]
    has_vip_profile = any(kw in text_lower for kw in vip_profiles)
    if has_vip_profile:
        reasons.append("+50: Đối tượng khách hàng VIP (Chủ DN, Đầu tư chuyên nghiệp, Mua sỉ/Số lượng lớn)")
        vip_matched = True

    # 5. Urgency & Transparency
    urgency_transparency = ["pháp lý chuẩn 100%", "sổ hồng riêng", "gặp trực tiếp chủ đầu tư", "gặp trực tiếp giám đốc", "đàm phán trực tiếp"]
    has_urgency = any(kw in text_lower for kw in urgency_transparency)
    if has_urgency:
        reasons.append("+50: Tính cấp thiết & Minh bạch cao (Sổ hồng riêng, Pháp lý 100%, Đàm phán trực tiếp)")
        vip_matched = True

    if vip_matched:
        score += 50

    # SPAM / UNQUALIFIED INDICATORS (-50)
    spam_matched = False
    
    # 1. Unrealistic requirements
    unrealistic_requests = ["giá thấp vô lý", "q1 giá 1 tỷ", "quận 1 giá 1", "quận 1 giá 2 tỷ", "thuê nguyên căn giá 2 triệu", "yêu cầu phi thực tế", "ngân sách rất thấp"]
    has_unrealistic = any(kw in text_lower for kw in unrealistic_requests) or ("q1" in text_lower and "1 tỷ" in text_lower)
    if has_unrealistic:
        reasons.append("-50: Yêu cầu phi thực tế (Giá quá rẻ so với thị trường)")
        spam_matched = True

    # 2. No intent
    no_intent = ["nhầm số", "không có nhu cầu", "dữ liệu cũ", "nhầm ngành"]
    has_no_intent = any(kw in text_lower for kw in no_intent)
    if has_no_intent:
        reasons.append("-50: Không có nhu cầu thực tế (Nhầm số, Dữ liệu cũ, Nhầm ngành)")
        spam_matched = True

    # 3. Uncooperative
    uncooperative = ["hỏi giá cho vui", "chưa có ý định mua", "thái độ không hợp tác", "không hợp tác"]
    has_uncooperative = any(kw in text_lower for kw in uncooperative)
    if has_uncooperative:
        reasons.append("-50: Khách hàng không thiện chí (Hỏi giá vui, Chưa muốn mua, Thái độ kém)")
        spam_matched = True

    # 4. Spam/Advertising (excluding requests for loans)
    advertising_keywords = ["quảng cáo ngược", "bảo hiểm", "vay vốn", "mời chào", "chào mời", "quảng cáo ngược lại dịch vụ"]
    has_ad = any(kw in text_lower for kw in advertising_keywords) and not ("cần hỗ trợ vay" in text_lower or "cần vay" in text_lower)
    if has_ad or "spam" in text_lower:
        reasons.append("-50: Tin nhắn Spam / Quảng cáo dịch vụ")
        spam_matched = True

    # 5. Contact failure
    contact_failure = ["thuê bao", "gọi nhiều lần không bắt máy", "không phản hồi zalo", "gọi không liên lạc được"]
    has_contact_error = any(kw in text_lower for kw in contact_failure)
    if has_contact_error:
        reasons.append("-50: Thông tin liên lạc lỗi (Thuê bao, Không nhấc máy, Không trả lời Zalo)")
        spam_matched = True

    if spam_matched:
        score -= 50

    # NORMAL / MID-RANGE CASES (If neither VIP nor Spam criteria matched)
    if not vip_matched and not spam_matched:
        normal_keywords = ["chung cư", "nhà phố", "căn hộ 2pn", "căn hộ 3pn", "tầm trung", "3-10 tỷ", "vay ngân hàng", "tư vấn thêm", "nhà phố liền kề"]
        if any(kw in text_lower for kw in normal_keywords):
            reasons.append("+10: Khách hàng nhu cầu thực phân khúc trung cấp (Căn hộ, Nhà phố 3-10 tỷ, Cần vay ngân hàng)")
            score = 60
        else:
            reasons.append("0: Khách hàng bình thường / Cần tư vấn thêm")
            score = 50

    # Boundary validation
    score = max(0, min(100, score))

    # Determine status
    if score >= 80:
        status = "VIP"
    elif score < 30:
        status = "SPAM"
    else:
        status = "NORMAL"

    return score, "; ".join(reasons), status

def export_to_excel(leads, file_path):
    """
    Generates a beautifully styled corporate Excel spreadsheet from the processed leads.
    """
    print(f"[*] Exporting report to Excel file: {file_path}")
    
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet("Leads Scored Summary")
    
    # Styles Setup
    header_format = workbook.add_format({
        'bold': True,
        'font_name': 'Segoe UI',
        'font_size': 11,
        'font_color': 'white',
        'bg_color': '#1F4E78',  # Classic Dark Blue
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    cell_format = workbook.add_format({
        'font_name': 'Segoe UI',
        'font_size': 10,
        'border': 1,
        'valign': 'vcenter'
    })
    
    center_format = workbook.add_format({
        'font_name': 'Segoe UI',
        'font_size': 10,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    # Row color configurations
    vip_format = workbook.add_format({
        'font_name': 'Segoe UI',
        'font_size': 10,
        'bold': True,
        'font_color': '#375623',
        'bg_color': '#E2EFDA',  # Pastel Green
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    normal_format = workbook.add_format({
        'font_name': 'Segoe UI',
        'font_size': 10,
        'bold': True,
        'font_color': '#7F6000',
        'bg_color': '#FFF2CC',  # Pastel Yellow
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    spam_format = workbook.add_format({
        'font_name': 'Segoe UI',
        'font_size': 10,
        'bold': True,
        'font_color': '#C00000',
        'bg_color': '#FCE4D6',  # Pastel Red
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    # Set Column Widths
    worksheet.set_column('A:A', 8)    # ID
    worksheet.set_column('B:B', 20)   # Tên khách hàng
    worksheet.set_column('C:C', 15)   # Số điện thoại
    worksheet.set_column('D:D', 65)   # Nhu cầu chi tiết
    worksheet.set_column('E:E', 12)   # Điểm số
    worksheet.set_column('F:F', 15)   # Trạng thái
    worksheet.set_column('G:G', 50)   # Lý do chấm điểm

    # Set Row Heights
    worksheet.set_row(0, 30)

    # Write Headers
    headers = ["ID", "Khách hàng", "Số điện thoại", "Nhu cầu chi tiết", "Điểm số", "Trạng thái", "Lý do chấm điểm"]
    for col_idx, header in enumerate(headers):
        worksheet.write(0, col_idx, header, header_format)

    # Write Data
    for row_idx, lead in enumerate(leads, start=1):
        worksheet.set_row(row_idx, 22)
        worksheet.write(row_idx, 0, lead["id"], center_format)
        worksheet.write(row_idx, 1, lead["name"], cell_format)
        worksheet.write(row_idx, 2, lead["phone"], center_format)
        worksheet.write(row_idx, 3, lead["description"], cell_format)
        worksheet.write(row_idx, 4, lead["score"], center_format)
        
        # Color conditional formatting based on status
        status = lead["status"]
        if status == "VIP":
            worksheet.write(row_idx, 5, status, vip_format)
        elif status == "SPAM":
            worksheet.write(row_idx, 5, status, spam_format)
        else:
            worksheet.write(row_idx, 5, status, normal_format)
            
        worksheet.write(row_idx, 6, lead["reason"], cell_format)
        
    workbook.close()
    print(f"[+] Excel report exported successfully to: {file_path}")

def main():
    print("="*60)
    print(" AUTOMATED REAL ESTATE LEAD SCORING PIPELINE")
    print("="*60)
    
    # 1. Download sheet
    csv_file = download_lead_data(SHEET_CSV_URL)
    
    # 2. Process rows
    processed_leads = []
    stats = {"total": 0, "vip": 0, "normal": 0, "spam": 0}
    
    with open(csv_file, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            desc = row["nhu_cau_mo_ta"]
            score, reason, status = calculate_score(desc)
            
            processed_leads.append({
                "id": int(row["id"]),
                "name": row["ten_khach"],
                "phone": row["sdt"],
                "description": desc,
                "score": score,
                "reason": reason,
                "status": status
            })
            
            # Count statistics
            stats["total"] += 1
            if status == "VIP":
                stats["vip"] += 1
            elif status == "SPAM":
                stats["spam"] += 1
            else:
                stats["normal"] += 1

    # 3. Export to Excel
    export_to_excel(processed_leads, OUTPUT_EXCEL_PATH)
    
    # 4. Print console statistics summary
    print("\n" + "="*40)
    print(" SCORING PIPELINE RUN SUMMARY")
    print("="*40)
    print(f" - Total Leads processed:   {stats['total']}")
    print(f" - Super Potential (VIP):   {stats['vip']} ({stats['vip']/stats['total']*100:.1f}%)")
    print(f" - Mid Potential (NORMAL):  {stats['normal']} ({stats['normal']/stats['total']*100:.1f}%)")
    print(f" - Low Potential (SPAM):    {stats['spam']} ({stats['spam']/stats['total']*100:.1f}%)")
    print("="*40)
    
    # Clean up local backup file
    if os.path.exists(csv_file):
        os.remove(csv_file)
        
    print("\n[+] Scoring execution completed successfully!")

if __name__ == "__main__":
    main()

