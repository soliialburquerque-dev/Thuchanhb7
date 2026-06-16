import os
import csv
import json
import re
import urllib.request
import urllib.parse

# Define file paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(BASE_DIR, "data.csv")
JSON_PATH = os.path.join(BASE_DIR, "leads_processed.json")
CRITERIA_PATH = os.path.join(BASE_DIR, "tieu_chi_cham_diem.txt")
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/16tCAf_qqtgYZxoumYQKMEOdBhKE0wg5A/export?format=csv&gid=1542775777"

def read_criteria():
    if os.path.exists(CRITERIA_PATH):
        with open(CRITERIA_PATH, "r", encoding="utf-8") as f:
            return f.read()
    return "Rules not found. Please refer to tieu_chi_cham_diem.txt."

def rule_based_score(text):
    """
    Evaluates score based on local keyword matching from tieu_chi_cham_diem.txt.
    Base score: 50.
    Adds 50 for VIP indicators, subtracts 50 for spam/unqualified indicators.
    """
    text_lower = text.lower()
    score = 50
    reasons = []
    
    # VIP Criteria Indicators (+50)
    vip_matched = False
    
    # Budget check (>= 20B or "tài chính mạnh", "không thành vấn đề", "tài chính cực mạnh", "tài chính mạnh", "ngân sách trên 30 tỷ")
    budget_keywords = ["tài chính mạnh", "không thành vấn đề", "tài chính cực mạnh", "ngân sách cực mạnh"]
    has_large_budget = any(kw in text_lower for kw in budget_keywords)
    # Check for numbers >= 20 followed by "tỷ" or "tỉ"
    budget_nums = re.findall(r'(\d+)\s*(?:tỷ|tỉ)', text_lower)
    for num in budget_nums:
        if int(num) >= 20:
            has_large_budget = True
            break
            
    if has_large_budget:
        reasons.append("+50: Ngân sách lớn (>= 20 tỷ hoặc tài chính mạnh/không thành vấn đề)")
        vip_matched = True
        
    # Luxury property check
    luxury_keywords = ["biệt thự đơn lập", "penthouse", "shophouse mặt đường lớn", "quỹ đất công nghiệp", "sàn văn phòng diện tích lớn", "shophouse", "sản văn phòng", "đất công nghiệp"]
    has_luxury = any(kw in text_lower for kw in luxury_keywords)
    if has_luxury:
        reasons.append("+50: Loại hình cao cấp (Biệt thự, Penthouse, Shophouse, Đất công nghiệp/Văn phòng)")
        vip_matched = True
        
    # Prime location check
    location_keywords = ["quận 1", "ven sông", "vinhomes ocean park", "phú mỹ hưng", "q1", "q.1"]
    has_prime_location = any(kw in text_lower for kw in location_keywords)
    # Avoid matching Q1 in negative context (e.g. Q1 giá 1 tỷ) unless it's indeed prime
    # If the text is negative, we handle it in SPAM rules.
    if has_prime_location:
        reasons.append("+50: Vị trí đắc địa (Quận 1, Ven sông, Vinhomes Ocean Park, Phú Mỹ Hưng)")
        vip_matched = True
        
    # Target profile check
    profile_keywords = ["chủ doanh nghiệp", "nhà đầu tư chuyên nghiệp", "mua sỉ", "mua số lượng lớn", "gom sỉ"]
    has_vip_profile = any(kw in text_lower for kw in profile_keywords)
    if has_vip_profile:
        reasons.append("+50: Đối tượng khách hàng VIP (Chủ DN, Mua sỉ, Đầu tư chuyên nghiệp)")
        vip_matched = True
        
    # Urgency & Transparency check
    urgency_keywords = ["pháp lý chuẩn 100%", "sổ hồng riêng", "gặp trực tiếp chủ đầu tư", "gặp trực tiếp giám đốc", "gặp trực tiếp chủ đầu tư để đàm phán"]
    has_urgency = any(kw in text_lower for kw in urgency_keywords)
    if has_urgency:
        reasons.append("+50: Tính cấp thiết & Minh bạch cao (Sổ hồng riêng, Pháp lý 100%, Đàm phán trực tiếp)")
        vip_matched = True

    if vip_matched:
        score += 50

    # Spam/Unqualified Criteria Indicators (-50)
    spam_matched = False
    
    # Unrealistic requirements (Q1 price 1-2B, central pool villa for cheap, price 2 million central)
    unrealistic_keywords = ["giá thấp vô lý", "q1 giá 1 tỷ", "quận 1 giá 1", "quận 1 giá 2 tỷ", "thuê nguyên căn giá 2 triệu", "yêu cầu phi thực tế", "ngân sách rất thấp"]
    has_unrealistic = any(kw in text_lower for kw in unrealistic_keywords) or ("q1" in text_lower and "1 tỷ" in text_lower)
    if has_unrealistic:
        reasons.append("-50: Yêu cầu phi thực tế (Giá quá rẻ so với thị trường)")
        spam_matched = True
        
    # No intent
    no_intent_keywords = ["nhầm số", "không có nhu cầu", "dữ liệu cũ", "nhầm ngành"]
    has_no_intent = any(kw in text_lower for kw in no_intent_keywords)
    if has_no_intent:
        reasons.append("-50: Không có nhu cầu (Nhầm số, dữ liệu cũ, nhầm ngành)")
        spam_matched = True
        
    # Uncooperative
    uncooperative_keywords = ["hỏi giá cho vui", "chưa có ý định mua", "thái độ không hợp tác", "không hợp tác"]
    has_uncooperative = any(kw in text_lower for kw in uncooperative_keywords)
    if has_uncooperative:
        reasons.append("-50: Khách hàng không thiện chí (Hỏi vui, chưa muốn mua, thái độ kém)")
        spam_matched = True
        
    # Spam/Advertising (Loans, Insurance, Advertising)
    spam_keywords = ["quảng cáo ngược", "bảo hiểm", "vay vốn", "mời chào", "chào mời", "quảng cáo ngược lại dịch vụ"]
    # Be careful not to mark "cần vay ngân hàng" as spam
    has_spam = any(kw in text_lower for kw in spam_keywords) and not ("cần hỗ trợ vay" in text_lower or "cần vay" in text_lower)
    if has_spam or "spam" in text_lower:
        reasons.append("-50: Tin nhắn Spam / Quảng cáo dịch vụ")
        spam_matched = True
        
    # Contact error
    contact_keywords = ["thuê bao", "gọi nhiều lần không bắt máy", "không phản hồi zalo", "gọi không liên lạc được"]
    has_contact_error = any(kw in text_lower for kw in contact_keywords)
    if has_contact_error:
        reasons.append("-50: Thông tin liên lạc lỗi (Thuê bao, không bắt máy, không trả lời Zalo)")
        spam_matched = True

    if spam_matched:
        score -= 50

    # Normal / Mid-range cases (e.g. Townhouse/Condo 3-10B, bank loan)
    # If no VIP and no SPAM matched, keep score or adjust slightly
    if not vip_matched and not spam_matched:
        normal_keywords = ["chung cư", "nhà phố", "căn hộ 2pn", "căn hộ 3pn", "tầm trung", "3-10 tỷ", "vay ngân hàng", "tư vấn thêm", "nhà phố liền kề"]
        if any(kw in text_lower for kw in normal_keywords):
            reasons.append("+10: Khách hàng nhu cầu thực phân khúc trung cấp (Căn hộ, Nhà phố 3-10 tỷ, Cần vay ngân hàng)")
            score = 60
        else:
            reasons.append("0: Khách hàng bình thường / Cần tư vấn thêm")
            score = 50

    # Ensure bounds
    score = max(0, min(100, score))
    
    # Categorize
    if score >= 80:
        status = "VIP"
    elif score < 30:
        status = "SPAM"
    else:
        status = "NORMAL"
        
    return score, "; ".join(reasons), status

def score_with_gemini(text, api_key):
    """
    Calls Gemini API to score the lead based on tieu_chi_cham_diem.txt.
    """
    criteria = read_criteria()
    prompt = f"""
Bạn là chuyên gia phân loại và chấm điểm khách hàng tiềm năng (Lead Scoring AI) cho ngành Bất động sản.
Nhiệm vụ của bạn là đọc mô tả nhu cầu của khách hàng, phân tích và chấm điểm dựa trên bộ quy tắc sau:

{criteria}

Hãy chấm điểm theo thang điểm từ 0 đến 100:
- Điểm mặc định bắt đầu: 50 điểm.
- Cộng thêm 50 điểm nếu có ít nhất một tiêu chí VIP/SIÊU TIỀM NĂNG.
- Trừ đi 50 điểm nếu có ít nhất một tiêu chí RÁC/SPAM/KHÔNG TIỀM NĂNG.
- Phân khúc trung cấp (chung cư, nhà phố 3-10 tỷ, cần vay ngân hàng, có nhu cầu thực cần tư vấn thêm): Giữ nguyên điểm (50 điểm) hoặc cộng thêm 10 điểm (thành 60 điểm).

Phân loại trạng thái (status):
- "VIP" nếu Điểm >= 80
- "SPAM" nếu Điểm < 30
- "NORMAL" nếu Điểm từ 30 đến 79

Hãy trả về kết quả dưới dạng JSON duy nhất, có cấu trúc như sau:
{{
  "score": <số nguyên từ 0 đến 100>,
  "status": "<VIP, NORMAL hoặc SPAM>",
  "reason": "<Nêu chi tiết các tiêu chí khớp, ví dụ: +50 vì yêu cầu biệt thự đơn lập, hoặc -50 vì nhầm số, thuê bao>"
}}

Nhu cầu khách hàng cần đánh giá:
"{text}"
"""
    
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    data = {
        "contents": [{
            "parts": [{
                "text": prompt
            }]
        }],
        "generationConfig": {
            "responseMimeType": "application/json"
        }
    }
    
    req_body = json.dumps(data).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=req_body,
        headers={"Content-Type": "application/json"}
    )
    
    try:
        with urllib.request.urlopen(req, timeout=15) as response:
            res_body = response.read().decode("utf-8")
            res_json = json.loads(res_body)
            text_response = res_json['candidates'][0]['content']['parts'][0]['text']
            
            # Clean response text if it has markdown ticks
            text_response = text_response.strip()
            if text_response.startswith("```json"):
                text_response = text_response[7:]
            if text_response.endswith("```"):
                text_response = text_response[:-3]
            text_response = text_response.strip()
            
            parsed = json.loads(text_response)
            score = int(parsed.get("score", 50))
            status = parsed.get("status", "NORMAL").upper()
            reason = parsed.get("reason", "Phân tích bằng AI.")
            
            # Validation
            score = max(0, min(100, score))
            if status not in ["VIP", "NORMAL", "SPAM"]:
                if score >= 80: status = "VIP"
                elif score < 30: status = "SPAM"
                else: status = "NORMAL"
                
            return score, f"[AI] {reason}", status
    except Exception as e:
        print(f"Gemini API request failed: {e}. Falling back to Rule-Based Scorer.")
        return rule_based_score(text)

def run_scoring(api_key=None, overwrite=False):
    """
    Loads data.csv, scores all leads, merges with existing leads_processed.json to keep manual edits
    unless overwrite is True, and writes back to leads_processed.json.
    """
    # Load existing processed data if any
    existing_data = {}
    if not overwrite and os.path.exists(JSON_PATH):
        try:
            with open(JSON_PATH, "r", encoding="utf-8") as f:
                leads_list = json.load(f)
                existing_data = {str(item["id"]): item for item in leads_list}
        except Exception as e:
            print(f"Error loading existing processed leads: {e}")

    # Read data.csv
    if not os.path.exists(CSV_PATH):
        print(f"Source file {CSV_PATH} not found!")
        return False

    processed_leads = []
    
    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            lead_id = str(row["id"])
            name = row["ten_khach"]
            phone = row["sdt"]
            desc = row["nhu_cau_mo_ta"]
            
            # Check if this lead is already manually reviewed/edited in existing_data
            if lead_id in existing_data and existing_data[lead_id].get("reviewed", False):
                # Keep the manually reviewed state
                processed_leads.append(existing_data[lead_id])
                continue
                
            # Perform scoring
            if api_key:
                score, reason, status = score_with_gemini(desc, api_key)
            else:
                score, reason, status = rule_based_score(desc)
                
            processed_leads.append({
                "id": int(lead_id),
                "ten_khach": name,
                "sdt": phone,
                "nhu_cau_mo_ta": desc,
                "score": score,
                "reason": reason,
                "status": status,
                "reviewed": False,
                "notes": existing_data.get(lead_id, {}).get("notes", "")
            })
            
    # Sort leads by id
    processed_leads.sort(key=lambda x: x["id"])
    
    # Save back to JSON
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(processed_leads, f, ensure_ascii=False, indent=2)
        
    print(f"Successfully processed and saved {len(processed_leads)} leads.")
    return True

# Alias for backward compatibility
calculate_score = rule_based_score

if __name__ == "__main__":
    # Test local keyword logic
    run_scoring()
