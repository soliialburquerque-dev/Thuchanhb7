---
name: lead-scoring-skill
description: |
  Lead Scoring and Customer Classification Skill for Real Estate.
  Automatically fetches leads from Google Sheets, scores them based on custom
  business rules, provides a human-in-the-loop validation interface, and
  exports final results to Excel.
---

# Lead Scoring & Automation Skill (Real Estate)

This skill describes the methodology, ruleset, and pipeline execution for classifying and scoring real estate leads.

## 1. Business Scoring Criteria

The leads are evaluated against a set of business rules from `tieu_chi_cham_diem.txt` to calculate potential scores (0 to 100 points, starting with a base of 50).

### A. VIP / Super Potential Leads (+50 Points)
Leads matching at least one of these criteria are awarded an additional 50 points (capped at 100 total):
- **Large Budget**: Mentions specific amounts of 20 Billion VND or more, or phrases like *"tài chính mạnh"*, *"không thành vấn đề"*.
- **Luxury Property Types**: Mentions *"Biệt thự đơn lập"*, *"Penthouse"*, *"Shophouse mặt đường lớn"*, *"Quỹ đất công nghiệp"*, *"Sàn văn phòng diện tích lớn"*.
- **Prime Locations**: Requesting properties in *"Quận 1"*, *"Ven sông"*, *"Vinhomes Ocean Park"*, *"Phú Mỹ Hưng"*.
- **VIP Customer Profiles**: Mentions being a *"Chủ doanh nghiệp"*, *"Nhà đầu tư chuyên nghiệp"*, *"Mua sỉ"*, or *"Mua số lượng lớn"*.
- **Urgency & Transparency**: Demands *"Pháp lý chuẩn 100%"*, *"Sổ hồng riêng"*, or *"Muốn gặp trực tiếp chủ đầu tư để đàm phán"*.

### B. Spam / Unqualified Leads (-50 Points)
Leads matching at least one of these criteria are penalized by 50 points (capped at 0 total):
- **Unrealistic Requests**: Requesting properties at ridiculously low prices (e.g., house in District 1 for 1-2 Billion, central pool villa for a few hundred million VND, or renting a whole house in center for 2 Million).
- **No Intent**: Expressions like *"Nhầm số"*, *"Không có nhu cầu"*, *"Dữ liệu cũ"*, or *"Nhầm ngành"*.
- **Uncooperative**: Expressions like *"Hỏi giá cho vui"*, *"Chưa có ý định mua"*, or *"Thái độ không hợp tác"*.
- **Spam / Advertisement**: Advertising unrelated services like *"Bảo hiểm"*, *"Vay vốn"*, or *"Mời chào dịch vụ".*
- **Contact Errors**: Tagged with *"Thuê bao"*, *"Gọi nhiều lần không bắt máy"*, or *"Không phản hồi Zalo"*.

### C. Normal / Mid-range Leads (0 to +10 Points)
Leads showing genuine intent for mid-market properties are kept at base score (50) or slightly increased (+10 points to 60):
- Apartment or townhouse search in the range of 3-10 Billion.
- Needs bank loan assistance or details on interest rates.
- Genuine interest but requires further consulting on legal or location aspects.

---

## 2. Technical Pipeline Architecture

The scoring system is implemented as a hybrid engine:

1. **Scoring Logic (`scorer.py`)**: Runs locally using custom Regex and string matching, OR calls the Google Gemini API to analyze descriptions using structured schema outputs if an API key is provided.
2. **REST API Server (`server.py`)**: A zero-dependency Python script using standard `http.server` serving the dashboard and supporting endpoints for listing, updating, recalculating, and exporting leads.
3. **Glassmorphic Frontend Dashboard (`public/`)**: Serves as the Human-in-the-loop review app where users can change classifications, adjust scores, type comments, and download the Excel file.

### REST API Endpoints:
- `GET /api/leads?search=...&status=vip|normal|spam&reviewed=true|false&page=1&limit=20`: Fetches filtered paginated leads and summary stats.
- `POST /api/leads/update`: Updates a lead's score, status, manual notes, and reviewed tag.
- `POST /api/leads/recalculate`: Triggers scorer pipeline (accepts optional JSON body `{"api_key": "...", "overwrite": true}`).
- `POST /api/leads/export`: Automatically creates a beautifully formatted Excel sheet using `xlsxwriter` and returns it as a direct download.

---

## 3. Execution & Deployment Guide

### Prerequisites
Make sure Python 3.x and the `xlsxwriter` library are installed.

### How to Run:
1. Navigate to the project directory:
   ```powershell
   cd c:\Users\Administrator\Documents\WORK\Software\AI\AGENTIC\Thuchanhb7
   ```
2. Start the HTTP Server:
   ```powershell
   python server.py
   ```
3. Open your browser and navigate to:
   [http://localhost:8000](http://localhost:8000)

### Integrating Gemini AI:
- In the dashboard, click the **"Chấm điểm lại"** (Recalculate) button.
- Paste your **Gemini API Key** into the text input.
- Click **"Bắt đầu chấm điểm"** (Start Scoring) to run the AI-based analysis.
- If no key is entered, it automatically falls back to local regex matching.

---

## 4. Excel Delivery Export Layout
The exported Excel spreadsheet features:
- A professional corporate navy theme header row.
- Alternating light blue background shading for table rows (Zebra striping).
- Explicit status-based conditional coloring (Green background for **VIP**, Yellow for **NORMAL**, Red for **SPAM**).
- Automatic width adjustments for columns.
- Dedicated columns showing the system's reasoning vs human notes.
