import http.server
import socketserver
import json
import os
import urllib.parse
from scorer import run_scoring, JSON_PATH, CSV_PATH
import xlsxwriter

PORT = 8000
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PUBLIC_DIR = os.path.join(BASE_DIR, "public")

class LeadScoringRequestHandler(http.server.SimpleHTTPRequestHandler):
    def translate_path(self, path):
        # Serve static files from public/ directory
        path = super().translate_path(path)
        rel_path = os.path.relpath(path, os.getcwd())
        return os.path.join(PUBLIC_DIR, rel_path)

    def end_headers(self):
        # Add CORS headers
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200, "OK")
        self.end_headers()

    def do_GET(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path

        if path == "/api/leads":
            self.handle_get_leads(parsed_url.query)
        else:
            # Handle static files
            if path == "/" or path == "":
                self.path = "/index.html"
            super().do_GET()

    def do_POST(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path

        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length).decode('utf-8') if content_length > 0 else ""

        if path == "/api/leads/update":
            self.handle_update_lead(body)
        elif path == "/api/leads/recalculate":
            self.handle_recalculate(body)
        elif path == "/api/leads/export":
            self.handle_export(body)
        else:
            self.send_error(404, "Endpoint not found")

    def handle_get_leads(self, query_str):
        query_params = urllib.parse.parse_qs(query_str)
        
        # Load leads
        if not os.path.exists(JSON_PATH):
            # Run initial scorer if file doesn't exist
            run_scoring()

        try:
            with open(JSON_PATH, "r", encoding="utf-8") as f:
                leads = json.load(f)
        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": f"Failed to load leads: {e}"}).encode('utf-8'))
            return

        # Filtering
        search = query_params.get("search", [""])[0].lower()
        status_filter = query_params.get("status", ["all"])[0].upper()
        reviewed_filter = query_params.get("reviewed", ["all"])[0]

        filtered = []
        for lead in leads:
            # Search filter
            if search:
                name_match = search in lead["ten_khach"].lower()
                phone_match = search in lead["sdt"].lower()
                desc_match = search in lead["nhu_cau_mo_ta"].lower()
                if not (name_match or phone_match or desc_match):
                    continue

            # Status filter
            if status_filter != "ALL" and lead["status"] != status_filter:
                continue

            # Reviewed filter
            if reviewed_filter != "all":
                is_reviewed = lead.get("reviewed", False)
                if reviewed_filter == "true" and not is_reviewed:
                    continue
                if reviewed_filter == "false" and is_reviewed:
                    continue

            filtered.append(lead)

        # Pagination
        try:
            page = int(query_params.get("page", [1])[0])
            limit = int(query_params.get("limit", [20])[0])
        except ValueError:
            page = 1
            limit = 20

        total = len(filtered)
        start_idx = (page - 1) * limit
        end_idx = start_idx + limit
        paginated_leads = filtered[start_idx:end_idx]

        # Calculate statistics
        stats = {
            "total": len(leads),
            "vip": sum(1 for l in leads if l["status"] == "VIP"),
            "normal": sum(1 for l in leads if l["status"] == "NORMAL"),
            "spam": sum(1 for l in leads if l["status"] == "SPAM"),
            "reviewed": sum(1 for l in leads if l.get("reviewed", False))
        }

        response_data = {
            "leads": paginated_leads,
            "total": total,
            "page": page,
            "limit": limit,
            "pages": (total + limit - 1) // limit,
            "stats": stats
        }

        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(json.dumps(response_data, ensure_ascii=False).encode('utf-8'))

    def handle_update_lead(self, body_str):
        try:
            data = json.loads(body_str)
            lead_id = int(data.get("id"))
            new_score = int(data.get("score"))
            new_status = data.get("status").upper()
            new_notes = data.get("notes", "")
            new_reviewed = bool(data.get("reviewed", True))
        except (ValueError, TypeError, KeyError) as e:
            self.send_response(400)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": f"Invalid request body: {e}"}).encode('utf-8'))
            return

        if not os.path.exists(JSON_PATH):
            self.send_response(404)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": "Leads file not found"}).encode('utf-8'))
            return

        try:
            with open(JSON_PATH, "r", encoding="utf-8") as f:
                leads = json.load(f)

            found = False
            for lead in leads:
                if lead["id"] == lead_id:
                    lead["score"] = max(0, min(100, new_score))
                    lead["status"] = new_status
                    lead["notes"] = new_notes
                    lead["reviewed"] = new_reviewed
                    # Append "(Manual Review)" if not already present in reason
                    if "reviewed" in lead and not lead.get("reason", "").startswith("[Manual]"):
                        lead["reason"] = f"[Manual] {lead.get('reason', '')}"
                    found = True
                    break

            if not found:
                self.send_response(404)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": f"Lead ID {lead_id} not found"}).encode('utf-8'))
                return

            with open(JSON_PATH, "w", encoding="utf-8") as f:
                json.dump(leads, f, ensure_ascii=False, indent=2)

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"success": True}).encode('utf-8'))

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode('utf-8'))

    def handle_recalculate(self, body_str):
        api_key = None
        overwrite = False
        if body_str:
            try:
                data = json.loads(body_str)
                api_key = data.get("api_key")
                overwrite = bool(data.get("overwrite", False))
            except Exception:
                pass

        try:
            success = run_scoring(api_key=api_key, overwrite=overwrite)
            if success:
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"success": True}).encode('utf-8'))
            else:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": "Failed to run scoring pipeline."}).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode('utf-8'))

    def handle_export(self, body_str):
        # We generate a formatted Excel spreadsheet using xlsxwriter and write it directly to the response
        if not os.path.exists(JSON_PATH):
            self.send_error(404, "No processed data to export")
            return

        try:
            with open(JSON_PATH, "r", encoding="utf-8") as f:
                leads = json.load(f)

            temp_xlsx_path = os.path.join(BASE_DIR, "temp_export.xlsx")
            
            # Create a new workbook
            workbook = xlsxwriter.Workbook(temp_xlsx_path)
            worksheet = workbook.add_worksheet("Leads Scoring")

            # Set up styles
            header_format = workbook.add_format({
                'bold': True,
                'font_name': 'Segoe UI',
                'font_size': 11,
                'font_color': 'white',
                'bg_color': '#1F4E78',
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

            # Style formats for VIP and SPAM status
            vip_format = workbook.add_format({
                'font_name': 'Segoe UI',
                'font_size': 10,
                'bold': True,
                'font_color': '#375623',
                'bg_color': '#E2EFDA',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            spam_format = workbook.add_format({
                'font_name': 'Segoe UI',
                'font_size': 10,
                'bold': True,
                'font_color': '#C00000',
                'bg_color': '#FCE4D6',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            normal_format = workbook.add_format({
                'font_name': 'Segoe UI',
                'font_size': 10,
                'bold': True,
                'font_color': '#333F48',
                'bg_color': '#FFF2CC',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            # Set column widths
            worksheet.set_column('A:A', 8)   # ID
            worksheet.set_column('B:B', 20)  # Tên khách hàng
            worksheet.set_column('C:C', 15)  # Số điện thoại
            worksheet.set_column('D:D', 65)  # Nhu cầu mô tả
            worksheet.set_column('E:E', 12)  # Điểm số
            worksheet.set_column('F:F', 15)  # Trạng thái
            worksheet.set_column('G:G', 12)  # Đã duyệt
            worksheet.set_column('H:H', 25)  # Ghi chú thủ công
            worksheet.set_column('I:I', 50)  # Lý do hệ thống

            # Set row heights
            worksheet.set_row(0, 30)

            # Write headers
            headers = ["ID", "Khách hàng", "Số điện thoại", "Nhu cầu mô tả", "Điểm số", "Trạng thái", "Đã duyệt", "Ghi chú thủ công", "Lý do hệ thống"]
            for col_idx, header in enumerate(headers):
                worksheet.write(0, col_idx, header, header_format)

            # Write data rows
            for row_idx, lead in enumerate(leads, start=1):
                worksheet.set_row(row_idx, 22)
                worksheet.write(row_idx, 0, lead["id"], center_format)
                worksheet.write(row_idx, 1, lead["ten_khach"], cell_format)
                worksheet.write(row_idx, 2, lead["sdt"], center_format)
                worksheet.write(row_idx, 3, lead["nhu_cau_mo_ta"], cell_format)
                worksheet.write(row_idx, 4, lead["score"], center_format)

                # Write status with conditional formatting
                status = lead["status"]
                if status == "VIP":
                    worksheet.write(row_idx, 5, status, vip_format)
                elif status == "SPAM":
                    worksheet.write(row_idx, 5, status, spam_format)
                else:
                    worksheet.write(row_idx, 5, status, normal_format)

                reviewed_str = "Đã duyệt" if lead.get("reviewed", False) else "Chưa duyệt"
                worksheet.write(row_idx, 6, reviewed_str, center_format)
                worksheet.write(row_idx, 7, lead.get("notes", ""), cell_format)
                worksheet.write(row_idx, 8, lead["reason"], cell_format)

            workbook.close()

            # Read workbook bytes
            with open(temp_xlsx_path, "rb") as f:
                xlsx_bytes = f.read()

            # Clean up temp file
            os.remove(temp_xlsx_path)

            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=Danh_sach_khach_hang_tiem_nang.xlsx")
            self.send_header("Content-Length", str(len(xlsx_bytes)))
            self.end_headers()
            self.wfile.write(xlsx_bytes)

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": f"Excel generation failed: {e}"}).encode('utf-8'))

if __name__ == "__main__":
    handler = LeadScoringRequestHandler
    print(f"Starting server on http://localhost:{PORT}")
    with socketserver.TCPServer(("", PORT), handler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down server.")
