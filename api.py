from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import pandas as pd
import io
import re
import base64

app = FastAPI()

# --- 1. 定義請求格式 ---
class ReportRequest(BaseModel):
    text: str
    company_name: str = "Company"

# --- 2. 核心解析邏輯 ---
def clean_text(text):
    if not text: return ""
    return text.replace("**", "").replace("###", "").strip()

def parse_copilot_final(text):
    pre_check, finance, group_a, other = [], [], [], []
    section = "other"
    
    if not text:
        return pre_check, finance, group_a, other

    lines = text.split('\n')
    current_row = []
    
    fin_start_keywords = ["財務指標表", "項目", "最新季"]
    fin_item_keywords = ["營業收入", "總資產", "負債比", "流動資產", "流動負債", "現金流", "EPS"]

    for line in lines:
        line = clean_text(line)
        if not line: continue
        
        # 區塊偵測
        if "Pre-check List" in line:
            section = "pre_check"; current_row = []; continue
        elif any(k in line for k in fin_start_keywords) and "財務指標" in line:
            section = "finance"; current_row = []; continue
        elif "非財務條件" in line:
            section = "group_a"; current_row = []; continue
        elif "Group A 判定" in line or "財務評分" in line or "綜合建議" in line:
            section = "other"
            other.append(("header", line, ""))
            continue

        # 填入邏輯
        try:
            if section == "pre_check":
                if "項次" in line or "檢核項目" in line: continue
                if line.isdigit() and len(line) < 3:
                    if current_row: 
                        while len(current_row) < 3: current_row.append("")
                        pre_check.append(current_row)
                    current_row = [line]
                elif current_row:
                    target_idx = 1 if len(current_row) == 1 else 2
                    if len(current_row) <= target_idx: current_row.append(line)
                    else: current_row[2] += f"\n{line}"

            elif section == "finance":
                if "最新季" in line or "去年同期" in line: continue
                is_new_item = any(k in line for k in fin_item_keywords) or (not any(c.isdigit() for c in line) and len(line) < 10)
                if is_new_item:
                    if current_row:
                        while len(current_row) < 5: current_row.append("")
                        finance.append(current_row)
                    current_row = [line]
                elif current_row and any(c.isdigit() for c in line):
                    if len(current_row) < 5: current_row.append(line)

            elif section == "group_a":
                if "項次" in line or "項目" in line: continue
                if line.isdigit() and len(line) < 3:
                    if current_row:
                        while len(current_row) < 3: current_row.append("")
                        group_a.append(current_row)
                    current_row = [line]
                elif current_row:
                    target_idx = 1 if len(current_row) == 1 else 2
                    if len(current_row) <= target_idx: current_row.append(line)
                    else: current_row[2] += f"\n{line}"

            else: # Other
                if "：" in line: parts = line.split("：", 1); other.append(("kv", parts[0], parts[1]))
                elif "=" in line: parts = line.split("=", 1); other.append(("kv", parts[0].strip(), parts[1].strip()))
                else: other.append(("full", line, ""))
        except Exception:
            continue # 如果這一行解析失敗，跳過就好，不要崩潰

    # 處理剩餘
    if section == "pre_check" and current_row: pre_check.append(current_row + [""]*(3-len(current_row)))
    elif section == "finance" and current_row: finance.append(current_row + [""]*(5-len(current_row)))
    elif section == "group_a" and current_row: group_a.append(current_row + [""]*(3-len(current_row)))

    return pre_check, finance, group_a, other

# --- 3. API 接口 (含強力防錯) ---
@app.post("/generate_excel")
async def generate_excel(request: ReportRequest):
    try:
        # A. 防呆檢查
        if not request.text:
            # 如果沒有文字，回傳一個空的 Excel，避免報錯
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                writer.book.add_worksheet("Error_No_Text")
            buffer.seek(0)
            return {"filename": "Error_No_Data.xlsx", "file_content_base64": base64.b64encode(buffer.read()).decode('utf-8')}

        # B. 解析文字
        pre, fin, grp, oth = parse_copilot_final(request.text)
        
        # C. 製作 Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            sheet_name = '核保評估表'
            workbook = writer.book
            worksheet = workbook.add_worksheet(sheet_name)
            
            # 定義樣式
            header_fmt = workbook.add_format({'bold': True, 'fg_color': '#0070C0', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            cell_fmt = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'left'})
            num_fmt = workbook.add_format({'border': 1, 'align': 'right', 'valign': 'vcenter'})
            section_fmt = workbook.add_format({'bold': True, 'fg_color': '#E0E0E0', 'border': 1})
            full_text_fmt = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'left', 'bg_color': '#FAFAFA'})
            formula_val_fmt = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'left', 'font_color': '#333333'})

            curr = 0
            
            # Helper function: 安全寫入，避免當機
            def safe_write(row, col, value, fmt):
                try:
                    worksheet.write(row, col, str(value), fmt)
                except:
                    pass

            # 1. Pre-check
            if pre:
                try:
                    worksheet.merge_range(curr, 0, curr, 2, "一、Pre-check List", section_fmt)
                    curr += 1
                    worksheet.write_row(curr, 0, ["項次", "檢核項目", "判斷結果"], header_fmt)
                    curr += 1
                    for row in pre:
                        for c, val in enumerate(row): 
                            if c < 3: safe_write(curr, c, val, cell_fmt)
                        curr += 1
                    curr += 1
                except: curr += 1
            
            # 2. Finance
            if fin:
                try:
                    worksheet.merge_range(curr, 0, curr, 4, "二、財務指標表", section_fmt)
                    curr += 1
                    worksheet.write_row(curr, 0, ["項目", "最新季", "去年同期", "前一年度", "前兩年度"], header_fmt)
                    curr += 1
                    for row in fin:
                        safe_write(curr, 0, row[0], cell_fmt)
                        for i in range(1, 5): 
                            val = row[i] if i < len(row) else ""
                            safe_write(curr, i, val, num_fmt)
                        curr += 1
                    curr += 1
                except: curr += 1

            # 3. Group A
            if grp:
                try:
                    worksheet.merge_range(curr, 0, curr, 2, "三、非財務條件", section_fmt)
                    curr += 1
                    worksheet.write_row(curr, 0, ["項次", "項目", "判斷"], header_fmt)
                    curr += 1
                    for row in grp:
                        for c, val in enumerate(row): 
                            if c < 3: safe_write(curr, c, val, cell_fmt)
                        curr += 1
                    curr += 1
                except: curr += 1
            
            # 4. Other (最容易出錯的地方，加強保護)
            if oth:
                for item_type, key, value in oth:
                    try:
                        key = str(key) if key else ""
                        value = str(value) if value else ""

                        if item_type == "header": 
                            # 嘗試合併，失敗則退回普通寫入
                            try: worksheet.merge_range(curr, 0, curr, 4, key, section_fmt)
                            except: worksheet.write(curr, 0, key, section_fmt)
                        
                        elif item_type == "full": 
                            try: worksheet.merge_range(curr, 0, curr, 4, key, full_text_fmt)
                            except: worksheet.write(curr, 0, key, full_text_fmt)
                        
                        elif item_type == "kv":
                            worksheet.write(curr, 0, key, cell_fmt)
                            try: worksheet.merge_range(curr, 1, curr, 4, value, formula_val_fmt)
                            except: worksheet.write(curr, 1, value, formula_val_fmt)
                        
                        curr += 1
                    except Exception as e:
                        # 如果這行真的沒救了，印出 Log 但不讓程式掛掉
                        print(f"Skipping row due to error: {e}")
                        curr += 1

            # 設定欄寬
            worksheet.set_column('A:A', 20)
            worksheet.set_column('B:B', 30)
            worksheet.set_column('C:E', 15)

        # D. 轉成 Base64
        buffer.seek(0)
        file_base64 = base64.b64encode(buffer.read()).decode('utf-8')
        
        # 確保檔名安全
        safe_filename = re.sub(r'[\\/*?:"<>|]', "", request.company_name)
        if not safe_filename: safe_filename = "Report"

        return {
            "filename": f"{safe_filename}_Report.xlsx",
            "file_content_base64": file_base64
        }

    except Exception as e:
        # 捕捉所有未知錯誤，回傳 200 但帶有錯誤訊息的 Excel，而不是 502
        print(f"CRITICAL ERROR: {str(e)}")
        
        # 嘗試回傳一個錯誤報告 Excel
        try:
            err_buffer = io.BytesIO()
            with pd.ExcelWriter(err_buffer, engine='xlsxwriter') as writer:
                wb = writer.book
                ws = wb.add_worksheet("Error_Log")
                ws.write(0, 0, "System Error Occurred")
                ws.write(1, 0, str(e))
            err_buffer.seek(0)
            err_b64 = base64.b64encode(err_buffer.read()).decode('utf-8')
            return {"filename": "Error_Log.xlsx", "file_content_base64": err_b64}
        except:
            # 真的沒救了才丟 500
            raise HTTPException(status_code=500, detail=f"Server Error: {str(e)}")
