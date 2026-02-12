from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import pandas as pd
import io
import re
import base64
from datetime import datetime

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
    
    # 定義關鍵字
    fin_start_keywords = ["財務指標表", "項目", "最新季"]
    
    # 這是舊邏輯用的，新邏輯主要靠 | 判斷
    fin_item_keywords = ["營業收入", "總資產", "負債比", "流動資產", "流動負債", "現金流", "EPS"]

    for line in lines:
        line = clean_text(line)
        if not line: continue
        
        # --- A. 區塊切換偵測 ---
        if "Pre-check List" in line:
            section = "pre_check"; current_row = []; continue
        elif any(k in line for k in fin_start_keywords) and "財務指標" in line:
            section = "finance"; current_row = []; continue
        elif "非財務條件" in line:
            section = "group_a"; current_row = []; continue
        elif any(marker in line for marker in ["3️⃣", "4️⃣", "5️⃣", "【Group A", "【核保結論", "【風險評級", "總結"]):
            section = "other"
            other.append(("header", line, ""))
            continue

        # --- B. 內容填入邏輯 ---
        try:
            # 1. Pre-check
            if section == "pre_check":
                if "項次" in line or "檢核項目" in line or "---" in line: continue
                # 嘗試處理表格格式 | 1 | 項目 | 結果 |
                if "|" in line:
                    parts = [p.strip() for p in line.split('|') if p.strip()]
                    if len(parts) >= 2: pre_check.append(parts[:3])
                else:
                    # 舊有的條列式解析
                    if line.isdigit() and len(line) < 3:
                        if current_row: 
                            while len(current_row) < 3: current_row.append("")
                            pre_check.append(current_row)
                        current_row = [line]
                    elif current_row:
                        target_idx = 1 if len(current_row) == 1 else 2
                        if len(current_row) <= target_idx: current_row.append(line)
                        else: current_row[2] += f"\n{line}"

            # 2. Finance (這裡做了重大修正！)
            elif section == "finance":
                if "最新季" in line or "去年同期" in line or "---" in line: continue
                
                # ★★★ 新增：優先檢查是否為 Markdown 表格 (有 | 符號)
                if "|" in line:
                    # 拆分並過濾空白
                    parts = [p.strip() for p in line.split('|') if p.strip()]
                    # 如果切出來有東西，就直接當作一行資料存進去
                    if len(parts) > 1:
                        # 確保長度補齊到 5 欄 (項目, 最新, 去年, 前一, 前二)
                        while len(parts) < 5: parts.append("")
                        finance.append(parts[:5]) # 只取前5個，避免格式跑掉
                    continue

                # 舊邏輯 (防呆用，處理沒用表格的情況)
                is_new_item = any(k in line for k in fin_item_keywords) or (not any(c.isdigit() for c in line) and len(line) < 10)
                if is_new_item:
                    if current_row:
                        while len(current_row) < 5: current_row.append("")
                        finance.append(current_row)
                    current_row = [line]
                elif current_row and any(c.isdigit() for c in line):
                    if len(current_row) < 5: current_row.append(line)

            # 3. Group A
            elif section == "group_a":
                if "項次" in line or "項目" in line or "---" in line: continue
                if "|" in line:
                    parts = [p.strip() for p in line.split('|') if p.strip()]
                    if len(parts) >= 2: group_a.append(parts[:3])
                else:
                    if line.isdigit() and len(line) < 3:
                        if current_row:
                            while len(current_row) < 3: current_row.append("")
                            group_a.append(current_row)
                        current_row = [line]
                    elif current_row:
                        target_idx = 1 if len(current_row) == 1 else 2
                        if len(current_row) <= target_idx: current_row.append(line)
                        else: current_row[2] += f"\n{line}"

            # 4. Other
            else:
                if "---" in line: continue
                if line.startswith("(") and ")" in line and len(line) < 15:
                    other.append(("subheader", line, ""))
                elif "|" in line:
                    parts = [p.strip() for p in line.split('|') if p.strip()]
                    if len(parts) >= 2: other.append(("kv", parts[0], parts[1]))
                elif line.startswith("* ") or line.startswith("- "):
                    clean_l = line.replace("* ", "").replace("- ", "")
                    if "：" in clean_l or "=" in clean_l or "≈" in clean_l:
                        for sep in ["：", "=", "≈"]:
                            if sep in clean_l:
                                parts = clean_l.split(sep, 1)
                                val_prefix = "≈ " if sep == "≈" else ""
                                other.append(("kv", parts[0].strip(), val_prefix + parts[1].strip()))
                                break
                    else:
                        other.append(("full", line, ""))
                elif "：" in line: 
                    parts = line.split("：", 1)
                    other.append(("kv", parts[0], parts[1]))
                elif "≈" in line: 
                    parts = line.split("≈", 1)
                    other.append(("kv", parts[0].strip(), "≈ " + parts[1].strip()))
                elif "=" in line: 
                    parts = line.split("=", 1)
                    other.append(("kv", parts[0].strip(), parts[1].strip()))
                else:
                    other.append(("full", line, ""))

        except Exception:
            continue 

    # 結尾補齊
    if section == "pre_check" and current_row: pre_check.append(current_row + [""]*(3-len(current_row)))
    elif section == "finance" and current_row: finance.append(current_row + [""]*(5-len(current_row)))
    elif section == "group_a" and current_row: group_a.append(current_row + [""]*(3-len(current_row)))

    return pre_check, finance, group_a, other

# --- 3. API 接口 ---
@app.post("/generate_excel")
async def generate_excel(request: ReportRequest):
    try:
        if not request.text:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                writer.book.add_worksheet("Error")
            buffer.seek(0)
            return {"filename": "Error_No_Text.xlsx", "file_content_base64": base64.b64encode(buffer.read()).decode('utf-8')}

        pre, fin, grp, oth = parse_copilot_final(request.text)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            sheet_name = '核保評估表'
            workbook = writer.book
            worksheet = workbook.add_worksheet(sheet_name)
            
            # 定義樣式
            header_fmt = workbook.add_format({'bold': True, 'fg_color': '#0070C0', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            subheader_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'left'}) 
            cell_fmt = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'left'})
            num_fmt = workbook.add_format({'border': 1, 'align': 'right', 'valign': 'vcenter'})
            section_fmt = workbook.add_format({'bold': True, 'fg_color': '#E0E0E0', 'border': 1})
            full_text_fmt = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'left', 'bg_color': '#FAFAFA'})
            formula_val_fmt = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'left', 'font_color': '#333333'})

            curr = 0
            
            def safe_write(row, col, value, fmt):
                try: worksheet.write(row, col, str(value), fmt)
                except: pass

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
                            # 如果是數字，用 num_fmt，如果是文字(如 N/A)，用 cell_fmt
                            fmt = num_fmt if str(val).replace(",","").replace(".","").replace("%","").isdigit() else num_fmt
                            safe_write(curr, i, val, fmt)
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
            
            # 4. Other
            if oth:
                for item_type, key, value in oth:
                    try:
                        key = str(key) if key else ""
                        value = str(value) if value else ""

                        if item_type == "header": 
                            try: worksheet.merge_range(curr, 0, curr, 4, key, section_fmt)
                            except: worksheet.write(curr, 0, key, section_fmt)
                        
                        elif item_type == "subheader": 
                            try: worksheet.merge_range(curr, 0, curr, 4, key, subheader_fmt)
                            except: worksheet.write(curr, 0, key, subheader_fmt)

                        elif item_type == "full": 
                            try: worksheet.merge_range(curr, 0, curr, 4, key, full_text_fmt)
                            except: worksheet.write(curr, 0, key, full_text_fmt)
                        
                        elif item_type == "kv":
                            worksheet.write(curr, 0, key, cell_fmt)
                            try: worksheet.merge_range(curr, 1, curr, 4, value, formula_val_fmt)
                            except: worksheet.write(curr, 1, value, formula_val_fmt)
                        
                        curr += 1
                    except Exception:
                        curr += 1

            worksheet.set_column('A:A', 25)
            worksheet.set_column('B:B', 20)
            worksheet.set_column('C:E', 18)

        buffer.seek(0)
        file_base64 = base64.b64encode(buffer.read()).decode('utf-8')
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_company = re.sub(r'[\\/*?:"<>|]', "", request.company_name)
        if not safe_company: safe_company = "Report"
        
        final_filename = f"{safe_company}_{timestamp}.xlsx"

        return {
            "filename": final_filename,
            "file_content_base64": file_base64
        }

    except Exception as e:
        try:
            err_buffer = io.BytesIO()
            with pd.ExcelWriter(err_buffer, engine='xlsxwriter') as writer:
                ws = writer.book.add_worksheet("Error")
                ws.write(0, 0, "System Error: " + str(e))
            err_buffer.seek(0)
            return {"filename": "Error_Log.xlsx", "file_content_base64": base64.b64encode(err_buffer.read()).decode('utf-8')}
        except:
            raise HTTPException(status_code=500, detail=str(e))
