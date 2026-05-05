import streamlit as st
import os
import json
import pandas as pd
from datetime import datetime, date
import calendar
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENTATION
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
from fpdf import FPDF
import io
from github import Github  # <-- Added for GitHub backups

# --- 1. CONFIGURATION & SETUP ---
st.set_page_config(page_title="AINP Grant Manager", page_icon="🌾", layout="wide")

# Directory Setup
DIRS = ['data', 'documents', 'fonts', 'logos']
for d in DIRS:
    if not os.path.exists(d):
        os.makedirs(d)

# AI Setup
api_key = st.secrets.get("GEMINI_API_KEY")
genai.configure(api_key=api_key)
model_pro = genai.GenerativeModel('gemini-3.1-pro-preview')   

# Load Logos if exist
NAU_LOGO = 'logos/nau_logo.png' if os.path.exists('logos/nau_logo.png') else None
ICAR_LOGO = 'logos/icar_logo.png' if os.path.exists('logos/icar_logo.png') else None
GUJARATI_FONT = 'fonts/NotoSansGujarati-Regular.ttf'

# Define standard Heads
BUDGET_HEADS = [
    "Pay and Allowances",
    "Travelling Allowances (TA)",
    "Other Recurring Contingencies (ORC)",
    "Non-Recurring Contingencies (Equipments/Works)",
    "TSP"
]

# --- 2. HELPER FUNCTIONS ---

# --- A. Data Persistence & GitHub Backup ---
def backup_to_github(filepath, content_str):
    """Pushes saved data back to GitHub so it isn't lost when Streamlit sleeps."""
    github_token = st.secrets.get("GITHUB_TOKEN")
    
    if not github_token:
        # If no token is found, just skip the backup (useful for local testing)
        return

    try:
        g = Github(github_token)
        repo = g.get_repo("vkcvaibhav/AINP-Grant-Manager") # Your exact repo name
        
        # Check if the file already exists in GitHub
        try:
            contents = repo.get_contents(filepath)
            # If it exists, UPDATE it
            repo.update_file(contents.path, f"Auto-backup {filepath}", content_str, contents.sha)
        except Exception:
            # If it does not exist yet, CREATE it
            repo.create_file(filepath, f"Auto-create {filepath}", content_str)
            
    except Exception as e:
        st.warning(f"Failed to backup to GitHub. Error: {e}")

def get_fy_filename(fy):
    return f'data/grant_data_{fy.replace("-", "_")}.json'

def load_data(fy):
    filename = get_fy_filename(fy)
    if os.path.exists(filename):
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # Ensure new quarterly structure exists in old data files
            if 'quarterly_allocations' not in data:
                data['quarterly_allocations'] = {"Q1": {}, "Q2": {}, "Q3": {}, "Q4": {}}
            return data
    else:
        return {
            "financial_year": fy,
            "allocation": {}, 
            "revised_allocation": {},
            "quarterly_allocations": {"Q1": {}, "Q2": {}, "Q3": {}, "Q4": {}},
            "installments": [], 
            "expenditure": [],
            "latest_quarter": "Full Year", 
            "latest_date": "N/A"
        }

def save_data(data, fy):
    filename = get_fy_filename(fy)
    
    # 1. Save locally for immediate app use
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
        
    # 2. Push to GitHub for permanent backup
    json_str = json.dumps(data, indent=4, ensure_ascii=False)
    backup_to_github(filename, json_str)


# --- B. Document Processing (AI PDF Reader) ---
def process_upload_with_ai(uploaded_file, prompt_task):
    """Uses Gemini to natively read the PDF and extract JSON structured data."""
    try:
        # Package the raw PDF file directly for Gemini
        pdf_data = {
            "mime_type": "application/pdf",
            "data": uploaded_file.getvalue()
        }

        full_prompt = f"""
        Analyze the attached document content and extract the required information 
        as a structured JSON object.

        DOCUMENT TYPE/CONTEXT: {prompt_task['context']}
        REQUIRED JSON STRUCTURE: {json.dumps(prompt_task['structure'], indent=2)}

        IMPORTANT: Only return the JSON object, nothing else. If a field cannot be found, set it to null.
        For currency, return numbers only (e.g., 100000). Convert dates to YYYY-MM-DD format.
        """
        
        # Send the text prompt and the raw PDF directly to Gemini
        response = model_pro.generate_content([full_prompt, pdf_data])
        
        # Parse JSON from response (MUST BE ON ONE LINE)
        json_str = response.text.replace('```json', '').replace('```', '').strip()
        extracted_data = json.loads(json_str)
        return extracted_data

    except Exception as e:
        st.error(f"Error processing document with AI: {e}")
        return None

# --- C. Output Generation (PDF/WORD) ---
def add_bottom_border(paragraph, size='24'):
    """Adds a bottom border to a paragraph. Size '24' is a thick 3pt line, '8' is a thin 1pt line."""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), size) 
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)

def generate_comptroller_docx(ref_no, letter_date, body_text, amt_words, pay_amt, total_rec, non_rec_amt, total_amt):
    """Generates the Native Microsoft Word format for the Comptroller Letter matching the exact layout."""
    doc = Document()
    
    # Set document margins tight to match letterhead
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    # Strip default spacing to ensure tables pack tightly together
    style = doc.styles['Normal']
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)
        
    # 1. Header Table for Logos and Center Text
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    table.columns[0].width = Inches(1.8)
    table.columns[1].width = Inches(3.6)
    table.columns[2].width = Inches(1.4)
    
    # Left Logo (NAU)
    if 'NAU_LOGO' in globals() and NAU_LOGO and os.path.exists(NAU_LOGO):
        cell_left = table.cell(0, 0)
        cell_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p_left = cell_left.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_left.paragraph_format.space_after = Pt(0)
        p_left.paragraph_format.line_spacing = 0.5
        
        r_left = p_left.add_run()
        r_left.add_picture(NAU_LOGO, width=Inches(1.8))
        
    # Center Text
    p_center = table.cell(0, 1).paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_center.paragraph_format.line_spacing = 0.85 # Tight gap below the main heading
    
    r1 = p_center.add_run("કીટકશાસ્ત્ર વિભાગ\n")
    r1.bold = True
    r1.font.size = Pt(22) # Large, prominent heading
    
    r2 = p_center.add_run("ન. મ. કૃષિ મહાવિદ્યાલય\nનવસારી કૃષિ યુનિવર્સિટી\nનવસારી- ૩૯૬ ૪૫૦ (ગુજરાત)")
    r2.bold = True
    r2.font.size = Pt(14)
    
    # Right Logo (ICAR)
    if 'ICAR_LOGO' in globals() and ICAR_LOGO and os.path.exists(ICAR_LOGO):
        p_right = table.cell(0, 2).paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_right = p_right.add_run()
        r_right.add_picture(ICAR_LOGO, width=Inches(1.5))
        
    # Draw First Thick Black Separator Line (Squashed height to remove gap)
    p_thick1 = doc.add_paragraph()
    p_thick1.paragraph_format.space_before = Pt(0)
    p_thick1.paragraph_format.space_after = Pt(0)
    p_thick1.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_thick1.paragraph_format.line_spacing = Pt(1)
    p_thick1.add_run().font.size = Pt(1) 
    add_bottom_border(p_thick1, size='24')
    
    # 2. Sender Info Block
    table2 = doc.add_table(rows=1, cols=2)
    table2.autofit = False
    table2.columns[0].width = Inches(3.4)
    table2.columns[1].width = Inches(3.4)
    
    p1 = table2.cell(0,0).paragraphs[0]
    p1.add_run("ડૉ. જે. જે. પસ્તાગિયા\nપ્રાધ્યાપક અને વડા (ઈ/ચા.)")
    
    p2 = table2.cell(0,1).paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run("મોબાઇલ: +૯૧ ૯૮૭૯૦ ૩૮૫૩૯\nઇમેલ: headentonau@gmail.com")
    
    # Ensure inner paragraphs of table2 have zero spacing
    for cell in table2.rows[0].cells:
        for p in cell.paragraphs:
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 0.8  # <--- This line squashes the \n gap
    
    # Draw Second Thick Black Separator Line (Squashed height to remove gap)
    p_thick2 = doc.add_paragraph()
    p_thick2.paragraph_format.space_before = Pt(0)
    p_thick2.paragraph_format.space_after = Pt(0)
    p_thick2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_thick2.paragraph_format.line_spacing = Pt(1)
    p_thick2.add_run().font.size = Pt(1) 
    add_bottom_border(p_thick2, size='24')
    
    # 3. Reference No & Date
    table3 = doc.add_table(rows=1, cols=2)
    table3.autofit = False
    table3.columns[0].width = Inches(4.5)
    table3.columns[1].width = Inches(2.3)
    
    p_ref = table3.cell(0,0).paragraphs[0]
    p_ref.add_run(f"જા.નં. એસીએન/એન્ટો/{ref_no}/૨૦૨૬, નવસારી")
    
    p_date = table3.cell(0,1).paragraphs[0]
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_date.add_run(f"તારીખ: {letter_date}")
    
    for cell in table3.rows[0].cells:
        for p in cell.paragraphs:
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            
    # Add a buffer space before the main letter body
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # 4. Recipient
    p_to = doc.add_paragraph()
    p_to.add_run("પ્રતિ,\n").bold = True
    p_to.add_run("હિસાબ નિયામકશ્રી\nનવસારી કૃષિ યુનિવર્સિટી\nનવસારી- ૩૯૬ ૪૫૦")
    p_to.paragraph_format.space_after = Pt(6)
    
    # 5. Through
    p_through = doc.add_paragraph()
    p_through.add_run("મારફત સવિનય: ").bold = True
    p_through.add_run("આચાર્ય અને ડિનશ્રી , ન. મ. કૃષિ મહાવિદ્યાલય, ન.કૃ.યુ., નવસારી ૩૯૬ ૪૫૦")
    p_through.paragraph_format.space_after = Pt(6)
    
    # 6. Subject
    p_sub = doc.add_paragraph()
    p_sub.add_run("વિષય:- ").bold = True
    p_sub.add_run("બ.સ. ૩૦૩/ ૨૦૯૨ અને ૩૦૩/ ૨૦૯૨/A માં ICAR – NCIPM તરફથી આવેલ ગ્રાન્ટ ફાળવવા બાબત...")
    p_sub.paragraph_format.space_after = Pt(12)
    
    # 7. Body Text (With Official Letter Indentation)
    p_body = doc.add_paragraph(body_text)
    p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_body.paragraph_format.first_line_indent = Inches(0.5) 
    p_body.paragraph_format.space_after = Pt(12)
    
    # 8. Finance Table
    table4 = doc.add_table(rows=2, cols=5)
    table4.style = 'Table Grid'
    headers = ["Name of Centre (Scheme)", "Pay And allowance", "Recurring Contingencies", "Non-Recurring Contingencies", "Total Amount"]
    
    for i, h in enumerate(headers):
        p = table4.cell(0, i).paragraphs[0]
        r_h = p.add_run(h)
        r_h.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    table4.cell(1, 0).text = "AINP on Agril Acarology\n(BH.303/2092)"
    table4.cell(1, 1).text = f"{int(pay_amt):,}/-" if pay_amt > 0 else "-"
    table4.cell(1, 2).text = f"{int(total_rec):,}/-" if total_rec > 0 else "-"
    table4.cell(1, 3).text = f"{int(non_rec_amt):,}/-" if non_rec_amt > 0 else "-"
    table4.cell(1, 4).text = f"{int(total_amt):,}/-"
    
    for i in range(5):
        table4.cell(1, i).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # 9. Footer Amount & Enclosure
    p_amt = doc.add_paragraph()
    p_amt.add_run(f"In Rupees: {amt_words}").bold = True
    p_amt.paragraph_format.space_after = Pt(12)
    
    p_enc = doc.add_paragraph("સામેલ: ઉપર મુજબ")
    p_enc.paragraph_format.space_after = Pt(24)
    
    # 10. Signatures
    table5 = doc.add_table(rows=1, cols=2)
    table5.autofit = False
    table5.columns[0].width = Inches(3.4)
    table5.columns[1].width = Inches(3.4)
    
    p_sig_left = table5.cell(0,0).paragraphs[0]
    p_sig_left.add_run("પ્રોજેક્ટ ઈન્ચાર્જ").bold = True
    
    p_sig_right = table5.cell(0,1).paragraphs[0]
    p_sig_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig_right.add_run("પ્રાધ્યાપક અને વડા").bold = True
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io
# ---------------------------------------------------------
# 👇 PASTE THIS NEW FUNCTION RIGHT HERE (Around Line 240) 👇
# ---------------------------------------------------------
def create_word_doc(dataframe, month_name, year_num, last_day):
    doc = Document()
    
    # Set narrow margins and change orientation to LANDSCAPE (Horizontal)
    sections = doc.sections
    for section in sections:
        section.orientation = WD_ORIENTATION.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
        
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
    
    # Headers - Dynamic Month and Year
    title = doc.add_paragraph()
    run_title = title.add_run(f"Statement of Expenditure for the month of {month_name} {year_num}")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title.bold = True
    run_title.font.size = Pt(12)
    
    doc.add_paragraph("Name of the Centre: Navsari").runs[0].bold = True
    doc.add_paragraph("Name of the Scheme: AICRP/AINP on Agricultural Acarology, NAU, Navsari").runs[0].bold = True
    
    # Create Table with 2 header rows
    table = doc.add_table(rows=2, cols=8)
    table.style = 'Table Grid'
    
    # --- ROW 0: Top Level Headers ---
    hdr0 = table.rows[0].cells
    hdr0[0].text = "Sr.\nNo."
    hdr0[1].text = "Head"
    hdr0[2].text = "Opening Balance\nas on 01.04.2025"
    hdr0[3].text = "Funds Received\nfrom the Council\nduring 2025-26"
    
    # Dynamic Headers
    hdr0[4].text = f"Expenditure up to\nthe month of {month_name}\n{year_num}"
    
    # Convert month name back to a digit for the DD.MM.YYYY format
    month_digit = list(calendar.month_name).index(month_name)
    hdr0[5].text = f"Cumulative Expenditure\nup to {last_day}.{month_digit:02d}.{year_num}"
    
    hdr0[7].text = "Total"
    
    # Merge "Cumulative Expenditure" across the two share columns
    hdr0[5].merge(hdr0[6])
    
    # --- ROW 1: Sub Headers ---
    hdr1 = table.rows[1].cells
    hdr1[5].text = "75%\nICAR Share"
    hdr1[6].text = "25%\nState Share"
    
    # Merge vertical columns for headers that span both rows
    for c in [0, 1, 2, 3, 4, 7]:
        table.cell(0, c).merge(table.cell(1, c))
    
    # Format all header cells
    for row_idx in [0, 1]:
        for cell in table.rows[row_idx].cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # --- Add Data Rows ---
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        is_yellow_header = row.iloc[1] in ["A. Recurring Contingencies", "B. Non Recurring Contingencies"]
        
        for i, cell_data in enumerate(row):
            text_val = str(cell_data) if pd.notna(cell_data) else ""
            row_cells[i].text = text_val
            
            if i == 1:
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
            row_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            if is_yellow_header or text_val in ["Total - A", "Total - B", "Grand Total A+B"]:
                row_cells[i].paragraphs[0].runs[0].bold = True
                
        # Apply yellow background shading
        if is_yellow_header:
            for cell in row_cells:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), 'FFFF00')
                tcPr.append(shd)
            
    # Add footer note
    doc.add_paragraph() 
    footer = doc.add_paragraph("In 2025-26 State share released only in Pay and allowances", style='List Bullet')
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
    # ---------------------------------------------------------
# 👇 PASTE THIS NEW FUNCTION RIGHT AFTER `create_word_doc`
# ---------------------------------------------------------
def create_yearly_word_doc(dataframe, fy_string):
    doc = Document()
    
    # Set narrow margins and change orientation to LANDSCAPE
    sections = doc.sections
    for section in sections:
        section.orientation = WD_ORIENTATION.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
        
        section.left_margin = Inches(0.4)
        section.right_margin = Inches(0.4)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
    
    # Headers
    title = doc.add_paragraph()
    run_title = title.add_run(f"12-Month Expenditure Summary for FY {fy_string}")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title.bold = True
    run_title.font.size = Pt(14)
    
    doc.add_paragraph("Name of the Centre: Navsari").runs[0].bold = True
    doc.add_paragraph("Name of the Scheme: AICRP/AINP on Agricultural Acarology, NAU, Navsari").runs[0].bold = True
    
    columns = dataframe.columns.tolist()
    
    # Create Table
    table = doc.add_table(rows=1, cols=len(columns))
    table.style = 'Table Grid'
    
    # --- Add Column Headers ---
    hdr_cells = table.rows[0].cells
    for i, column_name in enumerate(columns):
        hdr_cells[i].text = column_name
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # --- Add Data Rows ---
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        head_val = str(row.iloc[0])
        is_bold_row = head_val in ["A. Recurring Contingencies", "B. Non Recurring Contingencies", "Total - A", "Total - B", "Grand Total A+B"]
        
        for i, cell_data in enumerate(row):
            text_val = str(cell_data) if pd.notna(cell_data) else ""
            row_cells[i].text = text_val
            
            # Left align the Head column, center the numbers
            if i == 0:
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
            row_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            if is_bold_row:
                row_cells[i].paragraphs[0].runs[0].bold = True
                
        # Apply yellow background shading to Category Rows
        if head_val in ["A. Recurring Contingencies", "B. Non Recurring Contingencies"]:
            for cell in row_cells:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), 'FFFF00')
                tcPr.append(shd)
            
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
    # ---------------------------------------------------------
# 👇 PASTE THIS NEW FUNCTION RIGHT BELOW `create_word_doc` 👇
# ---------------------------------------------------------
def create_annual_word_doc(dataframe, fy_string):
    doc = Document()
    
    # Set narrow margins and change orientation to LANDSCAPE
    sections = doc.sections
    for section in sections:
        section.orientation = WD_ORIENTATION.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
        
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
    
    fy_start = fy_string.split('-')[0]
    fy_end = "20" + fy_string.split('-')[1]
    
    # Headers - Dynamic for the Full Year
    title = doc.add_paragraph()
    run_title = title.add_run(f"Statement of Expenditure for the Financial Year {fy_string}")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title.bold = True
    run_title.font.size = Pt(12)
    
    doc.add_paragraph("Name of the Centre: Navsari").runs[0].bold = True
    doc.add_paragraph("Name of the Scheme: AICRP/AINP on Agricultural Acarology, NAU, Navsari").runs[0].bold = True
    
    # Create Table with 2 header rows
    table = doc.add_table(rows=2, cols=8)
    table.style = 'Table Grid'
    
    # --- ROW 0: Top Level Headers ---
    hdr0 = table.rows[0].cells
    hdr0[0].text = "Sr.\nNo."
    hdr0[1].text = "Head"
    hdr0[2].text = f"Opening Balance\nas on 01.04.{fy_start}"
    hdr0[3].text = f"Funds Received\nfrom the Council\nduring {fy_string}"
    hdr0[4].text = f"Expenditure during\nthe FY {fy_string}"
    hdr0[5].text = f"Cumulative Expenditure\nup to 31.03.{fy_end}"
    hdr0[7].text = "Total"
    
    # Merge "Cumulative Expenditure" across the two share columns
    hdr0[5].merge(hdr0[6])
    
    # --- ROW 1: Sub Headers ---
    hdr1 = table.rows[1].cells
    hdr1[5].text = "75%\nICAR Share"
    hdr1[6].text = "25%\nState Share"
    
    # Merge vertical columns for headers that span both rows
    for c in [0, 1, 2, 3, 4, 7]:
        table.cell(0, c).merge(table.cell(1, c))
    
    # Format all header cells
    for row_idx in [0, 1]:
        for cell in table.rows[row_idx].cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # --- Add Data Rows ---
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        is_yellow_header = row.iloc[1] in ["A. Recurring Contingencies", "B. Non Recurring Contingencies"]
        
        for i, cell_data in enumerate(row):
            text_val = str(cell_data) if pd.notna(cell_data) else ""
            row_cells[i].text = text_val
            
            if i == 1:
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
            row_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            if is_yellow_header or text_val in ["Total - A", "Total - B", "Grand Total A+B"]:
                row_cells[i].paragraphs[0].runs[0].bold = True
                
        # Apply yellow background shading
        if is_yellow_header:
            for cell in row_cells:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), 'FFFF00')
                tcPr.append(shd)
            
    # Add footer note
    doc.add_paragraph() 
    footer = doc.add_paragraph("In 2025-26 State share released only in Pay and allowances", style='List Bullet')
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
    # ---------------------------------------------------------
# 👇 PASTE THESE NEW FUNCTIONS RIGHT BELOW `create_annual_word_doc` 👇
# ---------------------------------------------------------
def generate_auc_forwarding_docx(ref_no, letter_date, subject_text, body_text):
    """Generates the Forwarding Letter specifically for the AUC."""
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    style = doc.styles['Normal']
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)
        
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    table.columns[0].width = Inches(1.8)
    table.columns[1].width = Inches(3.6)
    table.columns[2].width = Inches(1.4)
    
    if 'NAU_LOGO' in globals() and NAU_LOGO and os.path.exists(NAU_LOGO):
        cell_left = table.cell(0, 0)
        cell_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_left = cell_left.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_left = p_left.add_run()
        r_left.add_picture(NAU_LOGO, width=Inches(1.8))
        
    p_center = table.cell(0, 1).paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_center.paragraph_format.line_spacing = 0.85
    r1 = p_center.add_run("કીટકશાસ્ત્ર વિભાગ\n")
    r1.bold = True
    r1.font.size = Pt(22)
    r2 = p_center.add_run("ન. મ. કૃષિ મહાવિદ્યાલય\nનવસારી કૃષિ યુનિવર્સિટી\nનવસારી- ૩૯૬ ૪૫૦ (ગુજરાત)")
    r2.bold = True
    r2.font.size = Pt(14)
    
    if 'ICAR_LOGO' in globals() and ICAR_LOGO and os.path.exists(ICAR_LOGO):
        p_right = table.cell(0, 2).paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_right = p_right.add_run()
        r_right.add_picture(ICAR_LOGO, width=Inches(1.5))
        
    p_thick1 = doc.add_paragraph()
    add_bottom_border(p_thick1, size='24')
    
    table2 = doc.add_table(rows=1, cols=2)
    p1 = table2.cell(0,0).paragraphs[0]
    p1.add_run("ડૉ. જે. જે. પસ્તાગિયા\nપ્રાધ્યાપક અને વડા (ઈ/ચા.)")
    p2 = table2.cell(0,1).paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run("મોબાઇલ: +૯૧ ૯૮૭૯૦ ૩૮૫૩૯\nઇમેલ: headentonau@gmail.com")
    
    p_thick2 = doc.add_paragraph()
    add_bottom_border(p_thick2, size='24')
    
    table3 = doc.add_table(rows=1, cols=2)
    p_ref = table3.cell(0,0).paragraphs[0]
    p_ref.add_run(f"જા.નં. એસીએન/એન્ટો/{ref_no}/૨૦૨૬, નવસારી")
    p_date = table3.cell(0,1).paragraphs[0]
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_date.add_run(f"તારીખ: {letter_date}")
            
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    p_to = doc.add_paragraph()
    p_to.add_run("પ્રતિ,\n").bold = True
    p_to.add_run("હિસાબ નિયામકશ્રી\nનવસારી કૃષિ યુનિવર્સિટી\nનવસારી- ૩૯૬ ૪૫૦")
    p_to.paragraph_format.space_after = Pt(6)
    
    p_through = doc.add_paragraph()
    p_through.add_run("મારફત સવિનય: ").bold = True
    p_through.add_run("આચાર્ય અને ડિનશ્રી , ન. મ. કૃષિ મહાવિદ્યાલય, ન.કૃ.યુ., નવસારી ૩૯૬ ૪૫૦")
    p_through.paragraph_format.space_after = Pt(6)
    
    p_sub = doc.add_paragraph()
    p_sub.add_run("વિષય:- ").bold = True
    p_sub.add_run(subject_text)
    p_sub.paragraph_format.space_after = Pt(12)
    
    p_body = doc.add_paragraph(body_text)
    p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_body.paragraph_format.first_line_indent = Inches(0.5) 
    p_body.paragraph_format.space_after = Pt(24)
    
    p_enc = doc.add_paragraph("સામેલ: ઉપર મુજબ (AUC)")
    p_enc.paragraph_format.space_after = Pt(36)
    
    table5 = doc.add_table(rows=1, cols=2)
    p_sig_left = table5.cell(0,0).paragraphs[0]
    p_sig_left.add_run("પ્રોજેક્ટ ઈન્ચાર્જ").bold = True
    p_sig_right = table5.cell(0,1).paragraphs[0]
    p_sig_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig_right.add_run("પ્રાધ્યાપક અને વડા").bold = True
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ---------------------------------------------------------
# 👇 REPLACE YOUR CURRENT `generate_auc_certificate` FUNCTION
# ---------------------------------------------------------
def generate_auc_certificate(inst_data, t1_data, t2_data, cert_text_1, fy_string):
    """Generates the official Audit Utilization Certificate."""
    doc = Document()
    
    # Set Margins
    for section in doc.sections:
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
    
    # Force Times New Roman Font globally
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11)
   
    # Certification Text
    p_cert1 = doc.add_paragraph()
    run_cert1 = p_cert1.add_run("Form of Utilization Certificate & Audit Utilization Certificate")
    run_cert1.bold = True
    run_cert1.font.size = Pt(12)
    run_cert1.font.name = 'Times New Roman'
    
    # Installment Table
    inst_table = doc.add_table(rows=1, cols=3)
    inst_table.style = 'Table Grid'
    hdr = inst_table.rows[0].cells
    hdr[0].text = "Sr.No"
    hdr[1].text = "Letter No and Date"
    hdr[2].text = "Amount"
    
    for row in inst_data:
        cells = inst_table.add_row().cells
        cells[0].text = str(row[0])
        cells[1].text = str(row[1])
        cells[2].text = str(row[2])
        
    for row in inst_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.style = doc.styles['Normal']
    
    doc.add_paragraph()
        
    # Inject the user-edited text paragraph here
    p_text = doc.add_paragraph(f"{cert_text_1}\n")
    p_text.add_run("2. Certified that I have satisfied myself that the condition on which the expenditure was made have dully fulfilled/are being fulfilled and that I have exercised the following check to see that the money was actually utilized for the purpose for which it was sanctioned.")
    
    p_t1_title = doc.add_paragraph("\nTable 1: Showing the details of receipt and expenditure figure (in Rupees)")
    p_t1_title.runs[0].bold = True
    
    # Table 1
    t1 = doc.add_table(rows=2, cols=5)
    t1.style = 'Table Grid'
    fy_start = fy_string.split('-')[0]
    fy_end = "20" + fy_string.split('-')[1]
    
    t1.cell(0,0).text = f"Opening balance as on 1st April {fy_start}"
    t1.cell(0,1).text = f"Remittance Received {fy_string}"
    t1.cell(0,2).text = "Receipt"
    t1.cell(0,3).text = f"ICAR share of Expenditure during the year {fy_string}"
    t1.cell(0,4).text = f"Closing balance as on 31st March {fy_end}"
    
    for i in range(5):
        t1.cell(1,i).text = str(t1_data[i])
        
    for row in t1.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.style = doc.styles['Normal']
        
    p_t2_title = doc.add_paragraph("\nTable 2: Showing the head wise details of expenditure figure (in Rupees)")
    p_t2_title.runs[0].bold = True
    
    # Table 2
    t2 = doc.add_table(rows=1, cols=5)
    t2.style = 'Table Grid'
    h2 = t2.rows[0].cells
    h2[0].text = "Head"
    h2[1].text = f"Allocation for the Year {fy_string} (100%)"
    h2[2].text = "ICAR share of Expenditure (75%)"
    h2[3].text = "State Share (25%)"
    h2[4].text = "Total Expenditure"
    
    for row in t2_data:
        cells = t2.add_row().cells
        for i in range(5):
            cells[i].text = str(row[i])
            cells[i].paragraphs[0].style = doc.styles['Normal']
            if str(row[0]) in ["Recurring", "Non Recurring Contingencies", "Total:-"]:
                cells[i].paragraphs[0].runs[0].bold = True

    doc.add_paragraph("\n\n")
    
    # Signatures
    sig_table = doc.add_table(rows=1, cols=4)
    sig_table.cell(0,0).text = "PI of the Scheme\nNAU, Navsari"
    sig_table.cell(0,1).text = "Director of Research\nNAU, Navsari"
    sig_table.cell(0,2).text = "Comptroller\nNAU, Navsari"
    sig_table.cell(0,3).text = "Duly audited and signed by the Chartered Accountant\n\nChartered Accountant"
    
    for row in sig_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.style = doc.styles['Normal']
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 3. THE UI APPLICATION ---

def main():
    st.title("🌾 AINP Grant Management System - NAU Navsari")
    
    current_year = datetime.now().year
    fy_options = [f"{y}-{str(y+1)[2:]}" for y in range(current_year-2, current_year+2)]
    selected_fy = st.sidebar.selectbox("Select Financial Year (Apr 1 - Mar 31)", fy_options, index=2)
    
    data = load_data(selected_fy)
    
    tabs = st.tabs([
        "📊 Dashboard", 
        "📤 1. Budget Intake", 
        "💰 2. Installments (PFMS)", 
        "📝 3. Generated Letters", 
        "💸 4. Monthly Spend", 
        "📄 5. SOE Generation", 
        "📜 6. AUC Generation",  # <-- NEW TAB ADDED
        "🤖 AI Chatbot"
    ])
    
    # --- TAB 1: DASHBOARD ---
    with tabs[0]:
        st.header(f"Financial Year {selected_fy}")
        
        current_qtr = data.get('latest_quarter', 'Full Year')
        last_date = data.get('latest_date', 'N/A')
        st.markdown(f"**Current Active Budget Period:** {current_qtr} | **Last Document Date:** {last_date}")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            total_allocated = sum(v.get('total', 0) for k,v in data['revised_allocation'].items()) if data['revised_allocation'] else (sum(v.get('total', 0) for k,v in data['allocation'].items()) if data['allocation'] else 0)
            st.metric("Total Sanctioned Budget", f"₹{total_allocated:,}")
        with col2:
            total_received = sum(inst['amount'] for inst in data['installments'])
            st.metric("Total Funds Received (PFMS)", f"₹{total_received:,}")
        with col3:
            st.metric("Pending Received", f"₹{total_allocated - total_received:,}")

    # --- TAB 2: BUDGET INTAKE ---
    with tabs[1]:
        st.header("Upload Budget Allocation / Revision PDF")
        
        st.markdown("""
        **Instructions for Uploading:**
        Please describe the document in the text box below (e.g., 'Initial Allocation', 'Revised Allocation', 'Q1 Release', 'Q3 Release') and upload the corresponding PDF.
        """)
        
        doc_type_input = st.text_input("Describe the document you are uploading:")
        budget_file = st.file_uploader("Upload Budget Document", type=['pdf'], key="budget_up")
        
        if budget_file and st.button("Analyze & Process Budget"):
            if not doc_type_input.strip():
                st.warning("⚠️ Please provide a description in the text box above before processing.")
            else:
                with st.spinner(f"AI is analyzing the PDF..."):
                    
                    budget_prompt = {
                        "context": f"Document Description: {doc_type_input}. Scheme: AICRP/AINP Acarology. Financial Year: {selected_fy}. Extract the budget allocation table. IMPORTANT: Map the head names exactly to: 'Pay and Allowances', 'Travelling Allowances (TA)', 'Other Recurring Contingencies (ORC)', 'Non-Recurring Contingencies (Equipments/Works)', 'TSP'. Convert all Lakh values to absolute Rupees (e.g., 1.33 Lakhs becomes 133000, 40 Lakhs becomes 4000000).",
                        "structure": {
                            "is_revision": "boolean (true if this revises a previous allocation)",
                            "date": "YYYY-MM-DD",
                            "heads": [
                                {"head_name": "string", "icar_share": 0.0, "state_share": 0.0, "total": 0.0}
                            ]
                        }
                    }
                    
                    extracted_budget = process_upload_with_ai(budget_file, budget_prompt)
                    
                    if extracted_budget:
                        st.success("Document analyzed successfully!")
                        
                        # Force creation of a dictionary with ALL 5 standard heads (defaulted to 0)
                        budget_dict = {h: {'icar': 0.0, 'state': 0.0, 'total': 0.0} for h in BUDGET_HEADS}
                        
                        # --- STRICT 75:25 MATHEMATICAL ENFORCEMENT (WITH 100% TSP EXCEPTION) ---
                        for head in extracted_budget.get('heads', []):
                            extracted_total = float(head.get('total') or 0.0)
                            extracted_icar = float(head.get('icar_share') or 0.0)
                            raw_name = head.get('head_name', '').upper()
                            
                            # Match the AI's extracted name to the standard BUDGET_HEADS
                            matched_head = None
                            if "PAY" in raw_name: matched_head = BUDGET_HEADS[0]
                            elif "TRAV" in raw_name or "TA" in raw_name: matched_head = BUDGET_HEADS[1]
                            elif "NON" in raw_name or "EQUIP" in raw_name or "WORK" in raw_name: matched_head = BUDGET_HEADS[3]
                            elif "RECURR" in raw_name or "ORC" in raw_name: matched_head = BUDGET_HEADS[2]
                            elif "TSP" in raw_name: matched_head = BUDGET_HEADS[4]
                            
                            if matched_head:
                                # Exception: TSP is always 100% ICAR share
                                if "TSP" in matched_head.upper():
                                    final_total = extracted_total if extracted_total > 0 else extracted_icar
                                    budget_dict[matched_head]['total'] = round(final_total, 2)
                                    budget_dict[matched_head]['icar'] = round(final_total, 2)
                                    budget_dict[matched_head]['state'] = 0.0
                                else:
                                    # Standard 75:25 Split
                                    if extracted_total == 0 and extracted_icar > 0:
                                        final_total = extracted_icar / 0.75
                                    else:
                                        final_total = extracted_total

                                    budget_dict[matched_head]['total'] = round(final_total, 2)
                                    budget_dict[matched_head]['icar'] = round(final_total * 0.75, 2)
                                    budget_dict[matched_head]['state'] = round(final_total * 0.25, 2)
                        
                        doc_date = extracted_budget.get('date', 'Unknown')
                        desc_lower = doc_type_input.lower()
                        assigned_type = "Unknown"
                        dict_key = ""
                        
                        if "revis" in desc_lower:
                            data['revised_allocation'] = budget_dict
                            assigned_type = "Revised Allocation"
                            dict_key = "revised_allocation"
                        elif "q1" in desc_lower or "1st" in desc_lower:
                            data['quarterly_allocations']["Q1"] = budget_dict
                            assigned_type = "Q1 Release"
                            dict_key = "Q1"
                        elif "q2" in desc_lower or "2nd" in desc_lower:
                            data['quarterly_allocations']["Q2"] = budget_dict
                            assigned_type = "Q2 Release"
                            dict_key = "Q2"
                        elif "q3" in desc_lower or "3rd" in desc_lower:
                            data['quarterly_allocations']["Q3"] = budget_dict
                            assigned_type = "Q3 Release"
                            dict_key = "Q3"
                        elif "q4" in desc_lower or "4th" in desc_lower:
                            data['quarterly_allocations']["Q4"] = budget_dict
                            assigned_type = "Q4 Release"
                            dict_key = "Q4"
                        else:
                            data['allocation'] = budget_dict
                            assigned_type = "Initial Allocation"
                            dict_key = "allocation"
                            
                        # Save the PDF locally for later download
                        pdf_path = f"documents/{selected_fy}_{dict_key}.pdf"
                        with open(pdf_path, "wb") as f:
                            f.write(budget_file.getvalue())
                            
                        st.info(f"📅 **Document Date:** {doc_date} | 🔄 **Detected Type:** {assigned_type}")
                        
                        data['latest_quarter'] = assigned_type
                        data['latest_date'] = doc_date
                        
                        save_data(data, selected_fy)
                        st.toast("Budget Data Saved to GitHub & PDF saved locally!")

        # --- BUDGET VISUALIZATION & EDITING ---
        st.divider()
        st.subheader("📑 Uploaded Budget Documents (Editable)")
        
        # Interactive budget editor function
        def interactive_budget_editor(b_dict, dict_key, doc_type_name):
            df_data = []
            for h in BUDGET_HEADS:
                if h in b_dict:
                    df_data.append({
                        "Budget Head": h,
                        "ICAR Share (₹)": b_dict[h]['icar'],
                        "State Share (₹)": b_dict[h]['state'],
                        "Total (₹)": b_dict[h]['total']
                    })
            
            if not df_data:
                return

            df = pd.DataFrame(df_data)
            
            # Display interactive Data Editor
            edited_df = st.data_editor(df, use_container_width=True, hide_index=True, key=f"editor_{dict_key}")
            
            # Read-only totals for display below the editor
            tot_icar = edited_df["ICAR Share (₹)"].astype(float).sum()
            tot_state = edited_df["State Share (₹)"].astype(float).sum()
            tot_all = edited_df["Total (₹)"].astype(float).sum()
            st.markdown(f"**Calculated Totals ➡️ ICAR:** ₹{tot_icar:,.2f} | **State:** ₹{tot_state:,.2f} | **Grand Total:** ₹{tot_all:,.2f}")
            
            col1, col2 = st.columns([1, 3])
            with col1:
                if st.button(f"💾 Save & Recalculate", key=f"save_{dict_key}"):
                    new_dict = {}
                    for index, row in edited_df.iterrows():
                        head = row["Budget Head"]
                        raw_total = float(row["Total (₹)"])
                        raw_icar = float(row["ICAR Share (₹)"])
                        
                        if "TSP" in head.upper():
                            final_total = raw_total if raw_total > 0 else raw_icar
                            new_dict[head] = {'icar': final_total, 'state': 0.0, 'total': final_total}
                        else:
                            if raw_total == 0 and raw_icar > 0:
                                final_total = raw_icar / 0.75
                            else:
                                final_total = raw_total

                            new_dict[head] = {
                                'icar': round(final_total * 0.75, 2),
                                'state': round(final_total * 0.25, 2),
                                'total': round(final_total, 2)
                            }
                    
                    if dict_key == 'allocation':
                        data['allocation'] = new_dict
                    elif dict_key == 'revised_allocation':
                        data['revised_allocation'] = new_dict
                    else:
                        data['quarterly_allocations'][dict_key] = new_dict
                        
                    save_data(data, selected_fy)
                    st.success("Changes Saved & Recalculated!")
                    st.rerun()

            with col2:
                pdf_path = f"documents/{selected_fy}_{dict_key}.pdf"
                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as pdf_file:
                        st.download_button(label=f"📥 Download Uploaded PDF", 
                                           data=pdf_file, 
                                           file_name=f"{doc_type_name}_{selected_fy}.pdf", 
                                           mime="application/pdf",
                                           key=f"dl_{dict_key}")

        col_a, col_b = st.columns(2)
        with col_a:
            if data.get('allocation'):
                with st.expander("📄 Initial Full-Year Budget Allocation"):
                    interactive_budget_editor(data['allocation'], 'allocation', "Initial_Allocation")
        with col_b:
            if data.get('revised_allocation'):
                with st.expander("📄 Revised Full-Year Budget Allocation"):
                    interactive_budget_editor(data['revised_allocation'], 'revised_allocation', "Revised_Allocation")
                    
        st.write("📅 **Quarterly Releases (Click to edit data)**")
        q_cols = st.columns(4)
        for i, q in enumerate(["Q1", "Q2", "Q3", "Q4"]):
            q_data = data['quarterly_allocations'].get(q)
            if q_data:
                with q_cols[i]:
                    with st.expander(f"📄 {q} Release Data"):
                        interactive_budget_editor(q_data, q, f"{q}_Release")
                        
        st.divider()
        
        # --- Total Mismatch/Status Table ---
        active_budget = data.get('revised_allocation') or data.get('allocation')
        if active_budget:
            st.subheader("⚖️ Allocation vs. Quarterly Release Mismatch")
            mismatch_data = {}
            
            # Force standard order
            for head in BUDGET_HEADS:
                vals = active_budget.get(head, {})
                tot_alloc = vals.get('total', 0)
                
                # Sum the quarterly totals for this specific head
                q_sum = 0
                for q in ["Q1", "Q2", "Q3", "Q4"]:
                    if data['quarterly_allocations'].get(q):
                        q_sum += data['quarterly_allocations'][q].get(head, {}).get('total', 0)
                        
                mismatch = tot_alloc - q_sum
                
                # Only show rows if there is allocated budget OR if money was mysteriously released for it
                if tot_alloc > 0 or q_sum > 0:
                    mismatch_data[head] = {
                        "Total Sanctioned (₹)": round(tot_alloc, 2),
                        "Released Q1-Q4 (₹)": round(q_sum, 2),
                        "Pending/Mismatch (₹)": round(mismatch, 2)
                    }
            
            if mismatch_data:
                df_mismatch = pd.DataFrame.from_dict(mismatch_data, orient='index')
                df_mismatch.loc['TOTAL'] = df_mismatch.sum(numeric_only=True)
                st.dataframe(df_mismatch, use_container_width=True)
            else:
                st.info("No allocated budget values to show yet.")
        else:
            st.info("Upload budget allocation to view Mismatch Table.")

    # --- TAB 3: INSTALLMENTS ---
    with tabs[2]:
        st.header("Received Grant Installment (PFMS & Email)")
        st.write("Upload the official release email and the PFMS attachment to log funds.")
        
        col1, col2 = st.columns(2)
        with col1:
            email_file = st.file_uploader("1. Upload Email PDF", type=['pdf'], key="email_up")
        with col2:
            pfms_file = st.file_uploader("2. Upload PFMS Receipt PDF", type=['pdf'], key="pfms_up")

        if email_file and pfms_file and st.button("Analyze & Process Installment"):
            with st.spinner("Analyzing Email and PFMS documents..."):
                
                # Package BOTH PDFs as native files for Gemini
                pdf_data_email = {"mime_type": "application/pdf", "data": email_file.getvalue()}
                pdf_data_pfms = {"mime_type": "application/pdf", "data": pfms_file.getvalue()}
                
                full_prompt = f"""
                Analyze the provided Email and PFMS documents and extract the installment information.
                Map the monetary values exactly to these heads: 'Pay and Allowances', 'Travelling Allowances (TA)', 'Other Recurring Contingencies (ORC)', 'Non-Recurring Contingencies (Equipments/Works)', 'TSP'.
                
                REQUIRED JSON STRUCTURE:
                {{
                    "date": "YYYY-MM-DD",
                    "installment_number": "String (e.g., I, II, III, IV, V, etc.)",
                    "purpose": "String (e.g., GIA-Salary, GIA-General, GIA-TSP, GIA-General & Capital)",
                    "pfms_transaction_id": "String (e.g., C022...)",
                    "heads": [
                        {{"head_name": "string", "amount": 0.0}}
                    ]
                }}
                IMPORTANT: Return ONLY valid JSON.
                """
                
                try:
                    # Pass the prompt and BOTH files directly to the model
                    response = model_pro.generate_content([full_prompt, pdf_data_email, pdf_data_pfms])
                    json_str = response.text.replace('```json', '').replace('```', '').strip()
                    extracted_inst = json.loads(json_str)
                    st.session_state['pending_installment'] = extracted_inst
                    st.success("Documents analyzed successfully!")
                except Exception as e:
                    st.error(f"Error processing documents with AI: {e}")

        # If we have a pending installment in session state, show the editor
        if 'pending_installment' in st.session_state:
            extracted = st.session_state['pending_installment']
            
            # Determine Quarter based on strict rules
            try:
                dt_obj = datetime.strptime(extracted['date'], "%Y-%m-%d")
                m = dt_obj.month
                if m in [4, 5, 6]: q_str = "Q1"
                elif m in [7, 8, 9]: q_str = "Q2"
                elif m in [10, 11, 12]: q_str = "Q3"
                else: q_str = "Q4"
            except:
                q_str = "Unknown Quarter"
                
            st.info(f"📅 **Date:** {extracted.get('date')} | 🕒 **Quarter:** {q_str} | 🔢 **Installment:** {extracted.get('installment_number')} | 🎯 **Purpose:** {extracted.get('purpose')} | 🆔 **PFMS ID:** {extracted.get('pfms_transaction_id')}")
            
            # Prepare dataframe for editing
            inst_dict = {h: 0.0 for h in BUDGET_HEADS}
            for h in extracted.get('heads', []):
                raw_name = h.get('head_name', '').upper()
                amt = float(h.get('amount') or 0.0)
                matched_head = None
                if "PAY" in raw_name or "SALARY" in raw_name: matched_head = BUDGET_HEADS[0]
                elif "TRAV" in raw_name or "TA" in raw_name: matched_head = BUDGET_HEADS[1]
                elif "NON" in raw_name or "EQUIP" in raw_name or "WORK" in raw_name or "CAPITAL" in raw_name: matched_head = BUDGET_HEADS[3]
                elif "RECURR" in raw_name or "ORC" in raw_name or "GENERAL" in raw_name: matched_head = BUDGET_HEADS[2]
                elif "TSP" in raw_name: matched_head = BUDGET_HEADS[4]
                
                if matched_head:
                    inst_dict[matched_head] += amt
                    
            df_data = [{"Budget Head": k, "Amount (₹)": v} for k, v in inst_dict.items()]
            df = pd.DataFrame(df_data)
            
            st.write("✏️ **Review and Edit Installment Amounts:**")
            edited_df = st.data_editor(df, use_container_width=True, hide_index=True, key="inst_editor")
            
            total_amt = edited_df["Amount (₹)"].astype(float).sum()
            st.markdown(f"**Total Installment Amount:** ₹{total_amt:,.2f}")
            
            if st.button("💾 Save Installment"):
                final_heads = {row["Budget Head"]: float(row["Amount (₹)"]) for _, row in edited_df.iterrows()}
                
                new_inst = {
                    "date": extracted.get('date'),
                    "quarter": q_str,
                    "installment_num": extracted.get('installment_number'),
                    "purpose": extracted.get('purpose'),
                    "pfms_id": extracted.get('pfms_transaction_id'),
                    "amount": total_amt, 
                    "heads": final_heads,
                    "type": extracted.get('installment_number', 'I'), 
                    "available": False
                }
                
                if not any(inst.get('pfms_id') == new_inst['pfms_id'] for inst in data['installments']):
                    data['installments'].append(new_inst)
                    save_data(data, selected_fy)
                    
                    # Save PDFs locally
                    email_path = f"documents/{selected_fy}_Inst_{new_inst['pfms_id']}_Email.pdf"
                    pfms_path = f"documents/{selected_fy}_Inst_{new_inst['pfms_id']}_PFMS.pdf"
                    with open(email_path, "wb") as f: f.write(email_file.getvalue())
                    with open(pfms_path, "wb") as f: f.write(pfms_file.getvalue())
                    
                    st.toast("Installment Saved and Backed up!")
                    del st.session_state['pending_installment']
                    st.rerun()
                else:
                    st.warning("An installment with this PFMS ID already exists.")
                    
        # Display existing installments
        st.divider()
        st.subheader("📁 Saved Installments (Editable)")
        if data['installments']:
            # Sort by Date
            data['installments'] = sorted(data['installments'], key=lambda x: x.get('date', '2000-01-01'))
            
            for q in ["Q1", "Q2", "Q3", "Q4"]:
                q_insts = [i for i in data['installments'] if i.get('quarter') == q]
                if q_insts:
                    with st.expander(f"📦 {q} Installments ({len(q_insts)})"):
                        for inst in q_insts:
                            st.markdown(f"**Date:** {inst.get('date')} | **Inst No:** {inst.get('installment_num')} | **Purpose:** {inst.get('purpose')} | **PFMS ID:** {inst.get('pfms_id')}")
                            
                            col_a, col_b = st.columns([3, 1])
                            with col_a:
                                # Provide an editable table for the saved installment
                                inst_heads = inst.get('heads', {h: 0.0 for h in BUDGET_HEADS})
                                df_data = [{"Budget Head": k, "Amount (₹)": v} for k, v in inst_heads.items()]
                                df_saved = pd.DataFrame(df_data)
                                
                                edited_saved_df = st.data_editor(df_saved, use_container_width=True, hide_index=True, key=f"edit_saved_{inst['pfms_id']}")
                                tot_amt_saved = edited_saved_df["Amount (₹)"].astype(float).sum()
                                st.markdown(f"**Total Amount:** ₹{tot_amt_saved:,.2f}")
                                
                                if st.button("💾 Save Changes", key=f"save_btn_{inst['pfms_id']}"):
                                    final_heads_saved = {row["Budget Head"]: float(row["Amount (₹)"]) for _, row in edited_saved_df.iterrows()}
                                    for main_inst in data['installments']:
                                        if main_inst['pfms_id'] == inst['pfms_id']:
                                            main_inst['heads'] = final_heads_saved
                                            main_inst['amount'] = tot_amt_saved
                                            break
                                    save_data(data, selected_fy)
                                    st.toast("Installment Updated!")
                                    st.rerun()
                                    
                            with col_b:
                                email_path = f"documents/{selected_fy}_Inst_{inst['pfms_id']}_Email.pdf"
                                pfms_path = f"documents/{selected_fy}_Inst_{inst['pfms_id']}_PFMS.pdf"
                                if os.path.exists(email_path):
                                    with open(email_path, "rb") as f:
                                        st.download_button("📥 Download Email PDF", f, file_name=f"{inst['pfms_id']}_Email.pdf", key=f"dl_e_{inst['pfms_id']}")
                                if os.path.exists(pfms_path):
                                    with open(pfms_path, "rb") as f:
                                        st.download_button("📥 Download PFMS PDF", f, file_name=f"{inst['pfms_id']}_PFMS.pdf", key=f"dl_p_{inst['pfms_id']}")
                            st.divider()
            
            # --- SUMMARY TABLE ---
            st.divider()
            st.subheader("📊 Summary of Received Installments (Q1 - Q4)")
            
            summary_data = {h: {"Q1 (₹)": 0.0, "Q2 (₹)": 0.0, "Q3 (₹)": 0.0, "Q4 (₹)": 0.0, "Total (₹)": 0.0} for h in BUDGET_HEADS}
            
            for inst in data['installments']:
                q = inst.get('quarter')
                if q in ["Q1", "Q2", "Q3", "Q4"]:
                    q_key = f"{q} (₹)"
                    for head, amt in inst.get('heads', {}).items():
                        if head in summary_data:
                            summary_data[head][q_key] += float(amt)
                            summary_data[head]["Total (₹)"] += float(amt)
                            
            df_summary = pd.DataFrame.from_dict(summary_data, orient='index')
            df_summary.loc['GRAND TOTAL'] = df_summary.sum(numeric_only=True)
            st.dataframe(df_summary, use_container_width=True)

        else:
            st.info("No installments recorded yet.")

    # --- TAB 4: GENERATED LETTERS ---
    with tabs[3]:
        st.header("Draft Letters based on PFMS Receipts")
        st.write("Generate the Gujarati letter (Word Document) to send to Comptroller.")
        
        pending_utilization = [inst for inst in data['installments'] if not inst.get('utilization_letter_generated')]
        
        if pending_utilization:
            options = {f"{inst['type']} (₹{inst['amount']:,} - {inst['pfms_id']})": inst for inst in pending_utilization}
            selected_inst_str = st.selectbox("Select PFMS Receipt to draft letter for:", list(options.keys()))
            selected_inst_data = options[selected_inst_str]

            # Extract heads for the template
            inst_heads = selected_inst_data.get('heads', {})
            pay_amt = inst_heads.get('Pay and Allowances', 0)
            rec_amt = inst_heads.get('Other Recurring Contingencies (ORC)', 0)
            ta_amt = inst_heads.get('Travelling Allowances (TA)', 0)
            total_rec = rec_amt + ta_amt
            non_rec_amt = inst_heads.get('Non-Recurring Contingencies (Equipments/Works)', 0)
            total_amt = selected_inst_data.get('amount', 0)

            # Editor Fields
            st.subheader("✏️ Edit Letter Details")
            col_a, col_b = st.columns(2)
            with col_a:
                ref_no = st.text_input("Reference No. (જા.નં. એસીએન/એન્ટો/___/૨૦૨૬):", value="")
                letter_date = st.text_input("Date (તારીખ):", value=datetime.now().strftime("%d/%m/%Y"))
            with col_b:
                amt_words = st.text_input("Amount in Words (In Rupees):", value="One lakh rupee only")
                
            body_text = st.text_area("Body Text:", value="જય ભારત સહ ઉપરોક્ત વિષય અન્વયે જણાવવાનું કે, અત્રેના કીટકશાસ્ત્ર વિભાગ ખાતે ચાલતી આઈ.સી.એ.આર. યોજના AINP on Agricultural Acarology (75:25%) (BH.303/2092) માં આવેલ ગ્રાન્ટને કોષ્ટકમાં જણાવ્યાનુસાર ફાળવી આપવા આપ સાહેબશ્રીને નમ્ર વિનંતી.", height=100)

            st.divider()
            st.subheader("👀 Letter Preview")

            # Inline CSS matching the requested exact layout
            st.markdown("""
            <style>
                .block-container { padding-top: 2rem; padding-bottom: 2rem; background-color: #ffffff; }
                .letter-body { font-family: 'Arial', sans-serif; font-size: 16px; color: #000000; line-height: 1.5; }
                .bold { font-weight: bold; }
                .center { text-align: center; }
                .right { text-align: right; }
                .justify { text-align: justify; }
                .indent { text-indent: 50px; }
                .table-custom { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 15px; }
                .table-custom th, .table-custom td { border: 1px solid black; padding: 8px; text-align: center; color: black; }
            </style>
            """, unsafe_allow_html=True)

            col1, col2, col3 = st.columns([1.2, 3, 1.2])
            with col1:
                if NAU_LOGO and os.path.exists(NAU_LOGO):
                    st.image(NAU_LOGO, use_container_width=True)
            with col2:
                st.markdown("""
                <div class="letter-body center">
                    <span class="bold">કીટકશાસ્ત્ર વિભાગ</span><br>
                    ન. મ. કૃષિ મહાવિદ્યાલય<br>
                    નવસારી કૃષિ યુનિવર્સિટી<br>
                    નવસારી- ૩૯૬ ૪૫૦ (ગુજરાત)
                </div>
                """, unsafe_allow_html=True)
            with col3:
                if ICAR_LOGO and os.path.exists(ICAR_LOGO):
                    st.image(ICAR_LOGO, use_container_width=True)

            st.markdown("<hr style='border: 1px solid black; margin-top: 10px; margin-bottom: 15px;' />", unsafe_allow_html=True)
            
            # Safely format numbers with commas
            pay_val = f"{int(pay_amt):,}/-" if pay_amt > 0 else "-"
            rec_val = f"{int(total_rec):,}/-" if total_rec > 0 else "-"
            non_rec_val = f"{int(non_rec_amt):,}/-" if non_rec_amt > 0 else "-"
            tot_val = f"{int(total_amt):,}/-"

            st.markdown(f"""
            <div class="letter-body">
                <div style="display: flex; justify-content: space-between;">
                    <div><span class="bold">ડૉ. જે. જે. પસ્તાગિયા</span><br>પ્રાધ્યાપક અને વડા (ઈ/ચા.)</div>
                    <div class="right">મોબાઇલ: +૯૧ ૯૮૭૯૦ ૩૮૫૩૯<br>ઇમેલ: headentonau@gmail.com</div>
                </div>
                <br>
                <div style="display: flex; justify-content: space-between;">
                    <div>જા.નં. એસીએન/એન્ટો/{ref_no}/૨૦૨૬, નવસારી</div>
                    <div class="right">તારીખ: {letter_date}</div>
                </div>
                <br>
                <div><span class="bold">પ્રતિ,</span><br>હિસાબ નિયામકશ્રી<br>નવસારી કૃષિ યુનિવર્સિટી<br>નવસારી- ૩૯૬ ૪૫૦</div>
                <br>
                <div><span class="bold">મારફત સવિનય:</span> આચાર્ય અને ડિનશ્રી , ન. મ. કૃષિ મહાવિદ્યાલય, ન.કૃ.યુ., નવસારી ૩૯૬ ૪૫૦</div>
                <br>
                <div><span class="bold">વિષય:-</span> બ.સ. ૩૦૩/ ૨૦૯૨ અને ૩૦૩/ ૨૦૯૨/A માં ICAR – NCIPM તરફથી આવેલ ગ્રાન્ટ ફાળવવા બાબત...</div>
                <br>
                <div class="justify indent">{body_text.replace(chr(10), '<br>')}</div>
                <table class="table-custom">
                    <tr>
                        <th style="text-align: left;">Name of Centre (Scheme)</th>
                        <th>Pay And allowance</th>
                        <th>Recurring Contingencies</th>
                        <th>Non-Recurring Contingencies</th>
                        <th>Total Amount</th>
                    </tr>
                    <tr>
                        <td style="text-align: left;">AINP on Agril Acarology<br>(BH.303/2092)</td>
                        <td>{pay_val}</td>
                        <td>{rec_val}</td>
                        <td>{non_rec_val}</td>
                        <td>{tot_val}</td>
                    </tr>
                </table>
                <div>In Rupees: {amt_words}</div>
                <br><br>
                <div>સામેલ: ઉપર મુજબ</div>
                <br><br><br>
                <div style="display: flex; justify-content: space-between;">
                    <div class="bold">પ્રોજેક્ટ ઈન્ચાર્જ</div>
                    <div class="bold right">પ્રાધ્યાપક અને વડા</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            st.divider()
            is_approved = st.checkbox("✅ I approve this letter format and content.")
            
            if is_approved:
                # Generate native DOCX Word file based on live edits
                doc_io = generate_comptroller_docx(ref_no, letter_date, body_text, amt_words, pay_amt, total_rec, non_rec_amt, total_amt)
                
                st.download_button(
                    label="📥 Download Approved Letter (.docx)",
                    data=doc_io,
                    file_name=f"Letter_to_Comptroller_{selected_inst_data['type']}_{selected_inst_data['pfms_id']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.info("No pending PFMS receipts to generate letters for.")

        st.divider()
        st.subheader("Activate Funds (Upload Comptroller Order)")
        st.write("Once received, upload the *'compotrollar grant relased latter'* to activate funds for utilization.")
        
        comp_file = st.file_uploader("Upload Comptroller Office Order PDF", type=['pdf'], key="comp_up")
        
        inst_to_activate = st.selectbox("This order relates to installment:", [inst['type'] for inst in data['installments']], key="act_type")

        if comp_file and st.button("Verify and Activate Funds"):
            with st.spinner("Activating..."):
                for inst in data['installments']:
                    if inst['type'] == inst_to_activate:
                        inst['available'] = True
                        inst['comptroller_order_uploaded'] = True
                save_data(data, selected_fy)
                st.success(f"Funds for Installment {inst_to_activate} are now READY FOR UTILIZATION.")


# --- TAB 5: MONTHLY SPEND ---
    with tabs[4]:
        st.header("Monthly Expenditure Tracking")
        
        today = date.today()
        month_to_process = st.selectbox("Month", [datetime(2000, m, 1).strftime('%B') for m in range(1, 13)], index=today.month-1)
        year_to_process = st.number_input("Year", value=today.year, min_value=2024, max_value=2030)

        with st.expander("Add New Expenditure Entry", expanded=True):
            with st.form("spend_form", clear_on_submit=True):
                col1, col2 = st.columns(2)
                exp_date = col1.date_input("Expenditure Date")
                exp_head = col1.selectbox("Budget Head", BUDGET_HEADS)
                exp_amt = col2.number_input("Amount Spent (₹)", min_value=0.0)
                exp_detail = col2.text_area("Expenditure Details/Voucher Info")
                
                if st.form_submit_button("Add Entry"):
                    new_exp = {
                        "date": exp_date.strftime("%Y-%m-%d"),
                        "head": exp_head,
                        "detail": exp_detail,
                        "amount": exp_amt
                    }
                    data['expenditure'].append(new_exp)
                    save_data(data, selected_fy)
                    st.toast("Expenditure Added and Backed up.")

        # Re-initialize df_exp here so both this tab and the chatbot can see it
        df_exp = pd.DataFrame(data.get('expenditure', []))
        
        if not df_exp.empty:
            df_exp['date'] = pd.to_datetime(df_exp['date'])
            current_month_exp = df_exp[
                (df_exp['date'].dt.strftime('%B') == month_to_process) & 
                (df_exp['date'].dt.year == year_to_process)
            ]
            st.subheader(f"Spend list for {month_to_process} {year_to_process}")
            st.dataframe(current_month_exp, use_container_width=True)
            
            st.divider()
            
# --- SUMMARY TABLES SECTION ---
            col_sum1, col_sum2 = st.columns(2)
            
            # 1. Monthly Summary Table
            with col_sum1:
                st.markdown(f"**📊 Monthly Summary ({month_to_process} {year_to_process})**")
                
                # Initialize dictionary with all standard heads set to 0
                monthly_summary = {head: 0.0 for head in BUDGET_HEADS}
                
                # If there are expenses this month, update the dictionary
                if not current_month_exp.empty:
                    month_grouped = current_month_exp.groupby('head')['amount'].sum().to_dict()
                    monthly_summary.update(month_grouped)
                
                # Convert to DataFrame for clean display
                df_month_sum = pd.DataFrame(list(monthly_summary.items()), columns=['Budget Head', 'Amount (₹)'])
                
                # Add "Grand Total" Row at the bottom for the Month
                total_month_amount = df_month_sum['Amount (₹)'].sum()
                df_month_sum.loc[len(df_month_sum)] = ['**GRAND TOTAL**', total_month_amount]
                
                st.dataframe(df_month_sum, use_container_width=True, hide_index=True)

            # 2. Yearly (FY) Summary Table with Grand Total
            with col_sum2:
                st.markdown(f"**📈 Yearly Summary (FY {selected_fy})**")
                
                # Initialize dictionary with all standard heads set to 0
                yearly_summary = {head: 0.0 for head in BUDGET_HEADS}
                
                # Update with actual yearly sums
                year_grouped = df_exp.groupby('head')['amount'].sum().to_dict()
                yearly_summary.update(year_grouped)
                
                df_year_sum = pd.DataFrame(list(yearly_summary.items()), columns=['Budget Head', 'Amount (₹)'])
                
                # Add "Grand Total" Row at the bottom for the Year
                total_amount = df_year_sum['Amount (₹)'].sum()
                df_year_sum.loc[len(df_year_sum)] = ['**GRAND TOTAL**', total_amount]
                
                st.dataframe(df_year_sum, use_container_width=True, hide_index=True)


# --- TAB 6: SOE GENERATION ---
    with tabs[5]:
        st.header("Statement of Expenditure (SOE) Generation")
        
        # 1. Custom function to format numbers as Indian Rupees (e.g., 24,74,691.00)
        def format_inr(number):
            if number == 0: return "-"
            is_neg = number < 0
            num_abs = abs(number)
            s, *d = str(f"{num_abs:.2f}").partition(".")
            r = ",".join([s[x-2:x] for x in range(-3, -len(s), -2)][::-1] + [s[-3:]]) if len(s) > 3 else s
            val = f"{r}{d[0]}{d[1]}"
            return f"(-) {val}" if is_neg else val

        # 2. Controls for Date Selection (SMART LINKED TO FY)
        fy_start_year = int(selected_fy.split('-')[0])
        fy_end_year = fy_start_year + 1
        fy_start_date = datetime(fy_start_year, 4, 1) # April 1st of the selected FY
        
        soe_month = st.selectbox("Select SOE Month", [calendar.month_name[m] for m in range(1, 13)], index=date.today().month-1, key="soe_m")
        month_idx = list(calendar.month_name).index(soe_month)
        
        # SMART YEAR LOGIC: Jan, Feb, Mar belong to the end year of the FY. April-Dec belong to the start year.
        if month_idx in [1, 2, 3]:
            soe_year = fy_end_year
        else:
            soe_year = fy_start_year
            
        st.info(f"📅 **Generating SOE for:** {soe_month} {soe_year} (Automatically linked to FY {selected_fy})")
        
        # Calculate last day of the selected month
        last_day = calendar.monthrange(soe_year, month_idx)[1]
        end_date = datetime(soe_year, month_idx, last_day) # Last day of selected month

        # 3. Setup Opening Balances Dictionary
        if 'opening_balances' not in data:
            data['opening_balances'] = {
                "Establishment Charges": 0.0, "TA": 0.0, "Contingencies": 0.0, 
                "TSP": 0.0, "Equipments": 0.0, "Works": 0.0
            }
            
        with st.expander(f"⚙️ Set Opening Balances for FY {selected_fy}"):
            st.write(f"Save the opening balances carried forward as on 01.04.{fy_start_year}.")
            
            with st.form(f"ob_form_{selected_fy}"):
                cols = st.columns(3)
                new_obs = {}
                for idx, (k, v) in enumerate(data['opening_balances'].items()):
                    # The unique key forces Streamlit to wipe the box clean when the FY changes
                    new_obs[k] = cols[idx % 3].number_input(
                        f"{k} (₹)", 
                        value=float(v), 
                        step=1000.0,
                        key=f"ob_input_{k}_{selected_fy}"
                    )
                if st.form_submit_button("💾 Save Opening Balances"):
                    data['opening_balances'] = new_obs
                    save_data(data, selected_fy)
                    st.success(f"Opening Balances Saved for {selected_fy}!")
                    st.rerun() 

        st.divider()
        st.markdown(f"<h3 style='text-align: center;'>Statement of Expenditure for the month of {soe_month} {soe_year}</h3>", unsafe_allow_html=True)
        st.markdown("**Name of the Centre:** Navsari")
        st.markdown("**Name of the Scheme:** AICRP/AINP on Agricultural Acarology, NAU, Navsari")

        # 4. Data Gathering & SMART Mapping Logic
        ob = data['opening_balances']
        funds = {k: 0.0 for k in ob.keys()}
        exp = {k: 0.0 for k in ob.keys()}
        
        # Smart Keyword Matcher: Converts any saved name into the strict SOE categories
        def get_smart_soe_head(raw_string):
            if not raw_string: return None
            rs = str(raw_string).upper()
            if "PAY" in rs or "ESTABLISHMENT" in rs: return "Establishment Charges"
            if "TA" in rs or "TRAVELLING" in rs: return "TA"
            if "TSP" in rs: return "TSP"
            if "NON" in rs or "EQUIP" in rs or "WORK" in rs: return "Equipments" 
            if "ORC" in rs or "CONTINGENC" in rs or "RECURRING" in rs: return "Contingencies"
            return None

        # Cumulative Funds Received (From April 1 up to end of selected month)
        for inst in data.get('installments', []):
            inst_date = datetime.strptime(inst['date'], "%Y-%m-%d")
            if fy_start_date <= inst_date <= end_date:
                for h_name, amt in inst.get('heads', {}).items():
                    soe_head = get_smart_soe_head(h_name)
                    if soe_head in funds:
                        funds[soe_head] += float(amt)

        # Cumulative Expenditure (From April 1 up to end of selected month)
        for e in data.get('expenditure', []):
            e_date = datetime.strptime(e['date'], "%Y-%m-%d")
            if fy_start_date <= e_date <= end_date:
                # Combine head and sub_head just in case, ensuring we catch the keyword
                combined_head = f"{e.get('head', '')} {e.get('sub_head', '')}"
                soe_head = get_smart_soe_head(combined_head)
                    
                if soe_head in exp:
                    exp[soe_head] += float(e.get('amount', 0.0))

        # 5. Build the Data Table Array
        data_table = []
        data_table.append(["", "A. Recurring Contingencies", "", "", "", "", "", ""])

        # Calculate A
        rec_heads = [("1.", "Establishment Charges"), ("2.", "TA"), ("3.", "Contingencies"), ("4.", "TSP")]
        tot_A = [0.0]*5 # [OB, Funds, Exp, ICAR, State]
        for sr, head in rec_heads:
            o = ob[head]; f = funds[head]; e = exp[head]
            i = e * 1.0 if head == "TSP" else e * 0.75
            s = 0.0 if head == "TSP" else e * 0.25
            data_table.append([sr, head, format_inr(o), format_inr(f), format_inr(e), format_inr(i), format_inr(s), format_inr(e)])
            tot_A[0] += o; tot_A[1] += f; tot_A[2] += e; tot_A[3] += i; tot_A[4] += s

        data_table.append(["", "Total - A", format_inr(tot_A[0]), format_inr(tot_A[1]), format_inr(tot_A[2]), format_inr(tot_A[3]), format_inr(tot_A[4]), format_inr(tot_A[2])])
        data_table.append(["", "B. Non Recurring Contingencies", "", "", "", "", "", ""])

        # Calculate B
        non_rec_heads = [("1.", "Equipments"), ("2.", "Works")]
        tot_B = [0.0]*5
        for sr, head in non_rec_heads:
            o = ob[head]; f = funds[head]; e = exp[head]
            i = e * 0.75
            s = e * 0.25
            data_table.append([sr, head, format_inr(o), format_inr(f), format_inr(e), format_inr(i), format_inr(s), format_inr(e)])
            tot_B[0] += o; tot_B[1] += f; tot_B[2] += e; tot_B[3] += i; tot_B[4] += s

        data_table.append(["", "Total - B", format_inr(tot_B[0]), format_inr(tot_B[1]), format_inr(tot_B[2]), format_inr(tot_B[3]), format_inr(tot_B[4]), format_inr(tot_B[2])])

        # Grand Total
        g_tot = [tot_A[j] + tot_B[j] for j in range(5)]
        data_table.append(["", "Grand Total A+B", format_inr(g_tot[0]), format_inr(g_tot[1]), format_inr(g_tot[2]), format_inr(g_tot[3]), format_inr(g_tot[4]), format_inr(g_tot[2])])

        # 6. Render Data
        columns = [
            "Sr. No.", "Head", f"Opening Balance as on 01.04.{fy_start_year}", f"Funds Received from Council", 
            f"Expenditure up to {soe_month} {soe_year}", "75% ICAR Share", "25% State Share", "Total"
        ]
        
        df_soe = pd.DataFrame(data_table, columns=columns)

        st.markdown("💡 **Tip: Your SOE is now calculated automatically. You can still click to edit values if needed before downloading.**")
        edited_df = st.data_editor(df_soe, use_container_width=True, hide_index=True)
        st.markdown("<ul><li>In 2025-26 State share released only in Pay and allowances</li></ul>", unsafe_allow_html=True)

        st.divider()
        if st.button("Generate SOE Word Document", key="soe_btn_new"):
            with st.spinner("Creating formatted Word file..."):
                soe_doc_buffer = create_word_doc(edited_df, soe_month, soe_year, last_day)
                st.success("SOE Generated with dynamic dates and calculations!")
                
                st.download_button(
                    label="📥 Download SOE Word File",
                    data=soe_doc_buffer,
                    file_name=f"SOE_{soe_month}_{soe_year}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                # =====================================================================
        # 👇 NEW SECTION: 12-MONTH YEARLY SUMMARY TABLE 👇
        # =====================================================================
        st.divider()
        st.markdown(f"<h3 style='text-align: center;'>12-Month Expenditure Summary (FY {selected_fy})</h3>", unsafe_allow_html=True)
        
        # Define the 12 months from April to March
        months_order = ['April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'January', 'February', 'March']
        short_months = [m[:3] for m in months_order] # e.g., 'Apr', 'May' for table columns
        
        # Initialize dictionary to hold 12 months of data for each head
        yearly_exp = {
            "Establishment Charges": {m: 0.0 for m in months_order},
            "TA": {m: 0.0 for m in months_order},
            "Contingencies": {m: 0.0 for m in months_order},
            "TSP": {m: 0.0 for m in months_order},
            "Equipments": {m: 0.0 for m in months_order},
            "Works": {m: 0.0 for m in months_order},
        }
        
        fy_end_date = datetime(fy_end_year, 3, 31) # Ends March 31st

        # Distribute all expenditures from the FY into their exact month buckets
        for e in data.get('expenditure', []):
            e_date = datetime.strptime(e['date'], "%Y-%m-%d")
            if fy_start_date <= e_date <= fy_end_date:
                m_name = calendar.month_name[e_date.month]
                combined_head = f"{e.get('head', '')} {e.get('sub_head', '')}"
                soe_head = get_smart_soe_head(combined_head) # Reuses your smart matcher!
                
                if soe_head in yearly_exp:
                    yearly_exp[soe_head][m_name] += float(e.get('amount', 0.0))

        # Build the 14-column layout for the Yearly Table
        y_table = []
        y_cols = ["Budget Head"] + short_months + ["Total"]

        y_table.append(["A. Recurring Contingencies"] + [""] * 13)

        # Calculate A (Recurring)
        tot_A_y = [0.0] * 13 # 12 months + 1 total
        for head in ["Establishment Charges", "TA", "Contingencies", "TSP"]:
            row = [head]
            row_tot = 0.0
            for i, m in enumerate(months_order):
                val = yearly_exp[head][m]
                row.append(format_inr(val))
                row_tot += val
                tot_A_y[i] += val
            row.append(format_inr(row_tot))
            tot_A_y[12] += row_tot
            y_table.append(row)

        y_table.append(["Total - A"] + [format_inr(v) for v in tot_A_y])
        y_table.append(["B. Non Recurring Contingencies"] + [""] * 13)

        # Calculate B (Non-Recurring)
        tot_B_y = [0.0] * 13
        for head in ["Equipments", "Works"]:
            row = [head]
            row_tot = 0.0
            for i, m in enumerate(months_order):
                val = yearly_exp[head][m]
                row.append(format_inr(val))
                row_tot += val
                tot_B_y[i] += val
            row.append(format_inr(row_tot))
            tot_B_y[12] += row_tot
            y_table.append(row)

        y_table.append(["Total - B"] + [format_inr(v) for v in tot_B_y])

        # Grand Total Calculation
        g_tot_y = [tot_A_y[i] + tot_B_y[i] for i in range(13)]
        y_table.append(["Grand Total A+B"] + [format_inr(v) for v in g_tot_y])

        # Render Yearly Data
        df_yearly = pd.DataFrame(y_table, columns=y_cols)
        
        st.markdown("💡 **Tip: Click to edit values before downloading your 12-Month Summary.**")
        edited_yearly_df = st.data_editor(df_yearly, use_container_width=True, hide_index=True, key="yearly_summary_editor")

        # Separate Download Button for 12-Month Summary
        if st.button("Generate 12-Month Summary Word Doc", key="yearly_btn_new"):
            with st.spinner("Creating wide landscape Word file..."):
                yearly_doc_buffer = create_yearly_word_doc(edited_yearly_df, selected_fy)
                st.success("12-Month Summary Generated!")
                
                st.download_button(
                    label="📥 Download Yearly Summary Word File",
                    data=yearly_doc_buffer,
                    file_name=f"12_Month_Summary_FY_{selected_fy}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                # =====================================================================
        # 🔵 SECTION 2: ANNUAL SOE (FULL FY)
        # =====================================================================
        st.divider()
        st.markdown(f"<h3 style='text-align: center;'>Annual Statement of Expenditure for FY {selected_fy}</h3>", unsafe_allow_html=True)
        st.info(f"Calculated from 01.04.{fy_start_year} up to 31.03.{fy_end_year}")

        funds_y = {k: 0.0 for k in ob.keys()}
        exp_y = {k: 0.0 for k in ob.keys()}

        # Annual Cumulative Funds
        for inst in data.get('installments', []):
            inst_date = datetime.strptime(inst['date'], "%Y-%m-%d")
            if fy_start_date <= inst_date <= fy_end_date:
                for h_name, amt in inst.get('heads', {}).items():
                    soe_head = get_smart_soe_head(h_name)
                    if soe_head in funds_y: funds_y[soe_head] += float(amt)

        # Annual Cumulative Exp
        for e in data.get('expenditure', []):
            e_date = datetime.strptime(e['date'], "%Y-%m-%d")
            if fy_start_date <= e_date <= fy_end_date:
                combined_head = f"{e.get('head', '')} {e.get('sub_head', '')}"
                soe_head = get_smart_soe_head(combined_head)
                if soe_head in exp_y: exp_y[soe_head] += float(e.get('amount', 0.0))

        # Build Annual Table Array
        data_table_y = [["", "A. Recurring Contingencies", "", "", "", "", "", ""]]
        tot_A_y = [0.0]*5
        for sr, head in rec_heads:
            o = ob[head]; f = funds_y[head]; e = exp_y[head]
            i = e * 1.0 if head == "TSP" else e * 0.75; s = 0.0 if head == "TSP" else e * 0.25
            data_table_y.append([sr, head, format_inr(o), format_inr(f), format_inr(e), format_inr(i), format_inr(s), format_inr(e)])
            tot_A_y[0] += o; tot_A_y[1] += f; tot_A_y[2] += e; tot_A_y[3] += i; tot_A_y[4] += s

        data_table_y.append(["", "Total - A", format_inr(tot_A_y[0]), format_inr(tot_A_y[1]), format_inr(tot_A_y[2]), format_inr(tot_A_y[3]), format_inr(tot_A_y[4]), format_inr(tot_A_y[2])])
        data_table_y.append(["", "B. Non Recurring Contingencies", "", "", "", "", "", ""])

        tot_B_y = [0.0]*5
        for sr, head in non_rec_heads:
            o = ob[head]; f = funds_y[head]; e = exp_y[head]
            i = e * 0.75; s = e * 0.25
            data_table_y.append([sr, head, format_inr(o), format_inr(f), format_inr(e), format_inr(i), format_inr(s), format_inr(e)])
            tot_B_y[0] += o; tot_B_y[1] += f; tot_B_y[2] += e; tot_B_y[3] += i; tot_B_y[4] += s

        data_table_y.append(["", "Total - B", format_inr(tot_B_y[0]), format_inr(tot_B_y[1]), format_inr(tot_B_y[2]), format_inr(tot_B_y[3]), format_inr(tot_B_y[4]), format_inr(tot_B_y[2])])
        g_tot_y = [tot_A_y[j] + tot_B_y[j] for j in range(5)]
        data_table_y.append(["", "Grand Total A+B", format_inr(g_tot_y[0]), format_inr(g_tot_y[1]), format_inr(g_tot_y[2]), format_inr(g_tot_y[3]), format_inr(g_tot_y[4]), format_inr(g_tot_y[2])])

        columns_y = [
            "Sr. No.", "Head", f"Opening Balance as on 01.04.{fy_start_year}", f"Funds Received during FY {selected_fy}", 
            f"Expenditure during FY {selected_fy}", "75% ICAR Share", "25% State Share", "Total"
        ]
        
        df_soe_y = pd.DataFrame(data_table_y, columns=columns_y)
        edited_df_y = st.data_editor(df_soe_y, use_container_width=True, hide_index=True, key="annual_edit")

        if st.button("Generate Annual SOE Word Document", key="soe_btn_y"):
            with st.spinner("Creating Annual Word file..."):
                annual_doc_buffer = create_annual_word_doc(edited_df_y, selected_fy)
                st.success("Annual SOE Generated!")
                st.download_button(
                    label="📥 Download Annual SOE Word File",
                    data=annual_doc_buffer,
                    file_name=f"Annual_SOE_FY_{selected_fy}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

# =====================================================================
    # 👇 NEW SECTION: TAB 6 - AUC GENERATION 👇
    # =====================================================================
    with tabs[6]:
        st.header("Audit Utilization Certificate (AUC) & Forwarding Letter")
        st.write("Generate the final year-end AUC and its corresponding Gujarati forwarding letter.")
        
        # Calculate Base Variables for the FY
        fy_start_year = int(selected_fy.split('-')[0])
        fy_end_year = int("20" + selected_fy.split('-')[1])
        fy_start_date = datetime(fy_start_year, 4, 1)
        fy_end_date = datetime(fy_end_year, 3, 31)
        
        # 1. Gather Installment Data
        inst_data = []
        tot_remittance = 0.0
        for idx, inst in enumerate(data.get('installments', [])):
            inst_date = datetime.strptime(inst['date'], "%Y-%m-%d")
            if fy_start_date <= inst_date <= fy_end_date:
                amt = float(inst.get('amount', 0.0))
                inst_data.append([idx+1, f"DSC Transaction payment advice report Dated: {inst.get('date')}", f"{amt:,.2f}"])
                tot_remittance += amt
        inst_data.append(["", "Total :-", f"{tot_remittance:,.2f}"])
        
        # 2. Gather Expenditure Data
        def format_inr_auc(number):
            if number == 0: return "-"
            is_neg = number < 0
            num_abs = abs(number)
            s, *d = str(f"{num_abs:.2f}").partition(".")
            r = ",".join([s[x-2:x] for x in range(-3, -len(s), -2)][::-1] + [s[-3:]]) if len(s) > 3 else s
            return f"(-) {r}{d[0]}{d[1]}" if is_neg else f"{r}{d[0]}{d[1]}"

        # Calculate Total ICAR Expense
        ob = data.get('opening_balances', {})
        total_opening_bal_icar = sum(float(v) for v in ob.values())
        
        exp_y = { "Establishment Charges": 0.0, "TA": 0.0, "Contingencies": 0.0, "TSP": 0.0, "Equipments": 0.0, "Works": 0.0 }
        
        def get_smart_soe_head(raw_string):
            rs = str(raw_string).upper()
            if "PAY" in rs or "ESTABLISHMENT" in rs: return "Establishment Charges"
            if "TA" in rs or "TRAVELLING" in rs: return "TA"
            if "TSP" in rs: return "TSP"
            if "NON" in rs or "EQUIP" in rs or "WORK" in rs: return "Equipments" 
            if "ORC" in rs or "CONTINGENC" in rs or "RECURRING" in rs: return "Contingencies"
            return None
            
        for e in data.get('expenditure', []):
            e_date = datetime.strptime(e['date'], "%Y-%m-%d")
            if fy_start_date <= e_date <= fy_end_date:
                combined_head = f"{e.get('head', '')} {e.get('sub_head', '')}"
                soe_head = get_smart_soe_head(combined_head)
                if soe_head in exp_y: exp_y[soe_head] += float(e.get('amount', 0.0))

        # Calculate total ICAR share of expenses
        tot_icar_exp = 0.0
        for head, amt in exp_y.items():
            if head == "TSP": tot_icar_exp += amt * 1.0
            else: tot_icar_exp += amt * 0.75
            
        closing_balance = (total_opening_bal_icar + tot_remittance) - tot_icar_exp
        
        t1_data = [
            format_inr_auc(total_opening_bal_icar),
            format_inr_auc(tot_remittance),
            "0.00",
            format_inr_auc(tot_icar_exp),
            format_inr_auc(closing_balance)
        ]
        
        # Build Table 2
        alloc = data.get('allocation', {})
        def get_alloc(head_kw):
            for k,v in alloc.items():
                if head_kw in k.upper(): return v.get('total', 0.0)
            return 0.0
            
        t2_data = [
            ["Recurring", "", "", "", ""],
            ["Pay and Allowance", format_inr_auc(get_alloc("PAY")), format_inr_auc(exp_y["Establishment Charges"]*0.75), format_inr_auc(exp_y["Establishment Charges"]*0.25), format_inr_auc(exp_y["Establishment Charges"])],
            ["Travelling Allowance", format_inr_auc(get_alloc("TA")), format_inr_auc(exp_y["TA"]*0.75), format_inr_auc(exp_y["TA"]*0.25), format_inr_auc(exp_y["TA"])],
            ["Recurring Contingencies", format_inr_auc(get_alloc("ORC")), format_inr_auc(exp_y["Contingencies"]*0.75), format_inr_auc(exp_y["Contingencies"]*0.25), format_inr_auc(exp_y["Contingencies"])],
            ["HRD", "-", "-", "-", "-"],
            ["Non Recurring Contingencies", "", "", "", ""],
            ["Equipment", format_inr_auc(get_alloc("EQUIP")), format_inr_auc(exp_y["Equipments"]*0.75), format_inr_auc(exp_y["Equipments"]*0.25), format_inr_auc(exp_y["Equipments"])],
            ["Works", "-", "-", "-", "-"],
            ["Vehicle (IT)", "-", "-", "-", "-"],
        ]
        
        tot_alloc = sum(get_alloc(k) for k in ["PAY", "TA", "ORC", "EQUIP"])
        tot_state_exp = sum(exp_y[k]*0.25 for k in exp_y if k != "TSP")
        tot_all_exp = sum(exp_y.values())
        t2_data.append(["Total:-", format_inr_auc(tot_alloc), format_inr_auc(tot_icar_exp), format_inr_auc(tot_state_exp), format_inr_auc(tot_all_exp)])
        
        # --- EDITABLE LIVE PREVIEW SECTION ---
        st.divider()
        st.subheader("👀 AUC Live Preview & Editor")
        st.info("💡 Make any adjustments below before generating the final Word document.")
        
        # 1. Editable Installment Table
        st.markdown("**Received Installments:**")
        df_inst = pd.DataFrame(inst_data, columns=["Sr.No", "Letter No and Date", "Amount"])
        edited_df_inst = st.data_editor(df_inst, use_container_width=True, hide_index=True, key="auc_edit_inst")
        final_inst_data = edited_df_inst.values.tolist()
        
        # 2. Editable Main Paragraph
        st.markdown("**Form of Utilization Certificate & Audit Utilization Certificate**")
        default_cert_text = f"1. Certified that the out of Rs. {format_inr_auc(tot_remittance)} sanctioned during the year {selected_fy} in favour of Comptroller, NAU, Navsari under this Ministry/Department Letter No. given in the margin and Rs. {format_inr_auc(total_opening_bal_icar)} on account of unspent balance of the previous year, a sum of Rs. {format_inr_auc(tot_icar_exp)} has been Utilized for the purpose of Agril. Acarology Research and remaining unutilized at the end of the year has been surrendered (vide No......Dated......) will be adjusted (to be payable the next year)."
        
        edited_cert_text = st.text_area("Certificate Paragraph 1 (Editable):", value=default_cert_text, height=120)
        st.markdown("2. Certified that I have satisfied myself that the condition on which the expenditure was made have dully fulfilled/are being fulfilled and that I have exercised the following check to see that the money was actually utilized for the purpose for which it was sanctioned.")
        
        # 3. Editable Table 1
        st.markdown("**Table 1: Showing the details of receipt and expenditure figure (in Rupees)**")
        t1_columns = [f"Opening balance as on 1st April {fy_start_year}", f"Remittance Received {selected_fy}", "Receipt", f"ICAR share of Exp during {selected_fy}", f"Closing balance as on 31st March {fy_end_year}"]
        df_t1 = pd.DataFrame([t1_data], columns=t1_columns)
        edited_df_t1 = st.data_editor(df_t1, use_container_width=True, hide_index=True, key="auc_edit_t1")
        final_t1_data = edited_df_t1.values.tolist()[0]
        
        # 4. Editable Table 2
        st.markdown("**Table 2: Showing the head wise details of expenditure figure (in Rupees)**")
        t2_columns = ["Head", f"Allocation for the Year {selected_fy} (100%)", "ICAR share of Expenditure (75%)", "State Share (25%)", "Total Expenditure"]
        df_t2 = pd.DataFrame(t2_data, columns=t2_columns)
        edited_df_t2 = st.data_editor(df_t2, use_container_width=True, hide_index=True, key="auc_edit_t2")
        final_t2_data = edited_df_t2.values.tolist()
        
        st.divider()

        # --- DOWNLOAD BUTTONS SECTION ---
        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("1. Audit Utilization Certificate (AUC)")
            st.write("Generates using the edited data from the preview above.")
            if st.button("📥 Generate & Download AUC Document"):
                with st.spinner("Generating AUC..."):
                    # Pass the final edited values into the generator
                    auc_doc = generate_auc_certificate(final_inst_data, final_t1_data, final_t2_data, edited_cert_text, selected_fy)
                    st.download_button("Download AUC (.docx)", data=auc_doc, file_name=f"AUC_FY_{selected_fy}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    
        with col_b:
            st.subheader("2. AUC Forwarding Letter")
            ref_no = st.text_input("Reference No. (જા.નં. એસીએન/એન્ટો/___/૨૦૨૬):", value="AUC", key="auc_ref_no")
            letter_date = st.text_input("Date (તારીખ):", value=datetime.now().strftime("%d/%m/%Y"), key="auc_letter_date")
            subj = st.text_area("Subject (વિષય):", value=f"AINP on Agricultural Acarology (BH.303/2092) નું વર્ષ {selected_fy} નુ Audit Utilization Certificate (AUC) મોકલવા બાબત.", key="auc_subject")
            body = st.text_area("Body Text:", value=f"જય ભારત સહ ઉપરોક્ત વિષય અન્વયે જણાવવાનું કે, અત્રેના કીટકશાસ્ત્ર વિભાગ ખાતે ચાલતી આઈ.સી.એ.આર. યોજના AINP on Agricultural Acarology (BH.303/2092) નું વર્ષ {selected_fy} નુ Audit Utilization Certificate (AUC) આ સાથે સામેલ રાખી મોકલી આપીએ છીએ.", key="auc_body_text")
            
            if st.button("📥 Generate & Download Forwarding Letter", key="auc_download_btn"):
                with st.spinner("Generating Letter..."):
                    fw_doc = generate_auc_forwarding_docx(ref_no, letter_date, subj, body)
                    st.download_button("Download Forwarding Letter (.docx)", data=fw_doc, file_name=f"AUC_Forwarding_Letter_{selected_fy}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="auc_final_dl_btn")

    # --- TAB 7: AI CHATBOT ---
    with tabs[7]:  # <--- MAKE SURE THIS IS CHANGED TO 7
        st.header("Grant Smart-Assistant")
        st.write("Ask questions like: *'How much is remaining in ORC Recurring?'* or *'Generate a summary of spend for Quarter 3'*.")

        budget_summary = data['revised_allocation'] if data['revised_allocation'] else data['allocation']
        received_summary = sum(inst['amount'] for inst in data['installments'])
        
        # Safe loading for chatbot
        df_exp = pd.DataFrame(data.get('expenditure', []))
        spend_summary = {}
        if not df_exp.empty:
            spend_summary = df_exp.groupby('head')['amount'].sum().to_dict()
            
        system_context = f"""
        You are a smart financial assistant for Dr. Vaibhav Chaudhari, managing the AINP on Agricultural Acarology grant at NAU Navsari.
        The current Financial Year is {selected_fy}.
        Total Funds Received from ICAR: ₹{received_summary:,}
        
        Budget Head Allocation Summary: {json.dumps(budget_summary, indent=2)}
        Current Expenditure Totals per Head: {json.dumps(spend_summary, indent=2)}
        
        You have context about the budget, installments received (PFMS), and expenditures added.
        Answer user questions robustly, calculating remaining balances where needed (Allocation - Spend).
        If the user asks for a report, summarize the data clearly in Markdown table format.
        """

        if "messages" not in st.session_state:
            st.session_state.messages = []

        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        if prompt := st.chat_input("Ask about grant status..."):
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("assistant"):
                message_placeholder = st.empty()
                full_response = ""
                
                try:
                    chat = model_pro.start_chat(history=[])
                    chat.send_message(f"SYSTEM_CONTEXT_DO_NOT_REPLY: {system_context}")
                    
                    response = chat.send_message(prompt)
                    full_response = response.text
                except Exception as e:
                    full_response = f"Sorry, AI service is currently unavailable. Error: {e}"

                message_placeholder.markdown(full_response)
            
            st.session_state.messages.append({"role": "assistant", "content": full_response})

# --- 4. RUN APPLICATION ---
if __name__ == "__main__":
    if NAU_LOGO and ICAR_LOGO:
        main()
    else:
        st.error("Missing Logos. Please ensure 'nau_logo.png' and 'icar_logo.png' are in the 'logos/' folder.")
        st.write(f"Paths checked: {NAU_LOGO}, {ICAR_LOGO}")
