import streamlit as st
import os
import json
import pandas as pd
from datetime import datetime, date
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
import io
from github import Github  # <-- Added for GitHub backups

# --- 1. CONFIGURATION & SETUP ---
st.set_page_config(page_title="AINP Grant Management System", page_icon="🌾", layout="wide")

# Directory Setup
DIRS = ['data', 'documents', 'fonts', 'logos']
for d in DIRS:
    if not os.path.exists(d):
        os.makedirs(d)

# AI Setup
api_key = st.secrets.get("GEMINI_API_KEY")
genai.configure(api_key=api_key)
model_pro = genai.GenerativeModel('gemini-1.5-pro-latest')   

# Load Logos if exist
NAU_LOGO = 'logos/nau_logo.png' if os.path.exists('logos/nau_logo.png') else None
ICAR_LOGO = 'logos/icar_logo.png' if os.path.exists('icar_logo.png') else None
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
def backup_to_github(filepath, content):
    """Pushes saved data/PDFs back to GitHub so it isn't lost when Streamlit sleeps."""
    github_token = st.secrets.get("GITHUB_TOKEN")
    
    if not github_token:
        return

    try:
        g = Github(github_token)
        repo = g.get_repo("vkcvaibhav/AINP-Grant-Manager") # Your exact repo name
        
        try:
            contents = repo.get_contents(filepath)
            repo.update_file(contents.path, f"Auto-backup {filepath}", content, contents.sha)
        except Exception:
            repo.create_file(filepath, f"Auto-create {filepath}", content)
            
    except Exception as e:
        st.warning(f"Failed to backup to GitHub. Error: {e}")

def get_fy_filename(fy):
    return f'data/grant_data_{fy.replace("-", "_")}.json'

def load_data(fy):
    filename = get_fy_filename(fy)
    if os.path.exists(filename):
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # Ensure new structures exist in older save files
            if 'quarterly_allocations' not in data:
                data['quarterly_allocations'] = {"Q1": {}, "Q2": {}, "Q3": {}, "Q4": {}}
            if 'pdfs' not in data:
                data['pdfs'] = {}
            return data
    else:
        return {
            "financial_year": fy,
            "allocation": {}, 
            "revised_allocation": {},
            "quarterly_allocations": {"Q1": {}, "Q2": {}, "Q3": {}, "Q4": {}},
            "installments": [], 
            "expenditure": [],
            "pdfs": {}, # To store the paths of uploaded PDFs
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
        pdf_data = {
            "mime_type": "application/pdf",
            "data": uploaded_file.getvalue()
        }

        full_prompt = f"""
        Analyze the attached document content and extract the required information 
        as a structured JSON object.

        DOCUMENT TYPE/CONTEXT: {prompt_task['context']}
        REQUIRED JSON STRUCTURE: {json.dumps(prompt_task['structure'], indent=2)}

        IMPORTANT: Only return the JSON object, nothing else. If a field cannot be found, set it to 0.
        For currency, return numbers only (e.g., 100000). Convert dates to YYYY-MM-DD format.
        """
        
        response = model_pro.generate_content([full_prompt, pdf_data])
        
        json_str = response.text.replace('
```json', '').replace('```', '').strip()
        extracted_data = json.loads(json_str)
        return extracted_data

    except Exception as e:
        st.error(f"Error processing document with AI: {e}")
        return None

# --- C. Output Generation (PDF/WORD) ---

class GujaratiPDF(FPDF):
    def header(self):
        if ICAR_LOGO:
            self.image(ICAR_LOGO, 10, 8, 20)
        if NAU_LOGO:
            self.image(NAU_LOGO, 180, 8, 20)
        
        self.add_font('Gujarati', '', GUJARATI_FONT, uni=True)
        self.set_font('Gujarati', '', 12)
        
        self.cell(0, 5, 'કીટકશાસ્ત્ર વિભાગ', ln=True, align='C')
        self.cell(0, 5, 'ન. મ. કૃષિ મહાવિદ્યાલય', ln=True, align='C')
        self.cell(0, 5, 'નવસારી કૃષિ યુનિવર્સિટી', ln=True, align='C')
        self.cell(0, 5, 'નવસારી- ૩૯૬ ૪૫૦ (ગુજરાત)', ln=True, align='C')
        self.ln(10)

def generate_comptroller_letter(data, installment):
    if not os.path.exists(GUJARATI_FONT):
        st.error(f"Gujarati font not found at {GUJARATI_FONT}. Cannot generate letter.")
        return None

    today_str = datetime.now().strftime("%d/%m/%Y")
    
    pdf = GujaratiPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    
    pdf.cell(0, 10, f'જા.નં. એસીએન/એન્ટો/Grant/{installment["type"]}/2026, નવસારી', ln=False)
    pdf.cell(0, 10, f'તારીખ: {today_str}', ln=True, align='R')
    pdf.ln(5)
    
    pdf.cell(0, 7, 'પ્રતિ,', ln=True)
    pdf.cell(0, 7, 'હિસાબ નિયામકશ્રી', ln=True)
    pdf.cell(0, 7, 'નવસારી કૃષિ યુનિવર્સિટી, નવસારી', ln=True)
    pdf.ln(3)
    
    pdf.cell(0, 7, 'મારફત સવિનય: આચાર્ય અને ડિનશ્રી, ન. મ. કૃષિ મહાવિદ્યાલય, ન.કૃ.યુ., નવસારી ૩૯૬ ૪૫૦', ln=True)
    pdf.ln(5)

    subj = f"વિષય:- ICAR – NCIPM તરફથી આવેલ AINP Acrology Grant ({installment['type']}) ફાળવવા બાબત..."
    pdf.set_font('Gujarati', '', 12)
    pdf.multi_cell(0, 7, subj, 0, 'J')
    pdf.ln(5)

    body = f"""જય ભારત સહ ઉપરોક્ત વિષય અન્વયે જણાવવાનું કે, અત્રેના કીટકશાસ્ત્ર વિભાગ ખાતે ચાલતી આઈ.સી.એ.આર. યોજના AINP on Agricultural Acarology (75:25%) માં આવેલ ગ્રાન્ટ રૂ. {installment['amount']:,}/- (PFMS ID: {installment['pfms_id']}) ને કોષ્ટકમાં જણાવ્યાનુસાર ફાળવી આપવા આપ સાહેબશ્રીને નમ્ર વિનંતી."""
    pdf.multi_cell(0, 7, body, 0, 'J')
    pdf.ln(5)

    pdf.set_font('Helvetica', 'B', 10) 
    cols = [60, 40, 40, 40]
    pdf.cell(cols[0], 7, 'Head', 1, 0, 'C')
    pdf.cell(cols[1], 7, 'Pay', 1, 0, 'C')
    pdf.cell(cols[2], 7, 'Recurring', 1, 0, 'C')
    pdf.cell(cols[3], 7, 'Total Amount', 1, 1, 'C')
    
    pdf.set_font('Helvetica', '', 10)
    pdf.cell(cols[0], 7, 'AINP Acarology', 1, 0, 'L')
    pdf.cell(cols[1], 7, '-', 1, 0, 'C')
    pdf.cell(cols[2], 7, '-', 1, 0, 'C')
    pdf.cell(cols[3], 7, f'{installment["amount"]:,}/-', 1, 1, 'R')
    pdf.ln(10)

    pdf.set_font('Gujarati', '', 12)
    pdf.cell(0, 7, 'પ્રોજેક્ટ ઈન્ચાર્જ                                                                                   પ્રાધ્યાપક અને વડા', ln=True)

    letter_filename = f"documents/Letter_Comptroller_{installment['pfms_id']}.pdf"
    pdf.output(letter_filename)
    
    with open(letter_filename, "rb") as f:
        pdf_bytes = f.read()
    return pdf_bytes

def generate_soe_word(data, month, year):
    doc = Document()
    
    title = doc.add_paragraph('Statement of Expenditure for the month of ')
    title.add_run(f'{month} {year}').bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    details = doc.add_paragraph()
    details.add_run('Name of the Centre: ').bold = True
    details.add_run('Navsari\n')
    details.add_run('Name of the Scheme: ').bold = True
    details.add_run('AICRP/AINP on Agricultural Acarology, NAU, Navsari')
    details.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table = doc.add_table(rows=2, cols=9)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Sr. No.'
    hdr_cells[1].text = 'Head'
    hdr_cells[2].text = 'Opening Balance as on 01.04'
    hdr_cells[3].text = 'Funds Received during Year'
    hdr_cells[4].text = f'Expenditure up to month of {month}'
    hdr_cells[5].text = 'Cumulative Expenditure'
    
    row_cells = table.add_row().cells
    row_cells[0].text = '1'
    row_cells[1].text = BUDGET_HEADS[0]
    row_cells[2].text = '-1,38,340' 
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# --- 3. THE UI APPLICATION ---

def main():
    st.title("🌾 AINP Grant Management System - NAU Navsari")
    
    current_year = datetime.now().year
    fy_options = [f"{y}-{str(y+1)[2:]}" for y in range(current_year-2, current_year+2)]
    selected_fy = st.sidebar.selectbox("Select Financial Year (Apr 1 - Mar 31)", fy_options, index=2)
    
    data = load_data(selected_fy)
    
    tabs = st.tabs(["📊 Dashboard", "📤 1. Budget Intake", "💰 2. Installments (PFMS)", "📝 3. Generated Letters", "💸 4. Monthly Spend", "🤖 AI Chatbot"])
    
    # --- TAB 1: DASHBOARD ---
    with tabs[0]:
        st.header(f"Financial Year {selected_fy}")
        
        current_qtr = data.get('latest_quarter', 'Full Year')
        last_date = data.get('latest_date', 'N/A')
        st.markdown(f"**Current Active Budget Period:** {current_qtr} | **Last Document Date:** {last_date}")
        
        # 1. Show Document Data Blocks (Clickable Expanders)
        st.divider()
        st.subheader("📑 Uploaded Budget Documents")
        
        # Helper to render a table with a Total row and PDF download button
        def render_budget_expander(title, dict_data, doc_key):
            with st.expander(title):
                df = pd.DataFrame.from_dict(dict_data, orient='index')
                df.loc['Total'] = df.sum(numeric_only=True)
                st.dataframe(df, use_container_width=True)
                
                # Check for PDF
                pdf_path = data.get('pdfs', {}).get(doc_key)
                if pdf_path and os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f:
                        st.download_button(f"📥 Download {doc_key} PDF", f, file_name=pdf_path.split('/')[-1], mime="application/pdf", key=f"dl_{doc_key}")
        
        col_a, col_b = st.columns(2)
        with col_a:
            if data.get('allocation'):
                render_budget_expander("📄 Initial Full-Year Budget Allocation", data['allocation'], "Initial Allocation")
        with col_b:
            if data.get('revised_allocation'):
                render_budget_expander("📄 Revised Full-Year Budget Allocation", data['revised_allocation'], "Revised Allocation")
                    
        st.markdown("<br><b>🗓️ Quarterly Releases (Click to view data & PDFs)</b>", unsafe_allow_html=True)
        q_cols = st.columns(4)
        for i, q in enumerate(["Q1", "Q2", "Q3", "Q4"]):
            q_data = data['quarterly_allocations'].get(q)
            if q_data:
                with q_cols[i]:
                    render_budget_expander(f"📄 {q} Release", q_data, f"{q} Release")
                        
        st.divider()
        
        # 2. Total Mismatch/Status Table
        active_budget = data.get('revised_allocation') if data.get('revised_allocation') else data.get('allocation')
        if active_budget:
            st.subheader("⚖️ Allocation vs. Quarterly Release Mismatch")
            mismatch_data = {}
            
            for head in BUDGET_HEADS:
                # Get Sanctioned Total for the head
                tot_alloc = active_budget.get(head, {}).get('total', 0.0)
                
                # Sum the quarterly totals for this specific head
                q_sum = 0.0
                for q in ["Q1", "Q2", "Q3", "Q4"]:
                    if data['quarterly_allocations'].get(q):
                        q_sum += data['quarterly_allocations'][q].get(head, {}).get('total', 0.0)
                        
                mismatch = tot_alloc - q_sum
                mismatch_data[head] = {
                    "Total Sanctioned (Lakhs)": round(tot_alloc, 2),
                    "Released Q1-Q4 (Lakhs)": round(q_sum, 2),
                    "Pending/Mismatch (Lakhs)": round(mismatch, 2)
                }
            
            # Create DataFrame and add Grand Total row
            df_mismatch = pd.DataFrame.from_dict(mismatch_data, orient='index')
            df_mismatch.loc['Grand Total'] = df_mismatch.sum()
            
            # Display Table
            st.dataframe(df_mismatch, use_container_width=True)
        else:
            st.info("Upload an Initial or Revised Budget Allocation to view the Mismatch Table.")

    # --- TAB 2: BUDGET INTAKE ---
    with tabs[1]:
        st.header("Upload Budget Allocation / Revision PDF")
        
        st.markdown("""
        **Instructions for Uploading:**
        Select the specific type of document you are uploading to help the AI accurately extract the data.
        *   **Initial / Revised Allocation:** Full year budget documents with Total, ICAR, and State shares.
        *   **Quarterly Release:** Letters (e.g., *Q1, Q2, Q3, Q4*) that typically only show the **ICAR Share**. The system will automatically calculate the 25% State Share and the Total amount based on the ICAR value!
        """)
        
        doc_type = st.radio("Select Document Type:", [
            "Initial Allocation", 
            "Revised Allocation", 
            "Q1 Release", 
            "Q2 Release", 
            "Q3 Release", 
            "Q4 Release"
        ])
        
        budget_file = st.file_uploader("Upload Budget Document", type=['pdf'], key="budget_up")
        
        if budget_file and st.button("Analyze & Process Budget"):
            with st.spinner(f"AI is analyzing {doc_type} PDF..."):
                
                budget_prompt = {
                    "context": f"Document Type: {doc_type}. Scheme: AICRP/AINP Acarology. Financial Year: {selected_fy}. Extract the budget allocation table strictly using the exact 5 heads defined.",
                    "structure": {
                        "is_revision": "boolean (true if this revises a previous allocation)",
                        "date": "YYYY-MM-DD",
                        "heads": [
                            {"head_name": "Pay and Allowances", "icar_share": 0.0, "state_share": 0.0, "total": 0.0},
                            {"head_name": "Travelling Allowances (TA)", "icar_share": 0.0, "state_share": 0.0, "total": 0.0},
                            {"head_name": "Other Recurring Contingencies (ORC)", "icar_share": 0.0, "state_share": 0.0, "total": 0.0},
                            {"head_name": "Non-Recurring Contingencies (Equipments/Works)", "icar_share": 0.0, "state_share": 0.0, "total": 0.0},
                            {"head_name": "TSP", "icar_share": 0.0, "state_share": 0.0, "total": 0.0}
                        ]
                    }
                }
                
                extracted_budget = process_upload_with_ai(budget_file, budget_prompt)
                
                if extracted_budget:
                    st.success("Document analyzed successfully!")
                    
                    # 1. Initialize a master dictionary forcing all 5 heads to 0.0
                    budget_dict = {head: {'icar': 0.0, 'state': 0.0, 'total': 0.0} for head in BUDGET_HEADS}
                    
                    # 2. STRICT 75:25 MATHEMATICAL ENFORCEMENT & HEAD MAPPING
                    for head in extracted_budget.get('heads', []):
                        h_name = head.get('head_name', '').upper()
                        extracted_total = float(head.get('total') or 0.0)
                        extracted_icar = float(head.get('icar_share') or 0.0)
                        
                        # Match the AI's head name to our strict 5 standard BUDGET_HEADS
                        matched_head = None
                        if "PAY" in h_name: matched_head = BUDGET_HEADS[0]
                        elif "TRAVEL" in h_name or "TA" in h_name: matched_head = BUDGET_HEADS[1]
                        elif "OTHER" in h_name or "ORC" in h_name: matched_head = BUDGET_HEADS[2]
                        elif "NON" in h_name or "EQUIPMENT" in h_name: matched_head = BUDGET_HEADS[3]
                        elif "TSP" in h_name: matched_head = BUDGET_HEADS[4]
                        
                        if not matched_head:
                            continue # Skip unknown categories
                        
                        # Exception: TSP is always 100% ICAR share
                        if "TSP" in matched_head:
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
                    
                    
                    # --- Save the PDF File Locally and push to GitHub ---
                    safe_doc_type = doc_type.replace(" ", "_")
                    pdf_path = f"documents/{safe_doc_type}_{selected_fy}.pdf"
                    with open(pdf_path, "wb") as f:
                        f.write(budget_file.getvalue())
                    # Push PDF bytes to GitHub
                    backup_to_github(pdf_path, budget_file.getvalue())
                    
                    # Store PDF reference in data
                    data['pdfs'][doc_type] = pdf_path
                    
                    # --- Route data to correct storage location based on Radio Button ---
                    if doc_type == "Initial Allocation":
                        data['allocation'] = budget_dict
                    elif doc_type == "Revised Allocation":
                        data['revised_allocation'] = budget_dict
                    else:
                        q_key = doc_type.split(" ")[0] # Extracts "Q1", "Q2", etc.
                        data['quarterly_allocations'][q_key] = budget_dict
                    
                    doc_date = extracted_budget.get('date', 'Unknown')
                    data['latest_quarter'] = doc_type
                    data['latest_date'] = doc_date
                    
                    save_data(data, selected_fy)
                    
                    st.info(f"📅 **Document Date:** {doc_date} | 🔄 **Type:** {doc_type}")
                    st.toast("Budget Data & PDF Saved to GitHub!")

    # --- TAB 3: INSTALLMENTS ---
    with tabs[2]:
        st.header("Recieved Grant Installment (PFMS Advice)")
        st.write("Mail attachment like: *'12 02 2026 100000.pdf'*")
        
        pfms_file = st.file_uploader("Upload PFMS Receipt PDF", type=['pdf'], key="pfms_up")
        inst_type = st.selectbox("Installment Number (e.g. from mail body)", ["I", "II", "III", "IV", "V", "Revised I"], key="inst_type")

        if pfms_file and st.button("Process Installment"):
            with st.spinner("Analyzing PFMS document..."):
                
                pfms_prompt = {
                    "context": "PFMS GENERATED DSC TRANSACTION PAYMENT ADVICE REPORT.",
                    "structure": {
                        "date": "YYYY-MM-DD (DSC Signing Date)",
                        "amount": 0.0,
                        "pfms_transaction_id": "string (CO...)",
                        "bank_account_last_digits": "string (...3215)"
                    }
                }
                
                extracted_pfms = process_upload_with_ai(pfms_file, pfms_prompt)
                
                if extracted_pfms:
                    st.success(f"Received ₹{extracted_pfms['amount']:,}!")
                    
                    new_inst = {
                        "date": extracted_pfms['date'],
                        "amount": extracted_pfms['amount'],
                        "pfms_id": extracted_pfms['pfms_transaction_id'],
                        "type": inst_type,
                        "available": False 
                    }
                    
                    if not any(inst['pfms_id'] == new_inst['pfms_id'] for inst in data['installments']):
                        data['installments'].append(new_inst)
                        save_data(data, selected_fy)
                        st.toast("Installment Added and Backed up!")
                    else:
                        st.warning("Installment already exists.")

    # --- TAB 4: GENERATED LETTERS ---
    with tabs[3]:
        st.header("Draft Letters based on PFMS Receipts")
        st.write("Generate the Gujarati letter (image_8.png template) to send to Comptroller.")
        
        pending_utilization = [inst for inst in data['installments'] if not inst.get('utilization_letter_generated')]
        
        if pending_utilization:
            options = {f"{inst['type']} (₹{inst['amount']:,} - {inst['pfms_id']})": inst for inst in pending_utilization}
            selected_inst_str = st.selectbox("Select PFMS Receipt to draft letter for:", list(options.keys()))
            selected_inst_data = options[selected_inst_str]

            if st.button("Draft Letter in Gujarati (PDF)"):
                with st.spinner("Generating PDF based on template..."):
                    pdf_bytes = generate_comptroller_letter(data, selected_inst_data)
                    
                    if pdf_bytes:
                        st.success("Draft Generated!")
                        st.download_button(
                            label="Download Comptroller Letter (A4 PDF)",
                            data=pdf_bytes,
                            file_name=f"Letter_to_Comptroller_{selected_inst_data['type']}_{selected_inst_data['pfms_id']}.pdf",
                            mime="application/pdf"
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


    # --- TAB 5: MONTHLY SPEND & SOE ---
    with tabs[4]:
        st.header("Monthly Expenditure & SOE Generation")
        
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

        df_exp = pd.DataFrame(data['expenditure'])
        if not df_exp.empty:
            df_exp['date'] = pd.to_datetime(df_exp['date'])
            current_month_exp = df_exp[
                (df_exp['date'].dt.strftime('%B') == month_to_process) & 
                (df_exp['date'].dt.year == year_to_process)
            ]
            st.subheader(f"Spend in {month_to_process} {year_to_process}")
            st.dataframe(current_month_exp, use_container_width=True)

        st.divider()
        if st.button("Generate Monthly SOE (Word Doc)", key="soe_btn"):
            with st.spinner("Calculating balances and creating Word file..."):
                soe_doc_buffer = generate_soe_word(data, month_to_process, year_to_process)
                st.success("SOE Generated!")
                st.download_button(
                    label="Download SOE Word File",
                    data=soe_doc_buffer,
                    file_name=f"SOE_{selected_fy}_{month_to_process}_{year_to_process}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    # --- TAB 6: AI CHATBOT ---
    with tabs[5]:
        st.header("Grant Smart-Assistant")
        st.write("Ask questions like: *'How much is remaining in ORC Recurring?'* or *'Generate a summary of spend for Quarter 3'*.")

        budget_summary = data['revised_allocation'] if data['revised_allocation'] else data['allocation']
        received_summary = sum(inst['amount'] for inst in data['installments'])
        
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
