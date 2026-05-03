import streamlit as st
import os
import json
import pandas as pd
from datetime import datetime, date
from PIL import Image
import pytesseract
import pdfplumber
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
import io
from github import Github  # <-- NEW: Added for GitHub backups

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
            return json.load(f)
    else:
        return {
            "financial_year": fy,
            "allocation": {}, 
            "revised_allocation": {},
            "installments": [], 
            "expenditure": []   
        }

def save_data(data, fy):
    filename = get_fy_filename(fy)
    
    # 1. Save locally for immediate app use
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
        
    # 2. Push to GitHub for permanent backup
    json_str = json.dumps(data, indent=4, ensure_ascii=False)
    backup_to_github(filename, json_str)


# --- B. Document Processing (OCR & AI) ---
def process_upload_with_ai(uploaded_file, prompt_task):
    """Uses Gemini to read the file and extract JSON structured data."""
    try:
        if uploaded_file.type == "application/pdf":
            with pdfplumber.open(uploaded_file) as pdf:
                full_text = ""
                for page in pdf.pages:
                    full_text += page.extract_text() + "\n"
            
            if full_text.strip():
                content = full_text
            else:
                images = []
                with pdfplumber.open(uploaded_file) as pdf:
                    images.append(pdf.pages[0].to_image(resolution=200).original)
                content = images 

        elif "image" in uploaded_file.type:
            image = Image.open(uploaded_file)
            content = [image]

        full_prompt = f"""
        Analyze the attached document content (text or image) and extract the required information 
        as a structured JSON object.

        DOCUMENT TYPE/CONTEXT: {prompt_task['context']}
        REQUIRED JSON STRUCTURE: {json.dumps(prompt_task['structure'], indent=2)}

        IMPORTANT: Only return the JSON object, nothing else. If a field cannot be found, set it to null.
        For currency, return numbers only (e.g., 100000). Convert dates to YYYY-MM-DD format.
        """
        
        response = model_pro.generate_content([full_prompt, content] if isinstance(content, list) else [full_prompt, content])
        
        json_str = response.text.replace('```json', '').replace('```', '').strip()
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
        
        col1, col2, col3 = st.columns(3)
        with col1:
            total_allocated = sum(v.get('total', 0) for k,v in data['revised_allocation'].items()) if data['revised_allocation'] else 0
            st.metric("Total Sanctioned Budget", f"₹{total_allocated:,}")
        with col2:
            total_received = sum(inst['amount'] for inst in data['installments'])
            st.metric("Total Funds Received (PFMS)", f"₹{total_received:,}")
        with col3:
            st.metric("Pending Received", "Calculated later")
            
        st.subheader("Budget Head Overview")
        if data['revised_allocation']:
            df_budget = pd.DataFrame.from_dict(data['revised_allocation'], orient='index')
            st.dataframe(df_budget, use_container_width=True)
        else:
            st.info("Upload budget allocation to begin.")

    # --- TAB 2: BUDGET INTAKE ---
    with tabs[1]:
        st.header("Upload Budget Allocation / Revision PDF/Image")
        st.write("Emails like: *'Revised Budget Allocation 2025-26 - NAU Centre, Navsari'*")
        
        budget_file = st.file_uploader("Upload Budget Document", type=['pdf', 'png', 'jpg', 'jpeg'], key="budget_up")
        
        if budget_file and st.button("Analyze & Process Budget"):
            with st.spinner("AI is analyzing budget structure..."):
                
                budget_prompt = {
                    "context": f"Budget Allocation for AICRP/AINP Acarology scheme for Financial Year {selected_fy}.",
                    "structure": {
                        "is_revision": "boolean (true if this revises a previous allocation)",
                        "date": "YYYY-MM-DD",
                        "heads": [
                            {"head_name": "string (e.g., Pay & Allowances, Recurring)", "icar_share": 0.0, "state_share": 0.0, "total": 0.0}
                        ]
                    }
                }
                
                extracted_budget = process_upload_with_ai(budget_file, budget_prompt)
                
                if extracted_budget:
                    st.success("Data Extracted Successfully!")
                    st.json(extracted_budget)
                    
                    budget_dict = {}
                    for head in extracted_budget.get('heads', []):
                        budget_dict[head['head_name']] = {
                            'icar': head['icar_share'],
                            'state': head['state_share'],
                            'total': head['total']
                        }
                    
                    if extracted_budget.get('is_revision'):
                        data['revised_allocation'] = budget_dict
                    else:
                        data['allocation'] = budget_dict
                    
                    save_data(data, selected_fy)
                    st.toast("Budget Data Saved to GitHub!")

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
        st.write("Once received, upload the *'compotrollar grant relased latter'* (image_7.png) to activate funds for utilization.")
        
        comp_file = st.file_uploader("Upload Comptroller Office Order", type=['pdf', 'png', 'jpg'], key="comp_up")
        
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
