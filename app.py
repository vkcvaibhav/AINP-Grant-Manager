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
st.set_page_config(page_title="AINP Grant Manager", page_icon="🌾", layout="wide")

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
        
        col1, col2, col3 = st.columns(3)
        with col1:
            total_allocated = sum(v.get('total', 0) for k,v in data['revised_allocation'].items()) if data['revised_allocation'] else (sum(v.get('total', 0) for k,v in data['allocation'].items()) if data['allocation'] else 0)
            st.metric("Total Sanctioned Budget", f"₹{total_allocated:,}")
        with col2:
            total_received = sum(inst['amount'] for inst in data['installments'])
            st.metric("Total Funds Received (PFMS)", f"₹{total_received:,}")
        with col3:
            st.metric("Pending Received", f"₹{total_allocated - total_received:,}")

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
            
            # --- NEW ADDITION: SUMMARY TABLE ---
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
            # --- END NEW ADDITION ---

        else:
            st.info("No installments recorded yet.")

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
