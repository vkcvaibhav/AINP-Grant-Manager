import streamlit as st
import google.generativeai as genai
import sqlite3
import datetime
from PIL import Image
import io
import PyPDF2
from docx import Document as DocxReader
import urllib.request
import requests
import pandas as pd
import re
import os

st.set_page_config(page_title="સાદર નોંધ જનરેટર", layout="wide")

# Word Document Generation Imports
from docx import Document
from docx.shared import Mm, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Inches, Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Define global logo paths
NAU_LOGO = "logos/nau_logo.png"
ICAR_LOGO = "logos/icar_logo.png"

from github import Github


def get_secret_value(*names, default=""):
    """Read Streamlit secrets safely, with environment variables as a local fallback."""
    for name in names:
        value = None
        try:
            value = st.secrets.get(name)
        except Exception:
            value = None
        if not value:
            value = os.environ.get(name)
        if value:
            return str(value).strip()
    return default


DEFAULT_GITHUB_REPO = "vkcvaibhav/AINP-Grant-Manager"
DEFAULT_STATUTE_PDF_URL = "https://raw.githubusercontent.com/vkcvaibhav/Nodh-maker-/main/121_Statutes.pdf"
DEFAULT_SAMPLE_NONDH_URL = "https://raw.githubusercontent.com/vkcvaibhav/Nodh-maker-/main/sample_nondh.docx"

GEMINI_MAIN_MODEL = get_secret_value("GEMINI_MAIN_MODEL", "GEMINI_GENERATION_MODEL", default="gemini-3.1-pro-preview")
GEMINI_LIGHT_MODEL = get_secret_value("GEMINI_LIGHT_MODEL", "GEMINI_FLASH_LITE_MODEL", default="gemini-3.1-flash-lite")
GEMINI_HEAVY_MODEL = get_secret_value("GEMINI_HEAVY_MODEL", "GEMINI_PRO_MODEL", default="gemini-3.1-pro-preview")
STATUTE_PDF_URL = get_secret_value("STATUTE_PDF_URL", default=DEFAULT_STATUTE_PDF_URL)
SAMPLE_NONDH_URL = get_secret_value("SAMPLE_NONDH_URL", default=DEFAULT_SAMPLE_NONDH_URL)


def sanitize_filename(filename):
    name = os.path.basename(str(filename or "uploaded_file")).strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name or "uploaded_file"

# ==========================================
# GitHub Cloud Sync Engine
# ==========================================
def get_github_repo():
    token = get_secret_value("GITHUB_TOKEN")
    repo_name = get_secret_value("REPO_NAME", "GITHUB_REPO", default=DEFAULT_GITHUB_REPO)
    if not token:
        return None
    try:
        g = Github(token)
        return g.get_repo(repo_name)
    except Exception as e:
        print(f"GitHub Auth Error: {e}")
        return None

def pull_db_from_github():
    """Downloads the latest DB from GitHub when the app wakes up."""
    repo = get_github_repo()
    if not repo: return
    try:
        # Assuming we store it in a 'data' folder on GitHub
        file_content = repo.get_contents(f"data/{DB_FILE}")
        with open(DB_FILE, "wb") as f:
            f.write(file_content.decoded_content)
        print("Database successfully pulled from GitHub!")
    except Exception:
        print("No existing DB found on GitHub. A new one will be created.")

def push_db_to_github():
    """Uploads the local DB to GitHub after any changes."""
    repo = get_github_repo()
    if not repo: return
    try:
        with open(DB_FILE, "rb") as f:
            content = f.read()
        try:
            # Update existing file
            contents = repo.get_contents(f"data/{DB_FILE}")
            repo.update_file(contents.path, f"Auto-backup DB {datetime.datetime.now()}", content, contents.sha)
        except Exception:
            # Create new file if it doesn't exist
            repo.create_file(f"data/{DB_FILE}", "Initial DB backup", content)
    except Exception as e:
        print(f"Failed to push DB to GitHub: {e}")

def push_file_to_github(file_bytes, github_path):
    """Uploads documents/PDFs to the GitHub repo."""
    repo = get_github_repo()
    if not repo: return
    try:
        file_content = bytes(file_bytes)
        try:
            contents = repo.get_contents(github_path)
            repo.update_file(contents.path, f"Updated {github_path}", file_content, contents.sha)
        except Exception:
            repo.create_file(github_path, f"Uploaded {github_path}", file_content)
    except Exception as e:
        print(f"Failed to push file to GitHub: {e}")
def pull_file_from_github(github_path):
    """Downloads a missing document back from GitHub repository."""
    repo = get_github_repo()
    if not repo: return None
    try:
        file_content = repo.get_contents(github_path)
        return file_content.decoded_content
    except Exception as e:
        print(f"Failed to pull file from GitHub: {e}")
        return None

def load_vault_file_bytes(file_path):
    """Checks local storage first; if missing, pulls from GitHub and re-caches it."""
    if os.path.exists(file_path):
        with open(file_path, "rb") as f:
            return f.read()
    else:
        # File is missing locally due to cloud reset; pull from cloud storage
        github_path = file_path.replace("\\", "/")
        file_bytes = pull_file_from_github(github_path)
        if file_bytes:
            try:
                # Cache it locally again so subsequent calls are fast
                os.makedirs(os.path.dirname(file_path), exist_ok=True)
                with open(file_path, "wb") as f:
                    f.write(file_bytes)
            except Exception:
                pass
            return file_bytes
    return None
# ==========================================
# Database Setup for Archiving, Workflow & Digital Vault
# ==========================================
DB_FILE = "sadar_nondh_archive.db"

GUJARATI_DIGIT_TRANS = str.maketrans("૦૧૨૩૪૫૬૭૮૯", "0123456789")

def coerce_amount(value, default=0.0):
    """Convert DB/AI/UI currency values to a safe float."""
    try:
        default_value = float(default)
    except (TypeError, ValueError):
        default_value = 0.0

    if value is None:
        return default_value

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            if pd.isna(value):
                return default_value
        except Exception:
            pass
        return float(value)

    text = str(value).strip().translate(GUJARATI_DIGIT_TRANS)
    if not text:
        return default_value

    text = text.replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return default_value

    try:
        return float(match.group(0))
    except ValueError:
        return default_value

def format_amount(value, default=0.0):
    return f"{coerce_amount(value, default):.2f}"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # Table for Sadar Nondh
    c.execute('''CREATE TABLE IF NOT EXISTS archive 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  date TEXT, month TEXT, year TEXT, subject TEXT, content TEXT)''')
    
    # Table for Purchase Orders (Workflow Tracking) - Added nondh_id and payment_info
    c.execute('''CREATE TABLE IF NOT EXISTS purchase_orders 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  nondh_id INTEGER, vendor_name TEXT, out_no TEXT, date TEXT, amount REAL, status TEXT, payment_info TEXT)''')
                  
    # Table for Digital Vault (Uploaded PDFs/Images & Drafts) - Added nondh_id
    c.execute('''CREATE TABLE IF NOT EXISTS digital_vault 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  nondh_id INTEGER, file_name TEXT, file_path TEXT, upload_date TEXT, 
                  financial_year TEXT, month TEXT, doc_type TEXT, description TEXT)''')

    c.execute('''CREATE TABLE IF NOT EXISTS learning_memory
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  category TEXT, title TEXT, keywords TEXT, memory_text TEXT,
                  enabled INTEGER DEFAULT 1, created_at TEXT, updated_at TEXT)''')
                  
    # Safe Database Schema Upgrades for existing users
    try: c.execute("ALTER TABLE purchase_orders ADD COLUMN nondh_id INTEGER")
    except: pass
    try: c.execute("ALTER TABLE purchase_orders ADD COLUMN payment_info TEXT")
    except: pass
    try: c.execute("ALTER TABLE digital_vault ADD COLUMN nondh_id INTEGER")
    except: pass

    conn.commit()
    conn.close()

def save_to_db(subject, content):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    now = datetime.datetime.now()
    c.execute("INSERT INTO archive (date, month, year, subject, content) VALUES (?, ?, ?, ?, ?)", 
              (now.strftime("%d/%m/%Y"), now.strftime("%m"), now.strftime("%Y"), subject, content))
    nondh_id = c.lastrowid
    conn.commit()
    conn.close()
    push_db_to_github()
    return nondh_id

def save_po_to_db(nondh_id, vendor_name, out_no, date, amount):
    amount = coerce_amount(amount)
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO purchase_orders (nondh_id, vendor_name, out_no, date, amount, status) VALUES (?, ?, ?, ?, ?, 'Unfinished')", 
              (nondh_id, vendor_name, out_no, date, amount))
    po_id = c.lastrowid
    conn.commit()
    conn.close()
    push_db_to_github()
    return po_id

def get_unfinished_pos(statuses=('Unfinished',)):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # Create the correct number of placeholders (?) for our statuses
    placeholders = ','.join(['?'] * len(statuses))
    query = f"SELECT id, nondh_id, vendor_name, out_no, date, amount FROM purchase_orders WHERE status IN ({placeholders})"
    c.execute(query, statuses)
    data = c.fetchall()
    conn.close()
    return data

def mark_po_as_payment_generated(po_id):
    """Updates the status so it disappears from Tab 4 but stays in Tab 5/6"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("UPDATE purchase_orders SET status = 'Payment_Generated' WHERE id = ?", (po_id,))
    conn.commit()
    push_db_to_github()
    conn.close()

def mark_po_as_paid(po_id, payment_info=""):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("UPDATE purchase_orders SET status = 'Paid', payment_info = ? WHERE id = ?", (payment_info, po_id))
    conn.commit()
    push_db_to_github()
    conn.close()
    
def delete_po(po_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM purchase_orders WHERE id = ?", (po_id,))
    conn.commit()
    push_db_to_github()
    conn.close()

def get_archives(month, year, keyword=""):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    query = "SELECT id, date, subject, content FROM archive WHERE 1=1"
    params = []
    if year != "All":
        query += " AND year=?"
        params.append(year)
    if month != "All":
        query += " AND month=?"
        params.append(month)
    if keyword:
        query += " AND (subject LIKE ? OR content LIKE ?)"
        params.extend([f"%{keyword}%", f"%{keyword}%"])
    query += " ORDER BY id DESC"
    c.execute(query, tuple(params))
    data = c.fetchall()
    conn.close()
    return data

# --- Digital Vault Helpers ---
def get_financial_year(date_obj):
    if date_obj.month < 4: return f"{date_obj.year - 1}-{str(date_obj.year)[2:]}"
    else: return f"{date_obj.year}-{str(date_obj.year + 1)[2:]}"

def save_file_to_vault(file_bytes, original_name, doc_type, nondh_id=None, description="", upload_date=None):
    if upload_date is None: upload_date = datetime.date.today()
    fy = get_financial_year(upload_date)
    month_str = upload_date.strftime("%B")
    safe_original_name = sanitize_filename(original_name)
    
    safe_fy = fy.replace("-", "_")
    folder_path = os.path.join("digital_vault", safe_fy, doc_type.replace(" ", "_"))
    os.makedirs(folder_path, exist_ok=True)
    
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
    safe_name = f"{timestamp}_{safe_original_name}"
    file_path = os.path.join(folder_path, safe_name)
    
    # 1. Save locally
    with open(file_path, "wb") as f: f.write(file_bytes)
        
    # 2. Save metadata to database
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO digital_vault (nondh_id, file_name, file_path, upload_date, financial_year, month, doc_type, description) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
              (nondh_id, safe_original_name, file_path, upload_date.strftime("%Y-%m-%d"), fy, month_str, doc_type, description))
    conn.commit()
    conn.close()
    
    # 3. Sync SQLite DB changes to GitHub
    push_db_to_github()
    
    # FIX FIX FIX: Convert windows paths if any, and push the actual file to GitHub!
    github_path = file_path.replace("\\", "/")
    push_file_to_github(file_bytes, github_path)
    
def get_vault_files_by_nondh(nondh_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT file_name, file_path, upload_date, doc_type, description FROM digital_vault WHERE nondh_id = ? ORDER BY id ASC", (nondh_id,))
    data = c.fetchall()
    conn.close()
    return data

def get_vault_files(fy="All", doc_type="All", search_keyword=""):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # સુધારો: ડેટાબેઝમાંથી ડિલીટ કરવા માટે id (vault_id) ઉમેર્યો
    query = "SELECT id, nondh_id, file_name, file_path, upload_date, financial_year, month, doc_type, description FROM digital_vault WHERE 1=1"
    params = []
    
    if fy != "All":
        query += " AND financial_year=?"
        params.append(fy)
    if doc_type != "All":
        query += " AND doc_type=?"
        params.append(doc_type)
    if search_keyword:
        query += " AND (file_name LIKE ? OR description LIKE ?)"
        params.extend([f"%{search_keyword}%", f"%{search_keyword}%"])
        
    query += " ORDER BY upload_date DESC"
    c.execute(query, tuple(params))
    data = c.fetchall()
    conn.close()
    return data

def delete_vault_record(vault_id, nondh_id, file_path):
    """વોલ્ટમાંથી ફાઈલ કાઢે છે, અને જો Nondh જોડાયેલી હોય તો તેને પણ કાઢીને GitHub પર સિન્ક કરે છે."""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    
    # ૧. વોલ્ટ ડેટાબેઝમાંથી ડિલીટ
    c.execute("DELETE FROM digital_vault WHERE id = ?", (vault_id,))
    
    # ૨. જો Nondh ID હોય, તો આર્કાઇવ અને પેમેન્ટ ઓર્ડરમાંથી પણ આખી Nondh કાઢી નાખો
    if nondh_id:
        c.execute("DELETE FROM archive WHERE id = ?", (nondh_id,))
        c.execute("DELETE FROM purchase_orders WHERE nondh_id = ?", (nondh_id,))
        
    conn.commit()
    conn.close()
    
    # ૩. ડેટાબેઝ GitHub પર અપડેટ કરો (જેથી ત્યાંથી પણ ડિલીટ થઈ જાય)
    push_db_to_github()
    
    # ૪. લોકલ ફોલ્ડરમાંથી PDF/Word ફાઈલ કાઢી નાખો
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
        except:
            pass

# Check if we already pulled the DB in this session to avoid constant downloading
if "db_synced" not in st.session_state:
    pull_db_from_github()
    st.session_state.db_synced = True

init_db()
def get_recent_nondhs(days=30):
    """છેલ્લા 30 દિવસની સાદર નોંધ મેળવવા માટે"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT id, date, subject FROM archive ORDER BY id DESC LIMIT 50")
    records = c.fetchall()
    conn.close()
    
    recent_records = []
    cutoff_date = datetime.date.today() - datetime.timedelta(days=days)
    for row in records:
        try:
            # ડેટાબેઝમાં તારીખ DD/MM/YYYY ફોર્મેટમાં છે
            row_date = datetime.datetime.strptime(row[1], "%d/%m/%Y").date()
            if row_date >= cutoff_date:
                recent_records.append(row)
        except Exception:
            pass # જો તારીખમાં ભૂલ હોય તો સ્કીપ કરો
    return recent_records

def delete_nondh(nondh_id):
    """સાદર નોંધને ડેટાબેઝમાંથી કાયમ માટે ડિલીટ કરવા માટે"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM archive WHERE id = ?", (nondh_id,))
    conn.commit()
    conn.close()


LEARNING_MEMORY_CATEGORIES = [
    "General Preference",
    "Nondh Wording",
    "Statute Item",
    "Vendor Preference",
    "Bill/Payment",
    "Search Hint",
    "Other",
]


def extract_nondh_subject(content):
    for line in str(content or "").splitlines():
        if "વિષય:" in line:
            return line.replace("વિષય:", "", 1).strip() or "No Subject"
    return "No Subject"


def save_learning_memory(category, title, keywords, memory_text, enabled=True, memory_id=None):
    now = datetime.datetime.now().isoformat(timespec="seconds")
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    if memory_id:
        c.execute(
            """UPDATE learning_memory
               SET category = ?, title = ?, keywords = ?, memory_text = ?,
                   enabled = ?, updated_at = ?
               WHERE id = ?""",
            (category, title, keywords, memory_text, 1 if enabled else 0, now, memory_id),
        )
    else:
        c.execute(
            """INSERT INTO learning_memory
               (category, title, keywords, memory_text, enabled, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (category, title, keywords, memory_text, 1 if enabled else 0, now, now),
        )
    conn.commit()
    conn.close()
    push_db_to_github()


def get_learning_memories(include_disabled=False):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    query = "SELECT id, category, title, keywords, memory_text, enabled, created_at, updated_at FROM learning_memory"
    params = []
    if not include_disabled:
        query += " WHERE enabled = ?"
        params.append(1)
    query += " ORDER BY updated_at DESC, id DESC"
    c.execute(query, tuple(params))
    data = c.fetchall()
    conn.close()
    return data


def set_learning_memory_enabled(memory_id, enabled):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        "UPDATE learning_memory SET enabled = ?, updated_at = ? WHERE id = ?",
        (1 if enabled else 0, datetime.datetime.now().isoformat(timespec="seconds"), memory_id),
    )
    conn.commit()
    conn.close()
    push_db_to_github()


def delete_learning_memory(memory_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM learning_memory WHERE id = ?", (memory_id,))
    conn.commit()
    conn.close()
    push_db_to_github()


def get_matching_learning_memories(query_text, limit=6):
    query = str(query_text or "").strip().lower()
    query_terms = [
        term for term in re.split(r"[\s,;|/()]+", query.translate(GUJARATI_DIGIT_TRANS))
        if len(term) >= 3
    ]
    if not query_terms:
        return []

    matches = []
    for memory in get_learning_memories(include_disabled=False):
        memory_id, category, title, keywords, memory_text, enabled, created_at, updated_at = memory
        haystack = f"{category} {title} {keywords} {memory_text}".lower().translate(GUJARATI_DIGIT_TRANS)
        keyword_phrases = [k.strip().lower().translate(GUJARATI_DIGIT_TRANS) for k in re.split(r"[,;|]", keywords or "") if k.strip()]
        score = sum(2 for phrase in keyword_phrases if phrase and phrase in query)
        score += sum(1 for term in query_terms if term in haystack)
        if score:
            matches.append((score, memory))

    matches.sort(key=lambda item: (-item[0], item[1][0]))
    return [memory for _, memory in matches[:limit]]


def format_learning_memory_context(query_text):
    memories = get_matching_learning_memories(query_text)
    if not memories:
        return "No matching user-saved learning memories."

    lines = ["LEARNED MEMORY (user-approved preferences and corrections):"]
    for _, category, title, keywords, memory_text, enabled, created_at, updated_at in memories:
        clean_text = str(memory_text or "").strip()
        if len(clean_text) > 1200:
            clean_text = clean_text[:1200] + "..."
        lines.append(f"- [{category}] {title} | Keywords: {keywords}\n  {clean_text}")
    return "\n".join(lines)
init_db()

# ==========================================
# Permanent Attachments & Parsing (GitHub)
# ==========================================
@st.cache_data(ttl=3600) 
def load_permanent_context():
    statute_text = "Statute 121 Rules:\n"
    sample_text = "Sample Nondh Format:\n"
    try:
        r_pdf = requests.get(STATUTE_PDF_URL, timeout=25)
        if r_pdf.status_code == 200:
            f = io.BytesIO(r_pdf.content)
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                statute_text += (page.extract_text() or "") + "\n"
    except Exception: pass
    try:
        r_docx = requests.get(SAMPLE_NONDH_URL, timeout=25)
        if r_docx.status_code == 200:
            f = io.BytesIO(r_docx.content)
            doc = DocxReader(f)
            for para in doc.paragraphs: sample_text += para.text + "\n"
    except Exception: pass
    return statute_text, sample_text

def search_sample_nondh(keyword, month, year):
    try: _, sample_text = load_permanent_context()
    except Exception: return []
    if not sample_text: return []
    guj_to_eng = str.maketrans("૦૧૨૩૪૫૬૭૮૯", "0123456789")
    blocks = re.split(r'\n(?=તા\.\s*)', '\n' + sample_text)
    results = []
    for block in blocks:
        block = block.strip()
        if not block or "સાદર નોંધ" not in block: continue
        eng_block = block.translate(guj_to_eng)
        if keyword and keyword.lower() not in block.lower() and keyword.lower() not in eng_block.lower(): continue
        date_match = re.search(r'તા\.\s*([\d/ \-]+)', eng_block)
        date_str = date_match.group(1).strip() if date_match else "Unknown"
        if year != "All" and year not in date_str: continue
        if month != "All":
            if f"/{month}/" not in date_str and f"-{month}-" not in date_str and f"{month}/" not in date_str: continue
        sub_match = re.search(r'વિષય:\s*([^\n]+)', block)
        subject_str = sub_match.group(1).strip() if sub_match else "Historical Reference"
        orig_date_match = re.search(r'તા\.\s*([^\n]+)', block)
        display_date = orig_date_match.group(1).strip() if orig_date_match else "Unknown"
        results.append((None, display_date + " [Old Sample Ref]", subject_str, block))
    return results

def parse_markdown_to_parts(text):
    lines = text.split('\n')
    pre_text, table_lines, post_text = [], [], []
    in_table, table_done = False, False
    for line in lines:
        if line.strip().startswith('|'):
            in_table = True
            table_lines.append(line)
        else:
            if in_table:
                table_done = True
                in_table = False
            if not table_done: pre_text.append(line)
            else: post_text.append(line)
    df = pd.DataFrame()
    if table_lines:
        parts = table_lines[0].split('|')
        if len(parts) > 2:
            header = [x.strip() for x in parts[1:-1]]
            data = []
            for line in table_lines[2:]: 
                row_parts = line.split('|')
                if len(row_parts) > 2:
                    data.append([x.strip() for x in row_parts[1:-1]])
            if data:
                df = pd.DataFrame(data, columns=header)
                if 'Details' in df.columns: df = df[~df['Details'].astype(str).str.contains('Grand Total', case=False, na=False)]
    return "\n".join(pre_text), df, "\n".join(post_text)

def df_to_markdown_with_total(df):
    if df.empty: return ""
    grand_total = 0
    if 'Total Price' in df.columns:
        grand_total = pd.to_numeric(df['Total Price'], errors='coerce').fillna(0).sum()
    markdown = "|" + "|".join(df.columns) + "|\n|" + "|".join(["---"] * len(df.columns)) + "|\n"
    for _, row in df.iterrows():
        clean_row = [str(int(x)) if isinstance(x, float) and x.is_integer() else str(x) for x in row]
        markdown += "|" + "|".join(clean_row) + "|\n"
    if len(df.columns) >= 5:
        total_row = [""] * len(df.columns)
        total_row[1] = "**Grand Total**"
        total_row[-1] = f"**{grand_total:.2f}**"
        markdown += "|" + "|".join(total_row) + "|\n"
    return markdown

def add_bottom_border(paragraph, size='24'):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), size)
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)
    pPr.append(pBdr)

# ==========================================
# Word Generators
# ==========================================
def create_docx(content):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.left_margin, section.right_margin = Mm(42), Mm(15)
    section.top_margin, section.bottom_margin = Mm(15), Mm(15)
    
    font = doc.styles['Normal'].font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Shruti')
    font._element.append(rFonts)

    lines = content.split('\n')
    table_data = []
    in_table = False
    sig_buffer = []

    def flush_signatures():
        if sig_buffer:
            doc.add_paragraph().paragraph_format.space_before = Pt(20)
            sig_table = doc.add_table(rows=1, cols=3)
            for c in sig_table.columns: 
                for cell in c.cells: cell.width = Mm(51)
            for i, sig in enumerate(sig_buffer):
                if i < 3:
                    p = sig_table.cell(0, i).paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    parts = sig.split(',')
                    for j, part in enumerate(parts):
                        run = p.add_run(part.strip())
                        if j < len(parts) - 1: run.add_break()
            sig_buffer.clear()

    def build_and_format_table(data):
        num_cols = len(data[0])
        table = doc.add_table(rows=len(data), cols=num_cols)
        table.style = 'Table Grid'
        
        # Turn off autofit to force Word to respect your explicit column widths
        table.autofit = False 
        
        # Calculated to fit exactly within your page margins (153mm usable width)
        # Sr. No (10mm) + Details (75mm) + Qty (17mm) + Avail (17mm) + Unit Price (17mm) + Total (17mm) = 153mm
        widths = [Mm(10), Mm(75), Mm(17), Mm(17), Mm(17), Mm(17)]
        
        for row_idx, row_data in enumerate(data):
            row_cells = table.rows[row_idx].cells
            for col_idx, cell_text in enumerate(row_data):
                cell = row_cells[col_idx]
                
                # Apply width to every cell individually to enforce it in MS Word
                if num_cols == 6: 
                    cell.width = widths[col_idx]
                    
                is_bold = (row_idx == 0) or ('**' in cell_text)
                cell.text = cell_text.replace('**', '')
                p = cell.paragraphs[0]
                
                # Center align all columns EXCEPT the Details column (col_idx == 1)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if (col_idx == 1 and row_idx > 0) else WD_ALIGN_PARAGRAPH.CENTER
                
                if is_bold:
                    for run in p.runs: run.bold = True
        doc.add_paragraph()

    for line in lines:
        line_stripped = line.strip()
        if not line_stripped: continue
        if line_stripped.startswith('|'):
            in_table = True
            if not line_stripped.replace('|', '').replace('-', '').replace(' ', ''): continue
            parts = line_stripped.split('|')
            if len(parts) > 2: table_data.append([cell.strip() for cell in parts[1:-1]])
        else:
            if in_table:
                if table_data: build_and_format_table(table_data)
                table_data, in_table = [], False

            if line_stripped.startswith("તા.") or line_stripped.startswith("સ્થળ:"):
                flush_signatures()
                p = doc.add_paragraph(line_stripped)
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.space_after = Pt(0) 
            elif "સાદર નોંધ" in line_stripped:
                flush_signatures()
                p = doc.add_paragraph()
                p.add_run(line_stripped).bold = True
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pPr = p._p.get_or_add_pPr()
                pBdr = OxmlElement('w:pBdr')
                bottom = OxmlElement('w:bottom')
                bottom.set(qn('w:val'), 'dotted')
                bottom.set(qn('w:sz'), '4')
                bottom.set(qn('w:space'), '1')
                bottom.set(qn('w:color'), '000000')
                pBdr.append(bottom)
                pPr.append(pBdr)
            elif line_stripped.startswith("વિષય:"):
                flush_signatures()
                p = doc.add_paragraph()
                p.add_run(line_stripped).bold = True
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            elif any(role in line_stripped for role in ["અધિકારી", "ઈન્ચાર્જ", "પ્રાધ્યાપક", "વડા"]) and not any(r in line_stripped for r in ["આચાર્ય", "ડીનશ્રી"]):
                sig_buffer.append(line_stripped)
            elif any(role in line_stripped for role in ["આચાર્ય", "ડીનશ્રી", "મહાવિધાયલય", "ન.કૃ.યુ"]):
                flush_signatures()
                doc.add_paragraph().paragraph_format.space_before = Pt(30)
                p_table = doc.add_table(rows=1, cols=2)
                p_table.columns[0].width, p_table.columns[1].width = Mm(79), Mm(74) 
                formatted_line = "\n".join([p.strip() for p in line_stripped.split(",")])
                p = p_table.cell(0, 1).paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(formatted_line)
            else:
                flush_signatures()
                doc.add_paragraph(line_stripped).alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 

    flush_signatures()
    if in_table and table_data: build_and_format_table(table_data)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_purchase_order_docx(vendor_name, vendor_address, out_no, po_date, df_items):
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(0.4), Inches(0.5)
        section.left_margin, section.right_margin = Inches(0.8), Inches(0.8)
        
    style = doc.styles['Normal']
    style.font.size = Pt(12)
    style.paragraph_format.space_after, style.paragraph_format.space_before = Pt(0), Pt(0)
        
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    table.columns[0].width, table.columns[1].width, table.columns[2].width = Inches(1.8), Inches(3.6), Inches(1.4)
    
    if 'NAU_LOGO' in globals() and NAU_LOGO and os.path.exists(NAU_LOGO):
        cell_left = table.cell(0, 0)
        cell_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_left = cell_left.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_left.add_run().add_picture(NAU_LOGO, width=Inches(1.8))
        
    p_center = table.cell(0, 1).paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_center.paragraph_format.line_spacing = 0.85
    r1 = p_center.add_run("કીટકશાસ્ત્ર વિભાગ\n")
    r1.bold = True
    r1.font.size = Pt(22)
    r2 = p_center.add_run("ન. મ. કૃષિ મહાવિદ્યાલય\nનવસારી કૃષિ યુનિવર્સિટી\nનવસારી- ૩૯૬ ૪૫૦ (ગુજરાત)")
    r2.bold, r2.font.size = True, Pt(14)
    
    if 'ICAR_LOGO' in globals() and ICAR_LOGO and os.path.exists(ICAR_LOGO):
        p_right = table.cell(0, 2).paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_right.add_run().add_picture(ICAR_LOGO, width=Inches(1.5))
        
    p_thick1 = doc.add_paragraph()
    p_thick1.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_thick1.paragraph_format.line_spacing = Pt(1)
    p_thick1.add_run().font.size = Pt(1) 
    add_bottom_border(p_thick1, size='24')
    
    table2 = doc.add_table(rows=1, cols=2)
    table2.columns[0].width, table2.columns[1].width = Inches(3.4), Inches(3.4)

    table2.cell(0,0).paragraphs[0].add_run("ડૉ. જે. જે. પસ્તાગીયા\nપ્રાધ્યાપક અને વડા")
    p2 = table2.cell(0,1).paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run("મોબાઇલ: +૯૧ ૯૪૨૭૮ ૬૭૯૨૫\nઇમેલ: headentonau@gmail.com")
    
    for cell in table2.rows[0].cells:
        for p in cell.paragraphs: p.paragraph_format.line_spacing = 0.8
            
    p_thick2 = doc.add_paragraph()
    p_thick2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_thick2.paragraph_format.line_spacing = Pt(1)
    p_thick2.add_run().font.size = Pt(1) 
    add_bottom_border(p_thick2, size='24')
    
    letter_year = po_date.split('/')[-1] if '/' in po_date else (po_date.split('.')[-1] if '.' in po_date else "૨૦૨૬")
    
    table3 = doc.add_table(rows=1, cols=2)
    table3.autofit = False
    
    # 6.9 Inches ની અંદર રહે તે રીતે પહોળાઈ સેટ કરી છે
    table3.columns[0].width = Inches(5.0)  
    table3.columns[1].width = Inches(1.9)  
    table3.cell(0, 0).width = Inches(5.0)
    table3.cell(0, 1).width = Inches(1.9)
    
    p_out = table3.cell(0,0).paragraphs[0]
    p_out.paragraph_format.space_before = Pt(0)
    p_out.paragraph_format.space_after = Pt(0)
    p_out.add_run(f"જા.નં. એસીએન/એન્ટો/એઆઈએનપી-એએ/{out_no}/{letter_year}, નવસારી")
    
    p_date = table3.cell(0,1).paragraphs[0]
    p_date.paragraph_format.space_before = Pt(0)
    p_date.paragraph_format.space_after = Pt(0)
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_date.add_run(f"તારીખ: {po_date}")
            
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    p_to = doc.add_paragraph()
    p_to.add_run("પ્રતિ,").bold = True
    doc.add_paragraph(vendor_name).runs[0].bold = True
    doc.add_paragraph(vendor_address)

    doc.add_paragraph()
    
    p_subj = doc.add_paragraph()
    p_subj.add_run("        વિષય: ખરીદી હુકમ").bold = True
    
    doc.add_paragraph("        જય ભારત સહ ઉપરોક્ત વિષય અન્વયે જણાવવાનું કે, અત્રેના કીટકશાસ્ત્ર વિભાગ ખાતે નિચેની વસ્તુઓ બિલ સહિત રજુ કરવા વિનંતી.").alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_paragraph() 

    # --- વસ્તુઓના લિસ્ટ વાળા ટેબલની ગોઠવણ ---
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    table.autofit = False  # <--- આ ઉમેરવું ખૂબ જરૂરી છે
    
    # કુલ 6.9 Inches થવું જોઈએ (0.5 + 3.4 + 1.0 + 1.0 + 1.0 = 6.9)
    widths = [Inches(0.5), Inches(3.4), Inches(1.0), Inches(1.0), Inches(1.0)]
    
    headers = ["અ.નં.", "વસ્તુઓના નામ", "જથ્થો", "ભાવ પ્રતિ નંગ", "કુલ રકમ"]
    for i, ht in enumerate(headers):
        table.columns[i].width = widths[i]
        table.cell(0, i).width = widths[i] # હેડર સેલની પહોળાઈ લોક કરો
        table.cell(0, i).text = ht
        p = table.cell(0, i).paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        table.cell(0, i).vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    total_amount = 0.0
    for index, row in df_items.iterrows():
        row_cells = table.add_row().cells
        
        # દરેક નવી રો ની પહોળાઈ લોક કરો
        for i in range(5): 
            row_cells[i].width = widths[i]

        row_cells[0].text = str(index + 1)
        row_cells[1].text = str(row.get('Details', ''))
        row_cells[2].text = str(row.get('Required Quantity', '')) + " " + str(row.get('Available Pkt/Unit', ''))
        
        unit_price = pd.to_numeric(row.get('Unit/Pkt Price', 0), errors='coerce')
        total_price = pd.to_numeric(row.get('Total Price', 0), errors='coerce')
        row_cells[3].text, row_cells[4].text = f"{unit_price:.2f}", f"{total_price:.2f}"
        total_amount += total_price
        
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    total_row = table.add_row().cells
    for i in range(5): 
        total_row[i].width = widths[i]

    total_row[3].text = "Total"
    total_row[3].paragraphs[0].runs[0].bold = True
    total_row[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    total_row[4].text = f"{total_amount:.2f}"
    total_row[4].paragraphs[0].runs[0].bold = True
    total_row[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    doc.add_paragraph()

    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.add_run("પ્રાધ્યાપક અને વડા\nકીટકશાસ્ત્ર વિભાગ").bold = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
    
# --- Bill Payment Form ---
def create_bill_payment_form(budget_head, bill_no, bill_date, party_name, amount, amount_words):
    amount = coerce_amount(amount)
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(0.8), Inches(0.8)
        section.left_margin, section.right_margin = Inches(1.0), Inches(1.0)
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(15)
    style.paragraph_format.space_after = Pt(0)
    
    p_header = doc.add_paragraph()
    p_header.add_run("No. ACN/ENTO/BILL/       /202\n").bold = True
    r_right = p_header.add_run(f"NAVSARI-396450, Date: {datetime.date.today().strftime('%d/%m/%Y')}").bold = True
    p_header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    
    doc.add_paragraph("To,").bold = True
    doc.add_paragraph("The Principal and Dean,").bold = True
    doc.add_paragraph("N.M. College of Agriculture,").bold = True
    doc.add_paragraph("Navsari").paragraph_format.space_after = Pt(12)
    
    p_sub = doc.add_paragraph()
    p_sub.add_run("Sub: Submission of bill(s) for payment............").bold = True
    p_sub.paragraph_format.space_after = Pt(12)
    
    p_body = doc.add_paragraph("With reference to the above subject, I am submitting herewith the following bill(s) for making payment to the respective party and debit the same in Budget Head No- ")
    p_body.add_run(budget_head).bold = True
    p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_paragraph()
    
    table = doc.add_table(rows=2, cols=4)
    table.style = 'Table Grid'
    
    hdr0 = table.rows[0].cells
    hdr1 = table.rows[1].cells
    
    hdr0[0].text = "Sr."
    hdr1[0].text = "No."
    hdr0[1].text = "No. of Bill/Date"
    hdr0[2].text = "Name of the Party"
    hdr0[3].text = "Amount"
    hdr1[3].text = "Rs.              Ps."
    
    for c in [0, 1, 2, 3]:
        if c in [1, 2]: table.cell(0, c).merge(table.cell(1, c))
        for r in [0, 1]:
            p = table.cell(r, c).paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs: run.bold = True
            table.cell(r, c).vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    table.columns[0].width = Inches(0.5)
    table.columns[1].width = Inches(1.5)
    table.columns[2].width = Inches(12.5)
    table.columns[3].width = Inches(1.5)

    row = table.add_row().cells
    row[0].text = "1"
    row[1].text = f"No: {bill_no}\nDt: {bill_date}"
    row[2].text = party_name
    row[3].text = format_amount(amount)
    
    for i in range(4): 
        row[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    total_row = table.add_row().cells
    table.cell(3,0).merge(table.cell(3,2))
    p_tot = total_row[0].paragraphs[0]
    p_tot.add_run("Total:   ").bold = True
    p_tot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    total_row[3].text = format_amount(amount)
    total_row[3].paragraphs[0].runs[0].bold = True
    total_row[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    word_row = table.add_row().cells
    table.cell(4,0).merge(table.cell(4,3))
    p_word = word_row[0].paragraphs[0]
    p_word.add_run("In words: ").bold = True
    p_word.add_run(f"Rupees {amount_words} Only.")
    
    party_row = table.add_row().cells
    table.cell(5,0).merge(table.cell(5,3))
    p_party = party_row[0].paragraphs[0]
    p_party.add_run("Name of Party for Payment: ").bold = True
    p_party.add_run(party_name)
         
    doc.add_paragraph("Encl: Cash/Credit Bill in original")
    doc.add_paragraph(f"No. {bill_no} with entry").paragraph_format.space_after = Pt(20)
    
    doc.add_paragraph(f"Copy F.W.C.S. to M/S: {party_name}").paragraph_format.space_after = Pt(30)
    
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.add_run("Professor and Head\nDepartment of Entomology\nN.M.C.A., N.A.U., Navsari").bold = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- PERFECT PDF REPLICA: Bill Pasting Form ---
def set_cell_border(cell, **kwargs):
    """Helper function to draw borders around specific table cells."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

import io
from docx import Document
from docx.shared import Inches, Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_border(cell, **kwargs):
    """Helper function to draw borders around specific table cells."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def create_bill_pasting_form(budget_head, grant_year, party_name, amount, amount_in_guj_words, reg_type, reg_page_no, bill_reg_date, bill_reg_page_no, bill_reg_sr_no, item_no, approval_no, approval_date):
    amount = coerce_amount(amount)
    doc = Document()
    
    # --- Helper to convert English digits to Gujarati digits ---
    def eng_to_guj(text):
        if not text: return ""
        return str(text).translate(str.maketrans("0123456789", "૦૧૨૩૪૫૬૭૮૯"))

    # Strict A4 Margins
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.2)
    section.bottom_margin = Inches(0.2)
    section.gutter = Inches(0)
    
    # Set base fonts
    style = doc.styles['Normal']
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Shruti')
    style.font._element.append(rFonts)

    # --- PAGE 1 ---
    header_table = doc.add_table(rows=1, cols=3)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.columns[0].width = Inches(2.0)
    header_table.columns[1].width = Inches(2.7) 
    header_table.columns[2].width = Inches(2.0)
    
    cell_left = header_table.cell(0, 0)
    p_office = cell_left.paragraphs[0]
    p_office.paragraph_format.space_before = Pt(12)
    run_office = p_office.add_run("Office No. 303")
    run_office.bold = True
    run_office.font.size = Pt(20) 
    p_office.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_border(cell_left, top={"sz": 12, "val": "single", "color": "000000"}, bottom={"sz": 12, "val": "single", "color": "000000"}, left={"sz": 12, "val": "single", "color": "000000"}, right={"sz": 12, "val": "single", "color": "000000"})
    
    cell_right = header_table.cell(0, 2)
    p_v1 = cell_right.paragraphs[0]
    run_v1 = p_v1.add_run("Voucher No:-....................")
    run_v1.font.size = Pt(15) 
    p_v1.paragraph_format.space_before = Pt(6) 
    p_v1.paragraph_format.space_after = Pt(8)  
    
    p_v2 = cell_right.add_paragraph()
    run_v2 = p_v2.add_run("          Date:-....................")
    run_v2.font.size = Pt(15) 
    p_v2.paragraph_format.space_after = Pt(0)
    set_cell_border(cell_right, top={"sz": 12, "val": "single", "color": "000000"}, bottom={"sz": 12, "val": "single", "color": "000000"}, left={"sz": 12, "val": "single", "color": "000000"}, right={"sz": 12, "val": "single", "color": "000000"})
        
    p_col = doc.add_paragraph()
    p_col.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_col.paragraph_format.space_before = Pt(0) 
    p_col.paragraph_format.space_after = Pt(0)
    run_col = p_col.add_run("N. M. COLLEGE OF AGRICULTURE")
    run_col.bold = True
    run_col.font.size = Pt(22) 
    
    p_uni = doc.add_paragraph()
    p_uni.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_uni = p_uni.add_run("Navsari Agricultural University, Navsari.-396450")
    run_uni.bold = True
    run_uni.font.size = Pt(20) 
    p_uni.paragraph_format.space_after = Pt(10)
    
    line_table = doc.add_table(rows=1, cols=1)
    line_table.columns[0].width = Inches(6.7)
    set_cell_border(line_table.cell(0, 0), top={"sz": 24, "val": "single", "color": "000000"})
    
    for _ in range(15): doc.add_paragraph()

    line_table_2 = doc.add_table(rows=1, cols=1)
    line_table_2.columns[0].width = Inches(6.7)
    set_cell_border(line_table_2.cell(0, 0), top={"sz": 24, "val": "single", "color": "000000"})
    
    p_note = doc.add_paragraph()
    run_note = p_note.add_run("Note:-")
    run_note.bold = True
    run_note.font.size = Pt(15) 
    p_note.paragraph_format.space_after = Pt(6)
    
    def add_bullet(num, text, size=15):
        p = doc.add_paragraph()
        run_num = p.add_run(num + "\t")
        run_num.font.size = Pt(15) 
        run_text = p.add_run(text)
        run_text.font.size = Pt(15) 
        p.paragraph_format.left_indent = Inches(0.5)
        p.paragraph_format.first_line_indent = Inches(-0.5)
        p.paragraph_format.space_after = Pt(2)
        
    add_bullet("1.", "Quotation from at least three parties for purchase above Rs. 1000/- should be obtained.", size=10)
    add_bullet("2.", "Purchase from authorized details and manufactures be certified on bill no other quotation were available due of purchase from manufacture or authorized dealers.", size=10)
    add_bullet("3.", "A special previous sanction of V.C of campus, Navsari should invariably be obtained for dead stock of other valuable articles before the purchase is made.", size=10)
    add_bullet("4.", "Purchase is made in the interested of University work.", size=10)
    
    doc.add_page_break()

    # --- PAGE 2 ---
    def add_p2_header_row(cell, text, size=10):
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.left_indent = Inches(0) 
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        run.font.size = Pt(size)
        
    top_table = doc.add_table(rows=4, cols=3)
    top_table.autofit = False
    top_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        
    for row in top_table.rows:
        row.cells[0].width = Inches(2.0)
        row.cells[1].width = Inches(0.3)
        row.cells[2].width = Inches(4.9)

    add_p2_header_row(top_table.cell(0,0), "બજેટ સદર")
    add_p2_header_row(top_table.cell(0,1), ":-")
    
    p_0_2 = top_table.cell(0,2).paragraphs[0]
    p_0_2.paragraph_format.space_before = Pt(0)
    p_0_2.paragraph_format.space_after = Pt(0)
    p_0_2.paragraph_format.left_indent = Inches(0)
    run_bh = p_0_2.add_run(f"{budget_head} \t\t")
    run_bh.font.size = Pt(10)
    run_exp = p_0_2.add_run("EXP. CODE NO._________")
    run_exp.font.size = Pt(10)
    run_exp.bold = True

    add_p2_header_row(top_table.cell(1,0), "ફાળવેલ ગ્રાન્ટ વર્ષ: ૨૦  - ૨૦")
    add_p2_header_row(top_table.cell(1,1), ":-")
    add_p2_header_row(top_table.cell(1,2), f"{grant_year}" if grant_year else "                     ")
    
    add_p2_header_row(top_table.cell(2,0), "બીલની કુલ રકમ")
    add_p2_header_row(top_table.cell(2,1), ":-")
    add_p2_header_row(top_table.cell(2,2), format_amount(amount)) 
    
    add_p2_header_row(top_table.cell(3,0), "ચુકવણું કરવામાં આવનાર પાર્ટીનું નામ (અંગ્રેજી કેપીટલ લેટર)")
    add_p2_header_row(top_table.cell(3,1), ":-")
    add_p2_header_row(top_table.cell(3,2), f"{party_name}")
    
    p_cert = doc.add_paragraph()
    p_cert.paragraph_format.space_before = Pt(6) 
    p_cert.paragraph_format.space_after = Pt(4)  
    run_cert = p_cert.add_run(":: પ્રમાણપત્ર ::")
    run_cert.bold = True
    run_cert.font.size = Pt(14)
    p_cert.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    table = doc.add_table(rows=9, cols=2)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
        
    for row in table.rows:
        row.cells[0].width = Inches(0.4)
        row.cells[1].width = Inches(6.8)
    
    def add_row(idx, no, text, size=11):
        p_no = table.rows[idx].cells[0].paragraphs[0]
        p_no.paragraph_format.left_indent = Inches(0) 
        p_no.paragraph_format.space_before = Pt(0)
        p_no.paragraph_format.space_after = Pt(0)
        run_no = p_no.add_run(no)
        run_no.font.size = Pt(9)
        
        p_text = table.rows[idx].cells[1].paragraphs[0]
        p_text.paragraph_format.left_indent = Inches(0)
        p_text.paragraph_format.space_before = Pt(0)
        p_text.paragraph_format.space_after = Pt(2) 
        p_text.paragraph_format.line_spacing = 1.0  
        p_text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # નોંધ: અહીથી જૂનો `run_text = p_text.add_run(text)` વાળો કોડ કાઢી નાખ્યો છે.
        
        # --- NEW: ** વચ્ચેના લખાણને Bold કરવા માટેનું સ્માર્ટ લોજીક ---
        parts = str(text).split("**")
        for i, part in enumerate(parts):
            if part:
                run_text = p_text.add_run(part)
                run_text.font.size = Pt(9)
                # એકી સંખ્યા વાળા ભાગ (અર્થાત ** ની વચ્ચે વાળા) Bold થશે
                if i % 2 != 0:  
                    run_text.bold = True

    # --- Convert inputs to Gujarati digits ---
    guj_amount = eng_to_guj(f"{amount:.2f}")
    guj_reg_page = eng_to_guj(reg_page_no)
    guj_bill_reg_page = eng_to_guj(bill_reg_page_no)
    guj_bill_reg_sr = eng_to_guj(bill_reg_sr_no)
    guj_bill_date = eng_to_guj(bill_reg_date)
    
    # New: Convert Approval Details to Gujarati
    guj_item_no = eng_to_guj(item_no) if item_no else "____________"
    guj_app_no = eng_to_guj(approval_no) if approval_no else "_________________________________________"
    guj_app_date = eng_to_guj(approval_date) if approval_date else "______/______/_________"

    # --- Register Setup ---
    blanks = {
        "સ્ટોર રોજમેળ": "____________",
        "ચીજવસ્તુ વપરાશ (કન્ઝયુમેબલ)": "____________",
        "ડેડસ્ટોક": "____________",
        "ટેલીફોન": "____________",
        "સ્ટેમ્પ": "____________",
        "સ્ટેશનરી": "____________",
        "પરચુરણ માલ સામાન": "____________",
        "રીપેરીંગ": "____________"
    }

    if guj_reg_page:
        if reg_type in blanks:
            blanks[reg_type] = f" {guj_reg_page} "

    # --- Dynamic Certificates ---
    cert_1 = (f"આ બીલમાં જણાવેલ વસ્તુ ખરીદવાની/રીપેરીંગના ખર્ચની મંજુરી આપવાની સત્તા ગુજરાત રાજય કૃષિ યુનિવર્સિટીઓ "
              f"(સતા સોપણી) નિયમ-૨૦૧૧ ના સ્ટેચ્યુટ નં. ૧૨૧ ની આઇટમ નં {guj_item_no} મુજબ એનાયત થયેલ સત્તા પ્રમાણે "
              f"હેડ ઓફિસ/હેડ ઓફ યુનિટ/યુનિ. ઓફિસર્સ/માન. કુલપતિશ્રીની મંજુરી નં: {guj_app_no} . "
              f"તારીખ: {guj_app_date} થી મંજુરી મળેલ છે. હુકમની નકલ સામેલ છે.")

    cert_2 = "આ બીલમાં જણાવેલ ખર્ચ આ વિભાગની આઇ.સી.એ.આર. યોજના બજેટ સદર ૩૦૩/૨૦૯૨ માં સમાવેશ કરવામાં આવેલ છે."
    
    cert_3 = (
        f"બીલમાં દર્શાવેલ માલની ખરીદી બજાર ભાવ તપાસી ભાવો મેળવી સૌથી ઓછા ભાવ મુજબ છે અને સારી સ્થિતિમાં મળેલ છે. "
        f"જે કચેરીના સ્ટોર રોજમેળ રજી પાના નં. {blanks['સ્ટોર રોજમેળ']}../ચીજવસ્તુ વપરાશ (કન્ઝયુમેબલ) "
        f"રજી. પાના નં. {blanks['ચીજવસ્તુ વપરાશ (કન્ઝયુમેબલ)']} ડેડસ્ટોક રજી. નં.... {blanks['ડેડસ્ટોક']} / ટેલીફોન રજી. પાના "
        f"નં {blanks['ટેલીફોન']} / સ્ટેમ્પ રજી. પાના નં {blanks['સ્ટેમ્પ']} / સ્ટેશનરી રજી. પાના નં. "
        f"{blanks['સ્ટેશનરી']} .રજીસ્ટરનાં ____________ / પરચુરણ માલ સામાન /.... {blanks['પરચુરણ માલ સામાન']}../ રીપેરીંગ "
        f"રજી. પાના નં.. {blanks['રીપેરીંગ']} નાં રોજ જમા કરવામાં આવેલ છે."
    )
    
    cert_4 = f"સદર બીલમાં દર્શાવવામાં આવેલ ખર્ચ સેલ્સ ટેક્ષ / એડી.ટેક્ષ / એકસાઇડયુટી / સેન્ટ્રલ ટેક્ષ વિગેરે પાર્ટીના માન્ય થયેલ ભાવ મુજબ ચકાસણી કરવામાં આવેલ છે. અને તે મુજબ પાર્ટીના બીલમાં દર્શાવ્યા મુજબની રકમ રૂ. {guj_amount}/- (અંકે રૂ. {amount_in_guj_words}) પુરા ચુકવવા ભલામણ કરવામાં આવે છે."
    cert_8 = f"તા. {guj_bill_date} સદર બીલની નોંધ કચેરી ખાતેના બીલ રજી પાના નં. {guj_bill_reg_page} અનુ.નં. {guj_bill_reg_sr} કરવામાં આવેલ છે."

    add_row(0, "૧.", cert_1, size=11)
    add_row(1, "૨.", cert_2, size=11)
    add_row(2, "૩.", cert_3, size=11)
    add_row(3, "૪.", cert_4, size=11)
    add_row(4, "૫.", "બીલમાં દર્શાવેલ મુજબ વાહન નં.... ____________ .ની રીપેરીંગ કામગીરી સંતોષકારક થયેલ છે જેની નોંધ હિસ્ટ્રીશીટ રજી. પાના નં....... ____________ .../ રીપેરીંગ રજી. પાના નં. ____________ થી કરેલ છે. જે ચાલુ નાણાંકીય વર્ષ દરમ્યાન આ બીલ સહીત કુલ ખર્ચ રૂ... ____________./- (અંકે ૩....................................) થયેલ છે.", size=11)
    add_row(5, "૬.", "બીલમાં દર્શાવેલ પેટ્રોલ / ડીઝલ / ઓઇલ વગેરે વાહન નં. ____________ .માટે ખરીદ કરવામાં આવેલ છે જે લોગબુક ભાગ નં. ____________ પાના નં. ____________ થી જમાં કરવામાં આવેલ છે.", size=11)
    add_row(6, "૭.", "તા. ____________ બીલમાં દર્શાવેલ વાહન નં. ____________ ના રીપેરીંગ કામ કરતી વખતે પરત આવેલ જુના સ્પેર પાર્ટસ મેળવીને આ કચેરીનાં રદ્દ રજી. પાના નં. ____________ ના રોજ જમાં લીધેલ છે.", size=11)
    add_row(7, "૮.", cert_8, size=11)
    add_row(8, "૯.", "સંશોધન નિયામકશ્રીના ૩૦/૧૦/૨૦૨૧ના પરિપત્રનો અમલ કરેલ છે.", size=11)
    
    p_special_note = doc.add_paragraph()
    p_special_note.paragraph_format.space_before = Pt(0)
    p_special_note.paragraph_format.space_after = Pt(2) 
    p_special_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_special_note = p_special_note.add_run("સદરહું ખર્ચ કચેરીની અગત્યની કામગીરીને ધ્યાને લઇ તેમજ યુનિવર્સિટીનાં હિતાર્થે કરવામાં આવેલ છે.")
    run_special_note.font.size = Pt(11)
    
    today_guj = eng_to_guj(datetime.date.today().strftime('%d/%m/%Y'))
    p_loc = doc.add_paragraph()
    p_loc.paragraph_format.space_before = Pt(2) 
    p_loc.paragraph_format.space_after = Pt(12)
    run_loc = p_loc.add_run(f"સ્થળ : નવસારી\nતારીખ : {today_guj}")
    run_loc.font.size = Pt(12)
    
    # --- સુધારેલો ભાગ: ૨ કોલમને બદલે ૩ સહીના કોલમ (3 Signature Columns) ---
    table_sig = doc.add_table(rows=1, cols=3)
    table_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # કુલ લિમિટ (6.8 Inches) ને ૩ સરખા ભાગમાં વહેંચી દીધી (આશરે 2.26 ઇંચ દરેક)
    col_widths = [Inches(2.26), Inches(2.26), Inches(2.26)]
    for i, cell in enumerate(table_sig.rows[0].cells): 
        cell.width = col_widths[i]
    
    # કોલમ ૧: ખેતીવાડી અધિકારી (Left Alignment)
    p_s1 = table_sig.cell(0, 0).paragraphs[0]
    p_s1.paragraph_format.space_before, p_s1.paragraph_format.space_after = Pt(0), Pt(0)
    p_s1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_s1 = p_s1.add_run("ખેતીવાડી અધિકારીની\nસહી અને હોદ્દો")
    run_s1.bold, run_s1.font.size = True, Pt(11)
    
    # કોલમ ૨: સિનિયર અકેરોલોજીસ્ટ (Center Alignment)
    p_s2 = table_sig.cell(0, 1).paragraphs[0]
    p_s2.paragraph_format.space_before, p_s2.paragraph_format.space_after = Pt(0), Pt(0)
    p_s2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_s2 = p_s2.add_run("સિનિયર અકેરોલોજીસ્ટની\nસહી અને હોદ્દો")
    run_s2.bold, run_s2.font.size = True, Pt(11)
    
    # કોલમ ૩: વિભાગીય વડા (Right Alignment)
    p_s3 = table_sig.cell(0, 2).paragraphs[0]
    p_s3.paragraph_format.space_before, p_s3.paragraph_format.space_after = Pt(0), Pt(0)
    p_s3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_s3 = p_s3.add_run("વિભાગીય વડાની\nસહી અને હોદ્દો")
    run_s3.bold, run_s3.font.size = True, Pt(11)
    
    # --- નીચેનો બાકીનો કોડ જેમ છે તેમ જ રહેશે ---
    p_passed = doc.add_paragraph()
    p_passed.paragraph_format.space_before, p_passed.paragraph_format.space_after = Pt(20), Pt(6)
    run_passed = p_passed.add_run("Passed for Payment Rs ........................................\nRupees: ........................................................................................")
    run_passed.font.size = Pt(12)
    
    p_aao = doc.add_paragraph()
    p_aao.paragraph_format.space_before, p_aao.paragraph_format.space_after = Pt(0), Pt(0)
    run_aao = p_aao.add_run("Assistant Administrative Officer\nN. M. College of Agriculture\nNavsari-396 450")
    run_aao.bold, run_aao.font.size = True, Pt(12)
    p_aao.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# Streamlit App UI
# ==========================================
st.title("સાદર નોંધ જનરેટર (Intelligent Sadar Nondh App)")

api_key = get_secret_value(
    "GEMINI_API_KEY",
    "GOOGLE_API_KEY",
    "GOOGLE_GENERATIVE_AI_API_KEY",
)

if api_key:
    st.sidebar.success("Gemini API key loaded from Streamlit secrets.")
else:
    api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")
    st.sidebar.caption('For Streamlit Cloud, add GEMINI_API_KEY = "your_key" in App settings > Secrets.')

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "નવી સાદર નોંધ (Create)", 
    "જુની નોંધ (Archives)", 
    "ખરીદી હુકમ (Purchase Order)",
    "બિલ પેમેન્ટ (Bill Payment)",
    "બિલ પેસ્ટિંગ (Bill Pasting)",
    "🗄️ ડિજિટલ આર્કાઇવ (Digital Vault)",
    "🧠 Learning Memory",
])

with tab1:
    st.markdown("### જરૂરિયાતની વિગત આપો (Provide Requirements)")
    col1, col2 = st.columns(2)
    with col1:
        text_prompt = st.text_area("તમારી જરૂરિયાત લખો:", placeholder="e.g., need 10 entomological pins...")
    with col2:
        uploaded_image = st.file_uploader("અથવા PDF/ફોટો અપલોડ કરો:", type=["pdf", "jpg", "jpeg", "png"])
    
    if st.button("જનરેટ કરો (Generate)"):
        if not api_key:
            st.error("Please add GEMINI_API_KEY in Streamlit secrets or enter your Gemini API Key in the sidebar.")
        elif not text_prompt and not uploaded_image:
            st.warning("Please provide either a text requirement or a PDF/image.")
        else:
            with st.spinner("સ્ટેચ્યુટ ૧૨૧ ની ચકાસણી અને નોંધ તૈયાર કરવામાં આવી રહી છે..."):
                try:
                    statute_context, sample_context = load_permanent_context()
                    memory_query = text_prompt or (uploaded_image.name if uploaded_image else "")
                    learned_memory_context = format_learning_memory_context(memory_query)
                    
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel(GEMINI_MAIN_MODEL)
                    
                    sys_prompt = f"""
                    You are an expert administrative AI for the Department of Entomology, N. M. College of Agriculture, NAU, Navsari.
                    Your task is to generate a formal 'સાદર નોંધ' in Gujarati.
                    If relevant, apply the user-approved learned memories below. The user's current request always wins if there is a conflict.
                    
                    [CONTEXT START]
                    {statute_context[:15000]}
                    {sample_context}
                    [CONTEXT END]

                    [LEARNED MEMORY START]
                    {learned_memory_context}
                    [LEARNED MEMORY END]

                    Format REQUIRED:
                    તા. {datetime.date.today().strftime('%d/%m/%Y')}
                    સ્થળ: નવસારી
                    સાદર નોંધ:
                    વિષય: [Appropriate Subject...]
                    સવિનય ઉપરોક્ત વિષય અન્વયે જણાવવાનું કે, અત્રેનાં કીટકશાસ્ત્ર વિભાગની આઈ.સી.એ.આર. યોજના AINP on Agril Acarology બ.સ. ૩૦૩/૨૦૯૨ અંતર્ગત [Detailed logical reason]. સદર વસ્તુનો કુલ અંદાજિત ખર્ચ [Total Amount] થનાર છે.
                    જે આપ સાહેબશ્રીને સ્ટેચ્યુટ ૧૨૧ની આઈટમ નંબર [DETERMINED_ITEM_NUMBER] મુજબ એનાયત થયેલ સત્તા અનુસાર સૈદ્ધાંતિક મંજુરી આપવા વિનંતી. સદર ખર્ચ અત્રેના વિભાગમાં ચાલતી આઈ.સી.એ.આર યોજના (બ.સ. ૩૦૩/૨૦૯૨) માં કરવામાં આવશે.

                    STATUTE 121 ITEM NUMBER DETERMINATION (CRITICAL INSTRUCTION):
                    1. You MUST read the 'Sample Nondh Format' provided in the context to find the historically correct 'આઈટમ નંબર' (Item Number) for this type of purchase.
                    2. Treat the Sample Nondh as the ultimate precedent. If the purchase involves laboratory chemicals, research materials, or AINP scheme-related items, strictly use the exact item number found in the sample (e.g., "૫૪ (i)" or "54 (i)").
                    3. Only rely on the dense 'Statute 121 Rules' PDF text if the item category is completely new and not covered by the sample document precedent. 
                    4. Replace [DETERMINED_ITEM_NUMBER] with the exact number in Gujarati format (like ૫૪ (i)). Do not guess or hallucinate numbers like 45 (ii) (ii).

                    TABLE LOGIC:
                    Analyze the user request to determine if a table is required. 
                    - If the request is a general administrative note WITHOUT specific items to purchase, DO NOT include a table.
                    - If the request involves purchasing, requesting, or listing items with quantities and prices, YOU MUST include a markdown table.
                    
                    If a table is included, the headers MUST strictly be in ENGLISH and use these EXACT columns:
                    Sr. No. | Details | Required Quantity | Available Pkt/Unit | Unit/Pkt Price | Total Price
                    
                    CRITICAL TABLE RULES:
                    1. The 'Details' column MUST contain ONLY the item name without the package size (e.g., "ACETIC ACID GLACIAL 99.5% Extra Pure").
                    2. The 'Available Pkt/Unit' column MUST contain the package size and unit type (e.g., "500 ML", "25 GM", "1 Unit").
                    3. The 'Required Quantity' and 'Unit/Pkt Price' columns MUST contain pure numbers only (e.g., "1" or "335.95"). Do NOT put units like "ML" or "GM" in these two columns.
                    4. Do NOT generate a "Grand Total" row. The system will calculate it automatically.

                    ખેતીવાડી અધિકારી,કીટકશાસ્ત્ર વિભાગ
                    પ્રોજેકટ ઈન્ચાર્જ,કીટકશાસ્ત્ર વિભાગ
                    પ્રાધ્યાપક અને વડા,કીટકશાસ્ત્ર વિભાગ

                    આચાર્ય અને ડીનશ્રી, ન. મ. કૃષિ મહાવિધાયલય, ન.કૃ.યુ. નવસારી
                    
                    ==== AI STATUTE ANALYSIS ====
                    1. **Original Statute 121 Details:** - **Item Number Used:** [State the specific rule number you used].
                        - **Original Statute Text:** [Provide the EXACT quote/sentence directly from the attached Statute 121 PDF for this specific rule number].
                    2. **Justification:** [Explain exactly WHY this specific statute applies to the requested purchase. Relate the items being bought to the statute's wording].
                    3. **Similar Precedent from Sample Nondh:** [Find a similar past purchase in the uploaded 'Sample Nondh' context. List its Subject, Date, and the Statute Item Number it used to prove your choice is historically accurate].
                    4. **Rejected Alternative Statute:** [Find another statute item number from the PDF that looks similar but is INCORRECT (e.g., a rule for furniture instead of chemicals). Quote it and explicitly explain why it is NOT compatible with this purchase].
                    """
                    
                    inputs = [sys_prompt, text_prompt]
                    if uploaded_image:
                        if uploaded_image.type == "application/pdf":
                            inputs.append({"mime_type": "application/pdf", "data": uploaded_image.getvalue()})
                        else:
                            inputs.append(Image.open(uploaded_image))
                        
                    response = model.generate_content(inputs)
                    res_text = response.text
                    
                    # Intercept and Split the AI Response safely
                    if "==== AI STATUTE ANALYSIS ====" in res_text:
                        parts = res_text.split("==== AI STATUTE ANALYSIS ====")
                        st.session_state['generated_nondh'] = parts[0].strip()
                        st.session_state['statute_analysis'] = parts[1].strip()
                    else:
                        st.session_state['generated_nondh'] = res_text.strip()
                        st.session_state['statute_analysis'] = ""
                        
                    st.success("સાદર નોંધ સફળતાપૂર્વક તૈયાર થઈ ગઈ છે!")
                    
                except Exception as e:
                    st.error(f"Error generating document: {e}")

    if 'generated_nondh' in st.session_state:
        
        # FEATURE ADDITION: Display Statute Analysis clearly separated from the Document Draft
        if 'statute_analysis' in st.session_state and st.session_state['statute_analysis']:
            with st.expander("🔍 Statute 121 Analysis & Justification (AI Reasoning)", expanded=True):
                st.info("આ વિભાગ ફક્ત તમારી જાણકારી માટે છે અને વર્ડ ડોક્યુમેન્ટ (DOCX) માં પ્રિન્ટ થશે નહીં.")
                st.markdown(st.session_state['statute_analysis'])
        
        st.markdown("---")
        st.markdown("### ડ્રાફ્ટ એડિટિંગ (Smart Editor)")
        
        pre_text, df, post_text = parse_markdown_to_parts(st.session_state['generated_nondh'])
        
        edit_pre = st.text_area("ઉપરનું લખાણ:", pre_text, height=400)
        
        if not df.empty:
            st.markdown("#### સ્માર્ટ ટેબલ (Smart Table)")
            st.info("નોંધ: 'Required Quantity' અથવા 'Unit/Pkt Price' બદલશો તો 'Total Price' અને લખાણમાં રહેલ 'અંદાજિત ખર્ચ' આપોઆપ બદલાઈ જશે.")
            
            # ટેબલની કોલમની પહોળાઈ (width) સેટ કરવા માટેનું કન્ફિગરેશન
            custom_column_config = {
                "Sr. No.": st.column_config.TextColumn("Sr. No.", width="small"),
                "Details": st.column_config.TextColumn("Details", width="large"),
                "Required Quantity": st.column_config.TextColumn("Required Quantity", width="small"),
                "Available Pkt/Unit": st.column_config.TextColumn("Available Pkt/Unit", width="medium"),
                "Unit/Pkt Price": st.column_config.TextColumn("Unit/Pkt Price", width="medium"),
                "Total Price": st.column_config.TextColumn("Total Price", width="medium")
            }
            
            edited_df = st.data_editor(
                df, 
                num_rows="dynamic", 
                use_container_width=True,
                column_config=custom_column_config
            )
            
            # Intelligent Math Calculation using Regex to strip any accidental text
            if 'Required Quantity' in edited_df.columns and 'Unit/Pkt Price' in edited_df.columns and 'Total Price' in edited_df.columns:
                
                # Extract pure numbers safely
                req_qty = edited_df['Required Quantity'].astype(str).str.extract(r'(\d+\.?\d*)')[0].astype(float).fillna(0)
                unit_price = edited_df['Unit/Pkt Price'].astype(str).str.extract(r'(\d+\.?\d*)')[0].astype(float).fillna(0)
                
                # Perform the calculation
                edited_df['Total Price'] = (req_qty * unit_price).round(2)
                
                # Dynamic Grand Total Calculation
                grand_total_calc = edited_df['Total Price'].sum()
                st.success(f"**Grand Total (કુલ રકમ): ₹ {grand_total_calc:,.2f}**")
                
                # Automatically sync the paragraph text with the accurate Grand Total
                edit_pre = re.sub(r'(અંદાજિત ખર્ચ\s*).*?(\s*થનાર)', f'\g<1>{grand_total_calc:,.2f}\g<2>', edit_pre)
        else:
            edited_df = pd.DataFrame()
            st.info("આ નોંધમાં ટેબલની જરૂરિયાત જણાઈ નથી. (No table required for this note based on the context).")
            
        edit_post = st.text_area("નીચેનું લખાણ:", post_text, height=300)
        
        # Re-stitch using the custom markdown generator that handles the Grand Total
        final_document = f"{edit_pre}\n\n{df_to_markdown_with_total(edited_df)}\n{edit_post}" if not edited_df.empty else f"{edit_pre}\n\n{edit_post}"
        current_subject = extract_nondh_subject(final_document)

        st.markdown("---")
        st.markdown("### દસ્તાવેજ પ્રીવ્યુ (Visual Preview - 20/80 Layout)")
        
        with st.container(border=True):
            prev_blank, prev_content = st.columns([2, 8]) 
            with prev_content:
                st.markdown(final_document)

        with st.expander("🧠 Save this as learning memory", expanded=False):
            mem_col1, mem_col2 = st.columns(2)
            with mem_col1:
                draft_memory_category = st.selectbox(
                    "Memory category",
                    LEARNING_MEMORY_CATEGORIES,
                    key="draft_memory_category",
                )
                draft_memory_title = st.text_input(
                    "Memory title",
                    value=current_subject,
                    key="draft_memory_title",
                )
            with mem_col2:
                draft_memory_keywords = st.text_input(
                    "Matching keywords",
                    value=current_subject,
                    key="draft_memory_keywords",
                )
            draft_memory_text = st.text_area(
                "What should the app remember?",
                value=f"When a future request matches these keywords, prefer this approved wording/correction:\n{final_document[:3000]}",
                height=180,
                key="draft_memory_text",
            )
            if st.button("🧠 Save this as learning memory", key="save_draft_learning_memory"):
                if not draft_memory_title.strip() or not draft_memory_text.strip():
                    st.error("Please add a title and memory text before saving.")
                else:
                    save_learning_memory(
                        draft_memory_category,
                        draft_memory_title.strip(),
                        draft_memory_keywords.strip(),
                        draft_memory_text.strip(),
                    )
                    st.success("Learning memory saved.")
        
        st.markdown("---")
        col_save, col_down = st.columns(2)
        with col_save:
            if st.button("આર્કાઇવમાં સેવ કરો (Save Nondh)"):
                subj = current_subject
                
                # Save Nondh to DB and get the ID
                nondh_id = save_to_db(subj, final_document)
                st.session_state['recent_nondh_id'] = nondh_id
                
                # Automatically save the DOCX to the vault linked to this new Nondh
                docx_data = create_docx(final_document)
                save_file_to_vault(docx_data, f"Nondh_{nondh_id}_{datetime.date.today().strftime('%d_%m')}.docx", "Sadar Nondh Draft", nondh_id=nondh_id)
                st.success("નોંધ અને ડ્રાફ્ટ સાચવી લેવામાં આવ્યા છે! (હવે તમે Tab 3 માં જઈ શકો છો)")
                
        with col_down:
            docx_data = create_docx(final_document)
            st.download_button(label="Download as Word (DOCX)",
                               data=docx_data,
                               file_name=f"Sadar_Nondh_{datetime.date.today().strftime('%d_%m_%Y')}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        # --- NEW FEATURE: Delete Recent Nondh ---
        st.markdown("---")
        with st.expander("🗑️ તાજેતરની નોંધ રદ કરો (Delete Recent Nondh - Last 30 Days)", expanded=False):
            recent_nondhs = get_recent_nondhs(30)
            if not recent_nondhs:
                st.info("છેલ્લા 30 દિવસની કોઈ નોંધ ઉપલબ્ધ નથી.")
            else:
                for n_id, n_date, n_subj in recent_nondhs:
                    col_del1, col_del2 = st.columns([8, 2])
                    with col_del1:
                        st.write(f"**ID: {n_id}** | 🗓️ {n_date} | 📝 {n_subj}")
                    with col_del2:
                        if st.button("🗑️ Delete", key=f"del_nondh_{n_id}", type="secondary"):
                            delete_nondh(n_id)
                            st.error(f"નોંધ ID {n_id} સફળતાપૂર્વક રદ કરવામાં આવી છે!")
                            st.rerun()

with tab2:
    st.markdown("### 🗄️ જૂના રેકોર્ડ શોધો (Archive Search)")
    search_query = st.text_input("🔍 Smart Search (Type in English or Gujarati)")
    db_records = get_archives("All", "All")
    sample_records = search_sample_nondh("", "All", "All")
    all_records = db_records + sample_records

    def smart_search_gemini(query, records):
        if not query.strip(): return records
        if not api_key: return records
        
        # Using a fast model for search
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(GEMINI_HEAVY_MODEL)
        learned_memory_context = format_learning_memory_context(query)
        
        records_context = ""
        for i, record in enumerate(records):
            # FIXED: Handle 4-item database records properly
            if len(record) == 4:
                _, _, subject, content = record
            elif len(record) == 3: 
                date, subject, content = record
            elif len(record) == 2: 
                subject, content = record
            else: 
                continue 
            
            # INCREASED CONTEXT: Send up to 800 characters so AI can read the actual items
            clean_content = content[:800].replace('\n', ' ')
            records_context += f"ID: {i} | Subject: {subject} | Content: {clean_content}...\n"
            
        # ADVANCED PROMPT: Instruct the AI to act as a cross-lingual search engine
        prompt = f"""
        You are an intelligent bilingual search engine (English <-> Gujarati).
        The user is searching for: "{query}"
        
        User-approved learned memory that may help interpret this search:
        {learned_memory_context}
        Your Task: Find all record IDs that semantically match the user's query.
        - If the query is in English (e.g., "chemicals", "pins", "equipment"), translate the intent and find the matching Gujarati records.
        - If the query is in Gujarati, find the exact or contextually matching records.
        - Match based on meaning and translated keywords, not just exact text matches.
        
        Records Data:
        {records_context}
        
        Return ONLY a comma-separated list of matching IDs (e.g., 0,2,5). 
        If absolutely no records match the meaning, return the exact word: NONE
        """
        
        try:
            response = model.generate_content(prompt)
            result = response.text.strip()
            
            if "NONE" in result.upper() or not result: 
                return []
                
            # Safely extract only numbers using regex (prevents crashes if AI adds extra words)
            import re
            matched_ids_str = re.findall(r'\d+', result)
            matched_ids = [int(x) for x in matched_ids_str]
            
            # Return only valid records
            return [records[i] for i in matched_ids if i < len(records)]
            
        except Exception as e:
            return records

    display_records = smart_search_gemini(search_query, all_records) if search_query else all_records
    if display_records:
        st.success(f"કુલ {len(display_records)} રેકોર્ડ મળ્યા.")
        for idx, record in enumerate(display_records):
            if len(record) == 4:
                nondh_id, date, subject, content = record[0], record[1], record[2], record[3]
            else:
                continue # Skip invalid records
            
            with st.expander(f"ID #{nondh_id if nondh_id else 'Ref'} | {date} - {subject}"):
                st.markdown(content)
                st.download_button("Download Nondh (Word)", data=create_docx(content), file_name=f"Nondh_{nondh_id}.docx", key=f"dl_{idx}") 
                
                # --- Unified View of Vault Files linked to this Nondh ---
                if nondh_id:
                    vault_files = get_vault_files_by_nondh(nondh_id)
                    if vault_files:
                        st.markdown("---")
                        st.markdown("#### 📂 જોડાયેલ દસ્તાવેજો (Linked Vault Documents)")
                        for f_name, f_path, u_date, d_type, desc in vault_files:
                            col1, col2 = st.columns([8, 2])
                            with col1: st.caption(f"**{d_type}**: {f_name} ({u_date}) - {desc}")
                            with col2:
                                file_data_t2 = load_vault_file_bytes(f_path)
                                if file_data_t2:
                                        st.download_button("⬇️", data=file_data_t2, file_name=f_name, key=f"dl_v_{f_path}")
                                else: 
                                    st.error("Missing")

with tab3:
    st.markdown("### 📝 ખરીદી હુકમ બનાવો (Generate Purchase Order)")
    st.info("નોંધ મંજૂર થયા પછી સપ્લાયરને ખરીદીનો ઓર્ડર મોકલવા માટે અહીં વિગતો ભરો.")
    
    st.markdown("#### ૧. મંજૂર થયેલ નોંધ પસંદ કરો (Select Approved Nondh)")
    db_records = get_archives("All", "All")
    
    # --- NEW: Check the database for Nondh IDs that already have a PO ---
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT nondh_id FROM purchase_orders WHERE nondh_id IS NOT NULL")
    used_nondh_ids = [row[0] for row in c.fetchall()]
    conn.close()

    options = ["-- જાતે માહિતી ભરો (Manual Entry) --"]
    record_dict = {}
    for idx, row in enumerate(db_records):
        if len(row) == 4:
            nondh_id = row[0]
            # ONLY add to dropdown if it hasn't been used yet!
            if nondh_id not in used_nondh_ids:
                label = f"[{nondh_id}] {row[1]} - {row[2]}" 
                record_dict[label] = row[3] 
                options.append(label)
                
    selected_nondh = st.selectbox("અગાઉ સેવ કરેલ નોંધ પસંદ કરો:", options)
    
    # --- Extract Nondh ID early so we can link the upload to it ---
    current_nondh_id = None
    if selected_nondh != "-- જાતે માહિતી ભરો (Manual Entry) --":
        import re
        match = re.search(r'\[(\d+)\]', selected_nondh)
        if match:
            current_nondh_id = int(match.group(1))

    st.markdown("#### ૨. સહી કરેલ નોંધ અપલોડ કરો (Upload Signed Nondh - Optional)")
    uploaded_nondh = st.file_uploader("મંજૂર થયેલ/સહીવાળી નોંધ અપલોડ કરો:", type=["pdf", "jpg", "jpeg", "png"], key="nondh_up")
    
    if uploaded_nondh:
        if not current_nondh_id:
            st.warning("ફાઈલ અપલોડ કરતા પહેલા ઉપરથી 'અગાઉ સેવ કરેલ નોંધ' પસંદ કરો!")
        else:
            # Check if already uploaded this session to prevent duplicates!
            if st.session_state.get(f"uploaded_nondh_{current_nondh_id}") != uploaded_nondh.name:
                file_bytes = uploaded_nondh.getbuffer()
                # Use explicit nondh_id keyword to link it perfectly!
                save_file_to_vault(file_bytes, uploaded_nondh.name, "Signed Nondh", nondh_id=current_nondh_id, description="Auto-uploaded from Tab 3")
                st.session_state[f"uploaded_nondh_{current_nondh_id}"] = uploaded_nondh.name
                st.success("સહી કરેલ નોંધ વોલ્ટમાં સાચવી લેવાઈ છે!")

    st.markdown("---")
    st.markdown("#### ૩. સપ્લાયર અને ઓર્ડરની વિગત (Supplier & Order Details)")
    col_po1, col_po2 = st.columns(2)
    with col_po1:
        vendor_name = st.text_input("સપ્લાયરનું નામ (Vendor Name)", value="DUTT ENTERPRISE")
        outward_no = st.text_input("જાવક નંબર (Outward No.)", value="139")
    with col_po2:
        vendor_address = st.text_area("સપ્લાયરનું સરનામું", value="A/5, Krishna complex, borsad chokadi,\nAnand sojitra road, Anand 388 001", height=110)
        po_date = st.date_input("તારીખ (Date)", value=datetime.date.today())
        
    default_df = pd.DataFrame(columns=["Details", "Required Quantity", "Available Pkt/Unit", "Unit/Pkt Price", "Total Price"])
    if selected_nondh != "-- જાતે માહિતી ભરો (Manual Entry) --":
        _, session_df, _ = parse_markdown_to_parts(record_dict[selected_nondh])
        if not session_df.empty: default_df = session_df
    
    po_df = st.data_editor(default_df, num_rows="dynamic", use_container_width=True, key="po_editor")
    
   # ડાઉનલોડ બટન અને ડેટાબેઝ સેવ કરવાનું નવું લોજીક
    # ડાઉનલોડ બટન અને ડેટાબેઝ સેવ કરવાનું નવું લોજીક
    if vendor_name and not po_df.empty:
        formatted_date = po_date.strftime("%d.%m.%Y")
        grand_total = pd.to_numeric(po_df['Total Price'], errors='coerce').fillna(0).sum()
        
        current_nondh_id = None
        if selected_nondh != "-- જાતે માહિતી ભરો (Manual Entry) --":
            import re
            match = re.search(r'\[(\d+)\]', selected_nondh)
            if match:
                current_nondh_id = int(match.group(1))
        
        # 1. પહેલાં વર્ડ ફાઈલ બેકગ્રાઉન્ડમાં તૈયાર કરી લો
# 1. પહેલાં વર્ડ ફાઈલ બેકગ્રાઉન્ડમાં તૈયાર કરી લો
        po_docx = create_purchase_order_docx(vendor_name, vendor_address, outward_no, formatted_date, po_df)
        
        st.markdown("---")
        st.info("સૂચના: પહેલા 'સ્ટેપ ૧' બટન પર ક્લિક કરીને ફાઈલ ડાઉનલોડ કરો, ત્યારબાદ જ 'સ્ટેપ ૨' પર ક્લિક કરો.")
        
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            st.download_button(
                label="૧. 📄 ખરીદી હુકમ ડાઉનલોડ કરો (Step 1: Download PO)", 
                data=po_docx, 
                file_name=f"PO_{vendor_name}.docx"
            )

        with col_btn2:
            if st.button("૨. ➡️ પેમેન્ટ માટે આગળ મોકલો (Step 2: Send to Tab 4)", type="primary"):
                save_po_to_db(current_nondh_id, vendor_name, outward_no, formatted_date, grand_total)
                st.success("ઓર્ડર સફળતાપૂર્વક Tab 4 (Bill Payment) માં મોકલી દેવાયો છે!")
                st.rerun()
# --- TAB 4 (Bill Payment ONLY) ---
with tab4:
    st.markdown("### 💳 બિલ પેમેન્ટ ફોર્મ (Bill Payment Form)")
    st.info("જે ખરીદીના હુકમ (Purchase Orders) માટે બિલ ચૂકવવાનું બાકી છે, તે જ અહીં દેખાશે.")
    
    unfinished_pos = get_unfinished_pos(('Unfinished',))
    
    if not unfinished_pos:
        st.success("હાલમાં કોઈ બિલ પેમેન્ટ બાકી નથી! (No unfinished purchase orders).")
    else:
        po_dict_tab4 = {}
        po_options_tab4 = []
        for po in unfinished_pos:
            # FIXED: Added nondh_id_t4 to catch all 6 items from the database
            po_id, nondh_id_t4, v_name, o_no, p_date, amt = po
            amt = coerce_amount(amt)
            label = f"PO #{o_no} - {v_name} - ₹{format_amount(amt)} ({p_date})"
            po_options_tab4.append(label)
            po_dict_tab4[label] = (po_id, nondh_id_t4, v_name, o_no, p_date, amt)
            
        selected_po_label_t4 = st.selectbox("પેમેન્ટ ફોર્મ માટે ઓર્ડર પસંદ કરો (Select Pending PO):", po_options_tab4, key="po_tab4")
        
        if selected_po_label_t4:
            # FIXED: Added nondh_id_t4 here as well
            po_id, nondh_id_t4, v_name, o_no, p_date, amt = po_dict_tab4[selected_po_label_t4]
            
            if (
                st.session_state.get("current_po_id_t4") != po_id
                or "amt_t4" not in st.session_state
                or "ext_amt" not in st.session_state
                or "ext_bill_no" not in st.session_state
                or "ext_words" not in st.session_state
                or "last_invoice" not in st.session_state
            ):
                st.session_state.current_po_id_t4 = po_id
                st.session_state.ext_bill_no = "INV-"
                st.session_state.ext_amt = coerce_amount(amt)
                st.session_state.amt_t4 = coerce_amount(amt)
                st.session_state.ext_words = ""
                st.session_state.last_invoice = None

            st.markdown("#### ઇન્વોઇસ અને બજેટની વિગતો (Invoice Details)")
            
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                budget_head = st.text_input("Budget Head No.", value="303/2092 (AINP on Agril Acarology)", key="bh_t4")
                bill_no = st.text_input("ઇન્વોઇસ/બિલ નંબર (Vendor Bill No.)", value=st.session_state.ext_bill_no)
                invoice_upload = st.file_uploader("પાર્ટીનું બિલ અપલોડ કરો (Upload Vendor Invoice PDF/Img)", type=["pdf", "jpg", "png"])
                
                if invoice_upload:
                    # Prevent duplicate uploads on app rerun
                    if st.session_state.get(f"uploaded_inv_{po_id}") != invoice_upload.name:
                        file_bytes = invoice_upload.getbuffer()
                        
                        # Use explicit nondh_id keyword to link it perfectly!
                        save_file_to_vault(file_bytes, invoice_upload.name, "Party Invoice", nondh_id=nondh_id_t4, description=f"Auto-uploaded from Tab 4 for PO #{o_no}")
                        st.session_state[f"uploaded_inv_{po_id}"] = invoice_upload.name
                        st.success("ઇન્વોઇસ વોલ્ટમાં Nondh ID સાથે સેવ થઈ ગયું!")

                    if invoice_upload.name != st.session_state.last_invoice and api_key:
                        with st.spinner("AI દ્વારા બિલની વિગતો વાંચવામાં આવી રહી છે... (Extracting...)"):
                            try:
                                import json
                                import re  # ટેક્સ્ટ સાફ કરવા માટે
                                genai.configure(api_key=api_key)
                                model = genai.GenerativeModel(GEMINI_HEAVY_MODEL)
                                
                                prompt = f"""
                                You are an intelligent accounting AI. The approved Purchase Order (PO) amount for this transaction is ₹{amt}.
                                Carefully analyze the uploaded invoice and extract the following:
                                
                                1. Invoice/Bill Number.
                                2. Final Payable Amount: Look for terms like 'Grand Total', 'Invoice Total', 'Net Payable', or 'Total Amount'. Use your intelligence to understand the invoice structure and identify the final amount including taxes. Logically compare it with the PO amount (₹{amt}) to ensure you pick the correct total. Return it as a PURE NUMBER WITHOUT COMMAS (e.g. 3894.00).
                                3. Final Payable Amount in English words.
                                
                                Return ONLY a valid JSON object in this exact format:
                                {{"bill_no": "INV-123", "amount": 1234.50, "amount_words": "One Thousand Two Hundred..."}}
                                """
                                
                                # --- મોટો સુધારો: PyPDF2 કાઢીને PDF સીધી જ Gemini ને આપો ---
                                if invoice_upload.type == "application/pdf":
                                    doc_parts = [{"mime_type": "application/pdf", "data": invoice_upload.getvalue()}]
                                    response = model.generate_content([prompt, doc_parts[0]])
                                else:
                                    img = Image.open(invoice_upload)
                                    response = model.generate_content([prompt, img])
                                
                                res_text = response.text.strip().replace("```json", "").replace("```", "")
                                data = json.loads(res_text)
                                
                                # --- બીજો સુધારો: રકમમાંથી કોમા (,) અને અન્ય ચિન્હો સાફ કરો ---
                                raw_amount = coerce_amount(data.get("amount", amt), amt)
                                
                                st.session_state.ext_bill_no = str(data.get("bill_no", "INV-"))
                                st.session_state.ext_amt = raw_amount
                                st.session_state.amt_t4 = raw_amount
                                st.session_state.ext_words = str(data.get("amount_words", ""))
                                st.session_state.last_invoice = invoice_upload.name
                                st.rerun() 
                            except Exception as e:
                                st.warning(f"આપમેળે વિગત મેળવવામાં ભૂલ: {e}. કૃપા કરીને જાતે ભરો.")

            with col_b2:
                bill_date = st.date_input("ઇન્વોઇસની તારીખ (Bill Date)", value=datetime.date.today())
                final_amt = st.number_input("ચૂકવવા પાત્ર રકમ (Amount to Pay)", min_value=0.0, step=1.0, format="%.2f", key="amt_t4")
                
                # --- NEW FEATURE: Auto-fill English words based on Amount ---
                col_eng1, col_eng2 = st.columns([3, 1])
                with col_eng1:
                    amount_words = st.text_input("રકમ શબ્દોમાં (Amount in Words - English)", value=st.session_state.ext_words, placeholder="e.g., Four Thousand Two Hundred Forty Eight")
                with col_eng2:
                    st.write("") 
                    if st.button("✨ AI થી ભરો", key="auto_fill_eng_t4"):
                        if api_key:
                            with st.spinner("Converting to words..."):
                                try:
                                    genai.configure(api_key=api_key)
                                    model = genai.GenerativeModel(GEMINI_HEAVY_MODEL)
                                    prompt = f"Convert the number {final_amt} into English words (capitalize the first letter of each word). Return ONLY the text, nothing else. Example: for 3956.50 return 'Three Thousand Nine Hundred Fifty Six And Fifty Paise'."
                                    res = model.generate_content(prompt)
                                    st.session_state.ext_words = res.text.strip()
                                    st.rerun()
                                except Exception as e:
                                    st.error("AI Error.")
                        else:
                            st.warning("API Key is required!")
            
            st.markdown("---")
            if st.button("📄 Generate Bill Payment Form"):
                if not amount_words: st.error("Please enter the amount in words!")
                else:
                    bp_docx = create_bill_payment_form(budget_head, bill_no, bill_date.strftime("%d/%m/%Y"), v_name, final_amt, amount_words)
                    st.download_button(
                        label="Download Bill Payment Form", 
                        data=bp_docx, 
                        file_name=f"Payment_Form_{v_name}.docx",
                        on_click=mark_po_as_payment_generated,
                        args=(po_id,)
                    )

# --- TAB 5 (Bill Pasting & Mark Paid ONLY) ---
with tab5:
    st.markdown("### 📑 બિલ પેસ્ટિંગ અને પ્રમાણપત્ર (Bill Pasting Form)")
    
    if "auto_guj_words" not in st.session_state:
        st.session_state.auto_guj_words = ""

    unfinished_pos_t5 = get_unfinished_pos(('Unfinished', 'Payment_Generated'))
    
    if not unfinished_pos_t5:
        st.success("હાલમાં કોઈ બિલ પેમેન્ટ બાકી નથી! (No unfinished purchase orders).")
    else:
        po_dict_tab5 = {}
        po_options_tab5 = []
        for po in unfinished_pos_t5:
            # FIXED: Added nondh_id_t5 to catch all 6 items from the database
            po_id, nondh_id_t5, v_name, o_no, p_date, amt = po
            amt = coerce_amount(amt)
            label = f"PO #{o_no} - {v_name} - ₹{format_amount(amt)} ({p_date})"
            po_options_tab5.append(label)
            po_dict_tab5[label] = (po_id, nondh_id_t5, v_name, o_no, p_date, amt)
            
        selected_po_label_t5 = st.selectbox("પેસ્ટિંગ ફોર્મ માટે ઓર્ડર પસંદ કરો:", po_options_tab5, key="po_tab5")
        
        if selected_po_label_t5:
            # FIXED: Added nondh_id_t5 here as well
            po_id_t5, nondh_id_t5, v_name_t5, o_no_t5, p_date_t5, amt_t5 = po_dict_tab5[selected_po_label_t5]

            col_p1, col_p2 = st.columns(2)
            with col_p1:
                budget_head_pst = st.text_input("Budget Head No.", value="303/2092 (AINP on Agril Acarology)", key="bh_t5")
                grant_year = st.text_input("ફાળવેલ ગ્રાન્ટ વર્ષ (Grant Year)", value="", placeholder="હાથેથી લખવા માટે ખાલી છોડી દો")
                party_name_pst = st.text_input("પાર્ટીનું નામ (Party Name)", value=v_name_t5, key="party_t5")
            with col_p2:
                # --- નવો સુધારો: Streamlit ની મેમરીને સીધી અપડેટ કરવા માટેનું લોજીક ---
                # જો Tab 4 માં આ જ ઓર્ડર ખૂલ્યો હોય અને ત્યાં રકમ સેટ હોય, તો Tab 5 ની મેમરી ફરજિયાત ઓવરરાઈટ કરો
                amount_from_tab4 = st.session_state.get("current_po_id_t4") == po_id_t5 and "amt_t4" in st.session_state
                amount_for_t5 = coerce_amount(st.session_state.amt_t4) if amount_from_tab4 else coerce_amount(amt_t5)
                if st.session_state.get("current_po_id_t5") != po_id_t5:
                    st.session_state.current_po_id_t5 = po_id_t5
                    st.session_state['amt_t5'] = amount_for_t5
                elif amount_from_tab4:
                    st.session_state['amt_t5'] = amount_for_t5
                
                # નોંધ: અહીથી `value=...` કાઢી નાખ્યું છે, કારણ કે તે સીધું `key="amt_t5"` ની મેમરીમાંથી જ લેટેસ્ટ રકમ ખેંચી લેશે
                final_amt_pst = st.number_input("બીલની કુલ રકમ (Amount)", min_value=0.0, step=1.0, format="%.2f", key="amt_t5")
                
                # --- બટન વગર આપોઆપ ગુજરાતી અનુવાદ (Auto-Translate) ---
                @st.cache_data(show_spinner=False)
                def get_gujarati_words_auto(amount, key):
                    if not key or amount == 0: return ""
                    try:
                        genai.configure(api_key=key)
                        model = genai.GenerativeModel(GEMINI_LIGHT_MODEL)
                        prompt = f"Translate the number {amount} into Gujarati words. Return ONLY the Gujarati translation. Example: for 3956 return 'ત્રણ હજાર નવસો છપ્પન'."
                        res = model.generate_content(prompt)
                        return res.text.strip()
                    except:
                        return ""
                
                auto_gujarati_text = get_gujarati_words_auto(final_amt_pst, api_key)
                amt_words_guj = st.text_input("રકમ શબ્દોમાં (ગુજરાતીમાં)", value=auto_gujarati_text, placeholder="દા.ત., ત્રણ હજાર નવસો છપ્પન")
                            
            st.markdown("#### 📝 મંજુરીની વિગતો (Approval Details - મુદ્દા નં. ૧)")
            col_a1, col_a2, col_a3 = st.columns(3)
            with col_a1:
                item_no_pst = st.text_input("સ્ટેચ્યુટ નં. ૧૨૧ ની આઇટમ નં", value="", placeholder="દા.ત. 14")
            with col_a2:
                approval_no_pst = st.text_input("મંજુરી નં. (Approval No.)", value="", placeholder="દા.ત. ACN/123/2026")
            with col_a3:
                approval_date_pst = st.text_input("મંજુરી તારીખ (Approval Date)", value="", placeholder="DD/MM/YYYY")
                
            st.markdown("#### 📝 રજીસ્ટર અને નોંધની વિગતો (Register Details - મુદ્દા નં. ૩ અને ૮)")
            col_r1, col_r2 = st.columns(2)
            with col_r1:
                reg_type = st.selectbox("કયા રજીસ્ટરમાં નોંધ કરી? (મુદ્દા નં. ૩)", ["ચીજવસ્તુ વપરાશ (કન્ઝયુમેબલ)", "ડેડસ્ટોક", "સ્ટોર રોજમેળ", "ટેલીફોન", "સ્ટેમ્પ", "સ્ટેશનરી", "પરચુરણ માલ સામાન", "રીપેરીંગ"])
                reg_page_no = st.text_input("રજીસ્ટર પાના નં. (Register Page No.)", value="")
            with col_r2:
                bill_reg_date = st.date_input("બીલ રજીસ્ટરમાં નોંધની તારીખ (મુદ્દા નં. ૮)", value=datetime.date.today())
                bill_reg_page_no = st.text_input("બીલ રજી પાના નં. (Bill Reg. Page No.)", value="")
                bill_reg_sr_no = st.text_input("બીલ રજી અનુ. નં. (Bill Reg. Serial No.)", value="")
                
            st.markdown("---")
            col_btn_pst1, col_btn_pst2 = st.columns(2)
            
            with col_btn_pst1:
                if st.button("📑 Generate Exact Bill Pasting Form"):
                    if not amt_words_guj:
                        st.error("કૃપા કરીને રકમ શબ્દોમાં (ગુજરાતીમાં) લખો અથવા 'AI થી ભરો' બટન દબાવો.")
                    elif not reg_page_no or not bill_reg_page_no or not bill_reg_sr_no:
                        st.warning("કૃપા કરીને રજીસ્ટરના પાના નંબર અને અનુક્રમ નંબર ભરો.")
                    else:
                        bill_reg_date_str = bill_reg_date.strftime("%d/%m/%Y")
                        
                        # Added new parameters to function call
                        pst_docx = create_bill_pasting_form(
                            budget_head_pst, grant_year, party_name_pst, final_amt_pst, 
                            amt_words_guj, reg_type, reg_page_no, bill_reg_date_str, 
                            bill_reg_page_no, bill_reg_sr_no, item_no_pst, approval_no_pst, approval_date_pst
                        )
                        st.download_button("Download Pasting Form", data=pst_docx, file_name=f"Pasting_Form_{v_name_t5}.docx")
            
            with col_btn_pst2:
                if st.button("✅ બિલ પેમેન્ટ પૂરું કરો (Mark as Paid)", key="mark_paid"):
                    mark_po_as_paid(po_id_t5)
                    st.success("ઓર્ડર પેમેન્ટ લિસ્ટમાંથી દૂર કરવામાં આવ્યો છે! રિફ્રેશ કરો.")
                    st.rerun()
# --- TAB 6 (Digital Vault / Archive) ---
# --- TAB 6 (DIGITAL VAULT & PAYMENT CLOSURE) ---
with tab6:
    st.markdown("### 🗄️ ડિજિટલ વોલ્ટ અને પેમેન્ટ ક્લોઝર (Vault & Payment Closure)")
    
    # Section A: Mark as Paid
    with st.expander("✅ બાકી પેમેન્ટ ક્લિયર કરો (Pending Payments to Mark as Paid)", expanded=False):
        pending_pos = get_unfinished_pos(('Unfinished', 'Payment_Generated'))
        if not pending_pos: st.info("કોઈ પેમેન્ટ બાકી નથી.")
        else:
            p_dict = {f"PO #{p[3]} - {p[2]} (₹{format_amount(p[5])})": p for p in pending_pos}
            sel_pay = st.selectbox("પેમેન્ટ થયેલ ઓર્ડર પસંદ કરો:", list(p_dict.keys()))
            if sel_pay:
                po_data = p_dict[sel_pay]
                col_pay1, col_pay2, col_pay3 = st.columns(3)
                with col_pay1: pay_info = st.text_input("પેમેન્ટની વિગત (UTR / Cheque No. / Date) - Optional")
                with col_pay2:
                    st.write("")
                    if st.button("Mark as Paid & Close Workflow", type="primary"):
                        mark_po_as_paid(po_data[0], pay_info)
                        st.success("પેમેન્ટ નોંધાઈ ગયું છે અને ફાઈલ ક્લોઝ થઈ ગઈ છે!")
                        st.rerun()
                with col_pay3:
                    st.write("")
                    if st.button("🗑️ Delete Entry Forever"):
                        delete_po(po_data[0])
                        st.error("ઓર્ડર કાયમ માટે રદ કરવામાં આવ્યો છે!")
                        st.rerun()

    # --- NEW SECTION: Upload Physical Signed Documents ---
    st.markdown("---")
    st.markdown("#### 📤 દસ્તાવેજ અપલોડ કરો (Upload Signed Documents to Vault)")
    st.info("સહી કરેલ PO, બિલ પેમેન્ટ અને પેસ્ટિંગ ફોર્મ અહીં કાયમી સાચવવા માટે અપલોડ કરો.")
    
    # 1. Get all Nondhs and map their subjects
    all_nondhs_for_vault = get_archives("All", "All")
    nondh_subject_map = {}
    
    # 2. Find which Nondhs are "Closed" (they already have a Signed Bill Pasting)
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT nondh_id FROM digital_vault WHERE doc_type = 'Signed Bill Pasting' AND nondh_id IS NOT NULL")
    closed_nondh_ids = [row[0] for row in c.fetchall()]
    conn.close()
    
    # 3. Create the filtered dropdown options (excluding closed Nondhs)
    nondh_opts = ["-- Select Nondh --"]
    for n in all_nondhs_for_vault:
        if len(n) == 4:
            n_id, n_subj = n[0], n[2]
            nondh_subject_map[n_id] = n_subj  # Save for the expander titles later
            if n_id not in closed_nondh_ids:
                nondh_opts.append(f"[{n_id}] {n_subj}") # Display ID and Subject
    
    col_u1, col_u2, col_u3 = st.columns(3)
    with col_u1:
        sel_nondh_vault = st.selectbox("કઈ નોંધ સાથે જોડવું છે? (Select Nondh ID)", nondh_opts)
    with col_u2:
        # Restricted Document Types
        sel_doc_type = st.selectbox("દસ્તાવેજનો પ્રકાર (Document Type)", [
            "Signed PO", "Signed Bill Payment", "Signed Bill Pasting"
        ])
    with col_u3:
        up_file = st.file_uploader("ફાઈલ પસંદ કરો (PDF/Image)", type=['pdf', 'jpg', 'jpeg', 'png'])
        
    if st.button("💾 વોલ્ટમાં સેવ કરો (Save to Permanent Vault)"):
        if sel_nondh_vault == "-- Select Nondh --":
            st.error("કૃપા કરીને નોંધ પસંદ કરો!")
        elif not up_file:
            st.error("કૃપા કરીને ફાઈલ અપલોડ કરો!")
        else:
            import re
            match = re.search(r'\[(\d+)\]', sel_nondh_vault)
            if match:
                n_id_val = int(match.group(1))
                file_bytes = up_file.getbuffer()
                save_file_to_vault(file_bytes, up_file.name, sel_doc_type, nondh_id=n_id_val, description="Manually Uploaded to Vault")
                st.success(f"{sel_doc_type} વોલ્ટમાં Nondh #{n_id_val} હેઠળ કાયમ માટે સેવ થઈ ગયું!")
                st.rerun()

    # --- UPDATED SECTION: Grouped Vault Display with Expanders ---
    st.markdown("---")
    st.markdown("#### 🔍 નોંધ મુજબ વોલ્ટ (Vault Grouped by Nondh ID)")
    
    current_year = datetime.date.today().year
    fy_options = ["All"] + [f"{y}-{str(y+1)[2:]}" for y in range(current_year-2, current_year+3)][::-1]
    
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1: filter_fy = st.selectbox("નાણાકીય વર્ષ (Financial Year)", fy_options)
    with col_f2: filter_type = st.selectbox("પ્રકાર (Type)", ["All", "Signed Nondh", "Approval Letter", "Signed PO", "Party Invoice", "Signed Bill Payment", "Signed Bill Pasting", "Sadar Nondh Draft", "PO Draft", "Other"])
    with col_f3: search_kw = st.text_input("શબ્દથી શોધો (Search by Name/Tag)")

    vault_records = get_vault_files(filter_fy, filter_type, search_kw)
    
    if not vault_records: 
        st.info("કોઈ ડોક્યુમેન્ટ મળ્યા નથી.")
    else:
        st.success(f"કુલ {len(vault_records)} ડોક્યુમેન્ટ્સ મળ્યા.")
        
        # 1. Group records by Nondh ID
        grouped_vault = {}
        for record in vault_records:
            v_id, n_id, f_name, f_path, u_date, fy, month, d_type, desc = record
            group_key = n_id if n_id else "Unlinked"
            if group_key not in grouped_vault:
                grouped_vault[group_key] = []
            grouped_vault[group_key].append(record)
            
        # 2. Display them neatly under clickable Expanders
        for n_id_key, records in grouped_vault.items():
            
            # Determine the title of the expander
            if n_id_key == "Unlinked":
                expander_title = "📁 Unlinked / General Documents"
            else:
                subj = nondh_subject_map.get(n_id_key, "Unknown Subject")
                expander_title = f"📁 Nondh ID: {n_id_key} - {subj}"
            
            # Create a clickable dropdown panel (expander)
            # Create a clickable dropdown panel (expander)
            with st.expander(expander_title, expanded=False):
                for record in records:
                    v_id, n_id, f_name, f_path, u_date, fy, month, d_type, desc = record
                    with st.container(border=True):
                        col_info, col_btn1, col_btn2 = st.columns([6, 2, 2])
                        
                        with col_info:
                            st.markdown(f"**{d_type}**: {f_name}")
                            st.caption(f"🗓️ {u_date} | 📁 {fy} ({month})")
                            
                        with col_btn1:
                            # Use our intelligent cloud loader instead of direct os.path.exists
                            file_data = load_vault_file_bytes(f_path)
                            if file_data:
                                st.download_button("⬇️ Download", data=file_data, file_name=f_name, key=f"dl_vault_main_{v_id}")
                            else:
                                st.error("File missing completely")
                                
                        with col_btn2:
                            if st.button("🗑️ Delete", key=f"del_vault_{v_id}"):
                                conn = sqlite3.connect(DB_FILE)
                                c = conn.cursor()
                                c.execute("DELETE FROM digital_vault WHERE id = ?", (v_id,))
                                conn.commit()
                                conn.close()
                                push_db_to_github()
                                if os.path.exists(f_path):
                                    try: os.remove(f_path)
                                    except: pass
                                st.error(f"ફાઈલ '{f_name}' રદ કરવામાં આવી છે!")
                                st.rerun()

with tab7:
    st.markdown("### 🧠 Learning Memory")

    with st.expander("➕ Add Memory", expanded=True):
        with st.form("add_learning_memory_form"):
            add_col1, add_col2 = st.columns(2)
            with add_col1:
                new_memory_category = st.selectbox("Category", LEARNING_MEMORY_CATEGORIES, key="new_memory_category")
                new_memory_title = st.text_input("Title", key="new_memory_title")
            with add_col2:
                new_memory_keywords = st.text_input("Matching keywords", key="new_memory_keywords")
                new_memory_enabled = st.checkbox("Enabled", value=True, key="new_memory_enabled")
            new_memory_text = st.text_area("Memory text", height=160, key="new_memory_text")
            if st.form_submit_button("Save Memory"):
                if not new_memory_title.strip() or not new_memory_text.strip():
                    st.error("Please enter a title and memory text.")
                else:
                    save_learning_memory(
                        new_memory_category,
                        new_memory_title.strip(),
                        new_memory_keywords.strip(),
                        new_memory_text.strip(),
                        new_memory_enabled,
                    )
                    st.success("Learning memory saved.")
                    st.rerun()

    memories = get_learning_memories(include_disabled=True)
    st.markdown("#### Saved Memories")
    if not memories:
        st.info("No learning memories saved yet.")
    else:
        for memory in memories:
            memory_id, category, title, keywords, memory_text, enabled, created_at, updated_at = memory
            status_label = "Enabled" if enabled else "Disabled"
            with st.expander(f"#{memory_id} | {status_label} | {title}", expanded=False):
                edit_col1, edit_col2 = st.columns(2)
                with edit_col1:
                    category_index = LEARNING_MEMORY_CATEGORIES.index(category) if category in LEARNING_MEMORY_CATEGORIES else 0
                    edit_category = st.selectbox(
                        "Category",
                        LEARNING_MEMORY_CATEGORIES,
                        index=category_index,
                        key=f"memory_category_{memory_id}",
                    )
                    edit_title = st.text_input("Title", value=title, key=f"memory_title_{memory_id}")
                with edit_col2:
                    edit_keywords = st.text_input("Matching keywords", value=keywords or "", key=f"memory_keywords_{memory_id}")
                    st.caption(f"Created: {created_at or '-'} | Updated: {updated_at or '-'}")
                edit_memory_text = st.text_area(
                    "Memory text",
                    value=memory_text or "",
                    height=180,
                    key=f"memory_text_{memory_id}",
                )

                action_col1, action_col2, action_col3 = st.columns(3)
                with action_col1:
                    if st.button("Save Changes", key=f"save_memory_{memory_id}"):
                        if not edit_title.strip() or not edit_memory_text.strip():
                            st.error("Please enter a title and memory text.")
                        else:
                            save_learning_memory(
                                edit_category,
                                edit_title.strip(),
                                edit_keywords.strip(),
                                edit_memory_text.strip(),
                                bool(enabled),
                                memory_id=memory_id,
                            )
                            st.success("Learning memory updated.")
                            st.rerun()
                with action_col2:
                    toggle_label = "Disable" if enabled else "Enable"
                    if st.button(toggle_label, key=f"toggle_memory_{memory_id}"):
                        set_learning_memory_enabled(memory_id, not bool(enabled))
                        st.rerun()
                with action_col3:
                    if st.button("Delete", key=f"delete_memory_{memory_id}", type="secondary"):
                        delete_learning_memory(memory_id)
                        st.error("Learning memory deleted.")
                        st.rerun()
