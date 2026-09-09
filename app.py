import os
import tempfile
from flask import Flask, request, jsonify, render_template, redirect, url_for, session
import openpyxl
import msoffcrypto
import sqlite3
import pandas as pd

app = Flask(__name__, template_folder='templates')
app.secret_key = os.environ.get('SECRET_KEY', 'your_secret_key_here')
DB_NAME = 'database.db'

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bank_transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tx_date TEXT,
            details TEXT,
            ref_no TEXT,
            debit REAL,
            credit REAL,
            balance REAL
        )
    ''')
    conn.commit()
    conn.close()

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if username == 'admin' and password == 'admin123':
            session['logged_in'] = True
            return redirect(url_for('index'))
        return render_template('login.html', error='തെറ്റായ യൂസർനെയിം അല്ലെങ്കിൽ പാസ്‌വേഡ്!')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if not session.get('logged_in'):
        return jsonify({'error': 'Unauthorized'}), 401

    if 'file' not in request.files:
        return jsonify({'error': 'ഫയൽ അപ്‌ലോഡ് ചെയ്തിട്ടില്ല'}), 400
    
    file = request.files['file']
    excel_password = request.form.get('password', '').strip()

    if file.filename == '':
        return jsonify({'error': 'ഫയൽ തിരഞ്ഞെടുത്തിട്ടില്ല'}), 400

    try:
        fd, temp_path = tempfile.mkstemp(suffix=os.path.splitext(file.filename)[1])
        os.close(fd)
        file.save(temp_path)

        target_path = temp_path

        # 1. പാസ്‌വേഡ് ഉണ്ടെങ്കിൽ msoffcrypto ഉപയോഗിച്ച് ഡീക്രിപ്റ്റ് ചെയ്യാൻ ശ്രമിക്കുന്നു
        if excel_password:
            try:
                with open(temp_path, "rb") as f:
                    file_decrypted = msoffcrypto.OfficeFile(f)
                    if file_decrypted.is_encrypted():
                        file_decrypted.load_key(password=excel_password)
                        decrypted_fd, decrypted_path = tempfile.mkstemp(suffix=".xlsx")
                        os.close(decrypted_fd)
                        with open(decrypted_path, "wb") as decrypted_file:
                            file_decrypted.decrypt(decrypted_file)
                        target_path = decrypted_path
            except Exception as pwd_err:
                return jsonify({'error': f'പാസ്‌വേഡ് തെറ്റാണ് അല്ലെങ്കിൽ ഫയൽ അൺലോക്ക് ചെയ്യാൻ കഴിഞ്ഞില്ല: {str(pwd_err)}'}), 400

        # 2. ബാങ്ക് ഫയൽ HTML ഫോർമാറ്റാണോ എന്ന് പരിശോധിക്കുന്നു (ചില ബാങ്കുകൾ .xls എക്സ്റ്റൻഷനിൽ HTML ആണ് തരുന്നത്)
        try:
            with open(target_path, 'rb') as f:
                header_bytes = f.read(100)
                if b'<html' in header_bytes.lower() or b'<table' in header_bytes.lower():
                    # HTML ടേബിൾ ആയിട്ടുള്ള ഫയൽ പാണ്ടസ് ഉപയോഗിച്ച് റീഡ് ചെയ്യുന്നു
                    dfs = pd.read_html(target_path)
                    if dfs:
                        df = dfs[0]
                        # ഡാറ്റ ക്ലീൻ ചെയ്ത് ഡാറ്റാബേസിലേക്ക് മാറ്റുന്നു
                        return save_dataframe_to_db(df)
        except Exception:
            pass

        # 3. സാധാരണ എക്സൽ ഫയൽ ആണെങ്കിൽ openpyxl ഉപയോഗിച്ച് റീഡ് ചെയ്യുന്നു
        try:
            wb = openpyxl.load_workbook(target_path, data_only=True)
            sheet = wb.active
        except Exception as e:
            return jsonify({'error': f'ഫയൽ ഫോർമാറ്റ് വായിക്കാൻ കഴിഞ്ഞില്ല. ദയവായി ഫയൽ മൈക്രോസോഫ്റ്റ് എക്സലിൽ തുറന്ന് സാധാരണ .xlsx ഫോർമാറ്റിൽ സേവ് ചെയ്ത് അപ്‌ലോഡ് ചെയ്യുക.'}), 400

        headers = []
        header_row_idx = 1
        for i, row in enumerate(sheet.iter_rows(values_only=True), 1):
            row_str = [str(cell).strip() for cell in row if cell is not None]
            if any('date' in s.lower() for s in row_str) and any('detail' in s.lower() for s in row_str):
                headers = row_str
                header_row_idx = i
                break
        
        if not headers:
            header_row_idx = 8
            headers = [str(cell.value).strip() if cell.value else '' for cell in sheet[header_row_idx]]

        col_map = {}
        for idx, h in enumerate(headers):
            h_lower = h.lower()
            if 'date' in h_lower: col_map['Date'] = idx
            elif 'detail' in h_lower: col_map['Details'] = idx
            elif 'ref' in h_lower or 'cheque' in h_lower: col_map['Ref No'] = idx
            elif 'debit' in h_lower: col_map['Debit'] = idx
            elif 'credit' in h_lower: col_map['Credit'] = idx
            elif 'balance' in h_lower: col_map['Balance'] = idx

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM bank_transactions')

        for row in sheet.iter_rows(min_row=header_row_idx + 1, values_only=True):
            if not any(row): continue
            
            get_val = lambda key: str(row[col_map[key]]).strip() if key in col_map and col_map[key] < len(row) and row[col_map[key]] is not None else ''
            
            tx_date = get_val('Date')
            details = get_val('Details')
            ref_no = get_val('Ref No')
            
            if not tx_date or ('/' not in tx_date and '-' not in tx_date):
                continue

            try: debit = float(str(get_val('Debit')).replace(',', ''))
            except: debit = 0.0
            
            try: credit = float(str(get_val('Credit')).replace(',', ''))
            except: credit = 0.0
            
            try: balance = float(str(get_val('Balance')).replace(',', ''))
            except: balance = 0.0

            cursor.execute('''
                INSERT INTO bank_transactions (tx_date, details, ref_no, debit, credit, balance)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (tx_date, details, ref_no, debit, credit, balance))
        
        conn.commit()
        conn.close()

        if os.path.exists(temp_path): os.remove(temp_path)
        if 'decrypted_path' in locals() and decrypted_path and os.path.exists(decrypted_path):
            os.remove(decrypted_path)

        return jsonify({'success': 'സ്റ്റേറ്റ്‌മെന്റ് വിജയകരമായി സേവ് ചെയ്തു!'})

    except Exception as e:
        return jsonify({'error': f'പ്രോസസ്സ് ചെയ്യുന്നതിൽ പിശക്: {str(e)}'}), 500

def save_dataframe_to_db(df):
    # HTML ടേബിൾ ആയി വരുന്ന ബാങ്ക് സ്റ്റേറ്റ്‌മെന്റുകൾ കൈകാര്യം ചെയ്യാൻ
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM bank_transactions')

    # ഡാറ്റയിലെ കോളം പേരുകൾ കണ്ടെത്തുന്നു
    df.columns = [str(c).strip().lower() for c in df.columns]
    
    date_col = next((c for c in df.columns if 'date' in c), None)
    detail_col = next((c for c in df.columns if 'detail' in c or 'particular' in c), None)
    ref_col = next((c for c in df.columns if 'ref' in c or 'cheque' in c), None)
    debit_col = next((c for c in df.columns if 'debit' in c), None)
    credit_col = next((c for c in df.columns if 'credit' in c), None)
    balance_col = next((c for c in df.columns if 'balance' in c), None)

    for _, row in df.iterrows():
        tx_date = str(row[date_col]) if date_col and pd.notna(row[date_col]) else ''
        if not tx_date or ('/' not in tx_date and '-' not in tx_date):
            continue
            
        details = str(row[detail_col]) if detail_col and pd.notna(row[detail_col]) else ''
        ref_no = str(row[ref_col]) if ref_col and pd.notna(row[ref_col]) else ''
        
        try: debit = float(str(row[debit_col]).replace(',', '')) if debit_col and pd.notna(row[debit_col]) else 0.0
        except: debit = 0.0

        try: credit = float(str(row[credit_col]).replace(',', '')) if credit_col and pd.notna(row[credit_col]) else 0.0
        except: credit = 0.0

        try: balance = float(str(row[balance_col]).replace(',', '')) if balance_col and pd.notna(row[balance_col]) else 0.0
        except: balance = 0.0

        cursor.execute('''
            INSERT INTO bank_transactions (tx_date, details, ref_no, debit, credit, balance)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (tx_date, details, ref_no, debit, credit, balance))

    conn.commit()
    conn.close()
    return jsonify({'success': 'സ്റ്റേറ്റ്‌മെന്റ് വിജയകരമായി സേവ് ചെയ്തു!'})

@app.route('/get_transactions', methods=['GET'])
def get_transactions():
    if not session.get('logged_in'):
        return jsonify({'error': 'Unauthorized'}), 401

    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute('SELECT tx_date as Date, details as Details, ref_no as RefNo, debit as Debit, credit as Credit, balance as Balance FROM bank_transactions')
    rows = cursor.fetchall()
    conn.close()

    result = [dict(row) for row in rows]
    return jsonify(result)

if __name__ == '__main__':
    init_db()
    app.run(debug=True, port=5000)
