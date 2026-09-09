import os
from flask import Flask, request, jsonify, render_template, redirect, url_for, session
import openpyxl
import sqlite3

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
    if file.filename == '':
        return jsonify({'error': 'ഫയൽ തിരഞ്ഞെടുത്തിട്ടില്ല'}), 400

    try:
        if not file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({'error': 'ദയവായി Excel (.xlsx) ഫയൽ മാത്രം അപ്‌ലോഡ് ചെയ്യുക!'}), 400

        wb = openpyxl.load_workbook(file, data_only=True)
        sheet = wb.active

        # ഹെഡർ കണ്ടെത്തുന്നു
        headers = []
        header_row_idx = 1
        for i, row in enumerate(sheet.iter_rows(values_only=True), 1):
            row_str = [str(cell).strip() for cell in row if cell is not None]
            if 'Date' in row_str and 'Details' in row_str:
                headers = row_str
                header_row_idx = i
                break
        
        if not headers:
            # ഡിഫോൾട്ട് കോളങ്ങൾ എടുക്കുന്നു (അല്ലെങ്കിൽ ആദ്യ റോ)
            header_row_idx = 1
            headers = [str(cell.value).strip() if cell.value else '' for cell in sheet[1]]

        # കോളങ്ങളുടെ ഇൻഡക്സ് കണ്ടുപിടിക്കുന്നു
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

        # ഡാറ്റ റീഡ് ചെയ്ത് ഇൻസേർട്ട് ചെയ്യുന്നു
        for row in sheet.iter_rows(min_row=header_row_idx + 1, values_only=True):
            if not any(row): continue
            
            get_val = lambda key: str(row[col_map[key]]).strip() if key in col_map and col_map[key] < len(row) and row[col_map[key]] is not None else ''
            
            tx_date = get_val('Date')
            details = get_val('Details')
            ref_no = get_val('Ref No')
            
            try: debit = float(str(get_val('Debit')).replace(',', ''))
            except: debit = 0.0
            
            try: credit = float(str(get_val('Credit')).replace(',', ''))
            except: credit = 0.0
            
            try: balance = float(str(get_val('Balance')).replace(',', ''))
            except: balance = 0.0

            if tx_date or details:
                cursor.execute('''
                    INSERT INTO bank_transactions (tx_date, details, ref_no, debit, credit, balance)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (tx_date, details, ref_no, debit, credit, balance))
        
        conn.commit()
        conn.close()

        return jsonify({'success': 'സ്റ്റേറ്റ്‌മെന്റ് വിജയകരമായി സേവ് ചെയ്തു!'})

    except Exception as e:
        return jsonify({'error': f'പ്രോസസ്സ് ചെയ്യുന്നതിൽ പിശക്: {str(e)}'}), 500

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
