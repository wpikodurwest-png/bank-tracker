import os
from flask import Flask, request, jsonify, render_template, redirect, url_for, session
import pandas as pd
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
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file)
        elif file.filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file)
        else:
            return jsonify({'error': 'തെറ്റായ ഫയൽ ഫോർമാറ്റ്!'}), 400

        df.columns = df.columns.str.strip()
        expected_cols = ['Date', 'Details', 'Ref No/Cheque No', 'Debit', 'Credit', 'Balance']
        for col in expected_cols:
            if col not in df.columns:
                df[col] = ''

        df = df.fillna('')

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM bank_transactions')

        for _, row in df.iterrows():
            deb = float(str(row['Debit']).replace(',', '')) if str(row['Debit']).replace('.','',1).isdigit() else 0
            cred = float(str(row['Credit']).replace(',', '')) if str(row['Credit']).replace('.','',1).isdigit() else 0
            bal = float(str(row['Balance']).replace(',', '')) if str(row['Balance']).replace('.','',1).isdigit() else 0

            cursor.execute('''
                INSERT INTO bank_transactions (tx_date, details, ref_no, debit, credit, balance)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (str(row['Date']), str(row['Details']), str(row['Ref No/Cheque No']), deb, cred, bal))
        
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