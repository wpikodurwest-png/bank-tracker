import tempfile

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

        # ഫയൽ സെർവറിലെ ടെമ്പ് ഫോൾഡറിലേക്ക് താൽക്കാലികമായി സേവ് ചെയ്യുന്നു
        fd, temp_path = tempfile.mkstemp()
        os.close(fd)
        file.save(temp_path)

        wb = openpyxl.load_workbook(temp_path, data_only=True)
        sheet = wb.active

        # താൽക്കാലിക ഫയൽ ഡിലീറ്റ് ചെയ്യാൻ തയ്യാറാക്കുക
        try:
            headers = []
            header_row_idx = 1
            for i, row in enumerate(sheet.iter_rows(values_only=True), 1):
                row_str = [str(cell).strip() for cell in row if cell is not None]
                if 'Date' in row_str and 'Details' in row_str:
                    headers = row_str
                    header_row_idx = i
                    break
            
            if not headers:
                header_row_idx = 1
                headers = [str(cell.value).strip() if cell.value else '' for cell in sheet[1]]

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

        finally:
            # താൽക്കാലിക ഫയൽ ക്ലീൻ ചെയ്യുന്നു
            if os.path.exists(temp_path):
                os.remove(temp_path)

        return jsonify({'success': 'സ്റ്റേറ്റ്‌മെന്റ് വിജയകരമായി സേവ് ചെയ്തു!'})

    except Exception as e:
        return jsonify({'error': f'പ്രോസസ്സ് ചെയ്യുന്നതിൽ പിശക്: {str(e)}'}), 500
