import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import io
import xlsxwriter

# --- Core Algorithm from v4 ---

date_pattern = re.compile(r"^\d{2}/\d{2}/\d{2}$")

def clean_amount(val):
    val = val.strip().replace(',', '')
    if not val: return None
    match = re.search(r"(-?\d+\.\d{2})", val)
    if match:
        return float(match.group(1))
    return val

def format_date(val):
    val = val.strip()
    if not val: return None
    try:
        dt = datetime.strptime(val, "%d/%m/%y")
        return dt 
    except:
        return val

def process_statement(pdf):
    txns = []
    current_txn = None
    account_no = ""
    period_from = ""
    period_to = ""

    last_page_text = pdf.pages[-1].extract_text()
    first_page_text = pdf.pages[0].extract_text()

    for page_idx, page in enumerate(pdf.pages):
        words = page.extract_words()
        
        if page_idx == 0:
            acc_match = re.search(r"AccountNo\s*:\s*(\d+)", first_page_text)
            if acc_match: account_no = acc_match.group(1)
            per_match = re.search(r"From\s*:\s*(\d{2}/\d{2}/\d{4})\s*To\s*:\s*(\d{2}/\d{2}/\d{4})", first_page_text)
            if per_match:
                period_from = per_match.group(1)
                period_to = per_match.group(2)
        
        # Group words into lines
        words.sort(key=lambda w: (w['top'], w['x0']))
        lines = []
        current_line = []
        current_top = None
        
        for w in words:
            # Fixed bounds that perfectly isolate transactions on all pages
            if w['top'] < 225 or w['top'] > 780:
                continue
                
            if current_top is None:
                current_top = w['top']
                current_line.append(w)
            elif abs(w['top'] - current_top) <= 3:
                current_line.append(w)
            else:
                lines.append(current_line)
                current_line = [w]
                current_top = w['top']
        if current_line:
            lines.append(current_line)
            
        for line in lines:
            line.sort(key=lambda w: w['x0'])
            if not line:
                continue
                
            first_word = line[0]
            is_new_txn = False
            if first_word['x0'] < 60 and date_pattern.match(first_word['text']):
                is_new_txn = True
                
            if is_new_txn:
                if current_txn:
                    txns.append(current_txn)
                current_txn = {
                    'Date': '', 'Narration': '', 'Chq./Ref.No.': '',
                    'Value Dt': '', 'Withdrawal Amt.': '', 'Deposit Amt.': '', 'Closing Balance': ''
                }
            
            if current_txn is not None:
                for w in line:
                    x = w['x0']
                    t_text = w['text']
                    if x < 60:
                        current_txn['Date'] += t_text + " "
                    elif 60 <= x < 280:
                        current_txn['Narration'] += t_text + " "
                    elif 280 <= x < 360:
                        current_txn['Chq./Ref.No.'] += t_text + " "
                    elif 360 <= x < 420:
                        current_txn['Value Dt'] += t_text + " "
                    elif 420 <= x < 500:
                        current_txn['Withdrawal Amt.'] += t_text + " "
                    elif 500 <= x < 570:
                        current_txn['Deposit Amt.'] += t_text + " "
                    else:
                        current_txn['Closing Balance'] += t_text + " "
                        
    if current_txn:
        txns.append(current_txn)

    cleaned_txns = []
    for t in txns:
        d = {}
        d['Date'] = format_date(t['Date'])
        
        nar = t['Narration'].strip()
        if 'STATEMENTSUMMARY' in nar.replace(" ", ""):
            continue # Drop the row completely as requested
            
        d['Narration'] = nar
        d['Chq./Ref.No.'] = t['Chq./Ref.No.'].strip()
        d['Value Dt'] = format_date(t['Value Dt'])
        d['Withdrawal Amt.'] = clean_amount(t['Withdrawal Amt.'])
        d['Deposit Amt.'] = clean_amount(t['Deposit Amt.'])
        d['Closing Balance'] = clean_amount(t['Closing Balance'])
        
        if isinstance(d['Closing Balance'], float):
            cleaned_txns.append(d)

    # Calculations
    first_txn = cleaned_txns[0]
    first_close = first_txn['Closing Balance'] or 0.0
    first_with = first_txn['Withdrawal Amt.'] or 0.0
    first_dep = first_txn['Deposit Amt.'] or 0.0

    calc_opening_balance = first_close + first_with - first_dep
    calc_dr_count = sum(1 for t in cleaned_txns if isinstance(t['Withdrawal Amt.'], float))
    calc_cr_count = sum(1 for t in cleaned_txns if isinstance(t['Deposit Amt.'], float))
    calc_debits = sum(t['Withdrawal Amt.'] for t in cleaned_txns if isinstance(t['Withdrawal Amt.'], float))
    calc_credits = sum(t['Deposit Amt.'] for t in cleaned_txns if isinstance(t['Deposit Amt.'], float))
    calc_closing_bal = cleaned_txns[-1]['Closing Balance']

    try:
        op_dt = datetime.strptime(period_from, "%d/%m/%Y")
    except:
        op_dt = period_from

    op_row = {
        'Date': op_dt,
        'Narration': 'Opening balance as per bank statement',
        'Chq./Ref.No.': '',
        'Value Dt': '',
        'Withdrawal Amt.': None,
        'Deposit Amt.': None,
        'Closing Balance': calc_opening_balance
    }
    cleaned_txns.insert(0, op_row)

    final_row = {
        'Date': '',
        'Narration': 'Total / Closing Balance',
        'Chq./Ref.No.': '',
        'Value Dt': '',
        'Withdrawal Amt.': calc_debits,
        'Deposit Amt.': calc_credits,
        'Closing Balance': calc_closing_bal
    }
    cleaned_txns.append(final_row)

    df_txns = pd.DataFrame(cleaned_txns)
    df_txns.rename(columns={'Closing Balance': 'Openingbalance /closing balance'}, inplace=True)

    # Extract dynamic Account Details using simple regex on first page
    def ext_val(regex, text):
        m = re.search(regex, text)
        return m.group(1).strip() if m else ""

    bank_name = ext_val(r"(HDFC\s*BANK\s*LIMITED)", last_page_text) or "HDFC BANK LIMITED"
    
    # Dynamic crop for Customer Name & Address (Left block)
    try:
        left_block = pdf.pages[0].crop((0, 85, 300, 170)).extract_text()
        if left_block:
            lines = left_block.split('\n')
            cust_name = lines[0].strip() if len(lines) > 0 else ""
            cust_address = " ".join([l.strip() for l in lines[1:]])
        else:
            cust_name = ""
            cust_address = ""
    except:
        cust_name = ""
        cust_address = ""

    # Dynamic crop for Branch Address (Right block)
    try:
        right_block = pdf.pages[0].crop((300, 20, 600, 100)).extract_text()
        # Clean up the branch address text
        branch_address = right_block.replace('\n', ' ').replace('ageNo.:1', '').strip() if right_block else ""
    except:
        branch_address = ""
    
    account_details = [
        ("Bank Name", bank_name),
        ("Customer Name", cust_name),
        ("Customer Address", cust_address),
        ("Account Branch", ext_val(r"AccountBranch\s*:\s*([^\n]+)", first_page_text)),
        ("Branch Address", branch_address),
        ("Phone no.", ext_val(r"Phoneno\.\s*:\s*([^\n]+)", first_page_text)),
        ("OD Limit", ext_val(r"ODLimit\s*:\s*([^\n]+)", first_page_text)),
        ("Currency", ext_val(r"Currency\s*:\s*([^\n]+)", first_page_text)),
        ("Email", ext_val(r"Email\s*:\s*([^\n]+)", first_page_text)),
        ("Cust ID", ext_val(r"CustID\s*:\s*([^\n]+)", first_page_text)),
        ("Account No", account_no),
        ("A/C Open Date", ext_val(r"A/COpenDate\s*:\s*([^\n]+)", first_page_text)),
        ("Account Status", ext_val(r"AccountStatus\s*:\s*([^\n]+)", first_page_text)),
        ("RTGS/NEFT IFSC", ext_val(r"RTGS/NEFTIFSC:\s*([^\n]+)", first_page_text)),
        ("Branch Code", ext_val(r"BranchCode\s*:\s*([^\n]+)", first_page_text)),
        ("Account Type", ext_val(r"AccountType\s*:\s*([^\n]+)", first_page_text))
    ]
    df_acc = pd.DataFrame(account_details, columns=["Field", "Value"])

    # Extract stated summary
    def ext_num(regex, text):
        m = re.search(regex, text)
        if m:
            val = m.group(1).replace(',', '').strip()
            try: return float(val)
            except: pass
        return 0.0

    stated_opening = ext_num(r"OpeningBalance\s+([\d,.]+)", last_page_text)
    stated_dr = ext_num(r"DrCount\s+([\d,.]+)", last_page_text)
    stated_cr = ext_num(r"CrCount\s+([\d,.]+)", last_page_text)
    stated_debits = ext_num(r"Debits\s+([\d,.]+)", last_page_text)
    stated_credits = ext_num(r"Credits\s+([\d,.]+)", last_page_text)
    stated_closing = ext_num(r"ClosingBal\s+([\d,.]+)", last_page_text)

    summary_data = [
        ("Metric", "Statement PDF", "Script Calculated", "Difference"),
        ("Opening Balance", stated_opening, calc_opening_balance, stated_opening - calc_opening_balance),
        ("Dr Count", stated_dr, calc_dr_count, stated_dr - calc_dr_count),
        ("Cr Count", stated_cr, calc_cr_count, stated_cr - calc_cr_count),
        ("Debits", stated_debits, calc_debits, stated_debits - calc_debits),
        ("Credits", stated_credits, calc_credits, stated_credits - calc_credits),
        ("Closing Bal", stated_closing, calc_closing_bal, stated_closing - calc_closing_bal)
    ]
    df_sum = pd.DataFrame(summary_data[1:], columns=summary_data[0])

    return df_acc, df_txns, df_sum, account_no, period_from, period_to, cleaned_txns

def create_excel_buffer(df_acc, df_txns, df_sum, account_no, period_from, period_to):
    buffer = io.BytesIO()
    writer = pd.ExcelWriter(buffer, engine='xlsxwriter')

    df_acc.to_excel(writer, index=False, sheet_name='Account Details', header=False)
    df_txns.to_excel(writer, index=False, sheet_name='Transactions')
    df_sum.to_excel(writer, index=False, sheet_name='Summary Analysis')

    workbook  = writer.book
    ws_acc = writer.sheets['Account Details']
    ws_txns = writer.sheets['Transactions']
    ws_sum = writer.sheets['Summary Analysis']

    header_format = workbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 12})
    cell_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 12, 'valign': 'vcenter'})
    date_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 12, 'valign': 'vcenter', 'num_format': 'dd-mmm-yyyy'})
    text_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 12, 'valign': 'vcenter', 'text_wrap': True})
    rupee_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 12, 'valign': 'vcenter', 'num_format': '#,##0.00'})
    diff_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 12, 'valign': 'vcenter', 'num_format': '0.00'})

    ws_acc.set_column('A:A', 25, header_format)
    ws_acc.set_column('B:B', 80, cell_format)

    ws_txns.set_column('A:A', 15, date_format)
    ws_txns.set_column('B:B', 45, text_format)
    ws_txns.set_column('C:C', 20, text_format)
    ws_txns.set_column('D:D', 15, date_format)
    ws_txns.set_column('E:E', 18, rupee_format)
    ws_txns.set_column('F:F', 18, rupee_format)
    ws_txns.set_column('G:G', 22, rupee_format)
    for col_num, value in enumerate(df_txns.columns.values):
        ws_txns.write(0, col_num, value, header_format)

    ws_sum.set_column('A:A', 20, header_format)
    ws_sum.set_column('B:B', 20, rupee_format)
    ws_sum.set_column('C:C', 20, rupee_format)
    ws_sum.set_column('D:D', 20, diff_format)
    for col_num, value in enumerate(df_sum.columns.values):
        ws_sum.write(0, col_num, value, header_format)

    writer.close()
    return buffer

def create_ca_template_buffer(cleaned_txns):
    buffer = io.BytesIO()
    writer = pd.ExcelWriter(buffer, engine='xlsxwriter')

    # Prepare CA Template Data
    ca_data = []
    # Drop the first (opening balance) and last (totals) rows to only include actual transactions
    actual_txns = cleaned_txns[1:-1]
    
    for t in actual_txns:
        # Date string formatting to strictly DD/MM/YYYY
        dt = t['Date']
        dt_str = dt.strftime("%d/%m/%Y") if isinstance(dt, datetime) else dt
        
        debit = t['Withdrawal Amt.'] if pd.notnull(t['Withdrawal Amt.']) and t['Withdrawal Amt.'] else ""
        credit = t['Deposit Amt.'] if pd.notnull(t['Deposit Amt.']) and t['Deposit Amt.'] else ""
        
        ca_data.append({
            "Date": dt_str,
            "Transaction": t['Narration'],
            "ID": t['Chq./Ref.No.'],
            "Debit": debit,
            "Credit": credit
        })
        
    df_ca = pd.DataFrame(ca_data)
    
    # Sheet 1: Template
    df_ca.to_excel(writer, index=False, sheet_name='Template')
    
    # Sheet 2: Instructions (blank or minimal instructions as requested)
    df_inst = pd.DataFrame({
        "Instructions": [
            "1. Do not modify the header row in the 'Template' sheet.",
            "2. Date Format: All entries must be in DD/MM/YYYY format.",
            "3. Enter amounts in either Debit or Credit, but never both.",
            "4. Leave non-applicable cells completely blank.",
            "5. Provide clear descriptions in Transaction column.",
            "6. Provide references in ID column."
        ]
    })
    df_inst.to_excel(writer, index=False, sheet_name='Instructions')

    # Formatting
    workbook = writer.book
    ws_temp = writer.sheets['Template']
    ws_inst = writer.sheets['Instructions']

    header_format = workbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 11, 'bottom': 1})
    cell_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11})
    num_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'num_format': '0.00'})

    ws_temp.set_column('A:A', 15, cell_format)
    ws_temp.set_column('B:B', 50, cell_format)
    ws_temp.set_column('C:C', 20, cell_format)
    ws_temp.set_column('D:D', 18, num_format)
    ws_temp.set_column('E:E', 18, num_format)
    
    for col_num, value in enumerate(df_ca.columns.values):
        ws_temp.write(0, col_num, value, header_format)

    ws_inst.set_column('A:A', 80, cell_format)

    writer.close()
    return buffer

# --- Streamlit UI ---

st.set_page_config(page_title="HDFC Statement Converter", page_icon="🏦", layout="centered")

st.title("🏦 HDFC Bank Statement to Excel Converter")
st.markdown("Upload your HDFC bank statement PDF below to accurately convert it to a fully-formatted, 3-sheet Excel workbook. This tool uses advanced dynamic boundary algorithms to prevent missing data.")

uploaded_file = st.file_uploader("Upload PDF Statement", type=["pdf"])
password = st.text_input("Document Password (if any)", type="password", help="If your statement is locked, enter the password here.")

if uploaded_file is not None:
    if st.button("Process PDF Statement", type="primary"):
        with st.spinner("Analyzing and parsing PDF..."):
            try:
                with pdfplumber.open(uploaded_file, password=password) as pdf:
                    df_acc, df_txns, df_sum, account_no, period_from, period_to, raw_txns = process_statement(pdf)
                    
                    st.success(f"✅ Successfully processed {len(df_txns) - 2} transactions!")
                    st.markdown(f"**Account**: `{account_no}` | **Period**: `{period_from}` to `{period_to}`")
                    
                    st.subheader("Summary Analysis Check")
                    st.dataframe(df_sum)
                    
                    # Buffer for the main 3-sheet report
                    buffer_main = create_excel_buffer(df_acc, df_txns, df_sum, account_no, period_from, period_to)
                    
                    # Buffer for the CA Tech Club template
                    buffer_ca = create_ca_template_buffer(raw_txns)
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.download_button(
                            label="📥 Download Full Analysis Report",
                            data=buffer_main.getvalue(),
                            file_name=f"HDFC_Statement_{account_no}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            help="The full 3-sheet workbook containing Account Details, Transactions, and the Summary Analysis."
                        )
                        
                    with col2:
                        st.download_button(
                            label="📥 Download CA Tech Club Template",
                            data=buffer_ca.getvalue(),
                            file_name=f"Transaction_Template_CA_Tech_Club_{account_no}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="secondary",
                            help="Strictly formatted for your internal processing system."
                        )
            except Exception as e:
                error_msg = str(e)
                if "PasswordIncorrect" in str(type(e).__name__) or "password" in error_msg.lower() or "decipher" in error_msg.lower():
                    st.error("🔒 **This PDF is password protected.** Please enter the correct password in the field above and try again.")
                else:
                    st.error(f"❌ **An error occurred while processing:** {e}")
                    st.exception(e)

st.markdown("---")
st.caption("Developed securely. All processing happens entirely within this environment, and files are never permanently saved.")
