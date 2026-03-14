"""
Customer Payment Manager
A simple interface to look up customers and record payments
"""

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import json
import calendar

import os

# === NEW 12-MONTH AGING SYSTEM ===
MONTHS_2026 = ['Jan_2026', 'Feb_2026', 'Mar_2026', 'Apr_2026', 'May_2026', 'Jun_2026', 
               'Jul_2026', 'Aug_2026', 'Sep_2026', 'Oct_2026', 'Nov_2026', 'Dec_2026']

# Map month number to column name
MONTH_MAP = {
    1: 'Jan_2026', 2: 'Feb_2026', 3: 'Mar_2026', 4: 'Apr_2026',
    5: 'May_2026', 6: 'Jun_2026', 7: 'Jul_2026', 8: 'Aug_2026',
    9: 'Sep_2026', 10: 'Oct_2026', 11: 'Nov_2026', 12: 'Dec_2026'
}

def get_monthly_balances(customer):
    balances = {}
    for month in MONTHS_2026:
        val = customer.get(month, 0)
        try:
            balances[month] = float(val) if val else 0
        except:
            balances[month] = 0
    return balances

def get_total_balance_from_months(balances):
    return sum(balances.values())

def update_amount_due_from_months(sheet_name, customer_name):
    """Update the Amount Due column (H) with sum of all 12 months (N-Y)"""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:
                total = 0
                for col_idx in range(13, 25):
                    try:
                        val = row[col_idx].value
                        if val:
                            total += float(val)
                    except:
                        pass
                row[7].value = total
                break
        
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        return False
      
def save_monthly_balance(sheet_name, customer_name, month_label, amount):
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        month_col_idx = None
        for col_idx, col in enumerate(ws[1], start=1):
            if col.value == month_label:
                month_col_idx = col_idx
                break
        if month_col_idx is None:
            return False
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:
                row[month_col_idx - 1].value = amount
                break
        wb.save(EXCEL_FILE)
        return True
    except:
        return False

# Configuration
EXCEL_FILE_LOCAL = "cleaned_billing_by_service.xlsx"

# Use persistent disk on Render, or local file for development
RENDER_DISK_PATH = "/app/data"
RENDER_SRC_PATH = "/opt/render/project/src"
if os.path.exists(RENDER_DISK_PATH):
    disk_file = os.path.join(RENDER_DISK_PATH, "cleaned_billing_by_service.xlsx")
    if os.path.exists(disk_file):
        EXCEL_FILE = disk_file
    else:
        # Check alternate location where uploads go
        src_file = os.path.join(RENDER_SRC_PATH, "cleaned_billing_by_service.xlsx")
        if os.path.exists(src_file):
            EXCEL_FILE = src_file
        else:
            EXCEL_FILE = EXCEL_FILE_LOCAL
elif os.path.exists(RENDER_SRC_PATH):
    EXCEL_FILE = os.path.join(RENDER_SRC_PATH, "cleaned_billing_by_service.xlsx")
else:
    EXCEL_FILE = EXCEL_FILE_LOCAL

st.set_page_config(page_title="Customer Payment Manager", page_icon="💳", layout="wide")

def save_customer_notes(sheet_name, customer_name, notes):
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:
                row[10].value = notes
                break
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Error saving notes: {e}")
        return False

def save_notes2(sheet_name, customer_name, notes2):
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:
                row[12].value = notes2
                break
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Error saving notes2: {e}")
        return False

def save_due_date(sheet_name, customer_name, due_day):
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:
                row[11].value = due_day
                break
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Error saving due date: {e}")
        return False

def save_customer_info(sheet_name, customer_name, new_name, phone, card_number, exp, cvv, plan_cost, original_phone=None, row_index=None):
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        target_row = None
        if row_index is not None:
            excel_row_num = row_index + 2
            for row in ws.iter_rows(min_row=excel_row_num, max_row=excel_row_num):
                target_row = row
                break
        else:
            for row in ws.iter_rows(min_row=2):
                if str(row[3].value).strip() == str(customer_name).strip():
                    target_row = row
                    break
        if target_row is None:
            return False
        target_row[2].value = plan_cost
        target_row[3].value = new_name
        target_row[4].value = card_number
        target_row[5].value = exp
        target_row[6].value = cvv
        target_row[9].value = phone
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Error saving customer info: {e}")
        return False

def get_balance(customer):
    # First check if Amount Due column has value, otherwise calculate from monthly
    amount_due = customer.get('Amount Due', '')
    
    # If Amount Due is empty/0, calculate from monthly balances
    if not amount_due or str(amount_due).strip() == '' or str(amount_due).strip().lower() == 'nan':
        total = 0
        for month in MONTHS_2026:
            val = customer.get(month, 0)
            try:
                total += float(val) if val else 0
            except:
                pass
        return total
    if amount_due and str(amount_due).strip() and str(amount_due).strip().lower() != 'nan':
        try:
            val = float(str(amount_due).strip())
            return val
        except:
            pass
    return 0

# Custom styling
st.markdown("""
    <style>
    .main { background-color: #f5f5f5; }
    .stButton>button { width: 100%; }
    .customer-card {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

def load_excel():
    try:
        excel_file = pd.ExcelFile(EXCEL_FILE)
        all_data = {}
        for sheet in excel_file.sheet_names:
            if sheet != "Summary":
                df = pd.read_excel(excel_file, sheet_name=sheet)
                df['Service'] = sheet
                if 'Due Day' not in df.columns:
                    if 'Charge Date' in df.columns:
                        try:
                            df['Charge Date'] = pd.to_datetime(df['Charge Date'], errors='coerce')
                            df['Due Day'] = df['Charge Date'].dt.day
                        except:
                            df['Due Day'] = None
                all_data[sheet] = df
        return all_data, excel_file
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None, None

def save_payment(sheet_name, customer_name, payment_amount, notes=""):
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:
                current_amount_due = row[7].value
                try:
                    current_balance = float(str(current_amount_due).strip()) if current_amount_due and str(current_amount_due).strip() else 0
                except:
                    current_balance = 0
                new_balance = current_balance - payment_amount
                row[7].value = new_balance
                row[8].value = "Paid" if new_balance <= 0 else "Partial"
                today = datetime.now().strftime('%Y-%m-%d')
                row[11].value = today
                row[12].value = (datetime.now() + timedelta(days=30)).day
                existing_notes = str(row[10].value) if row[10].value else ""
                payment_info = f"Payment ${payment_amount:.2f} on {today}"
                if notes:
                    payment_info += f": {notes}"
                row[10].value = (existing_notes + " | " + payment_info) if existing_notes else payment_info
                break
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Error saving: {e}")
        return False

def search_customers(all_data, query):
    results = []
    for service, df in all_data.items():
        if df is None or df.empty:
            continue
        mask = (
            df['Customer Name'].astype(str).str.contains(query, case=False, na=False) |
            df['Phone'].astype(str).str.contains(query, case=False, na=False) |
            df['Card Number'].astype(str).str.contains(query, case=False, na=False)
        )
        matches = df[mask]
        for idx, row in matches.iterrows():
            # Get monthly balances
            monthly = {}
            for m in MONTHS_2026:
                monthly[m] = row.get(m, 0)
            
            result = {
                'Service': service,
                'row_index': idx,
                'Customer Name': row.get('Customer Name', ''),
                'Phone': row.get('Phone', ''),
                'Amount Due': row.get('Amount Due', ''),
                'Status': row.get('Status', ''),
                'Card Number': row.get('Card Number', ''),
                'Exp': row.get('Exp', ''),
                'CVV': row.get('CVV', ''),
                'Plan Cost': row.get('Plan Cost', ''),
                'Charge Date': row.get('Charge Date', ''),
                'Due Day': row.get('Due Day', ''),
                'Notes': row.get('Notes', ''),
                'Notes2': row.get('Notes2', ''),
                'Payment Date': row.get('Payment Date', ''),
            }
            result.update(monthly)
            results.append(result)
    return results

def get_customers_by_due_day(all_data, due_day):
    results = []
    for service, df in all_data.items():
        if df is None or df.empty:
            continue
        if 'Due Day' in df.columns:
            matches = df[df['Due Day'] == due_day]
            for _, row in matches.iterrows():
                monthly = {}
                for m in MONTHS_2026:
                    monthly[m] = row.get(m, 0)
                result = {
                    'Service': service,
                    'Customer Name': row.get('Customer Name', ''),
                    'Phone': row.get('Phone', ''),
                    'Amount Due': row.get('Amount Due', ''),
                    'Status': row.get('Status', ''),
                    'Card Number': row.get('Card Number', ''),
                    'Exp': row.get('Exp', ''),
                    'CVV': row.get('CVV', ''),
                    'Plan Cost': row.get('Plan Cost', ''),
                    'Charge Date': row.get('Charge Date', ''),
                    'Due Day': row.get('Due Day', ''),
                    'Notes': row.get('Notes', ''),
                    'Notes2': row.get('Notes2', ''),
                    'Payment Date': row.get('Payment Date', ''),
                }
                result.update(monthly)
                results.append(result)
    return results

def auto_charge_due_today():
    """Auto-charge all customers whose due date matches today"""
    today = datetime.now()
    current_day = today.day
    current_month = today.month
    
    month_key = MONTH_MAP.get(current_month, 'Mar_2026')
    
    debug_info = [f"Today: day={current_day}, month={current_month}, month_key={month_key}", f"EXCEL_FILE: {EXCEL_FILE}"]
    charged = []
    try:
        wb = load_workbook(EXCEL_FILE)
        
        for sheet_name in wb.sheetnames:
            if sheet_name == 'Summary':
                continue
            ws = wb[sheet_name]
            
            headers = {col.value: idx for idx, col in enumerate(ws[1], start=1)}
            due_day_col = headers.get('Due Day')
            amount_due_col = headers.get('Amount Due')
            month_col = headers.get(month_key)
            
            # Add month columns if they don't exist
            if month_col is None:
                max_col = ws.max_column
                for m in MONTHS_2026:
                    ws.cell(row=1, column=max_col + 1, value=m)
                    max_col += 1
                headers = {col.value: idx for idx, col in enumerate(ws[1], start=1)}
                due_day_col = headers.get('Due Day')
                amount_due_col = headers.get('Amount Due')
                month_col = headers.get(month_key)
            
            if not all([due_day_col, month_col]):
                continue
            
            for row in ws.iter_rows(min_row=2):
                due_day = row[due_day_col - 1].value
                amount_due = row[amount_due_col - 1].value if amount_due_col else None
                customer_name = row[3].value
                current_month_val = row[month_col - 1].value
                
                if due_day == current_day:
                    try:
                        charge_amount = float(amount_due) if amount_due else 0
                        debug_info.append(f"Found {customer_name}: due_day={due_day}, amount_due={amount_due}, charge_amount={charge_amount}")
                        # Charge regardless of current value - just add the amount due
                        if charge_amount > 0:
                            debug_info.append(f"  -> Charging {customer_name} with ${charge_amount}")
                            # Add to existing balance (or set if empty)
                            current_val = float(current_month_val) if current_month_val else 0
                            new_val = current_val + charge_amount
                            row[month_col - 1].value = new_val
                            
                            # Update amount due (sum all months)
                            total = 0
                            for m in MONTHS_2026:
                                m_col = headers.get(m)
                                if m_col:
                                    val = row[m_col - 1].value
                                    total += float(val) if val else 0
                            if amount_due_col:
                                row[amount_due_col - 1].value = total
                            
                            # Advance due date by 30 days
                            current_due = row[due_day_col - 1].value
                            if current_due:
                                try:
                                    current_due = int(current_due)
                                    new_due = current_due + 30
                                    if new_due > 31:
                                        new_due = new_due % 30
                                        if new_due == 0:
                                            new_due = 30
                                    row[due_day_col - 1].value = new_due
                                except:
                                    pass
                            
                            charged.append(f"{customer_name} ({sheet_name}): ${charge_amount}")
                    except:
                        pass
        
        wb.save(EXCEL_FILE)
        if not charged:
            return debug_info
        return charged
    except Exception as e:
        return [f"Error: {str(e)}"]

def advance_due_date(sheet_name, customer_name, days=30):
    """Advance the due date by specified days (default 30)"""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        
        headers = {col.value: idx for idx, col in enumerate(ws[1], start=1)}
        due_day_col = headers.get('Due Day')
        
        if not due_day_col:
            return False
        
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:
                current_due = row[due_day_col - 1].value
                if current_due:
                    try:
                        current_due = int(current_due)
                        new_due = current_due + days
                        if new_due > 31:
                            new_due = new_due % 30
                            if new_due == 0:
                                new_due = 30
                        row[due_day_col - 1].value = new_due
                    except:
                        pass
                break
        
        wb.save(EXCEL_FILE)
        return True
    except:
        return False

def get_past_due_customers(all_data):
    results = []
    for service, df in all_data.items():
        if df is None or df.empty:
            continue
        if 'Due Day' in df.columns and 'Status' in df.columns:
            for _, row in df.iterrows():
                due_day = row.get('Due Day')
                status = row.get('Status', '')
                amount_due = row.get('Amount Due', 0)
                if due_day is None or pd.isna(due_day):
                    continue
                status_str = str(status).upper() if status else ''
                is_paid = 'PAID' in status_str or status_str == 'READY'
                try:
                    has_balance = float(amount_due) > 0 if amount_due else False
                except:
                    has_balance = False
                if not is_paid and has_balance:
                    monthly = {}
                    for m in MONTHS_2026:
                        monthly[m] = row.get(m, 0)
                    result = {
                        'Service': service,
                        'Customer Name': row.get('Customer Name', ''),
                        'Phone': row.get('Phone', ''),
                        'Amount Due': row.get('Amount Due', ''),
                        'Status': row.get('Status', ''),
                        'Card Number': row.get('Card Number', ''),
                        'Exp': row.get('Exp', ''),
                        'CVV': row.get('CVV', ''),
                        'Plan Cost': row.get('Plan Cost', ''),
                        'Charge Date': row.get('Charge Date', ''),
                        'Due Day': row.get('Due Day', ''),
                        'Notes': row.get('Notes', ''),
                        'Notes2': row.get('Notes2', ''),
                        'Payment Date': row.get('Payment Date', ''),
                    }
                    result.update(monthly)
                    results.append(result)
    return results

def get_collections_by_date(all_data, start_date, end_date):
    results = []
    for service, df in all_data.items():
        if df is None or df.empty:
            continue
        if 'Payment Date' in df.columns:
            for _, row in df.iterrows():
                payment_date = row.get('Payment Date')
                if payment_date is None or pd.isna(payment_date):
                    continue
                try:
                    payment_date = pd.to_datetime(payment_date)
                except:
                    continue
                if start_date <= payment_date.date() <= end_date:
                    results.append({
                        'Service': service,
                        'Customer Name': row.get('Customer Name', ''),
                        'Phone': row.get('Phone', ''),
                        'Amount Collected': row.get('Amount Due', 0),
                        'Status': row.get('Status', ''),
                        'Payment Date': payment_date.strftime('%Y-%m-%d'),
                        'Notes': row.get('Notes', '')
                    })
    return results

# Main UI
st.title("💳 Customer Payment Manager")

all_data, excel_file = load_excel()

if all_data is None:
    st.stop()

# Sidebar
with st.sidebar:
    st.header("📊 Summary")
    total_customers = 0
    total_revenue = 0
    for service, df in all_data.items():
        if df is not None and not df.empty:
            count = len(df)
            total_customers += count
            revenue = 0
            for amt in df['Amount Due'].fillna(0):
                try:
                    if amt and str(amt).strip():
                        revenue += float(str(amt).strip())
                except:
                    pass
            total_revenue += revenue
            st.metric(service, f"{count} customers", f"${revenue:,.2f}")
    st.divider()
    st.metric("Total Customers", total_customers)
    st.metric("Total Revenue", f"${total_revenue:,.2f}")

# Main content
st.header("🔍 Search & Filter")

# Auto-charge section
with st.expander("⚡ Auto-Charge Due Today"):
    st.write(f"Click to charge all customers with due date matching today (Day {datetime.now().day}, {MONTH_MAP.get(datetime.now().month)})")
    if st.button("⚡ Run Auto-Charge"):
        with st.spinner("Charging customers..."):
            results = auto_charge_due_today()
            if results:
                # Check if debug info
                if any("Today:" in str(r) or "EXCEL_FILE" in str(r) for r in results):
                    st.warning("Debug info:")
                    for r in results:
                        st.caption(r)
                else:
                    st.success(f"Charged {len(results)} customer(s)!")
                    for r in results:
                        st.write(f"• {r}")
            else:
                st.info("No customers to charge today")

tab1, tab2, tab3, tab4 = st.tabs(["📝 Search by Name", "📅 Filter by Due Date", "⚠️ Past Due", "💰 Collections History"])

with tab1:
    search_query = st.text_input("Search by name, phone, or account:", placeholder="Enter name or phone...", key="search_name")
    if search_query:
        results = search_customers(all_data, search_query)
        if results:
            st.write(f"Found **{len(results)}** customer(s)")
            for i, customer in enumerate(results):
                with st.container():
                    charge_date = customer.get('Charge Date', '')
                    if charge_date:
                        try:
                            charge_date = pd.to_datetime(charge_date).strftime('%Y-%m-%d')
                        except:
                            pass
                    
                    balance = get_balance(customer)
                    has_credit = balance < 0
                    status = str(customer.get('Status', '')).upper() if customer.get('Status') else ''
                    is_paid = 'PAID' in status or status == 'READY'
                    
                    if has_credit:
                        balance_display = f"💚 CREDIT: ${abs(balance):.2f}"
                    elif balance == 0:
                        balance_display = "$0.00"
                    else:
                        balance_display = f"${balance:.2f}"
                    
                    payment_date = customer.get('Payment Date', '') or ''
                    if payment_date:
                        try:
                            payment_date = pd.to_datetime(payment_date).strftime('%Y-%m-%d')
                        except:
                            pass
                    
                    # Get monthly balances
                    monthly_balances = get_monthly_balances(customer)
                    total_from_aging = get_total_balance_from_months(monthly_balances)
                    
                    # Show 12-month balance display
                    months_list = [
                        ('Jan_2026', 'January'), ('Feb_2026', 'February'), ('Mar_2026', 'March'),
                        ('Apr_2026', 'April'), ('May_2026', 'May'), ('Jun_2026', 'June'),
                        ('Jul_2026', 'July'), ('Aug_2026', 'August'), ('Sep_2026', 'September'),
                        ('Oct_2026', 'October'), ('Nov_2026', 'November'), ('Dec_2026', 'December')
                    ]
                    
                    st.write(f"### {customer['Customer Name']}")
                    st.write(f"**Service:** {customer['Service']} | **Phone:** {customer['Phone']} | **Status:** {customer['Status']}")
                    st.write(f"**Balance:** {balance_display} | **Due Day:** {customer.get('Due Day', 'N/A')}")
                    
                    # Show monthly balances
                    st.write("### 📊 Monthly Balance (2026)")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        for month_col, month_name in months_list[:4]:
                            val = monthly_balances.get(month_col, 0)
                            st.metric(month_name, f"${val:.2f}")
                    with col2:
                        for month_col, month_name in months_list[4:8]:
                            val = monthly_balances.get(month_col, 0)
                            st.metric(month_name, f"${val:.2f}")
                    with col3:
                        for month_col, month_name in months_list[8:]:
                            val = monthly_balances.get(month_col, 0)
                            st.metric(month_name, f"${val:.2f}")
                    
                    st.divider()
                    st.metric("💰 Total Balance", f"${total_from_aging:.2f}")
                    
                    # Edit monthly balance
                    st.write("### ✏️ Edit Monthly Balance")
                    col_edit1, col_edit2 = st.columns(2)
                    with col_edit1:
                        edit_month = st.selectbox("Select Month:", options=[m[0] for m in months_list], key=f"edit_month_{i}")
                    with col_edit2:
                        current_val = monthly_balances.get(edit_month, 0)
                        new_val = st.number_input("New Balance:", min_value=0.0, value=float(current_val), step=5.0, key=f"edit_val_{i}")
                    
                    if st.button(f"💾 Save {edit_month}", key=f"save_{i}"):
                        if save_monthly_balance(customer['Service'], customer['Customer Name'], edit_month, new_val):
                            # Also update total Amount Due
                            update_amount_due_from_months(customer['Service'], customer['Customer Name'])
                            st.success("Saved!")
                            st.rerun()
                        else:
                            st.error("Error saving!")
                    
                    # Post payment to specific month
                    st.write("### 💰 Post Payment to Month")
                    col_pay1, col_pay2 = st.columns(2)
                    with col_pay1:
                        pay_month = st.selectbox("Select Month to Pay:", options=[m[0] for m in months_list], key=f"pay_month_{i}")
                    with col_pay2:
                        pay_amount = st.number_input("Payment Amount:", min_value=0.0, value=float(monthly_balances.get(pay_month, 0)), step=5.0, key=f"pay_amt_{i}")
                    
                    if st.button(f"✅ Apply Payment to {pay_month}", key=f"apply_pay_{i}"):
                        if pay_amount > 0:
                            current = monthly_balances.get(pay_month, 0)
                            new_balance = max(0, current - pay_amount)
                            if save_monthly_balance(customer['Service'], customer['Customer Name'], pay_month, new_balance):
                                update_amount_due_from_months(customer['Service'], customer['Customer Name'])
                                
                                # Check if paying current month - advance due date by 30 days
                                current_month_col = MONTH_MAP.get(datetime.now().month)
                                if pay_month == current_month_col:
                                    advance_due_date(customer['Service'], customer['Customer Name'], days=30)
                                    st.success(f"Payment applied to {pay_month}! Due date advanced 30 days.")
                                else:
                                    st.success(f"Payment applied to {pay_month}!")
                                st.rerun()
                            else:
                                st.error("Error applying payment!")
                    
                    with st.expander("💳 Edit Customer Info"):
                        with st.form(f"edit_form_{i}"):
                            new_name = st.text_input("Name", value=customer.get('Customer Name', ''))
                            phone = st.text_input("Phone", value=customer.get('Phone', ''))
                            card = st.text_input("Card", value=customer.get('Card Number', ''))
                            exp = st.text_input("Exp", value=customer.get('Exp', ''))
                            cvv = st.text_input("CVV", value=customer.get('CVV', ''))
                            plan = st.number_input("Plan Cost", value=float(customer.get('Plan Cost', 0) or 0))
                            submitted = st.form_submit_button("Save Info")
                            if submitted:
                                if save_customer_info(customer['Service'], customer['Customer Name'], new_name, phone, card, exp, cvv, plan, row_index=customer.get('row_index')):
                                    st.success("Saved!")
                                    st.rerun()
                    
                    st.divider()

with tab2:
    due_day = st.selectbox("Select Due Day:", options=list(range(1, 32)))
    if due_day:
        customers = get_customers_by_due_day(all_data, due_day)
        st.write(f"Found **{len(customers)}** customers due on day {due_day}")
        for c in customers:
            card = c.get('Card Number', '')
            exp = c.get('Exp', '')
            cvv = c.get('CVV', '')
            st.write(f"- **{c['Customer Name']}** ({c['Service']}) | Balance: ${c.get('Amount Due', 0)} | Card: {card} | Exp: {exp} | CVV: {cvv}")

with tab3:
    past_due = get_past_due_customers(all_data)
    st.write(f"Found **{len(past_due)}** past due customers")
    for c in past_due:
        card = c.get('Card Number', '')
        exp = c.get('Exp', '')
        cvv = c.get('CVV', '')
        st.write(f"- **{c['Customer Name']}** ({c['Service']}) | Balance: ${c.get('Amount Due', 0)} | Card: {card} | Exp: {exp} | CVV: {cvv}")

with tab4:
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Start Date")
    with col2:
        end_date = st.date_input("End Date")
    if start_date and end_date:
        collections = get_collections_by_date(all_data, start_date, end_date)
        st.write(f"Found **{len(collections)}** payments")
        for c in collections:
            st.write(f"- {c['Customer Name']}: ${c.get('Amount Collected', 0)} on {c.get('Payment Date', 'N/A')}")

# Download/Upload Section
st.divider()
st.header("💾 Download/Upload Database")

col1, col2 = st.columns(2)

with col1:
    try:
        with open("/opt/render/project/src/cleaned_billing_by_service.xlsx", "rb") as f:
            st.download_button(label="📥 Download Excel File", data=f, file_name="cleaned_billing_by_service.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except:
        st.warning("File not found on server")

with col2:
    uploaded_file = st.file_uploader("📤 Upload Updated Excel File", type=["xlsx"])
    if uploaded_file is not None:
        try:
            with open("/opt/render/project/src/cleaned_billing_by_service.xlsx", "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success("✅ File uploaded! Refresh the page.")
        except Exception as e:
            st.error(f"Error: {e}")

# Footer
st.divider()
st.caption(f"💾 Data file: {EXCEL_FILE}")
