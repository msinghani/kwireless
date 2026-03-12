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

# Configuration
EXCEL_FILE_LOCAL = "cleaned_billing_by_service.xlsx"

# Use persistent disk on Render, or local file for development
RENDER_DISK_PATH = "/app/data"
if os.path.exists(RENDER_DISK_PATH):
    # Check if file exists on persistent disk
    disk_file = os.path.join(RENDER_DISK_PATH, "cleaned_billing_by_service.xlsx")
    if os.path.exists(disk_file):
        EXCEL_FILE = disk_file
    else:
        # Disk exists but no file yet - use local path (will prompt upload)
        EXCEL_FILE = EXCEL_FILE_LOCAL
else:
    # Not on Render (local development)
    EXCEL_FILE = EXCEL_FILE_LOCAL

st.set_page_config(page_title="Customer Payment Manager", page_icon="💳", layout="wide")

def save_customer_notes(sheet_name, customer_name, notes):
    """Save customer notes to the Excel file"""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        
        # Find the row with this customer
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:  # Column D is customer name
                # Update Notes column
                row[10].value = notes
                break
        
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Error saving notes: {e}")
        return False


def save_notes2(sheet_name, customer_name, notes2):
    """Save customer notes2 to the Excel file"""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        
        # Find the row with this customer
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:  # Column D is customer name
                # Update Notes2 column (column M = index 12)
                row[12].value = notes2
                break
        
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Error saving notes2: {e}")
        return False


def save_due_date(sheet_name, customer_name, due_day):
    """Save customer due date to the Excel file"""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        
        # Find the row with this customer
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:  # Column D is customer name
                # Update Due Day column (column L = index 11)
                row[11].value = due_day
                break
        
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Error saving due date: {e}")
        return False


def save_customer_info(sheet_name, customer_name, new_name, phone, card_number, exp, cvv, plan_cost, original_phone=None, row_index=None):
    """Save customer info (name, phone, card, plan cost) to the Excel file"""
    import traceback
    
    try:
        print(f"SAVE_DEBUG: Starting save - sheet={sheet_name}, name={customer_name}, new_name={new_name}, row_index={row_index}")
        
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        
        # Direct row access by index
        target_row = None
        
        if row_index is not None:
            # Convert pandas index to Excel row (pandas index + 2 because: +1 for header, +1 for 1-based)
            excel_row_num = row_index + 2
            print(f"SAVE_DEBUG: Looking for Excel row {excel_row_num}")
            
            # Access row directly
            for row in ws.iter_rows(min_row=excel_row_num, max_row=excel_row_num):
                target_row = row
                print(f"SAVE_DEBUG: Found row: {row[3].value}")
                break
        else:
            # Fallback: find by name + phone
            print(f"SAVE_DEBUG: No row_index, using fallback by name")
            for row in ws.iter_rows(min_row=2):
                if str(row[3].value).strip() == str(customer_name).strip():
                    target_row = row
                    print(f"SAVE_DEBUG: Found row by name: {row[3].value}")
                    break
        
        if target_row is None:
            print("SAVE_DEBUG: No matching row found!")
            return False
        
        # Update all fields
        target_row[2].value = plan_cost   # Column C - Plan Cost
        target_row[3].value = new_name    # Column D - Customer Name
        target_row[4].value = card_number  # Column E - Card Number
        target_row[5].value = exp        # Column F - Exp
        target_row[6].value = cvv         # Column G - CVV
        target_row[9].value = phone       # Column J - Phone
        
        print(f"SAVE_DEBUG: Updated row, saving...")
        wb.save(EXCEL_FILE)
        print(f"SAVE_DEBUG: Save complete!")
        return True
        
    except Exception as e:
        print(f"SAVE_DEBUG: ERROR - {e}")
        traceback.print_exc()
        st.error(f"Error saving customer info: {e}")
        return False



def get_balance(customer):
    """Get the current balance for a customer (negative = credit)"""
    # Check if already paid - but now check if there's a credit
    status = str(customer.get('Status', '')).upper() if customer.get('Status') else ''
    
    amount_due = customer.get('Amount Due', '')
    
    # Use Amount Due column only - never fallback to Plan Cost
    if amount_due and str(amount_due).strip() and str(amount_due).strip().lower() != 'nan':
        try:
            val = float(str(amount_due).strip())
            return val
        except:
            pass
    
    # If Amount Due is empty/invalid, return 0
    return 0


def get_balance_aging(customer):
    """Get balance breakdown by month"""
    balance_current = customer.get('Balance_Current', 0)
    balance_history = customer.get('Balance_History', '')
    
    aging = {
        'current': {'label': 'This Month', 'amount': 0},
        'past': []  # List of {'month': 'Feb 2026', 'amount': 50.00}
    }
    
    # Get current month
    now = datetime.now()
    current_month = now.strftime("%b %Y")
    
    # Current month balance
    try:
        aging['current']['amount'] = float(balance_current) if balance_current else 0
    except:
        aging['current']['amount'] = 0
    
    # Parse history
    if balance_history and str(balance_history).strip() and str(balance_history).strip().lower() != 'nan':
        try:
            history = json.loads(str(balance_history))
            for month, amount in history.items():
                aging['past'].append({'month': month, 'amount': float(amount)})
        except:
            pass
    
    # Sort past months (newest first)
    aging['past'].sort(key=lambda x: datetime.strptime(x['month'], "%b %Y"), reverse=True)
    
    return aging


def calculate_total_balance(aging):
    """Calculate total from aging breakdown"""
    total = aging['current']['amount']
    for past_month in aging['past']:
        total += past_month['amount']
    return total


def save_balance_by_month(sheet_name, customer_name, month_label, amount, is_current_month=False):
    """Save balance for a specific month"""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        
        # Find the row with this customer
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:  # Column D is customer name
                if is_current_month:
                    # Update Balance_Current (column 14 = index 13)
                    row[13].value = amount
                else:
                    # Update Balance_History (column 15 = index 14)
                    existing_history = row[14].value
                    history = {}
                    if existing_history and str(existing_history).strip():
                        try:
                            history = json.loads(str(existing_history))
                        except:
                            history = {}
                    
                    history[month_label] = amount
                    row[14].value = json.dumps(history)
                
                # Update total Amount Due
                # Need to recalculate total
                break
        
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Error saving balance: {e}")
        return False


def update_total_from_aging(sheet_name, customer_name):
    """Recalculate and update total Amount Due from aging breakdown"""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:
                # Get Balance_Current (column 14)
                current = row[13].value
                try:
                    current = float(current) if current else 0
                except:
                    current = 0
                
                # Get Balance_History (column 15)
                history_str = row[14].value
                history_total = 0
                if history_str and str(history_str).strip():
                    try:
                        history = json.loads(str(history_str))
                        history_total = sum(float(v) for v in history.values())
                    except:
                        history_total = 0
                
                # Update Amount Due (column 8)
                total = current + history_total
                row[7].value = total
                
                # Update Status
                if total <= 0:
                    row[8].value = "Paid"
                elif total < row[2].value:  # Less than plan cost
                    row[8].value = "Partial"
                else:
                    row[8].value = "Not Paid"
                
                break
        
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Error updating total: {e}")
        return False

# Custom styling
st.markdown("""
    <style>
    .main {
        background-color: #f5f5f5;
    }
    .stButton>button {
        width: 100%;
    }
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
    """Load all sheets from Excel file"""
    try:
        excel_file = pd.ExcelFile(EXCEL_FILE)
        all_data = {}
        for sheet in excel_file.sheet_names:
            if sheet != "Summary":
                df = pd.read_excel(excel_file, sheet_name=sheet)
                df['Service'] = sheet
                
                # Parse charge date to get day of month - only if Due Day column doesn't exist
                if 'Due Day' not in df.columns:
                    if 'Charge Date' in df.columns:
                        try:
                            df['Charge Date'] = pd.to_datetime(df['Charge Date'], errors='coerce')
                            df['Due Day'] = df['Charge Date'].dt.day
                        except:
                            df['Due Day'] = None
                # If Due Day exists in Excel, make sure it's available
                elif 'Due Day' in df.columns:
                    pass  # Already has Due Day column
                
                all_data[sheet] = df
        return all_data, excel_file
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None, None

def save_payment(sheet_name, customer_name, payment_amount, notes=""):
    """Save a payment to the Excel file"""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet_name]
        
        # Find the row with this customer
        for row in ws.iter_rows(min_row=2):
            if row[3].value == customer_name:  # Column D is customer name
                # Get current balance from Amount Due column only
                current_amount_due = row[7].value  # Column H is Amount Due
                
                # Calculate current balance - use Amount Due only
                try:
                    if current_amount_due and str(current_amount_due).strip():
                        current_balance = float(str(current_amount_due).strip())
                    else:
                        current_balance = 0
                except:
                    current_balance = 0
                
                # Calculate new balance after payment
                new_balance = current_balance - payment_amount
                
                # Update Amount Due (can be negative for credit)
                row[7].value = new_balance
                
                # Update Status based on new balance
                if new_balance <= 0:
                    row[8].value = "Paid"
                else:
                    row[8].value = "Partial"
                
                # Add payment date in column 11 (Payment Date)
                today = datetime.now().strftime('%Y-%m-%d')
                row[11].value = today
                
                # Set due date to 30 days from payment date
                from datetime import timedelta
                new_due_date = (datetime.now() + timedelta(days=30)).day
                row[12].value = new_due_date  # Column M is Due Day
                
                # Update notes in column 10 - append payment info
                existing_notes = str(row[10].value) if row[10].value else ""
                payment_info = f"Payment ${payment_amount:.2f} on {today} at {datetime.now().strftime('%I:%M %p')}"
                if notes:
                    payment_info += f": {notes}"
                
                if existing_notes and existing_notes != "None":
                    row[10].value = existing_notes + " | " + payment_info
                else:
                    row[10].value = payment_info
                
                break
        
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        st.error(f"Error saving: {e}")
        return False

def search_customers(all_data, query):
    """Search customers across all sheets"""
    results = []
    
    for service, df in all_data.items():
        if df is None or df.empty:
            continue
            
        # Search in name, phone, or account columns
        mask = (
            df['Customer Name'].astype(str).str.contains(query, case=False, na=False) |
            df['Phone'].astype(str).str.contains(query, case=False, na=False) |
            df['Card Number'].astype(str).str.contains(query, case=False, na=False)
        )
        
        matches = df[mask]
        for idx, row in matches.iterrows():
            results.append({
                'Service': service,
                'row_index': idx,  # Store pandas index for saving
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
                'Balance_Current': row.get('Balance_Current', 0),
                'Balance_History': row.get('Balance_History', '')
            })
    
    return results

def get_customers_by_due_day(all_data, due_day):
    """Get all customers due on a specific day of the month"""
    results = []
    
    for service, df in all_data.items():
        if df is None or df.empty:
            continue
        
        # Filter by due day
        if 'Due Day' in df.columns:
            matches = df[df['Due Day'] == due_day]
            
            for _, row in matches.iterrows():
                results.append({
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
                    'Balance_Current': row.get('Balance_Current', 0),
                    'Balance_History': row.get('Balance_History', '')
                })
    
    return results

def get_past_due_customers(all_data):
    """Get all customers who are past due (not paid and have an amount due)"""
    from datetime import datetime
    
    results = []
    today = datetime.now().day
    
    for service, df in all_data.items():
        if df is None or df.empty:
            continue
        
        if 'Due Day' in df.columns and 'Status' in df.columns:
            for _, row in df.iterrows():
                due_day = row.get('Due Day')
                status = row.get('Status', '')
                amount_due = row.get('Amount Due', 0)
                
                # Skip if no due day
                if due_day is None or pd.isna(due_day):
                    continue
                
                # Check if past due: not paid, has amount due, and due day has passed
                status_str = str(status).upper() if status else ''
                is_paid = 'PAID' in status_str or status_str == 'READY'
                
                # Also check if they have a balance (amount due > 0)
                try:
                    has_balance = float(amount_due) > 0 if amount_due else False
                except:
                    has_balance = False
                
                if not is_paid and has_balance:
                    results.append({
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
                        'Balance_Current': row.get('Balance_Current', 0),
                        'Balance_History': row.get('Balance_History', '')
                    })
    
    return results

def get_collections_by_date(all_data, start_date, end_date):
    """Get all payments collected within a date range"""
    from datetime import datetime
    
    results = []
    
    for service, df in all_data.items():
        if df is None or df.empty:
            continue
        
        if 'Payment Date' in df.columns:
            for _, row in df.iterrows():
                payment_date = row.get('Payment Date')
                
                # Skip if no payment date
                if payment_date is None or pd.isna(payment_date):
                    continue
                
                # Parse payment date
                try:
                    payment_date = pd.to_datetime(payment_date)
                except:
                    continue
                
                # Check if within date range
                if start_date <= payment_date.date() <= end_date:
                    # Calculate payment amount from the change in balance
                    # We'll get the current amount due and show it as collected
                    amount_due = row.get('Amount Due', 0)
                    status = row.get('Status', '')
                    
                    results.append({
                        'Service': service,
                        'Customer Name': row.get('Customer Name', ''),
                        'Phone': row.get('Phone', ''),
                        'Amount Collected': amount_due,
                        'Status': status,
                        'Payment Date': payment_date.strftime('%Y-%m-%d'),
                        'Notes': row.get('Notes', '')
                    })
    
    return results

# Main UI
st.title("💳 Customer Payment Manager")

# Load data
all_data, excel_file = load_excel()

if all_data is None:
    st.stop()

# Sidebar with summary
with st.sidebar:
    st.header("📊 Summary")
    
    total_customers = 0
    total_revenue = 0
    
    for service, df in all_data.items():
        if df is not None and not df.empty:
            count = len(df)
            total_customers += count
            
            # Calculate revenue (Amount Due or Plan Cost)
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

# Create tabs for different search options
tab1, tab2, tab3, tab4 = st.tabs(["📝 Search by Name", "📅 Filter by Due Date", "⚠️ Past Due", "💰 Collections History"])

with tab1:
    search_query = st.text_input("Search by name, phone, or account:", placeholder="Enter name or phone...", key="search_name")
    
    if search_query:
        results = search_customers(all_data, search_query)
        
        if results:
            st.write(f"Found **{len(results)}** customer(s)")
            
            # Display results
            for i, customer in enumerate(results):
                with st.container():
                    charge_date = customer.get('Charge Date', '')
                    if charge_date:
                        try:
                            charge_date = pd.to_datetime(charge_date).strftime('%Y-%m-%d')
                        except:
                            pass
                    
                    # Calculate balance
                    balance = get_balance(customer)
                    
                    # Check if overpayment/credit
                    has_credit = balance < 0
                    
                    # Check if already paid
                    status = str(customer.get('Status', '')).upper() if customer.get('Status') else ''
                    is_paid = 'PAID' in status or status == 'READY'
                    
                    # Format balance display
                    if has_credit:
                        balance_display = f"💚 CREDIT: ${abs(balance):.2f}"
                    elif balance == 0:
                        balance_display = "$0.00"
                    else:
                        balance_display = f"${balance:.2f}"
                    
                    # Get existing notes
                    existing_notes = customer.get('Notes', '') or ''
                    payment_date = customer.get('Payment Date', '') or ''
                    
                    # Format payment date
                    if payment_date:
                        try:
                            payment_date = pd.to_datetime(payment_date).strftime('%Y-%m-%d')
                        except:
                            pass
                    
                    # Get balance aging breakdown
                    aging = get_balance_aging(customer)
                    total_from_aging = calculate_total_balance(aging)
                    
                    # Show aging breakdown
                    st.markdown("""
                    <div class="customer-card">
                        <h3>{}</h3>
                        <p><strong>Service:</strong> {} | 
                           <strong>Phone:</strong> {} | 
                           <strong>Status:</strong> {}</p>
                        <p><strong>💰 Balance:</strong> {} | 
                           <strong>Due Date:</strong> Charge Date: {} | Due Day: {}</p>
                        <p><strong>💳 Card:</strong> {} | 
                           <strong>Exp:</strong> {} | 
                           <strong>CVV:</strong> {}</p>
                        <p><strong>📅 Last Payment:</strong> {}</p>
                    </div>
                    """.format(
                        customer['Customer Name'],
                        customer['Service'],
                        customer['Phone'],
                        customer['Status'],
                        balance_display,
                        charge_date,
                        customer.get('Due Day', 'N/A'),
                        customer.get('Card Number', 'N/A'),
                        customer.get('Exp', 'N/A'),
                        customer.get('CVV', 'N/A'),
                        payment_date if payment_date else 'N/A'
                    ), unsafe_allow_html=True)
                    
                    # Balance Aging Section
                    st.write("### 📊 Balance Aging")
                    
                    # Show current month
                    current_month = datetime.now().strftime("%b %Y")
                    col_aging1, col_aging2, col_aging3 = st.columns([2, 1, 1])
                    with col_aging1:
                        st.write(f"**This Month ({current_month}):**")
                    with col_aging2:
                        st.write(f"${aging['current']['amount']:.2f}")
                    with col_aging3:
                        pass
                    
                    # Show past months
                    if aging['past']:
                        for past in aging['past']:
                            col_aging1, col_aging2, col_aging3 = st.columns([2, 1, 1])
                            with col_aging1:
                                st.write(f"**{past['month']}:**")
                            with col_aging2:
                                st.write(f"${past['amount']:.2f}")
                            with col_aging3:
                                pass
                    
                    # Total
                    st.divider()
                    col_total1, col_total2 = st.columns([2, 1])
                    with col_total1:
                        st.write("**Total Balance:**")
                    with col_total2:
                        st.write(f"${total_from_aging:.2f}")
                    
                    # Billing Cycle Info
                    st.write("### 📅 Billing Cycle")
                    due_day = customer.get('Due Day', 1)
                    if not due_day or str(due_day) == 'None' or pd.isna(due_day):
                        due_day = 1
                    if hasattr(due_day, 'day'):
                        due_day = due_day.day
                    st.write(f"Charges on: Day {due_day} of each month")
                    st.write(f"Payment due: Day {due_day} of each month")
                    
                    # Post Payment to Specific Month
                    st.write("### 💰 Post Payment to Month")
                    
                    # Build list of months with balances
                    month_options = [("This Month", "current", aging['current']['amount'])]
                    for past in aging['past']:
                        month_options.append((past['month'], past['month'], past['amount']))
                    
                    # Let user select which month to post payment to
                    col_pay1, col_pay2, col_pay3 = st.columns([2, 1, 1])
                    with col_pay1:
                        selected_month_label = st.selectbox(
                            "Select month to post payment:",
                            options=[m[0] for m in month_options],
                            key=f"month_select_{i}"
                        )
                    with col_pay2:
                        # Find selected month's current balance
                        selected_balance = 0
                        for m in month_options:
                            if m[0] == selected_month_label:
                                selected_balance = m[2]
                                break
                        post_amount = st.number_input(
                            "Amount to post:",
                            min_value=0.0,
                            value=float(selected_balance),
                            step=5.0,
                            key=f"post_amt_{i}"
                        )
                    with col_pay3:
                        st.write("")
                        st.write("")
                        if st.button("✅ Post Payment", key=f"post_pay_{i}"):
                            if post_amount > 0:
                                # Determine if current month or past
                                is_current = selected_month_label == "This Month"
                                current_month = datetime.now().strftime("%b %Y")
                                
                                if is_current:
                                    # Deduct from Balance_Current
                                    new_current = aging['current']['amount'] - post_amount
                                    save_balance_by_month(customer['Service'], customer['Customer Name'], current_month, new_current, is_current_month=True)
                                else:
                                    # Deduct from Balance_History
                                    history = {}
                                    for past in aging['past']:
                                        if past['month'] == selected_month_label:
                                            history[past['month']] = max(0, past['amount'] - post_amount)
                                        else:
                                            history[past['month']] = past['amount']
                                    
                                    # Save updated history
                                    try:
                                        wb = load_workbook(EXCEL_FILE)
                                        ws = wb[customer['Service']]
                                        for row in ws.iter_rows(min_row=2):
                                            if row[3].value == customer['Customer Name']:
                                                row[14].value = json.dumps(history)
                                                break
                                        wb.save(EXCEL_FILE)
                                    except Exception as e:
                                        st.error(f"Error: {e}")
                                
                                # Update total
                                update_total_from_aging(customer['Service'], customer['Customer Name'])
                                st.success(f"Posted ${post_amount:.2f} to {selected_month_label}!")
                                st.rerun()
                            else:
                                st.warning("Enter an amount greater than 0")
                    
                    st.divider()
                    
                    # Due Date Edit section
                    st.write("📅 **Due Date:**")
                    current_due_day = customer.get('Due Day', 1)
                    if not current_due_day or str(current_due_day) == 'None' or pd.isna(current_due_day):
                        current_due_day = 1
                    # Extract day number if it's a datetime/date object or string date
                    if hasattr(current_due_day, 'day'):
                        current_due_day = current_due_day.day
                    elif isinstance(current_due_day, str) and '-' in str(current_due_day):
                        try:
                            current_due_day = int(str(current_due_day).split('-')[-1])
                        except:
                            current_due_day = 1
                    due_key = f"due_day_{customer['Service']}_{customer['Customer Name']}_{i}"
                    new_due_day = st.number_input("Due Day of Month (1-31):", min_value=1, max_value=31, value=int(current_due_day), key=due_key)
                    if st.button("💾 Save Due Date", key=f"save_due_{i}"):
                        if save_due_date(customer['Service'], customer['Customer Name'], new_due_day):
                            st.success("Due date saved!")
                            st.rerun()

                    # Customer Info Edit section - hidden by default
                    edit_info_key = f"edit_info_{customer['Service']}_{customer['Customer Name']}"
                    show_edit = st.checkbox("✏️ Edit Customer Info", key=edit_info_key)
                    
                    if show_edit:
                        # Show phone/card to help identify duplicate names
                        st.caption(f"📱 Phone: {customer.get('Phone', 'N/A')} | 💳 Card: {customer.get('Card Number', 'N/A')}")
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            edit_name = st.text_input("Customer Name", value=str(customer.get('Customer Name', '')), key=f"edit_name_{i}")
                            edit_phone = st.text_input("Phone", value=str(customer.get('Phone', '')), key=f"edit_phone_{i}")
                            edit_card = st.text_input("Card Number", value=str(customer.get('Card Number', '')), key=f"edit_card_{i}")
                        with col2:
                            edit_exp = st.text_input("Exp", value=str(customer.get('Exp', '')), key=f"edit_exp_{i}")
                            edit_cvv = st.text_input("CVV", value=str(customer.get('CVV', '')), key=f"edit_cvv_{i}")
                            edit_plan_cost = st.number_input("Plan Cost", min_value=0.0, value=float(customer.get('Plan Cost', 0) or 0), step=1.0, key=f"edit_plan_{i}")
                        
                        col_btn1, col_btn2 = st.columns(2)
                        with col_btn1:
                            if st.button("💾 Save Customer Info", key=f"save_info_{i}"):
                                saved = save_customer_info(customer['Service'], customer['Customer Name'], edit_name, edit_phone, edit_card, edit_exp, edit_cvv, edit_plan_cost, original_phone=customer.get('Phone'), row_index=customer.get('row_index'))
                                if saved:
                                    st.success("Customer info saved!")
                                    st.rerun()
                                else:
                                    st.error("Could not save - customer not found")
                        with col_btn2:
                            pass

                    # Notes section - separate from payment notes
                    existing_notes2 = customer.get('Notes2', '') or ''
                    
                    col_notes1, col_notes2 = st.columns(2)
                    with col_notes1:
                        st.write("📝 **Customer Notes:**")
                        with st.expander("View/Edit Notes"):
                            notes_key = f"notes_{customer['Service']}_{customer['Customer Name']}"
                            notes_text = st.text_area("Notes (saved to Excel):", value=str(existing_notes), height=100, key=notes_key)
                            if st.button("💾 Save Notes", key=f"save_notes_{i}"):
                                if save_customer_notes(customer['Service'], customer['Customer Name'], notes_text):
                                    st.success("Notes saved!")
                                    st.rerun()
                    with col_notes2:
                        st.write("📝 **Backend Account Info:**")
                        if existing_notes2 and str(existing_notes2) != 'None' and str(existing_notes2).strip():
                            st.info(f"{existing_notes2}")
                        else:
                            st.caption("No Notes2 yet")
                    
                    st.divider()
                    
                    # Payment form - only show if has balance owing
                    if is_paid and not has_credit:
                        st.success(f"✅ {customer['Customer Name']} is paid up! Balance: $0.00")
                    elif has_credit:
                        st.success(f"💚 {customer['Customer Name']} has a credit of ${abs(balance):.2f} on their account!")
                    else:
                        # Payment form with custom amount (allows negative to adjust balance)
                        with st.form(f"payment_form_{i}"):
                            col1, col2, col3 = st.columns([1, 2, 1])
                            with col1:
                                payment_amount = st.number_input(f"Payment amount", min_value=-10000.0, value=float(balance), step=5.0, key=f"amt_{i}")
                            with col2:
                                payment_notes = st.text_input("Payment Notes (optional)", key=f"notes_{i}")
                            with col3:
                                submitted = st.form_submit_button(f"💰 Record", type="primary")
                            
                            if submitted:
                                if payment_amount != 0:
                                    new_balance = balance - payment_amount
                                    if save_payment(customer['Service'], customer['Customer Name'], payment_amount, payment_notes):
                                        if new_balance < 0:
                                            st.success(f"Payment of ${payment_amount:.2f} recorded! New balance: -${abs(new_balance):.2f} (CREDIT)")
                                        elif new_balance == 0:
                                            st.success(f"Payment of ${payment_amount:.2f} recorded! Balance: $0.00 - PAID IN FULL!")
                                        else:
                                            st.success(f"Payment of ${payment_amount:.2f} recorded! New balance: ${new_balance:.2f}")
                                        st.rerun()
                                else:
                                    st.warning("Please enter a non-zero amount")
                    
                    st.divider()
                    
                    st.divider()
        else:
            st.warning("No customers found matching your search.")
    else:
        st.info("Enter a name, phone number, or account to search.")

with tab2:
    st.subheader("📅 Filter Customers by Due Date")
    
    # Day selector
    col1, col2 = st.columns([1, 2])
    
    with col1:
        selected_day = st.selectbox(
            "Select Due Day of Month:",
            options=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31],
            index=0,
            format_func=lambda x: f"{x}{'st' if x==1 else 'nd' if x==2 else 'rd' if x==3 else 'th'}"
        )
    
    with col2:
        st.write("")
        st.write("")
        if st.button("🔄 Show All Due Dates", key="show_all_days"):
            st.rerun()
    
    # Get customers due on that day
    results = get_customers_by_due_day(all_data, selected_day)
    
    if results:
        st.write(f"### 📋 Customers Due on the {selected_day}{'st' if selected_day==1 else 'nd' if selected_day==2 else 'rd' if selected_day==3 else 'th'} of the Month")
        st.write(f"**{len(results)}** customer(s) found")
        
        # Calculate totals
        total_amount = 0
        for c in results:
            try:
                amt = c['Amount Due'] if c['Amount Due'] else 0
                if amt:
                    amt_str = str(amt).strip()
                    if amt_str and amt_str.lower() != 'nan':
                        total_amount += float(amt_str)
            except:
                pass
        
        if total_amount > 0:
            st.info(f"💰 **Total Due: ${total_amount:,.2f}**")
        else:
            st.info("💰 **Total Due: See individual amounts below**")
        
        # Display as table
        display_data = []
        for customer in results:
            charge_date = customer.get('Charge Date', '')
            if charge_date:
                try:
                    charge_date = pd.to_datetime(charge_date).strftime('%Y-%m-%d')
                except:
                    pass
            
            balance = get_balance(customer)
            
            display_data.append({
                'Service': customer['Service'],
                'Customer Name': customer['Customer Name'],
                'Phone': customer['Phone'],
                'Balance': f"${balance:.2f}",
                'Status': customer['Status'],
                'Charge Date': charge_date,
                'Card Number': customer.get('Card Number', ''),
                'Exp': customer.get('Exp', ''),
                'Notes': customer.get('Notes', '')
            })
        
        display_df = pd.DataFrame(display_data)
        st.dataframe(display_df, hide_index=True, use_container_width=True)
        
        # Payment section for each
        st.write("### 💳 Record Payments")
        
        for i, customer in enumerate(results):
            balance = get_balance(customer)
            with st.expander(f"{customer['Customer Name']} ({customer['Service']}) - 💰 Balance: ${balance:.2f}"):
                # Allow custom amount (including negative for adjustments)
                col1, col2, col3 = st.columns([1, 2, 1])
                with col1:
                    adjust_amount = st.number_input("Amount", min_value=-10000.0, value=float(balance), step=5.0, key=f"adjust_{selected_day}_{i}")
                with col2:
                    payment_notes = st.text_input("Payment Notes (optional)", key=f"date_notes_{selected_day}_{i}")
                with col3:
                    if st.button(f"💰 Record", key=f"pay_{selected_day}_{i}"):
                        if adjust_amount != 0:
                            new_balance = balance - adjust_amount
                            if save_payment(customer['Service'], customer['Customer Name'], adjust_amount, payment_notes):
                                st.success(f"Recorded ${adjust_amount:.2f}! New balance: ${new_balance:.2f}")
                                st.rerun()
                        else:
                            st.warning("Enter a non-zero amount")
    else:
        st.warning(f"No customers due on the {selected_day}{'st' if selected_day==1 else 'nd' if selected_day==2 else 'rd' if selected_day==3 else 'th'} of the month.")

with tab3:
    from datetime import datetime
    today = datetime.now().day
    
    st.subheader("⚠️ Past Due Customers")
    st.write(f"Customers whose due date has passed (before the {today}{'st' if today==1 else 'nd' if today==2 else 'rd' if today==3 else 'th'} of the month) and haven't paid yet.")
    
    # Get past due customers
    results = get_past_due_customers(all_data)
    
    if results:
        st.write(f"### ⚠️ {len(results)} Past Due Customer(s)")
        
        # Calculate totals
        total_amount = 0
        for c in results:
            try:
                amt = c['Amount Due'] if c['Amount Due'] else 0
                if amt:
                    amt_str = str(amt).strip()
                    if amt_str and amt_str.lower() != 'nan':
                        total_amount += float(amt_str)
            except:
                pass
        
        if total_amount > 0:
            st.error(f"💸 **Total Past Due: ${total_amount:,.2f}**")
        else:
            st.warning("💸 **Total Past Due: See individual amounts below**")
        
        # Display as table
        display_data = []
        for customer in results:
            charge_date = customer.get('Charge Date', '')
            due_day = customer.get('Due Day', '')
            if charge_date:
                try:
                    charge_date = pd.to_datetime(charge_date).strftime('%Y-%m-%d')
                except:
                    pass
            
            balance = get_balance(customer)
            
            display_data.append({
                'Service': customer['Service'],
                'Customer Name': customer['Customer Name'],
                'Phone': customer['Phone'],
                'Balance': f"${balance:.2f}",
                'Due Day': f"{due_day}{'st' if due_day==1 else 'nd' if due_day==2 else 'rd' if due_day==3 else 'th'}",
                'Status': customer['Status'],
                'Card Number': customer.get('Card Number', ''),
                'Exp': customer.get('Exp', ''),
                'Notes': customer.get('Notes', '')
            })
        
        display_df = pd.DataFrame(display_data)
        st.dataframe(display_df, hide_index=True, use_container_width=True)
        
        # Payment section for each
        st.write("### 💳 Record Payments")
        
        for i, customer in enumerate(results):
            due_day = customer.get('Due Day', '')
            balance = get_balance(customer)
            with st.expander(f"⚠️ {customer['Customer Name']} ({customer['Service']}) - 💰 Balance: ${balance:.2f} - Due: {due_day}{'st' if due_day==1 else 'nd' if due_day==2 else 'rd' if due_day==3 else 'th'}"):
                # Allow custom amount (including negative for adjustments)
                col1, col2, col3 = st.columns([1, 2, 1])
                with col1:
                    adjust_amount = st.number_input("Amount", min_value=-10000.0, value=float(balance), step=5.0, key=f"adjust_pastdue_{i}")
                with col2:
                    payment_notes = st.text_input("Payment Notes (optional)", key=f"pastdue_notes_{i}")
                with col3:
                    if st.button(f"💰 Record", key=f"pay_pastdue_{i}"):
                        if adjust_amount != 0:
                            new_balance = balance - adjust_amount
                            if save_payment(customer['Service'], customer['Customer Name'], adjust_amount, payment_notes):
                                st.success(f"Recorded ${adjust_amount:.2f}! New balance: ${new_balance:.2f}")
                                st.rerun()
                        else:
                            st.warning("Enter a non-zero amount")
    else:
        st.success("✅ No past due customers! Everyone is paid up!")

with tab4:
    from datetime import datetime, timedelta
    
    st.subheader("💰 Collections History")
    st.write("View payments collected on a specific day or date range")
    
    # Date selection
    col1, col2, col3 = st.columns([1, 1, 1])
    
    with col1:
        # Default to today
        today = datetime.now().date()
        start_date = st.date_input("Start Date", value=today)
    
    with col2:
        end_date = st.date_input("End Date", value=today)
    
    with col3:
        st.write("")
        st.write("")
        # Quick select buttons
        col_q1, col_q2, col_q3 = st.columns(3)
        with col_q1:
            if st.button("Today"):
                start_date = today
                end_date = today
                st.rerun()
        with col_q2:
            if st.button("Yesterday"):
                yesterday = today - timedelta(days=1)
                start_date = yesterday
                end_date = yesterday
                st.rerun()
        with col_q3:
            if st.button("This Week"):
                start_date = today - timedelta(days=today.weekday())
                end_date = today
                st.rerun()
    
    # Get collections for date range
    collections = get_collections_by_date(all_data, start_date, end_date)
    
    if collections:
        st.write(f"### 📋 Payments from {start_date} to {end_date}")
        st.write(f"**{len(collections)}** payment(s) found")
        
        # Calculate total collected
        total_collected = 0
        for c in collections:
            try:
                amt = c['Amount Collected']
                if amt:
                    amt_str = str(amt).strip()
                    if amt_str and amt_str.lower() != 'nan':
                        total_collected += float(amt_str)
            except:
                pass
        
        if total_collected > 0:
            st.success(f"💵 **Total Collected: ${total_collected:,.2f}**")
        
        # Display as table
        display_data = []
        for payment in collections:
            display_data.append({
                'Service': payment['Service'],
                'Customer Name': payment['Customer Name'],
                'Phone': payment['Phone'],
                'Amount': f"${payment['Amount Collected']:.2f}" if payment['Amount Collected'] else "$0.00",
                'Payment Date': payment['Payment Date'],
                'Status': payment['Status']
            })
        
        display_df = pd.DataFrame(display_data)
        st.dataframe(display_df, hide_index=True, use_container_width=True)
        
        # Summary by service
        st.write("### 📊 Summary by Service")
        service_totals = {}
        for payment in collections:
            service = payment['Service']
            amt = payment['Amount Collected'] if payment['Amount Collected'] else 0
            try:
                amt = float(str(amt).strip())
            except:
                amt = 0
            if service in service_totals:
                service_totals[service] += amt
            else:
                service_totals[service] = amt
        
        for service, total in service_totals.items():
            st.metric(service, f"${total:,.2f}")
        
    else:
        st.info(f"No payments found between {start_date} and {end_date}")



# Admin section for database management
st.divider()
st.header("🗄️ Database Management")

# Show where data is being stored
st.info(f"📁 Data file location: `{EXCEL_FILE}`")

with st.expander("Upload New Database"):
    st.write("Upload your Excel file. It will be saved to the persistent disk.")
    uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx'])
    if uploaded_file is not None:
        if st.button("Replace Database"):
            # Save to persistent disk path
            target_path = os.path.join(RENDER_DISK_PATH, "cleaned_billing_by_service.xlsx") if os.path.exists(RENDER_DISK_PATH) else EXCEL_FILE
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            with open(target_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())
            st.success(f"Database saved to {target_path}! Restarting...")
            st.rerun()

# Footer
st.divider()
st.caption(f"💾 Data file: {EXCEL_FILE}")
