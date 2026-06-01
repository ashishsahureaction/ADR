import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
import os
import smtplib
from email.message import EmailMessage
import pywhatkit as kit
import time
import win32com.client as win32

# --- CONFIG ---
adr_companies = {
    "INFY": "Infosys Ltd.",
    "WIT": "Wipro Ltd.",
    "HDB": "HDFC Bank Ltd.",
    "IBN": "ICICI Bank Ltd.",
    "RDY": "Dr. Reddy's Laboratories Ltd.",
    "RELI": "Reliance Industries Ltd.",
    "SUN": "Sun Pharmaceutical Industries Ltd.",
    "AUBN": "Aurobindo Pharma Ltd.",
    "DIVI": "Divi's Laboratories Ltd.",
    "MIND": "Mindtree Ltd."
}

excel_file = r"C:\Users\deepc\OneDrive\Desktop\Final\Mail ADR\ADR_Tracker\Indian_ADR_Prices.xlsx"
whatsapp_first_run_file = r"C:\Users\deepc\OneDrive\Desktop\Final\Mail ADR\ADR_Tracker\whatsapp_logged_in.flag"

# --- FETCH DATA ---
data = []
for ticker, company in adr_companies.items():
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period="2d")
        if hist.empty or len(hist) < 2:
            continue
        prev_close = hist['Close'].iloc[-2]
        open_price = hist['Open'].iloc[-1]
        last_close = hist['Close'].iloc[-1]
        percent_change = round(((last_close - prev_close) / prev_close) * 100, 2)
        direction = "↑" if percent_change > 0 else "↓" if percent_change < 0 else "-"

        data.append({
            "DateTime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Company": company,
            "ADR Ticker": ticker,
            "Open Price (USD)": round(open_price, 2),
            "Previous Close (USD)": round(prev_close, 2),
            "Current Close (USD)": round(last_close, 2),
            "% Change": percent_change,
            "Direction": direction
        })
    except:
        pass

df = pd.DataFrame(data).sort_values("Company").reset_index(drop=True)
if df.empty:
    exit()

# --- UPDATE EXCEL SILENTLY ---
def update_excel_silent(df, file_path):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(file_path)
    except:
        wb = excel.Workbooks.Add()
        wb.SaveAs(file_path)

    ws_name = "ADR Prices"
    try:
        ws = wb.Sheets(ws_name)
        ws.Cells.Clear()
    except:
        ws = wb.Sheets.Add()
        ws.Name = ws_name

    for col_idx, col_name in enumerate(df.columns, start=1):
        ws.Cells(1, col_idx).Value = col_name

    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.Cells(row_idx, col_idx).Value = value

    # Format
    max_row = df.shape[0] + 1
    max_col = df.shape[1]
    for r in range(2, max_row + 1):
        direction = ws.Cells(r, 8).Value
        change = ws.Cells(r, 7).Value
        if direction == "↑":
            ws.Cells(r, 8).Font.Color = 0x008000
            ws.Cells(r, 7).Font.Color = 0x008000
        elif direction == "↓":
            ws.Cells(r, 8).Font.Color = 0xFF0000
            ws.Cells(r, 7).Font.Color = 0xFF0000
        try:
            if abs(float(change)) > 2:
                for c in range(1, max_col + 1):
                    cell = ws.Cells(r, c)
                    cell.Font.Bold = True
                    cell.Interior.Color = 0xFFFACD
        except:
            pass

    ws.Columns.AutoFit()
    wb.Save()
    wb.Close(SaveChanges=True)
    excel.Quit()

update_excel_silent(df, excel_file)

# --- SEND EMAIL ---
def send_excel_via_email(sender_email, sender_password, receiver_email, subject, body, attachment_path):
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg.set_content(body)
        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(attachment_path)
        msg.add_attachment(file_data, maintype='application',
                           subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           filename=file_name)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
    except:
        pass

sender_email = "ashishsahu.reaction@gmail.com"
sender_password = "xsny kjmt wqjk quqp"
receiver_email = "ashishsahu.itm@gmail.com"
subject = "Updated ADR Prices"
body = "Hello,\n\nPlease find attached the latest ADR Prices Excel sheet.\n\nRegards"
send_excel_via_email(sender_email, sender_password, receiver_email, subject, body, excel_file)

# --- WHATSAPP WITH FIRST-RUN DETECTION ---
message_text = "📊 Latest ADR Prices:\n\n"
for idx, row in df.iterrows():
    message_text += f"{row['Company']} ({row['ADR Ticker']}): {row['Current Close (USD)']} USD {row['Direction']} ({row['% Change']}%)\n"
message_text = message_text[:1000]
receiver_number = "+14378702975"
hour, minute = 22, 0  # 10 PM

max_retries = 5
for attempt in range(max_retries):
    try:
        # Check if first run
        if not os.path.exists(whatsapp_first_run_file):
            print("⚠️ First run: Please scan the QR code in the browser window.")
        kit.sendwhatmsg(receiver_number, message_text, hour, minute, wait_time=10, tab_close=True)
        # Mark first run completed
        if not os.path.exists(whatsapp_first_run_file):
            with open(whatsapp_first_run_file, "w") as f:
                f.write("logged_in")
        print(f"✅ WhatsApp message scheduled successfully on attempt {attempt+1}")
        break
    except Exception as e:
        print(f"⚠️ WhatsApp send failed on attempt {attempt+1}, retrying in 30 seconds...")
        time.sleep(30)
else:
    print("❌ Failed to schedule WhatsApp message after multiple attempts.")
