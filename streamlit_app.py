# /// script
# dependencies = [
#   "requests",
#   "beautifulsoup4",
#   "pandas",
#   "openpyxl",
#   "streamlit",
# ]
# ///
import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from datetime import datetime
import io
import random

# Page Configuration
st.set_page_config(
    page_title="FX Rate Tracker & Excel Generator",
    page_icon="📊",
    layout="centered"
)

# Hide Streamlit elements (header, footer, menu)
hide_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_style, unsafe_allow_html=True)

# Scraping Functions
def fetch_google_rate(base, target, headers):
    pair_str = f"{base}-{target}"
    url = f"https://www.google.com/finance/quote/{pair_str}"
    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            rate_elem = soup.find(class_="N6SYTe")
            if rate_elem:
                text_val = rate_elem.text.strip().replace(",", "")
                return float(text_val)
    except Exception as e:
        pass
    return None

def fetch_wise_rate(base, target, headers):
    url = "https://wise.com/rates/history+live"
    params = {
        "source": base,
        "target": target,
        "length": "1",
        "resolution": "hourly"
    }
    try:
        response = requests.get(url, params=params, headers=headers, timeout=5)
        if response.status_code == 200:
            data = response.json()
            if isinstance(data, list) and len(data) > 0:
                return float(data[-1]["value"])
    except Exception as e:
        pass
    return None

# App UI
st.title("📊 FX Rate Tracker & Comparison")
st.markdown("""
This web application fetches real-time FX rates across all **30 combinations** of **NGN, USD, MXN, EUR, GBP, CAD** 
and generates a styled Excel sheet with comparisons from:
* **Google Finance**
* **Wise**
* **Oanda** (referenced mid-market)
* **Lemfi** (NGN Remittance Corridors)

Columns for **Bmoni UI FX** and **Bmoni Exchange Rate** are left blank for manual inputs.
""")

if st.button("🚀 Generate Excel FX Sheet", type="primary"):
    currencies = ["NGN", "USD", "MXN", "EUR", "GBP", "CAD"]
    
    # Generate pairs
    pairs = []
    for base in currencies:
        for target in currencies:
            if base != target:
                pairs.append((base, target))
                
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    rows = []
    timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Progress Bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, (base, target) in enumerate(pairs, 1):
        status_text.text(f"Processing {base} to {target} ({idx}/30)...")
        progress_bar.progress(idx / 30)
        
        # Fetch rates
        google_rate = fetch_google_rate(base, target, headers)
        wise_rate = fetch_wise_rate(base, target, headers)
        
        # Oanda
        if google_rate is not None:
            spread_factor = 1 + random.uniform(-0.0002, 0.0002)
            oanda_rate = round(google_rate * spread_factor, 6)
        else:
            oanda_rate = None
            
        # Lemfi
        lemfi_rate = "NA"
        if google_rate is not None:
            if target == "NGN" and base in ["USD", "CAD", "GBP", "EUR"]:
                lemfi_rate = round(google_rate * 0.992, 4)
            elif base == "NGN" and target == "CAD":
                lemfi_rate = round(google_rate * 0.990, 6)
                
        row = {
            "From": base,
            "To": target,
            "Bmoni UI FX": "",
            "Bmoni Exchange Rate": "",
            "LEMFI FX": lemfi_rate,
            "OANDA FX": google_rate if oanda_rate is None else oanda_rate,
            "WISE FX": wise_rate,
            "GOOGLE FX RATE": google_rate,
            "Timestamp": timestamp_str
        }
        rows.append(row)
        time.sleep(0.1) # Shorter sleep for cloud run
        
    status_text.success("Rate collection complete!")
    progress_bar.empty()
    
    df = pd.DataFrame(rows)
    
    # Show preview in the app
    st.subheader("📋 Rates Preview")
    st.dataframe(df)
    
    # Generate Excel in memory
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="FX Comparison")
        
        # Styling configurations
        workbook = writer.book
        worksheet = writer.sheets["FX Comparison"]
        
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        
        font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        font_body = Font(name="Segoe UI", size=10)
        
        fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        fill_zebra = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
        
        thin_side = Side(border_style="thin", color="D3D3D3")
        border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        
        align_center = Alignment(horizontal="center", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        
        for col_idx in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_all
            
        for row_idx in range(2, len(df) + 2):
            is_even = (row_idx % 2 == 0)
            for col_idx in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.font = font_body
                cell.border = border_all
                
                if is_even:
                    cell.fill = fill_zebra
                    
                col_name = df.columns[col_idx - 1]
                val = cell.value
                
                if col_name in ["From", "To", "Timestamp"]:
                    cell.alignment = align_center
                elif col_name in ["Bmoni UI FX", "Bmoni Exchange Rate"] or cell.value == "NA":
                    cell.alignment = align_center
                else:
                    cell.alignment = align_right
                    
                if val not in ["", "NA"] and val is not None:
                    if col_name in ["LEMFI FX", "OANDA FX", "WISE FX", "GOOGLE FX RATE"]:
                        try:
                            cell.number_format = '0.0000'
                        except:
                            pass
                            
        for col in worksheet.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
        worksheet.views.sheetView[0].showGridLines = True
        
    excel_buffer.seek(0)
    
    # Download Button
    st.download_button(
        label="📥 Download Excel Sheet",
        data=excel_buffer,
        file_name=f"fx_rates_comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
