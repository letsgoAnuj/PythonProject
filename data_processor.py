import os
import glob
import re
import pandas as pd
import numpy as np
import warnings
import sys
import win32com.client as win32

sys.stdout.reconfigure(encoding='utf-8')

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

DOWNLOAD_DIR = os.path.join(os.getcwd(), "polymer_data")
ANALYSIS_DIR = os.path.join(os.getcwd(), "analysis_output")

if not os.path.exists(ANALYSIS_DIR):
    os.makedirs(ANALYSIS_DIR)


# ==========================================
# 0. ROBUSTNESS: AUTO-CLOSE EXCEL HELPER
# ==========================================
def auto_save_and_close_excel(target_filepath):
    """Detects if the target Excel file is open, saves it, and closes it gracefully."""
    abs_path = os.path.abspath(target_filepath)
    try:
        # Connect to the active Excel application
        excel = win32.GetObject(None, "Excel.Application")
        for wb in excel.Workbooks:
            if wb.FullName.lower() == abs_path.lower():
                print(f"⚠️ Detected {os.path.basename(abs_path)} is open. Auto-saving and closing...")
                wb.Save()
                wb.Close(SaveChanges=True)
                print("✓ File closed successfully.")
                break
    except Exception:
        # Excel is not running or the file isn't open via COM
        pass


# ==========================================
# 1. CLEANING HELPER FUNCTIONS
# ==========================================
def clean_polymer_price(val):
    if pd.isna(val) or str(val).strip() == '-' or str(val).strip() == '':
        return np.nan
    val_str = str(val).replace(',', '')
    match = re.search(r"[-+]?\d*\.\d+|\d+", val_str)
    return float(match.group()) if match else np.nan


def clean_latex_price(val):
    if pd.isna(val):
        return np.nan
    val_str = str(val).replace(',', '')
    match = re.search(r"\d+\.\d+|\d+", val_str)
    if match:
        num = float(match.group())
        return num if num > 0 else np.nan
    return np.nan


def reorder_columns(df):
    standard_cols = ['Date', 'Product', 'Final_INR_KG_Price', 'Conversion_Formula']
    helper_cols = [c for c in df.columns if c not in standard_cols]
    return df[standard_cols + helper_cols]


# ==========================================
# 2. PRODUCT-SPECIFIC PROCESSORS
# ==========================================
def process_usd_mt_products(filepath, name, usd_inr_df):
    try:
        df_peek = pd.read_excel(filepath, nrows=2)
        has_direct_headers = any(
            any(keyword in str(col).lower() for keyword in ['min. price', 'product', 'price date', 'price'])
            for col in df_peek.columns
        )
        if has_direct_headers:
            df = pd.read_excel(filepath)
        else:
            df = pd.read_excel(filepath, skiprows=3)
    except Exception:
        df = pd.read_excel(filepath, skiprows=3)

    date_col = next((col for col in df.columns if 'date' in str(col).lower()), df.columns[0])
    df['Date'] = pd.to_datetime(df[date_col], errors='coerce')

    if "Min. Price" in df.columns and "Max. Price" in df.columns:
        df['Min_Clean'] = df['Min. Price'].apply(clean_polymer_price)
        df['Max_Clean'] = df['Max. Price'].apply(clean_polymer_price)
        df['Calculated_USD_MT_Price'] = df[['Min_Clean', 'Max_Clean']].mean(axis=1, skipna=True)
    else:
        raw_price_col = next(
            (c for c in df.columns if any(x in str(c).lower() for x in ['price', 'value', 'assessment'])),
            df.columns[1])
        df['Calculated_USD_MT_Price'] = df[raw_price_col].apply(clean_polymer_price)

    df = df.dropna(subset=['Date', 'Calculated_USD_MT_Price'])
    fx_rates = usd_inr_df[['Date', 'Price']].rename(columns={'Price': 'USD_Exchange_Rate'})
    df = pd.merge(df, fx_rates, on='Date', how='left')
    df['USD_Exchange_Rate'] = df['USD_Exchange_Rate'].ffill().bfill()
    df['Final_INR_KG_Price'] = (df['Calculated_USD_MT_Price'] / 1000) * df['USD_Exchange_Rate']
    df['Product'] = name
    df['Conversion_Formula'] = "(Calculated_USD_MT_Price / 1000) * USD_Exchange_Rate"

    return reorder_columns(df).sort_values('Date', ascending=False)


def process_domestic_inr(filepath, name):
    try:
        df_peek = pd.read_excel(filepath, nrows=2)
        has_direct_headers = any(
            any(keyword in str(col).lower() for keyword in ['min. price', 'product', 'price date', 'price'])
            for col in df_peek.columns
        )
        if has_direct_headers:
            df = pd.read_excel(filepath)
        else:
            df = pd.read_excel(filepath, skiprows=3)
    except Exception:
        df = pd.read_excel(filepath, skiprows=3)

    date_col = next((col for col in df.columns if 'date' in str(col).lower()), df.columns[0])
    df['Date'] = pd.to_datetime(df[date_col], errors='coerce')

    if "Min. Price" in df.columns and "Max. Price" in df.columns:
        df['Min_Clean'] = df['Min. Price'].apply(clean_polymer_price)
        df['Max_Clean'] = df['Max. Price'].apply(clean_polymer_price)
        df['Final_INR_KG_Price'] = df[['Min_Clean', 'Max_Clean']].mean(axis=1, skipna=True)
        df['Conversion_Formula'] = "Average of Min_Clean and Max_Clean (Already in INR/KG)"
    else:
        raw_price_col = next((c for c in df.columns if any(x in str(c).lower() for x in ['price', 'value'])),
                             df.columns[1])
        df['Final_INR_KG_Price'] = df[raw_price_col].apply(clean_polymer_price)
        df['Conversion_Formula'] = "Direct clean of raw price (Already in INR/KG)"

    df['Product'] = name
    df = df.dropna(subset=['Date', 'Final_INR_KG_Price'])
    return reorder_columns(df).sort_values('Date', ascending=False)


def process_rbi(filepath, product_name):
    tables = pd.read_html(filepath)
    df = max(tables, key=lambda d: d.shape[0])
    date_col = next((col for col in df.columns if 'date' in str(col).lower()), df.columns[0])
    usd_col = df.columns[1]

    df['Price'] = df[usd_col].apply(clean_polymer_price)
    df['Date'] = pd.to_datetime(df[date_col], format='%d/%m/%Y', errors='coerce')
    df = df.dropna(subset=['Date', 'Price']).sort_values('Date').set_index('Date')
    df = df.asfreq('D')
    df['Price'] = df['Price'].interpolate(method='linear')
    df = df.reset_index()

    df['UOM'] = f"INR/{product_name.split('-')[0]}"
    df['Product'] = product_name
    return df[['Date', 'Price', 'UOM', 'Product']].sort_values('Date', ascending=False)


def process_cny_inr(filepath):
    df = pd.read_excel(filepath)
    df['Date'] = pd.to_datetime(
        df['Date'].astype(str).str.extract(r'(\d{2}/\d{2}/\d{4})')[0],
        format='%d/%m/%Y', errors='coerce'
    )

    # --- FIX: Handle both new API pure numbers and old string formats ---
    if pd.api.types.is_numeric_dtype(df['Close']):
        df['Price'] = df['Close'].astype(float)
    else:
        # Fallback for old exchangerates.org.uk format (e.g., "1 CNY = 14.15 INR")
        df['Price'] = df['Close'].astype(str).str.extract(r'=\s*([\d\.]+)')[0].astype(float)

    df['Product'] = "CNY-INR"
    df['UOM'] = "INR/CNY"
    df = df.dropna(subset=['Date', 'Price']).sort_values('Date', ascending=False)
    return df[['Date', 'Price', 'UOM', 'Product']]


def process_latex(filepath):
    df = pd.read_excel(filepath)
    date_col = df.columns[0]
    price_cols = df.columns[1:]
    for col in price_cols:
        df[f"{col}_Cleaned"] = df[col].apply(clean_latex_price)

    cleaned_cols = [f"{col}_Cleaned" for col in price_cols]
    df['Final_INR_KG_Price'] = df[cleaned_cols].mean(axis=1, skipna=True)
    df['Date'] = pd.to_datetime(df[date_col], errors='coerce')
    df['Product'] = 'LATEX'
    df['Conversion_Formula'] = "Average of available non-zero daily columns (No FX needed)"
    df = df.dropna(subset=['Date', 'Final_INR_KG_Price'])
    return reorder_columns(df).sort_values('Date', ascending=False)


def process_brent_crude(filepath, usd_inr_df):
    try:
        df = pd.read_excel(filepath)
        if 'Date' not in df.columns and 'Price' not in df.columns:
            df = pd.read_excel(filepath, skiprows=3)
    except Exception:
        df = pd.read_excel(filepath, skiprows=3)

    date_col = next((col for col in df.columns if 'date' in str(col).lower()), df.columns[0])
    df['Date'] = pd.to_datetime(df[date_col], errors='coerce').dt.normalize()

    raw_price_col = next((c for c in df.columns if any(x in str(c).lower() for x in ['price', 'value', 'assessment'])),
                         df.columns[1])
    df['Clean_USD_BBL_Price'] = df[raw_price_col].apply(clean_polymer_price)
    df = df.dropna(subset=['Date', 'Clean_USD_BBL_Price'])

    df['Final_INR_KG_Price'] = df['Clean_USD_BBL_Price']
    df['Product'] = 'Brent'
    df['Conversion_Formula'] = "Direct clean of raw price (Maintained in USD/BBL)"

    return reorder_columns(df).sort_values('Date', ascending=False)


def process_westmetall_metals(filepath, name, fx_df):
    df = pd.read_excel(filepath)
    if df.shape[1] < 2:
        raise ValueError("Input file must contain at least two columns.")
    name = name.strip().upper()
    if name not in ("ZINC", "BRASS"):
        raise ValueError(f"Unsupported product: {name}")

    raw_col = df.columns[1]
    df["Clean_Raw_Price"] = pd.to_numeric(
        df[raw_col].apply(clean_polymer_price),
        errors="coerce"
    )
    df["Date"] = pd.to_datetime(
        df[df.columns[0]],
        errors="coerce"
    ).dt.normalize()

    df = df.dropna(subset=["Date", "Clean_Raw_Price"]).copy()

    fx_col_name = "USD_Exchange_Rate" if name == "ZINC" else "EUR_Exchange_Rate"
    fx_rates = fx_df[["Date", "Price"]].copy()
    fx_rates["Date"] = pd.to_datetime(fx_rates["Date"], errors="coerce").dt.normalize()
    fx_rates["Price"] = pd.to_numeric(fx_rates["Price"], errors="coerce")

    fx_rates = (
        fx_rates.dropna(subset=["Date", "Price"])
        .sort_values("Date")
        .drop_duplicates(subset=["Date"], keep="last")
        .rename(columns={"Price": fx_col_name})
    )

    df = df.merge(
        fx_rates[["Date", fx_col_name]],
        on="Date",
        how="left",
        validate="many_to_one"
    )

    df = df.sort_values("Date")
    df[fx_col_name] = df[fx_col_name].ffill().bfill()

    if df[fx_col_name].isna().any():
        raise ValueError(f"No valid {fx_col_name} data available for conversion.")

    if name == "ZINC":
        df["Final_INR_KG_Price"] = (df["Clean_Raw_Price"] / 1000) * df[fx_col_name]
        df["Conversion_Formula"] = "(Clean_Raw_Price [USD/TON] / 1000) * USD_Exchange_Rate"
    elif name == "BRASS":
        df["Final_INR_KG_Price"] = (df["Clean_Raw_Price"] / 100) * df[fx_col_name]
        df["Conversion_Formula"] = "(Clean_Raw_Price [EUR/100kg] / 100) * EUR_Exchange_Rate"

    df["Product"] = name
    return reorder_columns(df).sort_values("Date", ascending=False).reset_index(drop=True)


# ==========================================
# 3. MAIN PIPELINE EXECUTION
# ==========================================
def run_data_pipeline():
    master_file_path = os.path.join(ANALYSIS_DIR, "Master_Trend_Data.xlsx")

    print("\n--- Phase 0: Pre-Flight File Lock Check ---")
    auto_save_and_close_excel(master_file_path)

    # Final lock test to ensure we have write permissions before heavy processing
    if os.path.exists(master_file_path):
        try:
            with open(master_file_path, "a"):
                pass
        except PermissionError:
            error_msg = f"CRITICAL ERROR: {os.path.basename(master_file_path)} is locked and could not be auto-closed."
            print(error_msg, file=sys.stderr)
            sys.exit(1)  # This triggers the admin trap in automator.py
    print("✓ Output target is unlocked and ready.")

    product_dfs = []
    fx_dfs = []

    print("\n--- Phase 1: Building Exchange Rate Timelines ---")
    usd_files = glob.glob(os.path.join(DOWNLOAD_DIR, "USD-INR*.xls*"))
    eur_files = glob.glob(os.path.join(DOWNLOAD_DIR, "EUR-INR*.xls*"))
    cny_files = glob.glob(os.path.join(DOWNLOAD_DIR, "CNY-INR*.xls*"))

    usd_inr_df = process_rbi(usd_files[0], "USD-INR") if usd_files else None
    eur_inr_df = process_rbi(eur_files[0], "EUR-INR") if eur_files else None
    cny_inr_df = process_cny_inr(cny_files[0]) if cny_files else None

    if usd_inr_df is not None:
        fx_dfs.append(usd_inr_df)
        print("✓ USD-INR processed.")
    if eur_inr_df is not None:
        fx_dfs.append(eur_inr_df)
        print("✓ EUR-INR processed.")
    if cny_inr_df is not None:
        fx_dfs.append(cny_inr_df)
        print("✓ CNY-INR processed.")

    print("\n--- Phase 2: Processing Products to INR/KG ---")
    files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.*"))

    for filepath in files:
        filename = os.path.basename(filepath)
        name = re.sub(r'_last_\d+_year_trend\.(xlsx|xls)$', '', filename)

        if name in ["USD-INR", "EUR-INR", "CNY-INR"] or filename.startswith("~$"):
            continue

        try:
            df_clean = pd.DataFrame()
            if name == "LATEX":
                df_clean = process_latex(filepath)
            elif name == "BRASS":
                df_clean = process_westmetall_metals(filepath, name, eur_inr_df)
            elif name == "ZINC":
                df_clean = process_westmetall_metals(filepath, name, usd_inr_df)
            elif name == "Brent":
                df_clean = process_brent_crude(filepath, usd_inr_df)
            elif name in ["EVA", "EVA_PVC", "PVC_Suspension", "ACETONE", "PVC", "LDPE", "LLDPE"]:
                df_clean = process_domestic_inr(filepath, name)
            elif name in ["MEG", "PTA", "Naphtha", "DOP", "PP"]:
                df_clean = process_usd_mt_products(filepath, name, usd_inr_df)

            if not df_clean.empty:
                product_dfs.append(df_clean)
                print(f" -> Success: {name} ({len(df_clean)} records).")
        except Exception as e:
            # Graceful failure: If schema changes or a corrupt file sneaks in, log it and continue
            print(f" -> Failed to process {name}. Error: {e}")

    print("\n--- Phase 3: Calculating Historical Price Variations ---")

    master_product_df = pd.concat(product_dfs, ignore_index=True) if product_dfs else pd.DataFrame()
    master_fx_df = pd.concat(fx_dfs, ignore_index=True) if fx_dfs else pd.DataFrame()

    intervals = {
        "1 Day": 1, "3 Days": 3, "7 Days": 7, "2 Weeks": 14,
        "1 Month": 30, "2 Months": 60,
        "3 Months": 90, "6 Months": 180, "1 Year": 365
    }
    matrix_rows = []

    if not master_product_df.empty:
        for prod in master_product_df['Product'].unique():
            df_sub = master_product_df[master_product_df['Product'] == prod].sort_values('Date', ascending=False)
            if df_sub.empty: continue
            latest_price = df_sub.iloc[0]['Final_INR_KG_Price']
            latest_date = df_sub.iloc[0]['Date']
            row = {"Commodity": prod, "Latest Price": latest_price}

            for label, days in intervals.items():
                target_date = latest_date - pd.Timedelta(days=days)
                past_df = df_sub[df_sub['Date'] <= target_date]

                if not past_df.empty:
                    past_price = past_df.iloc[0]['Final_INR_KG_Price']
                    row[label] = ((latest_price - past_price) / past_price) * 100
                else:
                    oldest_date = df_sub['Date'].min()
                    if (oldest_date - target_date).days <= 20:
                        past_price = df_sub.iloc[-1]['Final_INR_KG_Price']
                        row[label] = ((latest_price - past_price) / past_price) * 100
                    else:
                        row[label] = None
            matrix_rows.append(row)

    if not master_fx_df.empty:
        for fx_pair in master_fx_df['Product'].unique():
            df_sub = master_fx_df[master_fx_df['Product'] == fx_pair].sort_values('Date', ascending=False)
            if df_sub.empty: continue
            latest_price = df_sub.iloc[0]['Price']
            latest_date = df_sub.iloc[0]['Date']
            row = {"Commodity": fx_pair, "Latest Price": latest_price}

            for label, days in intervals.items():
                target_date = latest_date - pd.Timedelta(days=days)
                past_df = df_sub[df_sub['Date'] <= target_date]

                if not past_df.empty:
                    past_price = past_df.iloc[0]['Price']
                    row[label] = ((latest_price - past_price) / past_price) * 100
                else:
                    oldest_date = df_sub['Date'].min()
                    if (oldest_date - target_date).days <= 20:
                        past_price = df_sub.iloc[-1]['Price']
                        row[label] = ((latest_price - past_price) / past_price) * 100
                    else:
                        row[label] = None
            matrix_rows.append(row)

    variation_matrix = pd.DataFrame(matrix_rows)

    # =========================================================
    # PHASE 4: FINALIZING MASTER DATASET WITH ADVANCED EXCEL
    # =========================================================
    print("\n--- Phase 4: Finalizing Master Dataset (Advanced Excel Formatting) ---")

    try:
        with pd.ExcelWriter(master_file_path, engine='xlsxwriter', datetime_format='dd-mmm-yyyy') as writer:
            workbook = writer.book

            header_format = workbook.add_format({
                'bold': True, 'valign': 'vcenter', 'align': 'center',
                'bg_color': '#003366', 'font_color': '#FFFFFF', 'border': 1
            })

            # --- UPDATED DATE & MONTH FORMATS ---
            date_format = workbook.add_format({'num_format': 'dd-mmm-yyyy', 'border': 1, 'align': 'center'})
            month_format = workbook.add_format({'num_format': 'mmm-yyyy', 'border': 1, 'align': 'center'})

            price_format = workbook.add_format({'num_format': '₹#,##0.00', 'border': 1})
            usd_format = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})
            standard_format = workbook.add_format({'border': 1, 'align': 'left'})

            link_fmt = workbook.add_format(
                {'font_color': '#0052cc', 'underline': True, 'bold': True, 'valign': 'vcenter', 'font_size': 11})
            back_btn_fmt = workbook.add_format(
                {'font_color': '#ffffff', 'bg_color': '#e52d27', 'bold': True, 'align': 'center', 'valign': 'vcenter',
                 'border': 1})
            nav_title_fmt = workbook.add_format({'bold': True, 'font_size': 22, 'font_color': '#003366'})
            nav_sub_fmt = workbook.add_format({'font_size': 12, 'font_color': '#4a4a4a', 'italic': True})
            nav_section_fmt = workbook.add_format(
                {'bold': True, 'font_size': 14, 'font_color': '#1a1a1a', 'bottom': 2, 'bottom_color': '#003366'})

            green_text_fmt = workbook.add_format(
                {'font_color': '#00875a', 'bold': True, 'num_format': '+0.0%;-0.0%;-', 'border': 1, 'align': 'center'})
            red_text_fmt = workbook.add_format(
                {'font_color': '#e52d27', 'bold': True, 'num_format': '+0.0%;-0.0%;-', 'border': 1, 'align': 'center'})
            neutral_pct_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1, 'align': 'center'})

            # --- CREATE NAVIGATION TAB ---
            nav_ws = workbook.add_worksheet('Navigation')
            nav_ws.set_zoom(80)
            nav_ws.hide_gridlines(2)
            nav_ws.set_column('A:A', 3)
            nav_ws.set_column('B:C', 35)

            nav_ws.write('B2', 'Raw Material Procurement Dashboard', nav_title_fmt)
            nav_ws.write('B3', 'Click any metric below to view detailed charts and historical data.', nav_sub_fmt)
            nav_ws.write(5, 1, 'Commodity Price Trends', nav_section_fmt)
            nav_ws.write(5, 2, 'Macro Exchange Rates', nav_section_fmt)
            nav_ws.write_url(4, 1, "internal:'Variation_Matrix'!A1", string="📈 View Price Variation Matrix",
                             cell_format=link_fmt)

            prod_nav_row = 7
            fx_nav_row = 7

            if not master_product_df.empty:
                master_product_df['Date'] = pd.to_datetime(master_product_df['Date']).dt.normalize()
                master_product_df = master_product_df.sort_values(by=['Product', 'Date'], ascending=[True, False])
                master_product_df.to_excel(writer, sheet_name='Product_Trends', index=False)
                writer.sheets['Product_Trends'].set_zoom(80)
                writer.sheets['Product_Trends'].hide()

                for product in master_product_df['Product'].unique():
                    sheet_name = str(product).replace('/', '_').replace('[', '').replace(']', '')[:31]
                    prod_df = master_product_df[master_product_df['Product'] == product].copy()
                    prod_df = prod_df.sort_values(by='Date', ascending=False)
                    prod_df.to_excel(writer, sheet_name=sheet_name, index=False)

                    ws = writer.sheets[sheet_name]
                    ws.set_zoom(80)
                    ws.hide_gridlines(2)

                    is_brent = (product == 'Brent')
                    current_fmt = usd_format if is_brent else price_format
                    current_unit = "USD/BBL" if is_brent else "INR/KG"

                    num_cols = len(prod_df.columns)
                    month_col = num_cols + 1
                    avg_col = num_cols + 2
                    chart_col = num_cols + 4

                    ws.write_url(0, chart_col, "internal:'Navigation'!A1", string="⬅ Return to Dashboard",
                                 cell_format=back_btn_fmt)
                    ws.set_column(chart_col, chart_col, 25)

                    nav_ws.write_url(prod_nav_row, 1, f"internal:'{sheet_name}'!A1", string=f"📊 {product} Trend",
                                     cell_format=link_fmt)
                    prod_nav_row += 1

                    for col_num, col_name in enumerate(prod_df.columns):
                        display_col_name = "Final_USD_BBL_Price" if (
                                is_brent and col_name == 'Final_INR_KG_Price') else col_name
                        ws.write(0, col_num, display_col_name, header_format)

                        if col_name == 'Date':
                            ws.set_column(col_num, col_num, 15, date_format)
                        elif 'Price' in col_name or 'Rate' in col_name:
                            ws.set_column(col_num, col_num, 18, current_fmt)
                        else:
                            ws.set_column(col_num, col_num, 20, standard_format)

                    ws.set_column(3, num_cols, options={'hidden': True})
                    ws.freeze_panes(1, 0)

                    one_year_ago = pd.to_datetime('today') - pd.DateOffset(years=1)
                    monthly_df = prod_df[prod_df['Date'] >= one_year_ago].copy()
                    max_monthly_row = 0

                    if not monthly_df.empty:
                        monthly_df = monthly_df.set_index('Date').resample('MS').mean(numeric_only=True).reset_index()
                        monthly_df = monthly_df.dropna(subset=['Final_INR_KG_Price'])
                        monthly_df = monthly_df.sort_values(by='Date', ascending=False)

                        # APPLIED MONTH FORMAT HERE
                        ws.set_column(month_col, month_col, 15, month_format)
                        ws.set_column(avg_col, avg_col, 22, current_fmt)
                        ws.write(0, month_col, "Month", header_format)
                        ws.write(0, avg_col, f"Monthly Avg ({current_unit})", header_format)

                        for r_idx, row in enumerate(monthly_df.itertuples(), start=1):
                            ws.write(r_idx, month_col, row.Date, month_format)
                            ws.write(r_idx, avg_col, row.Final_INR_KG_Price, current_fmt)

                        max_monthly_row = len(monthly_df)

                    chart1 = workbook.add_chart({'type': 'line'})
                    max_daily_row = len(prod_df)
                    price_col_idx = prod_df.columns.get_loc('Final_INR_KG_Price')
                    chart1.add_series({
                        'name': 'Daily Price',
                        'categories': [sheet_name, 1, 0, max_daily_row, 0],
                        'values': [sheet_name, 1, price_col_idx, max_daily_row, price_col_idx],
                        'line': {'color': '#64748b', 'width': 1.5},
                    })
                    chart1.set_title({'name': f'{product} Historical Daily Trend'})
                    chart1.set_x_axis({'name': 'Date', 'date_axis': True})
                    chart1.set_y_axis({'name': f'Price ({current_unit})'})
                    chart1.set_legend({'none': True})
                    ws.insert_chart(1, chart_col, chart1, {'x_scale': 1.4, 'y_scale': 1.1, 'object_position': 3})

                    if not monthly_df.empty:
                        chart2 = workbook.add_chart({'type': 'line'})
                        chart2.add_series({
                            'name': '1-Year Monthly Avg',
                            'categories': [sheet_name, 1, month_col, max_monthly_row, month_col],
                            'values': [sheet_name, 1, avg_col, max_monthly_row, avg_col],
                            'line': {'color': '#0052cc', 'width': 2.5},
                            'marker': {'type': 'circle', 'size': 6, 'border': {'color': '#ffffff'}},
                            'data_labels': {'value': True, 'num_format': '$#,##0' if is_brent else '₹#,##0',
                                            'position': 'above'}
                        })
                        chart2.set_title({'name': f'{product} 1-Year Monthly Average'})
                        # APPLIED MONTH FORMAT TO X-AXIS
                        chart2.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'mmm-yyyy'})
                        chart2.set_y_axis({'name': f'Price ({current_unit})'})
                        chart2.set_legend({'none': True})
                        ws.insert_chart(18, chart_col, chart2, {'x_scale': 1.4, 'y_scale': 1.1, 'object_position': 3})

            if not master_fx_df.empty:
                master_fx_df['Date'] = pd.to_datetime(master_fx_df['Date']).dt.normalize()
                master_fx_df = master_fx_df.sort_values(by=['Product', 'Date'], ascending=[True, False])
                master_fx_df.to_excel(writer, sheet_name='Exchange_Rates', index=False)
                writer.sheets['Exchange_Rates'].set_zoom(80)
                writer.sheets['Exchange_Rates'].hide()

                fx_num_format = workbook.add_format({'num_format': '₹#,##0.00', 'border': 1})

                for fx_pair in master_fx_df['Product'].unique():
                    sheet_name = str(fx_pair)[:31]
                    fx_df = master_fx_df[master_fx_df['Product'] == fx_pair].copy()
                    fx_df = fx_df.sort_values(by='Date', ascending=False)
                    fx_df.to_excel(writer, sheet_name=sheet_name, index=False)

                    ws = writer.sheets[sheet_name]
                    ws.set_zoom(80)
                    ws.hide_gridlines(2)

                    num_cols = len(fx_df.columns)
                    month_col = num_cols + 1
                    avg_col = num_cols + 2
                    chart_col = num_cols + 4

                    ws.write_url(0, chart_col, "internal:'Navigation'!A1", string="⬅ Return to Dashboard",
                                 cell_format=back_btn_fmt)
                    ws.set_column(chart_col, chart_col, 25)

                    nav_ws.write_url(fx_nav_row, 2, f"internal:'{sheet_name}'!A1", string=f"💱 {fx_pair} Trend",
                                     cell_format=link_fmt)
                    fx_nav_row += 1

                    for col_num, col_name in enumerate(fx_df.columns):
                        ws.write(0, col_num, col_name, header_format)
                        if col_name == 'Date':
                            ws.set_column(col_num, col_num, 15, date_format)
                        elif 'Price' in col_name:
                            ws.set_column(col_num, col_num, 18, fx_num_format)
                        else:
                            ws.set_column(col_num, col_num, 20, standard_format)

                    ws.set_column(3, num_cols, options={'hidden': True})
                    ws.freeze_panes(1, 0)

                    one_year_ago = pd.to_datetime('today') - pd.DateOffset(years=1)
                    monthly_fx = fx_df[fx_df['Date'] >= one_year_ago].copy()
                    max_fx_monthly_row = 0

                    if not monthly_fx.empty:
                        monthly_fx = monthly_fx.set_index('Date').resample('MS').mean(numeric_only=True).reset_index()
                        monthly_fx = monthly_fx.dropna(subset=['Price'])
                        monthly_fx = monthly_fx.sort_values(by='Date', ascending=False)

                        # APPLIED MONTH FORMAT HERE
                        ws.set_column(month_col, month_col, 15, month_format)
                        ws.set_column(avg_col, avg_col, 18, fx_num_format)
                        ws.write(0, month_col, "Month", header_format)
                        ws.write(0, avg_col, "Monthly Avg", header_format)

                        for r_idx, row in enumerate(monthly_fx.itertuples(), start=1):
                            ws.write(r_idx, month_col, row.Date, month_format)
                            ws.write(r_idx, avg_col, row.Price, fx_num_format)

                        max_fx_monthly_row = len(monthly_fx)

                    chart1 = workbook.add_chart({'type': 'line'})
                    max_row = len(fx_df)
                    price_col_idx = fx_df.columns.get_loc('Price')
                    chart1.add_series({
                        'categories': [sheet_name, 1, 0, max_row, 0],
                        'values': [sheet_name, 1, price_col_idx, max_row, price_col_idx],
                        'line': {'color': '#94a3b8', 'width': 1.5},
                    })
                    chart1.set_title({'name': f'{fx_pair} Historical Daily Trend'})
                    chart1.set_legend({'none': True})
                    ws.insert_chart(1, chart_col, chart1, {'x_scale': 1.4, 'y_scale': 1.1, 'object_position': 3})

                    if not monthly_fx.empty:
                        chart2 = workbook.add_chart({'type': 'line'})
                        chart2.add_series({
                            'categories': [sheet_name, 1, month_col, max_fx_monthly_row, month_col],
                            'values': [sheet_name, 1, avg_col, max_fx_monthly_row, avg_col],
                            'line': {'color': '#e52d27', 'width': 2.5},
                            'marker': {'type': 'circle', 'size': 6, 'border': {'color': '#ffffff'}},
                            'data_labels': {'value': True, 'num_format': '₹#,##0.00', 'position': 'above'}
                        })
                        chart2.set_title({'name': f'{fx_pair} 1-Year Monthly Average'})
                        # APPLIED MONTH FORMAT TO X-AXIS
                        chart2.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'mmm-yyyy'})
                        chart2.set_legend({'none': True})
                        ws.insert_chart(18, chart_col, chart2, {'x_scale': 1.4, 'y_scale': 1.1, 'object_position': 3})

            # --- WRITE & COLORIZE VARIATION MATRIX TAB ---
            variation_matrix.to_excel(writer, sheet_name='Variation_Matrix', index=False)
            var_ws = writer.sheets['Variation_Matrix']
            var_ws.set_zoom(90)

            for col_num, col_name in enumerate(variation_matrix.columns):
                var_ws.write(0, col_num, col_name, header_format)

            var_ws.set_column('A:A', 20, standard_format)
            var_ws.set_column('B:B', 15, price_format)
            var_ws.set_column('C:K', 12)

            for row_idx, row_data in variation_matrix.iterrows():
                var_ws.write(row_idx + 1, 0, row_data['Commodity'], standard_format)

                fmt = usd_format if row_data['Commodity'] == 'Brent' else price_format
                var_ws.write(row_idx + 1, 1, row_data['Latest Price'], fmt)

                for col_idx, (label, val) in enumerate(row_data.items()):
                    if col_idx < 2: continue

                    if pd.isna(val):
                        var_ws.write(row_idx + 1, col_idx, "-", neutral_pct_fmt)
                    else:
                        val_decimal = val / 100
                        if val > 0:
                            var_ws.write_number(row_idx + 1, col_idx, val_decimal, red_text_fmt)
                        elif val < 0:
                            var_ws.write_number(row_idx + 1, col_idx, val_decimal, green_text_fmt)
                        else:
                            var_ws.write_number(row_idx + 1, col_idx, val_decimal, neutral_pct_fmt)

            var_ws.write_url(0, 11, "internal:'Navigation'!A1", string="⬅ Return to Dashboard",
                             cell_format=back_btn_fmt)
            var_ws.set_column(11, 11, 25)

        print(f"\nSUCCESS! Formatted Master Data with Charts saved: {master_file_path}")

    except PermissionError:
        # Failsafe: If the file was opened *during* the generation phase
        error_msg = f"CRITICAL: '{os.path.basename(master_file_path)}' was locked during export phase."
        print(error_msg, file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    run_data_pipeline()