import os
import glob
import re
import pandas as pd
import numpy as np
import warnings

# Suppress openpyxl/pandas default style warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

DOWNLOAD_DIR = os.path.join(os.getcwd(), "polymer_data")
ANALYSIS_DIR = os.path.join(os.getcwd(), "analysis_output")

if not os.path.exists(ANALYSIS_DIR):
    os.makedirs(ANALYSIS_DIR)


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
    df = pd.read_excel(filepath, skiprows=3)
    date_col = next((col for col in df.columns if 'date' in str(col).lower()), df.columns[0])
    df['Date'] = pd.to_datetime(df[date_col], errors='coerce')
    raw_price_col = next((c for c in df.columns if any(x in str(c).lower() for x in ['price', 'value', 'assessment'])),
                         df.columns[1])
    df['Clean_USD_BBL_Price'] = df[raw_price_col].apply(clean_polymer_price)
    df = df.dropna(subset=['Date', 'Clean_USD_BBL_Price'])

    fx_rates = usd_inr_df[['Date', 'Price']].rename(columns={'Price': 'USD_Exchange_Rate'})
    df = pd.merge(df, fx_rates, on='Date', how='left')
    df['USD_Exchange_Rate'] = df['USD_Exchange_Rate'].ffill().bfill()

    df['Final_INR_KG_Price'] = (df['Clean_USD_BBL_Price'] * df['USD_Exchange_Rate']) / 136.4
    df['Product'] = 'Brent'
    df['Conversion_Formula'] = "(Clean_USD_BBL_Price * USD_Exchange_Rate) / 136.4"
    return reorder_columns(df).sort_values('Date', ascending=False)


def process_westmetall_metals(filepath, name, fx_df):
    df = pd.read_excel(filepath)
    raw_col = df.columns[1]
    df['Clean_Raw_Price'] = df[raw_col].apply(clean_polymer_price)
    df['Date'] = pd.to_datetime(df[df.columns[0]], errors='coerce')
    df = df.dropna(subset=['Date', 'Clean_Raw_Price'])

    fx_col_name = 'USD_Exchange_Rate' if name == 'ZINC' else 'EUR_Exchange_Rate'
    fx_rates = fx_df[['Date', 'Price']].rename(columns={'Price': fx_col_name})
    df = pd.merge(df, fx_rates, on='Date', how='left')
    df[fx_col_name] = df[fx_col_name].ffill().bfill()

    if name == "ZINC":
        df['Final_INR_KG_Price'] = (df['Clean_Raw_Price'] / 1000) * df[fx_col_name]
        df['Conversion_Formula'] = "(Clean_Raw_Price [USD/TON] / 1000) * USD_Exchange_Rate"
    elif name == "BRASS":
        df['Final_INR_KG_Price'] = (df['Clean_Raw_Price'] / 100) * df[fx_col_name]
        df['Conversion_Formula'] = "(Clean_Raw_Price [EUR/100kg] / 100) * EUR_Exchange_Rate"

    df['Product'] = name
    return reorder_columns(df).sort_values('Date', ascending=False)


# ==========================================
# 3. MAIN PIPELINE EXECUTION
# ==========================================

def run_data_pipeline():
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
            elif name in ["EVA", "EVA_PVC", "PVC_Suspension", "ACETONE", "PVC"]:
                df_clean = process_domestic_inr(filepath, name)
            elif name in ["MEG", "PTA", "Naphtha", "DOP", "PP"]:
                df_clean = process_usd_mt_products(filepath, name, usd_inr_df)

            if not df_clean.empty:
                product_dfs.append(df_clean)
                print(f" -> Success: {name} ({len(df_clean)} records).")
        except Exception as e:
            print(f" -> Failed to process {name}. Error: {e}")

    # =========================================================
    # PHASE 3: FINALIZING MASTER DATASET WITH ADVANCED EXCEL
    # =========================================================
    print("\n--- Phase 3: Finalizing Master Dataset (Advanced Excel Formatting) ---")
    master_file_path = os.path.join(ANALYSIS_DIR, "Master_Trend_Data.xlsx")

    try:
        with pd.ExcelWriter(master_file_path, engine='xlsxwriter', datetime_format='dd-mmm-yyyy') as writer:
            workbook = writer.book

            # --- Corporate Formatting Styles ---
            header_format = workbook.add_format({
                'bold': True, 'valign': 'vcenter', 'align': 'center',
                'bg_color': '#003366', 'font_color': '#FFFFFF', 'border': 1
            })
            date_format = workbook.add_format({'num_format': 'dd-mmm-yyyy', 'border': 1, 'align': 'center'})
            price_format = workbook.add_format({'num_format': '₹#,##0.00', 'border': 1})
            standard_format = workbook.add_format({'border': 1, 'align': 'left'})

            # Navigation Formats
            link_fmt = workbook.add_format(
                {'font_color': '#0052cc', 'underline': True, 'bold': True, 'valign': 'vcenter', 'font_size': 11})
            back_btn_fmt = workbook.add_format(
                {'font_color': '#ffffff', 'bg_color': '#e52d27', 'bold': True, 'align': 'center', 'valign': 'vcenter',
                 'border': 1})
            nav_title_fmt = workbook.add_format({'bold': True, 'font_size': 22, 'font_color': '#003366'})
            nav_sub_fmt = workbook.add_format({'font_size': 12, 'font_color': '#4a4a4a', 'italic': True})
            nav_section_fmt = workbook.add_format(
                {'bold': True, 'font_size': 14, 'font_color': '#1a1a1a', 'bottom': 2, 'bottom_color': '#003366'})

            # --- CREATE NAVIGATION TAB FIRST ---
            nav_ws = workbook.add_worksheet('Navigation')
            nav_ws.set_zoom(80)  # 80% Zoom applied
            nav_ws.hide_gridlines(2)
            nav_ws.set_column('A:A', 3)
            nav_ws.set_column('B:C', 35)

            nav_ws.write('B2', 'Raw Material Procurement Dashboard', nav_title_fmt)
            nav_ws.write('B3', 'Click any metric below to view detailed charts and historical data.', nav_sub_fmt)
            nav_ws.write(5, 1, 'Commodity Price Trends', nav_section_fmt)
            nav_ws.write(5, 2, 'Macro Exchange Rates', nav_section_fmt)

            month_col = 29  # Column AD
            avg_col = 30  # Column AE

            prod_nav_row = 7
            fx_nav_row = 7

            # ---------------------------------------------------------
            # 1. PROCESS COMMODITIES
            # ---------------------------------------------------------
            if product_dfs:
                master_product_df = pd.concat(product_dfs, ignore_index=True)
                master_product_df['Date'] = pd.to_datetime(master_product_df['Date']).dt.normalize()
                master_product_df = master_product_df.sort_values(by=['Product', 'Date'], ascending=[True, False])

                # Write & Hide Raw Data Tab for Pipeline
                master_product_df.to_excel(writer, sheet_name='Product_Trends', index=False)
                writer.sheets['Product_Trends'].set_zoom(80)
                writer.sheets['Product_Trends'].hide()

                for product in master_product_df['Product'].unique():
                    sheet_name = str(product).replace('/', '_').replace('[', '').replace(']', '')[:31]
                    prod_df = master_product_df[master_product_df['Product'] == product].copy()

                    prod_df = prod_df.sort_values(by='Date', ascending=False)
                    prod_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]

                    ws.set_zoom(80)  # 80% Zoom applied
                    ws.hide_gridlines(2)

                    # Add Hyperlink back to Navigation
                    chart_col = len(prod_df.columns) + 4
                    ws.write_url(0, chart_col, "internal:'Navigation'!A1", string="⬅ Return to Dashboard",
                                 cell_format=back_btn_fmt)
                    ws.set_column(chart_col, chart_col, 25)  # Ensure button width

                    # Add Hyperlink on Navigation Page
                    nav_ws.write_url(prod_nav_row, 1, f"internal:'{sheet_name}'!A1", string=f"📊 {product} Trend",
                                     cell_format=link_fmt)
                    prod_nav_row += 1

                    for col_num, col_name in enumerate(prod_df.columns):
                        ws.write(0, col_num, col_name, header_format)
                        if col_name == 'Date':
                            ws.set_column(col_num, col_num, 15, date_format)
                        elif 'Price' in col_name or 'Rate' in col_name:
                            ws.set_column(col_num, col_num, 18, price_format)
                        else:
                            ws.set_column(col_num, col_num, 20, standard_format)

                    ws.set_column('D:AC', None, None, {'hidden': True})
                    ws.freeze_panes(1, 0)

                    one_year_ago = pd.to_datetime('today') - pd.DateOffset(years=1)
                    monthly_df = prod_df[prod_df['Date'] >= one_year_ago].copy()
                    max_monthly_row = 0

                    if not monthly_df.empty:
                        monthly_df = monthly_df.set_index('Date').resample('MS').mean(numeric_only=True).reset_index()
                        monthly_df = monthly_df.dropna(subset=['Final_INR_KG_Price'])
                        monthly_df = monthly_df.sort_values(by='Date', ascending=False)

                        ws.set_column(month_col, month_col, 15, date_format)
                        ws.set_column(avg_col, avg_col, 22, price_format)
                        ws.write(0, month_col, "Month", header_format)
                        ws.write(0, avg_col, "Monthly Avg (INR/KG)", header_format)

                        for r_idx, row in enumerate(monthly_df.itertuples(), start=1):
                            ws.write(r_idx, month_col, row.Date, date_format)
                            ws.write(r_idx, avg_col, row.Final_INR_KG_Price, price_format)

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
                    chart1.set_y_axis({'name': 'Price (INR/KG)'})
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
                            'data_labels': {'value': True, 'num_format': '₹#,##0', 'position': 'above'}
                        })
                        chart2.set_title({'name': f'{product} 1-Year Monthly Average'})
                        chart2.set_x_axis({'name': 'Date', 'date_axis': True})
                        chart2.set_y_axis({'name': 'Price (INR/KG)'})
                        chart2.set_legend({'none': True})
                        ws.insert_chart(18, chart_col, chart2, {'x_scale': 1.4, 'y_scale': 1.1, 'object_position': 3})

            # ---------------------------------------------------------
            # 2. PROCESS EXCHANGE RATES
            # ---------------------------------------------------------
            if fx_dfs:
                master_fx_df = pd.concat(fx_dfs, ignore_index=True)
                master_fx_df['Date'] = pd.to_datetime(master_fx_df['Date']).dt.normalize()
                master_fx_df = master_fx_df.sort_values(by=['Product', 'Date'], ascending=[True, False])

                # Write & Hide Raw Data Tab for Pipeline
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

                    ws.set_zoom(80)  # 80% Zoom applied
                    ws.hide_gridlines(2)

                    # Add Hyperlink back to Navigation
                    chart_col = len(fx_df.columns) + 4
                    ws.write_url(0, chart_col, "internal:'Navigation'!A1", string="⬅ Return to Dashboard",
                                 cell_format=back_btn_fmt)
                    ws.set_column(chart_col, chart_col, 25)  # Ensure button width

                    # Add Hyperlink on Navigation Page
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

                    ws.set_column('D:AC', None, None, {'hidden': True})
                    ws.freeze_panes(1, 0)

                    one_year_ago = pd.to_datetime('today') - pd.DateOffset(years=1)
                    monthly_fx = fx_df[fx_df['Date'] >= one_year_ago].copy()
                    max_fx_monthly_row = 0

                    if not monthly_fx.empty:
                        monthly_fx = monthly_fx.set_index('Date').resample('MS').mean(numeric_only=True).reset_index()
                        monthly_fx = monthly_fx.dropna(subset=['Price'])
                        monthly_fx = monthly_fx.sort_values(by='Date', ascending=False)

                        ws.set_column(month_col, month_col, 15, date_format)
                        ws.set_column(avg_col, avg_col, 18, fx_num_format)
                        ws.write(0, month_col, "Month", header_format)
                        ws.write(0, avg_col, "Monthly Avg", header_format)

                        for r_idx, row in enumerate(monthly_fx.itertuples(), start=1):
                            ws.write(r_idx, month_col, row.Date, date_format)
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
                        chart2.set_legend({'none': True})
                        ws.insert_chart(18, chart_col, chart2, {'x_scale': 1.4, 'y_scale': 1.1, 'object_position': 3})

        print(f"\nSUCCESS! Formatted Master Data with Charts saved: {master_file_path}")

    except PermissionError:
        print("\nPERMISSION ERROR: Please close 'Master_Trend_Data.xlsx' in Excel and run this script again.")


if __name__ == "__main__":
    run_data_pipeline()