import subprocess
import os
import sys
import logging
import pandas as pd
import win32com.client as win32
from datetime import datetime

# Paths
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
ANALYSIS_DIR = os.path.join(PROJECT_DIR, "analysis_output")

PDF_REPORT = os.path.join(ANALYSIS_DIR, "Footwear_Procurement_Forecast.pdf")
PPT_REPORT = os.path.join(ANALYSIS_DIR, "Footwear_Procurement_Forecast.pptx")
EXCEL_REPORT = os.path.join(ANALYSIS_DIR, "Master_Trend_Data.xlsx")
LOG_FILE = os.path.join(PROJECT_DIR, "pipeline.log")

# ==========================================
# EMAIL DISTRIBUTION LISTS
# ==========================================
EMAIL_TO = [
    "vishal.sharma1@campusshoes.com", "rmupper.purchase@campusshoes.com", "isha.khandelwal@campusshoes.com",
    "packaging.purchase@campusshoes.com", "mis@campusshoes.com"
]
EMAIL_CC = [
    "karattla.apparao@campusshoes.com"  # Add additional management CCs here
]

# Admin Isolation (For Errors Only)
ADMIN_EMAIL = [
    "anuj.gupta1@campusshoes.com"
]

# Dashboard hosted on your local machine
DASHBOARD_URL = "http://192.168.3.217:8501"

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


def get_volatility_indicator(val):
    """Helper function to generate the block-based visual gauge for HTML emails."""
    if pd.isna(val) or str(val).strip() == "-":
        return '<td style="border: 1px solid #dddddd; padding: 6px 12px; text-align: left; color: #333333;">➖</td>'

    try:
        pct_val = float(val) * 100
        abs_pct = abs(pct_val)

        # Determine the number of visual blocks
        if abs_pct < 1.0:
            return f'<td style="border: 1px solid #dddddd; padding: 6px 12px; text-align: left; color: #333333;">➖ {pct_val:+.1f}%</td>'
        elif abs_pct < 10.0:
            blocks = '🟥' if pct_val > 0 else '🟩'
        elif abs_pct < 30.0:
            blocks = '🟥🟥' if pct_val > 0 else '🟩🟩'
        else:
            blocks = '🟥🟥🟥' if pct_val > 0 else '🟩🟩🟩'

        sign = '+' if pct_val > 0 else ''
        return f'<td style="border: 1px solid #dddddd; padding: 6px 12px; text-align: left; font-weight: bold; color: #333333;">{blocks} {sign}{pct_val:.1f}%</td>'

    except (ValueError, TypeError):
        return '<td style="border: 1px solid #dddddd; padding: 6px 12px; text-align: left; color: #333333;">➖</td>'


def generate_email_matrix_html(excel_path):
    """Reads the Variation Matrix from Excel and builds an Outlook-friendly HTML table."""
    try:
        raw_df = pd.read_excel(excel_path, sheet_name='Variation_Matrix')

        # Filter down strictly to the requested executive horizons
        # Added '7 Days' (which represents 1 week) and '2 Months'
        target_columns = ['Commodity', 'Latest Price', '7 Days', '1 Week', '1 Month', '2 Months', '3 Months',
                          '6 Months', '1 Year']
        existing_columns = [col for col in target_columns if col in raw_df.columns]
        df = raw_df[existing_columns]

        # Rename '7 Days' to '1 Week' for a cleaner email header if applicable
        df = df.rename(columns={'7 Days': '1 Week'})

    except Exception as e:
        return f"<p><i>Variation Matrix could not be loaded: {e}</i></p>"

    # Build HTML table
    html = [
        '<div style="margin-bottom: 15px;">',
        '<h3 style="color: #003366; font-family: Arial, sans-serif; margin-bottom: 8px;">Historical Price Variation Matrix</h3>',
        '<p style="font-family: Arial, sans-serif; font-size: 12px; color: #475569; margin-top: 0; margin-bottom: 10px;">',
        '<b>Volatility Gauge Scale:</b> ➖ Neutral (0-1%) &nbsp;|&nbsp; ',
        '🟥 / 🟩 Moderate (1-10%) &nbsp;|&nbsp; ',
        '🟥🟥 / 🟩🟩 High (10-30%) &nbsp;|&nbsp; ',
        '🟥🟥🟥 / 🟩🟩🟩 Extreme (>30%)',
        '</p></div>',
        '<table style="border-collapse: collapse; width: auto; font-family: Arial, sans-serif; font-size: 12px; border: 1px solid #dddddd; white-space: nowrap;" cellpadding="6" cellspacing="0">',
        '<thead><tr style="background-color: #0052cc; color: #ffffff; font-weight: bold; text-align: left;">'
    ]

    # Table headers
    for col in df.columns:
        html.append(f'<th style="border: 1px solid #dddddd; padding: 8px 12px;">{col}</th>')

    html.append('</tr></thead><tbody>')

    # Table rows
    for idx, row in df.iterrows():
        bg_color = "#f8fafc" if idx % 2 == 1 else "#ffffff"
        html.append(f'<tr style="background-color: {bg_color};">')

        for col_name in df.columns:
            val = row[col_name]

            # Commodity column
            if col_name == 'Commodity':
                html.append(
                    f'<td style="border: 1px solid #dddddd; padding: 6px 12px; text-align: left; font-weight: bold; color: #333333;">{val}</td>'
                )

            # Latest Price column - FIXED ALIGNMENT
            elif col_name == 'Latest Price':
                try:
                    if pd.isna(val) or str(val).strip() == '-':
                        price_str = "-"
                    elif row['Commodity'] == 'Brent':
                        price_str = f"${float(val):,.2f}"
                    else:
                        price_str = f"&#8377;{float(val):,.2f}"
                except (ValueError, TypeError):
                    price_str = str(val) if pd.notna(val) else "-"

                html.append(
                    f'<td style="border: 1px solid #dddddd; padding: 6px 12px; color: #333333;">{price_str}</td>'
                )

            # Historical percentage variation columns (Block Gauge)
            else:
                html.append(get_volatility_indicator(val))

        html.append('</tr>')

    html.append('</tbody></table><br>')

    return "".join(html)


def save_outlook_draft(subject, html_body, to_emails, cc_emails, attachments=None):
    """Creates an email draft in local Outlook with HTML links, TO/CC fields, and multiple attachments."""
    try:
        logging.info("Connecting to local Outlook...")
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        # Apply TO and CC logic
        mail.To = "; ".join(to_emails)
        if cc_emails:
            mail.CC = "; ".join(cc_emails)

        mail.Subject = subject
        mail.HTMLBody = html_body

        if attachments:
            for file_path in attachments:
                abs_path = os.path.abspath(file_path)
                if os.path.exists(abs_path):
                    mail.Attachments.Add(Source=abs_path)
                    logging.info(f"Attached file: {abs_path}")
                    print(f"✓ Attached: {os.path.basename(abs_path)}")
                else:
                    logging.warning(f"Attachment not found, skipping: {abs_path}")
                    print(f"⚠️ Warning: File not found to attach: {abs_path}")

        mail.Save()
        logging.info(f"Draft saved successfully in Outlook: {subject}")
        print(f"\n✅ Draft successfully saved in Outlook Drafts: '{subject}'")

    except Exception as e:
        logging.error(f"Failed to create Outlook draft: {e}")
        print(f"❌ Failed to create Outlook draft: {e}")


def run_script(script_name):
    """Executes child scripts sequentially and logs stdout/stderr."""
    logging.info(f"Starting {script_name}...")
    print(f"Starting {script_name}...")
    script_path = os.path.join(PROJECT_DIR, script_name)

    result = subprocess.run(
        [sys.executable, script_path],
        capture_output=True,
        text=True,
        encoding='utf-8'
    )

    if result.returncode != 0:
        error_msg = f"CRITICAL FAILURE in {script_name}.\n\n{result.stderr}"
        logging.error(error_msg)

        error_html = f"""
        <p style="color:red; font-weight:bold;">🚨 PIPELINE FAILURE ALERT</p>
        <p>The automated procurement run failed during <b>{script_name}</b>.</p>
        <p><b>Note: This alert is sent ONLY to the Admin. The team has NOT been notified.</b></p>
        <pre style="background:#f4f4f4; padding:10px; border:1px solid #ddd;">{result.stderr}</pre>
        """
        save_outlook_draft(
            subject=f"🚨 PIPELINE ALERT: Failed at {script_name}",
            html_body=error_html,
            to_emails=ADMIN_EMAIL,
            cc_emails=[]
        )
        sys.exit(1)
    else:
        logging.info(f"Successfully completed {script_name}.")
        print(f"✓ Completed {script_name}")
        if result.stdout:
            print(result.stdout.strip())


if __name__ == "__main__":
    logging.info("=== INITIATING AUTOMATED PIPELINE ===")
    print("🚀 Initiating Procurement Data Pipeline...\n")

    # 1. Run Pipeline
    run_script("scraper.py")
    run_script("data_processor.py")
    run_script("report_generator.py")

    # =================================================================
    # GRACEFUL DEGRADATION CHECK (Admin-Only Missing Products Alert)
    # =================================================================
    missing_log_path = os.path.join(PROJECT_DIR, "missing_products.log")

    if os.path.exists(missing_log_path):
        with open(missing_log_path, "r") as f:
            missing_items = f.read()

        admin_alert_html = f"""
        <h3 style="color:#e52d27;">⚠️ Partial Pipeline Success</h3>
        <p>The pipeline completed, but the scraper failed to download data for:</p>
        <p><b>{missing_items}</b></p>
        <p>The team report was still generated successfully with the available data and saved to your drafts. Please check the scrapers for these specific commodities.</p>
        """

        print(f"\n⚠️ Alerting Admin about missing products: {missing_items}")
        save_outlook_draft(
            subject="⚠️ Pipeline Warning: Missing Commodity Data",
            html_body=admin_alert_html,
            to_emails=ADMIN_EMAIL,  # Sent exclusively to you
            cc_emails=[]
        )

    # 2. Build HTML Body
    current_date = datetime.now().strftime("%d-%b-%Y")

    matrix_html = generate_email_matrix_html(EXCEL_REPORT)

    email_html = f"""
            <html>
            <body style="font-family: 'Segoe UI', Arial, sans-serif; font-size: 14px; line-height: 1.6; color: #333333; margin: 0; padding: 10px;">

                <h2 style="color: #0f172a; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; margin-bottom: 20px; font-size: 18px;">
                    Footwear Raw Material Market Analytics <span style="color: #64748b; font-size: 16px;">({current_date})</span>
                </h2>

                <p>Hi Team,</p>

                <p>Please find the latest weekly update on raw material pricing and macro exchange trends. We've pulled the latest data for our key upstream commodities (petrochemicals, LME metals, and rubber) and normalized everything to <strong>INR/KG</strong> except Brent(USD/BBL) for a better comparison.</p>

                <p>Please find below a table summarizing current price trends, highlighting potential cost savings and possible price increases for your reference.:</p>

                <!-- Embedded Matrix Table -->
                {matrix_html}

                <p style="margin-top: 25px;"><strong style="color: #0f172a;">Attached Deliverables:</strong></p>
                <ul style="margin-top: 5px; padding-left: 20px;">
                    <li style="margin-bottom: 8px;"><strong>Footwear_Procurement_Forecast.pdf</strong></li>
                    <li style="margin-bottom: 8px;"><strong>Footwear_Procurement_Forecast.pptx</strong></li>
                    <li style="margin-bottom: 8px;"><strong>Master_Trend_Data.xlsx</strong></li>
                </ul>

                <p>Please let me know if any clarification is required regarding the data or sources.</li>
                <p>Thank you..</p>

                <p style="margin-top: 30px;">Best regards,<br>
                <strong>Anuj Kumar Gupta</strong><br>
                Purchase D-1 | Campus Activewear Ltd.</p>

            </body>
            </html>
            """

    # 3. Compile File List to Attach
    target_attachments = [
        PDF_REPORT,
        PPT_REPORT,
        EXCEL_REPORT
    ]

    # 4. Save to Outlook Drafts
    print("\nSaving final draft to Outlook...")
    save_outlook_draft(
        subject=f"Weekly Market Intelligence & Raw Material Cost Trends - {current_date}",
        html_body=email_html,
        to_emails=EMAIL_TO,
        cc_emails=EMAIL_CC,
        attachments=target_attachments
    )

    logging.info("=== PIPELINE EXECUTION COMPLETE ===")
    print("\n🎉 Pipeline complete. Open Outlook to inspect your Drafts folder.")

    # # 5. Automatically Sync Data to Cloud Dashboard (GitHub)
    # print("\nSyncing updated data to Cloud Dashboard...")
    # try:
    #     subprocess.run(["git", "add", EXCEL_REPORT], check=True, capture_output=True)
    #     subprocess.run(["git", "commit", "-m", f"Automated data sync: {current_date}"], check=True, capture_output=True)
    #     subprocess.run(["git", "push"], check=True, capture_output=True)
    #     logging.info("Successfully pushed updated Master_Trend_Data.xlsx to GitHub.")
    #     print("✓ Cloud Dashboard successfully updated!")
    # except subprocess.CalledProcessError as e:
    #     logging.error(f"Failed to sync with GitHub: {e.stderr.decode('utf-8')}")
    #     print("⚠️ Warning: Failed to sync data to Cloud Dashboard. Check pipeline.log.")
    #
    # logging.info("=== PIPELINE EXECUTION COMPLETE ===")