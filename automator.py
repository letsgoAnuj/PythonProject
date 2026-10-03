import subprocess
import os
import sys
import logging
import win32com.client as win32
from datetime import datetime

# Paths
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
ANALYSIS_DIR = os.path.join(PROJECT_DIR, "analysis_output")

PDF_REPORT = os.path.join(ANALYSIS_DIR, "Footwear_Procurement_Forecast.pdf")
PPT_REPORT = os.path.join(ANALYSIS_DIR, "Footwear_Procurement_Forecast.pptx")
EXCEL_REPORT = os.path.join(ANALYSIS_DIR, "Master_Trend_Data.xlsx")
LOG_FILE = os.path.join(PROJECT_DIR, "pipeline.log")

# Email Config
EMAIL_RECEIVERS = ["manager1@campusshoes.com", "team@campusshoes.com"]
# Dashboard hosted on your local machine
DASHBOARD_URL = "http://192.168.3.217:8501"

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def save_outlook_draft(subject, html_body, attachments=None):
    """Creates an email draft in local Outlook with HTML links and multiple attachments."""
    try:
        logging.info("Connecting to local Outlook...")
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.To = "; ".join(EMAIL_RECEIVERS)
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
        <pre style="background:#f4f4f4; padding:10px; border:1px solid #ddd;">{result.stderr}</pre>
        """
        save_outlook_draft(
            subject=f"🚨 PIPELINE ALERT: Failed at {script_name}",
            html_body=error_html
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

    # 2. Build HTML Body
    current_date = datetime.now().strftime("%d-%b-%Y")
    email_html = f"""
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <h3 style="color: #0052cc;">Footwear Raw Material Market Analytics ({current_date})</h3>
        <p>Hello Team,</p>
        <p>The automated weekly raw material analysis has successfully completed. All upstream commodity drivers (Petrochemicals, LME Metals, Rubber) and macro exchange rates (USD, EUR, CNY) have been extracted, mapped to our component categories, and normalized to <b>INR/KG</b>.</p>

        <div style="margin: 20px 0; padding: 15px; background-color: #f8fafc; border-left: 5px solid #0052cc; border-radius: 4px;">
            <b style="color: #0052cc;">Live Interactive Dashboard:</b><br>
            You can access the live market tracking dashboard via our internal network here:<br>
            <a href="{DASHBOARD_URL}" style="display: inline-block; margin-top: 10px; padding: 10px 18px; background-color: #0052cc; color: #ffffff; text-decoration: none; border-radius: 4px; font-weight: bold;">
                Open Web Dashboard ↗
            </a>
            <br><br>
            <span style="font-size: 12px; color: #64748b;">(Note: You must be connected to the office network/VPN to access this link.)</span>
        </div>

        <p><b>Attached Deliverables:</b></p>
        <ul>
            <li><b>Footwear_Procurement_Forecast.pdf</b>: Executive summary & 1-year monthly trend slides.</li>
            <li><b>Footwear_Procurement_Forecast.pptx</b>: Editable slide deck for management presentations.</li>
            <li><b>Master_Trend_Data.xlsx</b>: Clean dataset featuring the interactive UI navigation menu, conversion formulas, and full currency coverage.</li>
        </ul>

        <p style="font-size: 12px; color: #64748b; margin-top: 30px;"><i>Generated automatically via Python Procurement Pipeline.</i></p>
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
        attachments=target_attachments
    )

    logging.info("=== PIPELINE EXECUTION COMPLETE ===")
    print("\n🎉 Pipeline complete. Open Outlook to inspect your Drafts folder.")