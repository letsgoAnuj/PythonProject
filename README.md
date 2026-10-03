# Footwear Raw Material Procurement Pipeline

An automated, end-to-end data pipeline built in Python to track, normalize, and forecast raw material costs (Polymers, Petrochemicals, Metals, and Rubber) and macro exchange rates (USD, EUR, CNY).

## Features
* **Hybrid Web Scraping:** Utilizes Playwright and Selenium to bypass complex ASP.NET forms and extract historical pricing data.
* **Financial Normalization:** Automatically converts global metric tons, barrels, and foreign currencies into a standardized `INR/KG` metric.
* **Native PDF Generation:** Builds a multi-page executive PDF report with native `matplotlib` charting, 1-year monthly trend lines, and sorted daily data tables.
* **Automated Distribution:** Hooks into local Windows COM to silently draft and attach deliverables in Microsoft Outlook.

## Project Structure
* `automator.py` - Master orchestrator and email generation.
* `scraper.py` - Ingestion layer (Selenium + Playwright).
* `data_processor.py` - ETL engine and Excel formatting.
* `report_generator.py` - Native PDF and PPTX compilation.
* `dashboard.py` - Interactive Streamlit web application.

## Setup Instructions
1. Clone the repository.
2. Install dependencies: `pip install -r requirements.txt`
3. Install Playwright browsers: `playwright install chromium`
4. Create a `.env` file in the root directory with your credentials:
   ```text
   POLYMER_USERNAME=your_email@domain.com
   POLYMER_PASSWORD=your_password
   