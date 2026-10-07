import pandas as pd
import matplotlib.pyplot as plt
import os
import warnings
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
import matplotlib.dates as mdates
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.patches as patches
from datetime import datetime
import sys
sys.stdout.reconfigure(encoding='utf-8')

warnings.filterwarnings('ignore')

# Set up paths
BASE_DIR = os.getcwd()
ANALYSIS_DIR = os.path.join(BASE_DIR, "analysis_output")
CHART_DIR = os.path.join(ANALYSIS_DIR, "charts")
DATA_FILE = os.path.join(ANALYSIS_DIR, "Master_Trend_Data.xlsx")
PPT_OUTPUT = os.path.join(ANALYSIS_DIR, "Footwear_Procurement_Forecast.pptx")
PDF_OUTPUT = os.path.join(ANALYSIS_DIR, "Footwear_Procurement_Forecast.pdf")

if not os.path.exists(CHART_DIR):
    os.makedirs(CHART_DIR)

# ==========================================
# 1. EXPANDED SLIDE CONFIGURATION
# ==========================================
PAGES_CONFIG = [
    {
        "title": "Macro Impact: Exchange Rate Tracker",
        "insight": "FORECAST: Tracking USD and EUR fluctuations to gauge import cost inflation and macro-level procurement impacts on imported components.",
        "charts": [
            {"title": "USD Trend vs INR", "products": ["USD-INR"], "sheet": "Exchange_Rates", "price_col": "Price",
             "unit": "INR"},
            {"title": "EUR Trend vs INR", "products": ["EUR-INR"], "sheet": "Exchange_Rates", "price_col": "Price",
             "unit": "INR"}
        ]
    },
    {
        "title": "Regional Impact: RMB Exchange Tracker",
        "insight": "FORECAST: Tracking the Chinese Yuan (RMB) to INR trajectory. Direct impacts on landed costs for imported Chinese components, including synthetics, hardware, and specialized textiles.",
        "charts": [
            {"title": "CNY (RMB) Trend vs INR", "products": ["CNY-INR"], "sheet": "Exchange_Rates",
             "price_col": "Price", "unit": "INR"}
        ]
    },
    {
        "title": "Upstream Drivers: Items Impacting Yarn",
        "insight": "FORECAST: Brent Crude acts as the ultimate upstream driver. Track PTA, MEG, and PP for cascading price adjustments in yarn and woven labels.",
        "charts": [
            {"title": "PTA & MEG Trajectory", "products": ["PTA", "MEG"], "sheet": "Product_Trends",
             "price_col": "Final_INR_KG_Price", "unit": "INR/KG"},
            {"title": "PP & Brent Crude Trajectory", "products": ["PP", "Brent"], "sheet": "Product_Trends",
             "price_col": "Final_INR_KG_Price", "unit": "INR/KG"}
        ]
    },
    {
        "title": "Coated Fabrics: Items Impacting PVC",
        "insight": "FORECAST: DOP (plasticizer) remains highly volatile. While PVC Suspension offers a stable baseline for fabrics, DOP drives the cost of flexibility.",
        "charts": [
            {"title": "DOP & PVC Suspension Trend", "products": ["DOP", "PVC_Suspension"], "sheet": "Product_Trends",
             "price_col": "Final_INR_KG_Price", "unit": "INR/KG"}
        ]
    },
    {
        "title": "Raw Precursors: Naphtha & Acetone",
        "insight": "FORECAST: Naphtha and Acetone act as primary precursors. Spikes in these upstream commodities will cascade into PU foam and EVA sheet cost increases.",
        "charts": [
            {"title": "Naphtha Trend", "products": ["Naphtha"], "sheet": "Product_Trends",
             "price_col": "Final_INR_KG_Price", "unit": "INR/KG"},
            {"title": "Acetone Trend", "products": ["ACETONE"], "sheet": "Product_Trends",
             "price_col": "Final_INR_KG_Price", "unit": "INR/KG"}
        ]
    },
    {
        "title": "Sole Materials: EVA & Latex",
        "insight": "ANALYSIS: Divergence between natural Latex and synthetic EVA allows for strategic recipe blending.",
        "charts": [
            {"title": "EVA Trend", "products": ["EVA"], "sheet": "Product_Trends", "price_col": "Final_INR_KG_Price",
             "unit": "INR/KG"},
            {"title": "Latex Trend", "products": ["LATEX"], "sheet": "Product_Trends",
             "price_col": "Final_INR_KG_Price", "unit": "INR/KG"}
        ]
    },
    {
        "title": "Hardware Trims: Brass & Zinc",
        "insight": "ANALYSIS: Zinc (galvanizing) and Brass dictate hardware/accessory costs. Watch global LME trends to prevent domestic suppliers from inflating pricing.",
        "charts": [
            {"title": "Brass Trend", "products": ["BRASS"], "sheet": "Product_Trends",
             "price_col": "Final_INR_KG_Price", "unit": "INR/KG"},
            {"title": "Zinc Trend", "products": ["ZINC"], "sheet": "Product_Trends", "price_col": "Final_INR_KG_Price",
             "unit": "INR/KG"}
        ]
    }
]

RESOURCE_LINKS = [
    ("PolymerUpdate", "PVC, EVA, Naphtha, DOP, Acetone, PET, PP, Brent Crude"),
    ("RBI (Reserve Bank of India)", "USD-INR, EUR-INR Exchange Rates"),
    ("ExchangeRates.org.uk", "CNY-INR (Chinese Yuan) Exchange Rates"),
    ("Westmetall", "Zinc (LME), Brass (Domestic/Imported)"),
    ("PCK Ltd", "Natural Rubber / Latex")
]

COLORS = ['#0052cc', '#e52d27', '#00b8d9', '#6554c0', '#ff8b00']


# ==========================================
# 2. POWERPOINT GENERATOR ENGINE
# ==========================================
def create_slide_image(page_config, df_products, df_fx):
    num_charts = len(page_config["charts"])
    fig, axes = plt.subplots(1, num_charts, figsize=(12, 4.2))
    fig.patch.set_facecolor('white')

    if num_charts == 1:
        axes = [axes]

    one_year_ago = pd.to_datetime('today') - pd.DateOffset(years=1)

    for ax, chart_config in zip(axes, page_config["charts"]):
        ax.set_facecolor('white')
        y_max, y_min = 0, float('inf')
        target_df = df_fx if chart_config["sheet"] == "Exchange_Rates" else df_products

        if not target_df.empty and 'Date' in target_df.columns:
            df_1yr = target_df[target_df['Date'] >= one_year_ago]
            for idx, prod in enumerate(chart_config["products"]):
                mat_data = df_1yr[df_1yr['Product'] == prod].sort_values('Date')
                if not mat_data.empty:
                    price_col = chart_config["price_col"]
                    monthly_data = mat_data.resample('MS', on='Date').mean(numeric_only=True).reset_index().dropna(
                        subset=[price_col])

                    if monthly_data.empty: continue
                    if monthly_data[price_col].max() > y_max: y_max = monthly_data[price_col].max()
                    if monthly_data[price_col].min() < y_min: y_min = monthly_data[price_col].min()

                    color = COLORS[idx % len(COLORS)]
                    ax.plot(monthly_data['Date'], monthly_data[price_col], label=prod, linewidth=3.0, marker='o',
                            markersize=7, markeredgecolor='white', markeredgewidth=1.5, color=color, zorder=3)

                    for x, y in zip(monthly_data['Date'], monthly_data[price_col]):
                        offset = (y_max - y_min) * 0.06 if y_max != y_min else y_max * 0.06
                        ax.text(x, y + offset, f"{y:,.0f}", fontsize=9.5, ha='center', va='bottom', color=color,
                                fontweight='bold', zorder=4)

        ax.set_title(chart_config["title"], fontsize=12, fontweight='bold', color='#1a1a1a', pad=15)
        ax.set_ylabel(f"{chart_config['unit']}", fontsize=10, color='#4a4a4a', labelpad=10)

        if y_min == float('inf'):
            ax.set_ylim(bottom=0, top=100)
            ax.text(0.5, 0.5, "Data Unavailable", ha='center', va='center', transform=ax.transAxes, color='#e52d27',
                    fontsize=12, fontweight='bold')
        else:
            ax.set_ylim(bottom=y_min * 0.90, top=y_max * 1.18)

        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#e0e0e0')
        ax.spines['bottom'].set_color('#e0e0e0')
        ax.grid(axis='y', linestyle='-', color='#f0f0f0')
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %y'))
        ax.tick_params(axis='x', rotation=45, labelsize=9.5)
        ax.legend(loc="upper left", fontsize=9, frameon=False)

    plt.tight_layout()
    filename = page_config['title'].replace(" ", "_").replace(",", "").replace("&", "and")[:25] + ".png"
    chart_path = os.path.join(CHART_DIR, filename)
    plt.savefig(chart_path, dpi=400, bbox_inches='tight', facecolor='white')
    plt.close()
    return chart_path


def build_presentation():
    print("Generating PPTX presentation...")
    try:
        df_products = pd.read_excel(DATA_FILE, sheet_name="Product_Trends")
        df_fx = pd.read_excel(DATA_FILE, sheet_name="Exchange_Rates")
        df_products['Date'], df_fx['Date'] = pd.to_datetime(df_products['Date']), pd.to_datetime(df_fx['Date'])
    except Exception as e:
        df_products, df_fx = pd.DataFrame(), pd.DataFrame()

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Footwear Raw Material Market Analytics"
    slide.placeholders[1].text = "1-Year Monthly Average Trend Analysis\nGenerated via Automated Procurement Pipeline"

    for page in PAGES_CONFIG:
        chart_path = create_slide_image(page, df_products, df_fx)
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = page["title"]
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(28)
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(10, 37, 64)
        slide.shapes.add_picture(chart_path, Inches(0.2), Inches(1.2), width=Inches(12.9))

        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(5.9), Inches(12.333), Inches(1.1))
        txBox.fill.solid()
        txBox.fill.fore_color.rgb = RGBColor(246, 249, 252)
        txBox.line.color.rgb = RGBColor(226, 232, 240)
        p = txBox.text_frame.add_paragraph()
        p.text = page["insight"]
        p.font.size, p.font.bold = Pt(13), True
        p.font.color.rgb = RGBColor(10, 37, 64)

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Data Source & Resource Mapping"
    tf = slide.placeholders[1].text_frame
    for source, components in RESOURCE_LINKS:
        p = tf.add_paragraph()
        p.text = f"{source}: "
        p.font.bold, p.font.size, p.font.color.rgb = True, Pt(18), RGBColor(10, 37, 64)
        p_run = p.add_run()
        p_run.text, p_run.font.bold, p_run.font.size, p_run.font.color.rgb = components, False, Pt(16), RGBColor(71, 85,
                                                                                                                 105)

    prs.save(PPT_OUTPUT)
    print(f"✓ PowerPoint Presentation saved: {PPT_OUTPUT}")


# ==========================================
# 3. NATIVE PDF GENERATOR ENGINE (WITH TOC)
# ==========================================
def generate_native_pdf():
    print("\nGenerating Executive PDF Report (Cover, TOC + Charts)...")
    try:
        df_products = pd.read_excel(DATA_FILE, sheet_name="Product_Trends")
        df_fx = pd.read_excel(DATA_FILE, sheet_name="Exchange_Rates")
        df_products['Date'] = pd.to_datetime(df_products['Date'])
        df_fx['Date'] = pd.to_datetime(df_fx['Date'])

        all_items = []
        if not df_fx.empty:
            for p in sorted(df_fx['Product'].unique()):
                all_items.append({'name': p, 'df': df_fx[df_fx['Product'] == p], 'val_col': 'Price', 'unit': 'INR',
                                  'cat': 'Macro Exchange Rates'})
        if not df_products.empty:
            for p in sorted(df_products['Product'].unique()):
                all_items.append(
                    {'name': p, 'df': df_products[df_products['Product'] == p], 'val_col': 'Final_INR_KG_Price',
                     'unit': 'INR/KG', 'cat': 'Raw Material Commodities'})

        with PdfPages(PDF_OUTPUT) as pdf:

            # ----------------------------------------------------
            # PAGE 1: COVER PAGE
            # ----------------------------------------------------
            fig_cover = plt.figure(figsize=(8.27, 11.69))
            fig_cover.patch.set_facecolor('white')

            fig_cover.patches.extend([
                patches.Rectangle((0, 0.85), 1, 0.15, transform=fig_cover.transFigure, clip_on=False,
                                  facecolor='#0052cc'),
                patches.Rectangle((0, 0.0), 1, 0.05, transform=fig_cover.transFigure, clip_on=False,
                                  facecolor='#0f172a')
            ])

            # Replaced letter_spacing with manual string spacing to avoid matplotlib kwargs errors
            fig_cover.text(0.5, 0.92, "C A M P U S   A C T I V E W E A R", ha='center', va='center', fontsize=26,
                           fontweight='bold', color='white')
            fig_cover.text(0.5, 0.65, "Procurement & Sourcing", ha='center', va='center', fontsize=34,
                           fontweight='bold', color='#0f172a')
            fig_cover.text(0.5, 0.58, "Raw Material Cost & Market Analytics", ha='center', va='center', fontsize=20,
                           color='#334155')
            fig_cover.text(0.5, 0.45, f"Report Generated: {datetime.now().strftime('%d %B %Y')}", ha='center',
                           va='center', fontsize=14, color='#64748b')
            fig_cover.text(0.5, 0.40, "Automated 1-Year Market Intelligence Pipeline", ha='center', va='center',
                           fontsize=12, color='#94a3b8', style='italic')

            pdf.savefig(fig_cover)
            plt.close(fig_cover)

            # ----------------------------------------------------
            # PAGE 2: TABLE OF CONTENTS
            # ----------------------------------------------------
            fig_toc = plt.figure(figsize=(8.27, 11.69))
            fig_toc.patch.set_facecolor('white')

            fig_toc.text(0.1, 0.9, "Table of Contents", fontsize=24, fontweight='bold', color='#0052cc')
            fig_toc.text(0.1, 0.87, "_______________________________________________________________________________",
                         color='#e2e8f0', fontsize=14)

            start_page = 3
            y_pos = 0.8
            col_x = 0.1
            current_category = ""

            for idx, item in enumerate(all_items):
                if idx == 20:
                    col_x = 0.55
                    y_pos = 0.8
                    current_category = ""

                # Add category headers
                if item['cat'] != current_category:
                    y_pos -= 0.02
                    fig_toc.text(col_x, y_pos, item['cat'], fontsize=14, fontweight='bold', color='#e52d27')
                    y_pos -= 0.04
                    current_category = item['cat']

                page_num = start_page + idx
                fig_toc.text(col_x, y_pos, f"{item['name']}", fontsize=11, fontweight='bold', color='#334155')
                fig_toc.text(col_x + 0.35, y_pos, f"......... Page {page_num}", ha='right', fontsize=11,
                             color='#64748b')
                y_pos -= 0.03

            fig_toc.text(0.95, 0.02, "Page 2", ha='right', fontsize=9, color='#64748b')
            pdf.savefig(fig_toc)
            plt.close(fig_toc)

            # ----------------------------------------------------
            # PAGES 3+: INDIVIDUAL COMMODITY DATA PAGES
            # ----------------------------------------------------
            for idx, item in enumerate(all_items):
                current_page = start_page + idx
                name = item['name']
                df = item['df'].sort_values('Date', ascending=False)
                val_col = item['val_col']
                unit = item['unit']

                if df.empty: continue

                fig = plt.figure(figsize=(8.27, 11.69))
                fig.clf()
                fig.patch.set_facecolor('white')
                fig.suptitle(f"{name} - Market Trend Analysis", fontsize=18, fontweight='bold', color='#0f172a', y=0.96)

                # TOP: 1-Year Monthly Chart
                ax1 = fig.add_axes([0.1, 0.58, 0.8, 0.32])
                one_year_ago = pd.to_datetime('today') - pd.DateOffset(years=1)
                df_1yr = df[df['Date'] >= one_year_ago]

                monthly = pd.DataFrame()
                if not df_1yr.empty:
                    monthly = df_1yr.set_index('Date').resample('MS').mean(numeric_only=True).reset_index().dropna(
                        subset=[val_col])
                    ax1.plot(monthly['Date'], monthly[val_col], marker='o', linestyle='-', color='#0052cc',
                             linewidth=2.5, markersize=6, label="Monthly Avg")

                    y_max = monthly[val_col].max()
                    y_min = monthly[val_col].min()
                    y_range = y_max - y_min if y_max != y_min else y_max * 0.1

                    for i_val, (x, y) in enumerate(zip(monthly['Date'], monthly[val_col])):
                        format_str = f"{y:,.2f}" if unit == "INR" else f"{y:,.0f}"
                        offset_y = y_range * 0.05 if i_val % 2 == 0 else y_range * -0.07
                        va_align = 'bottom' if i_val % 2 == 0 else 'top'

                        ax1.text(x, y + offset_y, format_str, fontsize=8, ha='center', va=va_align, color='#0f172a',
                                 fontweight='bold',
                                 bbox=dict(facecolor='white', edgecolor='none', alpha=0.7, pad=1))

                    ax1.set_title("1-Year Monthly Average Trend", fontsize=12, fontweight='bold', color='#334155',
                                  pad=10)
                    ax1.set_ylabel(f"Price ({unit})", fontsize=10)
                    ax1.xaxis.set_major_locator(mdates.MonthLocator())
                    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%b %y'))

                    plt.setp(ax1.get_xticklabels(), rotation=45, ha='right', fontsize=9)
                    plt.setp(ax1.get_yticklabels(), fontsize=9)
                    ax1.grid(axis='y', linestyle='--', alpha=0.6)
                    ax1.spines['top'].set_visible(False)
                    ax1.spines['right'].set_visible(False)
                    ax1.legend(loc='upper right', frameon=False, fontsize=9)

                # BOTTOM LEFT: Monthly Table
                ax2 = fig.add_axes([0.1, 0.05, 0.35, 0.45])
                ax2.axis('off')
                ax2.set_title("1-Year Monthly Averages", fontsize=11, fontweight='bold', color='#334155', pad=8)

                if not monthly.empty:
                    monthly_disp = monthly.sort_values('Date', ascending=False).copy()
                    monthly_disp['Month'] = monthly_disp['Date'].dt.strftime('%b %Y')
                    monthly_disp['Price_Str'] = monthly_disp[val_col].apply(lambda x: f"₹ {x:,.2f}")

                    table_data_m = [['Month', f'Avg Price ({unit})']]
                    for _, row in monthly_disp.iterrows():
                        table_data_m.append([row['Month'], row['Price_Str']])

                    table_m = ax2.table(cellText=table_data_m, loc='center', cellLoc='center')
                    table_m.auto_set_font_size(False)
                    table_m.set_fontsize(9)
                    table_m.scale(1, 1.4)

                    for j in range(len(table_data_m[0])):
                        cell = table_m[0, j]
                        cell.set_text_props(weight='bold', color='white', size=10)
                        cell.set_facecolor('#0052cc')

                    for r_idx in range(1, len(table_data_m)):
                        for j in range(len(table_data_m[0])):
                            if r_idx % 2 == 0: table_m[r_idx, j].set_facecolor('#f8fafc')

                # BOTTOM RIGHT: Daily Table
                ax3 = fig.add_axes([0.55, 0.05, 0.35, 0.45])
                ax3.axis('off')

                latest_date = df['Date'].max()
                current_month_df = df[
                    (df['Date'].dt.year == latest_date.year) & (df['Date'].dt.month == latest_date.month)]
                current_month_df = current_month_df.sort_values('Date', ascending=False)

                ax3.set_title(f"Daily Data ({latest_date.strftime('%b %Y')})", fontsize=11, fontweight='bold',
                              color='#334155', pad=8)

                display_df = current_month_df.head(28).copy()
                if not display_df.empty:
                    display_df['Date_Str'] = display_df['Date'].dt.strftime('%d-%b-%y')
                    display_df['Price_Str'] = display_df[val_col].apply(lambda x: f"₹ {x:,.2f}")

                    table_data_d = [['Date', f'Daily Price ({unit})']]
                    for _, row in display_df.iterrows():
                        table_data_d.append([row['Date_Str'], row['Price_Str']])

                    table_d = ax3.table(cellText=table_data_d, loc='center', cellLoc='center')
                    table_d.auto_set_font_size(False)

                    row_count = len(table_data_d)
                    font_size_d = 9 if row_count < 18 else 8
                    scale_y = 1.4 if row_count < 18 else 1.1
                    table_d.set_fontsize(font_size_d)
                    table_d.scale(1, scale_y)

                    for j in range(len(table_data_d[0])):
                        cell = table_d[0, j]
                        cell.set_text_props(weight='bold', color='white', size=9)
                        cell.set_facecolor('#0052cc')

                    for r_idx in range(1, len(table_data_d)):
                        for j in range(len(table_data_d[0])):
                            if r_idx % 2 == 0: table_d[r_idx, j].set_facecolor('#f8fafc')

                # Add Page Number Footer
                fig.text(0.95, 0.02, f"Page {current_page}", ha='right', fontsize=9, color='#64748b')

                pdf.savefig(fig)
                plt.close(fig)

        print(f"✓ Native PDF Report successfully generated: {PDF_OUTPUT}")
    except Exception as e:
        print(f"\nCould not generate native PDF. Error: {e}")


if __name__ == "__main__":
    build_presentation()
    generate_native_pdf()