import streamlit as st
import pandas as pd
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

# ==========================================
# PAGE SETUP & BRANDING
# ==========================================
st.set_page_config(
    page_title="Footwear Procurement Market Analytics",
    page_icon="👟",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
        .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
        .executive-box {
            background-color: #f8fafc;
            border-left: 5px solid #0052cc;
            padding: 14px 18px;
            border-radius: 4px;
            margin-bottom: 1.5rem;
            color: #1e293b;
        }
        .executive-box h4 { margin: 0 0 6px 0; color: #0052cc; font-size: 1rem; }
        .executive-box p { margin: 0; font-size: 0.95rem; }
    </style>
""", unsafe_allow_html=True)

st.title("👟 Footwear Raw Material Market Analytics")
st.caption("1-Year Monthly Procurement Forecast & Trend Dashboard | Master Pipeline")


# ==========================================
# DATA LOADING ENGINE (Reads Hidden Tabs)
# ==========================================
@st.cache_data(ttl=3600)
def load_data():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(BASE_DIR, "analysis_output", "Master_Trend_Data.xlsx")

    df = pd.read_excel(file_path, sheet_name="Product_Trends")
    fx = pd.read_excel(file_path, sheet_name="Exchange_Rates")
    df['Date'] = pd.to_datetime(df['Date'])
    fx['Date'] = pd.to_datetime(fx['Date'])
    return df, fx


try:
    df, fx = load_data()
except Exception as e:
    st.error(
        f"Could not load master dataset. Ensure `Master_Trend_Data.xlsx` is generated in `analysis_output/`. Error: {e}")
    st.stop()

# ==========================================
# SIDEBAR CONTROLS & SOURCES
# ==========================================
st.sidebar.header("Analytics Controls")
view_mode = st.sidebar.radio(
    "Data Granularity:",
    ["1-Year Monthly Averages (Matches PPT)", "Daily Historical Prices"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.subheader("Data Sources")
st.sidebar.markdown("""
- **PolymerUpdate**: PVC, EVA, Naphtha, DOP, Acetone, MEG, PTA, PP, Brent, LDPE, LLDPE
- **RBI**: Reference FX Rates (USD, EUR)
- **ExchangeRates.org.uk**: CNY/INR History
- **Westmetall**: LME Zinc & Brass
- **PCK Ltd**: Natural Rubber / Latex
""")

CHART_COLORS = ['#0052cc', '#e52d27', '#00b8d9', '#6554c0', '#ff8b00']


def prepare_data(data_df, value_col):
    data_df = data_df.copy().sort_values('Date')
    one_year_ago = pd.to_datetime('today') - pd.DateOffset(years=1)
    data_df = data_df[data_df['Date'] >= one_year_ago]

    if data_df.empty:
        return data_df

    if view_mode.startswith("1-Year Monthly"):
        monthly = data_df.set_index('Date').resample('MS').mean(numeric_only=True).reset_index()
        monthly = monthly.dropna(subset=[value_col])
        return monthly

    return data_df.dropna(subset=[value_col])


def build_trend_chart(items_config, default_sheet_df, unit_label="INR/KG"):
    fig = go.Figure()

    for idx, item in enumerate(items_config):
        prod = item["prod"]
        source_df = fx if item.get("is_fx", False) else default_sheet_df
        mat_data = source_df[source_df['Product'] == prod]
        val_col = "Price" if item.get("is_fx", False) else "Final_INR_KG_Price"

        plot_df = prepare_data(mat_data, val_col)

        if not plot_df.empty:
            color = CHART_COLORS[idx % len(CHART_COLORS)]
            text_labels = [f"{v:,.0f}" for v in plot_df[val_col]]

            fig.add_trace(go.Scatter(
                x=plot_df['Date'],
                y=plot_df[val_col],
                mode='lines+markers+text' if view_mode.startswith("1-Year Monthly") else 'lines',
                name=prod,
                text=text_labels if view_mode.startswith("1-Year Monthly") else None,
                textposition="top center",
                textfont=dict(size=11, color=color, family="Arial Black"),
                line=dict(width=3, color=color),
                marker=dict(size=7, color=color, line=dict(width=1.5, color='white')),
                hovertemplate=f"<b>{prod}</b><br>Date: %{{x|%b %Y}}<br>Price: %{{y:,.0f}} {unit_label}<extra></extra>"
            ))

    fig.update_layout(
        template="plotly_white",
        hovermode="x unified",
        margin=dict(l=20, r=20, t=30, b=20),
        height=380,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(
            title=f"Price ({unit_label})",
            showgrid=True,
            gridcolor="#f1f5f9",
            zeroline=False
        ),
        xaxis=dict(
            showgrid=False,
            dtick="M1" if view_mode.startswith("1-Year Monthly") else None,
            tickformat="%b %y"
        )
    )
    return fig


def display_kpis(products, default_df, is_fx=False):
    cols = st.columns(len(products))
    for i, prod in enumerate(products):
        source = fx if is_fx else default_df
        val_col = "Price" if is_fx else "Final_INR_KG_Price"
        mat_data = source[source['Product'] == prod].sort_values('Date')

        if not mat_data.empty:
            latest = mat_data.iloc[-1][val_col]
            prev = mat_data.iloc[-2][val_col] if len(mat_data) > 1 else latest
            pct_change = ((latest - prev) / prev) * 100 if prev != 0 else 0.0

            unit = "INR" if is_fx else ("USD/bbl" if prod == "Brent" else "₹/kg")
            prefix = "$" if prod == "Brent" else ""

            cols[i].metric(
                label=f"{prod} Current",
                value=f"{prefix}{latest:,.2f} {unit}" if is_fx or prod == "Brent" else f"{latest:,.0f} {unit}",
                delta=f"{pct_change:+.2f}% vs last period" if is_fx else f"{pct_change:+.1f}% vs last period"
            )

# ==========================================
# SLIDE-ALIGNED TAB NAVIGATION
# ==========================================
tabs = st.tabs([
    "💱 Exchange Rates (USD/EUR/RMB)",
    "🧵 Yarn Impact (PTA, MEG, PP, Brent)",
    "🧪 PVC Coated (DOP, Suspension)",
    "📦 Packaging Films",
    "🧽 Precursors (EVA, Acetone, Naphtha)",
    "⚙️ Soles & Hardware (Brass, Latex, Zinc)",
    "📈 Price Variation Analysis",
    "📄 Master Audit Trail"
])

# ------------------------------------------
# TAB 1: EXCHANGE RATE TRACKER
# ------------------------------------------
with tabs[0]:
    st.subheader("Global & Regional Exchange Rate Tracker")
    st.markdown("""
        <div class="executive-box">
            <h4>Data Source: RBI & ExchangeRates</h4>
            <p>Tracking major global currencies (USD/EUR) alongside the Chinese Yuan (RMB) to gauge inflationary impacts on both commodity prices and landed costs for direct imported components.</p>
        </div>
    """, unsafe_allow_html=True)

    display_kpis(["USD-INR", "EUR-INR", "CNY-INR"], fx, is_fx=True)
    st.markdown("###")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**USD Trend vs INR**")
        st.plotly_chart(build_trend_chart([{"prod": "USD-INR", "is_fx": True}], fx, unit_label="INR"),
                        use_container_width=True)
    with col2:
        st.markdown("**EUR Trend vs INR**")
        st.plotly_chart(build_trend_chart([{"prod": "EUR-INR", "is_fx": True}], fx, unit_label="INR"),
                        use_container_width=True)
    with col3:
        st.markdown("**RMB Trend vs INR**")
        st.plotly_chart(build_trend_chart([{"prod": "CNY-INR", "is_fx": True}], fx, unit_label="INR"),
                        use_container_width=True)

# ------------------------------------------
# TAB 2: ITEM IMPACTING YARN
# ------------------------------------------
with tabs[1]:
    st.subheader("Items Impacting Yarn")
    st.markdown("""
        <div class="executive-box">
            <h4>Source: PolymerUpdate & RBI</h4>
            <p>Brent Crude acts as the ultimate upstream driver. Track PTA, MEG, and PP for cascading price adjustments in knitted vamps, airmesh, and woven labels.</p>
        </div>
    """, unsafe_allow_html=True)

    display_kpis(["PTA", "MEG", "PP", "Brent"], df)
    st.markdown("###")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**PTA & MEG Trajectory**")
        st.plotly_chart(build_trend_chart([{"prod": "PTA"}, {"prod": "MEG"}], df), use_container_width=True)
    with col2:
        st.markdown("**PP Trajectory**")
        st.plotly_chart(build_trend_chart([{"prod": "PP"}], df), use_container_width=True)
    with col3:
        st.markdown("**Brent Crude Trajectory**")
        st.plotly_chart(build_trend_chart([{"prod": "Brent"}], df, unit_label="USD/BBL"), use_container_width=True)
# ------------------------------------------
# TAB 3: ITEM IMPACTING PVC COATED
# ------------------------------------------
with tabs[2]:
    st.subheader("Items Impacting PVC Coated Fabric")
    st.markdown("""
        <div class="executive-box">
            <h4>Source: PolymerUpdate</h4>
            <p>PVC Suspension provides the baseline cost for coated synthetic leather and binding. DOP (plasticizer) is highly volatile and dictates flexibility and premium coating costs.</p>
        </div>
    """, unsafe_allow_html=True)

    display_kpis(["PVC_Suspension", "DOP"], df)
    st.markdown("###")
    st.plotly_chart(build_trend_chart([{"prod": "DOP"}, {"prod": "PVC_Suspension"}], df), use_container_width=True)

# ------------------------------------------
# TAB 4: PACKAGING FILMS (LDPE & LLDPE)
# ------------------------------------------
with tabs[3]:
    st.subheader("Packaging Films: LDPE & LLDPE")
    st.markdown("""
        <div class="executive-box">
            <h4>Source: PolymerUpdate</h4>
            <p>Tracking Low-Density and Linear Low-Density Polyethylene. These commodities dictate the baseline costs for flexible packaging, protective films, and shrink wraps.</p>
        </div>
    """, unsafe_allow_html=True)

    display_kpis(["LDPE", "LLDPE"], df)
    st.markdown("###")
    st.plotly_chart(build_trend_chart([{"prod": "LDPE"}, {"prod": "LLDPE"}], df), use_container_width=True)

# ------------------------------------------
# TAB 5: PRECURSORS
# ------------------------------------------
with tabs[4]:
    st.subheader("EVA, Acetone & Naphtha: Market Trend")
    st.markdown("""
        <div class="executive-box">
            <h4>Source: PolymerUpdate</h4>
            <p>Naphtha and Acetone act as primary petrochemical precursors. Spikes in these upstream commodities cascade into PU foam and EVA sheet cost increases within a 4-6 week lag.</p>
        </div>
    """, unsafe_allow_html=True)

    display_kpis(["Naphtha", "ACETONE", "EVA"], df)
    st.markdown("###")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**Naphtha Trend**")
        st.plotly_chart(build_trend_chart([{"prod": "Naphtha"}], df), use_container_width=True)
    with col2:
        st.markdown("**Acetone Trend**")
        st.plotly_chart(build_trend_chart([{"prod": "ACETONE"}], df), use_container_width=True)
    with col3:
        st.markdown("**EVA Trend**")
        st.plotly_chart(build_trend_chart([{"prod": "EVA"}], df), use_container_width=True)

# ------------------------------------------
# TAB 6: SOLES & HARDWARE
# ------------------------------------------
with tabs[5]:
    st.subheader("Brass, Latex & Zinc: Market Trend")
    st.markdown("""
        <div class="executive-box">
            <h4>Source: Westmetall & PCK Ltd.</h4>
            <p>Zinc (galvanizing) and Brass dictate hardware/trim costs (eyelets, buckles). Divergence between natural Latex and synthetic EVA allows for strategic sole recipe blending.</p>
        </div>
    """, unsafe_allow_html=True)

    display_kpis(["BRASS", "LATEX", "ZINC"], df)
    st.markdown("###")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**Brass Trend**")
        st.plotly_chart(build_trend_chart([{"prod": "BRASS"}], df), use_container_width=True)
    with col2:
        st.markdown("**Latex Trend**")
        st.plotly_chart(build_trend_chart([{"prod": "LATEX"}], df), use_container_width=True)
    with col3:
        st.markdown("**Zinc Trend**")
        st.plotly_chart(build_trend_chart([{"prod": "ZINC"}], df), use_container_width=True)

# ------------------------------------------
# TAB 7: PRICE VARIATION ANALYSIS
# ------------------------------------------
with tabs[6]:
    st.subheader("Historical Price Variation Matrix")
    st.markdown(
        "Track percentage changes in raw material costs across standard timeframes. *(Note: Uses the closest available trading day for historical comparison).*")

    all_products = []
    for prod in df['Product'].unique():
        all_products.append({'name': prod, 'df': df[df['Product'] == prod], 'val_col': 'Final_INR_KG_Price'})
    for fx_pair in fx['Product'].unique():
        all_products.append({'name': fx_pair, 'df': fx[fx['Product'] == fx_pair], 'val_col': 'Price'})

    intervals = {
        "1 Day": 1, "3 Days": 3, "7 Days": 7, "2 Weeks": 14,
        "1 Month": 30, "2 Months": 60,
        "3 Months": 90, "6 Months": 180, "1 Year": 365
    }

    matrix_rows = []
    for item in all_products:
        prod_df = item['df'].sort_values('Date', ascending=False)
        if prod_df.empty: continue

        latest_row = prod_df.iloc[0]
        latest_price = latest_row[item['val_col']]
        latest_date = latest_row['Date']

        row = {"Commodity": item['name'], "Latest Price": latest_price}

        for label, days in intervals.items():
            target_date = latest_date - pd.Timedelta(days=days)
            past_df = prod_df[prod_df['Date'] <= target_date]

            if not past_df.empty:
                past_price = past_df.iloc[0][item['val_col']]
                pct_change = ((latest_price - past_price) / past_price) * 100
                row[label] = pct_change
            else:
                # 20-day tolerance fallback for 1-Year edge cases
                oldest_date = prod_df['Date'].min()
                if (oldest_date - target_date).days <= 20:
                    past_price = prod_df.iloc[-1][item['val_col']]
                    pct_change = ((latest_price - past_price) / past_price) * 100
                    row[label] = pct_change
                else:
                    row[label] = None

        matrix_rows.append(row)

    matrix_df = pd.DataFrame(matrix_rows)


    def format_pct(val):
        return "-" if pd.isna(val) else f"{val:+.2f}%"


    format_dict = {col: format_pct for col in intervals.keys()}
    format_dict["Latest Price"] = "₹{:,.2f}"


    def color_procurement_risk(val):
        if isinstance(val, (int, float)) and not pd.isna(val):
            color = '#e52d27' if val > 0 else '#00875a' if val < 0 else 'black'
            return f'color: {color}; font-weight: bold;'
        return ''


    styled_matrix = matrix_df.style.format(format_dict).map(
        color_procurement_risk,
        subset=list(intervals.keys())
    )

    st.dataframe(styled_matrix, use_container_width=True, hide_index=True)
    st.caption(
        "🔍 **Procurement Insight:** **Red** percentages indicate a raw material cost increase over the period. **Green** percentages indicate favorable cost savings.")
    st.markdown("---")

    # Custom Date Range Calculator
    st.subheader("Custom Date Range Calculator")
    col1, col2, col3 = st.columns(3)
    with col1:
        sel_prod = st.selectbox("Select Commodity to Analyze:", [p['name'] for p in all_products])
    with col2:
        start_date = st.date_input("Start Date (Baseline)", value=pd.to_datetime('today') - pd.Timedelta(days=30))
    with col3:
        end_date = st.date_input("End Date (Current)", value=pd.to_datetime('today'))

    if start_date and end_date:
        target_item = next(p for p in all_products if p['name'] == sel_prod)
        t_df = target_item['df'].sort_values('Date', ascending=False)

        start_df = t_df[t_df['Date'] <= pd.to_datetime(start_date)]
        end_df = t_df[t_df['Date'] <= pd.to_datetime(end_date)]

        if not start_df.empty and not end_df.empty:
            s_price = start_df.iloc[0][target_item['val_col']]
            s_date_actual = start_df.iloc[0]['Date'].strftime('%d-%b-%Y')

            e_price = end_df.iloc[0][target_item['val_col']]
            e_date_actual = end_df.iloc[0]['Date'].strftime('%d-%b-%Y')

            custom_pct = ((e_price - s_price) / s_price) * 100
            curr = "$" if sel_prod == "Brent" else "₹"
            st.metric(
                label=f"{sel_prod} Price Shift: {s_date_actual} ➔ {e_date_actual}",
                value=f"₹{e_price:,.2f}",
                delta=f"{custom_pct:+.2f}% (from ₹{s_price:,.2f})",
                delta_color="inverse"
            )
        else:
            st.warning("Insufficient historical data to cover this specific date range.")

# ------------------------------------------
# TAB 8: AUDIT TRAIL
# ------------------------------------------
with tabs[7]:
    st.subheader("Data Audit Trail & Conversion Formulas")
    st.markdown("Full transparency on raw inputs, currency conversions, and calculation formulas.")

    selected_prod = st.selectbox("Filter Audit Data by Product:",
                                 ["All Products"] + sorted(df['Product'].unique().tolist()))

    audit_view = df if selected_prod == "All Products" else df[df['Product'] == selected_prod]
    st.dataframe(audit_view.sort_values(by=['Date'], ascending=False), use_container_width=True)

    csv_data = audit_view.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Filtered Data as CSV",
        data=csv_data,
        file_name="Footwear_Raw_Material_Audit.csv",
        mime="text/csv"
    )