import streamlit as st
import pandas as pd
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import os

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
    file_path = os.path.join(os.getcwd(), "analysis_output", "Master_Trend_Data.xlsx")
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
- **PolymerUpdate**: PVC, EVA, Naphtha, DOP, Acetone, MEG, PTA, PP, Brent
- **RBI**: Reference FX Rates (USD, EUR)
- **ExchangeRates.org.uk**: CNY/INR History
- **Westmetall**: LME Zinc & Brass
- **PCK Ltd**: Natural Rubber / Latex
""")

# High-contrast corporate colors matching report generator
CHART_COLORS = ['#0052cc', '#e52d27', '#00b8d9', '#6554c0', '#ff8b00']


# Helper to aggregate to monthly if selected
def prepare_data(data_df, value_col):
    data_df = data_df.copy().sort_values('Date')
    one_year_ago = pd.to_datetime('today') - pd.DateOffset(years=1)
    data_df = data_df[data_df['Date'] >= one_year_ago]

    # Safety check: If no data exists in the last year, return empty immediately
    if data_df.empty:
        return data_df

    if view_mode.startswith("1-Year Monthly"):
        # Explicitly set the Date as the index BEFORE resampling
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

            # Format text labels with zero decimals
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


# Metric KPI helper
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

            unit = "INR" if is_fx else "₹/kg"
            cols[i].metric(
                label=f"{prod} Current",
                value=f"{latest:,.2f} {unit}" if is_fx else f"{latest:,.0f} {unit}",
                delta=f"{pct_change:+.2f}% vs last period" if is_fx else f"{pct_change:+.1f}% vs last period"
            )


# ==========================================
# SLIDE-ALIGNED TAB NAVIGATION
# ==========================================
tabs = st.tabs([
    "💱 Exchange Rates (USD/EUR/RMB)",
    "🧵 Yarn Impact (PTA, MEG, PP, Brent)",
    "🧪 PVC Coated (DOP, Suspension)",
    "🧽 Precursors (EVA, Acetone, Naphtha)",
    "⚙️ Soles & Hardware (Brass, Latex, Zinc)",
    "📄 Master Audit Trail"
])

# ------------------------------------------
# TAB 1: EXCHANGE RATE TRACKER (PAGE 1)
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
# TAB 2: ITEM IMPACTING YARN (PAGE 2)
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

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**PTA & MEG Trajectory**")
        st.plotly_chart(build_trend_chart([{"prod": "PTA"}, {"prod": "MEG"}], df), use_container_width=True)
    with col2:
        st.markdown("**PP & Brent Crude Trajectory**")
        st.plotly_chart(build_trend_chart([{"prod": "PP"}, {"prod": "Brent"}], df), use_container_width=True)

# ------------------------------------------
# TAB 3: ITEM IMPACTING PVC COATED (PAGE 3)
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
# TAB 4: PRECURSORS (EVA, ACETONE, NAPHTHA) (PAGE 4)
# ------------------------------------------
with tabs[3]:
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
# TAB 5: SOLES & HARDWARE (BRASS, LATEX, ZINC) (PAGE 5)
# ------------------------------------------
with tabs[4]:
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
# TAB 6: AUDIT TRAIL & EXPORT
# ------------------------------------------
with tabs[5]:
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