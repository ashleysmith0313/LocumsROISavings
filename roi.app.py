import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from openpyxl import load_workbook

st.set_page_config(page_title="Commonwealth Cares Deployment Model", layout="wide")

st.title("📊 Commonwealth Cares Deployment Visualizer")
st.markdown("Upload your Excel model and adjust key inputs to visualize shift coverage and staffing cost savings over **24 months**.")

uploaded_file = st.file_uploader("Upload your Excel File", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    sheet = wb['Combo- Select & Float Pool']

    editable_values = []
    for row in sheet.iter_rows():
        for idx, cell in enumerate(row):
            if isinstance(cell.value, (int, float)):
                if cell.font and cell.font.color and cell.font.color.type == 'rgb':
                    if cell.font.color.rgb.upper().startswith('FFFF0000'):
                        description = None
                        for prev_cell in reversed(row[:idx]):
                            if isinstance(prev_cell.value, str) and prev_cell.value.strip() != "":
                                description = prev_cell.value.strip()
                                break
                        if description:
                            editable_values.append({'description': description, 'value': cell.value, 'cell': cell.coordinate})

    st.sidebar.header("🔧 Adjust Model Inputs")
    input_values = {}

    with st.sidebar.expander("📌 Permanent Staffing Inputs"):
        st.markdown("**Permanent staffing ramp-up assumptions.**")
        for item in editable_values:
            if 'Providers Onboarded per Month (B21)' in f"{item['description']} ({item['cell']})" or 'Average Days per provider per Month (B22)' in f"{item['description']} ({item['cell']})":
                label = f"{item['description']} ({item['cell']})"
                input_values[label] = st.slider(
                    label,
                    min_value=0,
                    max_value=int(item['value'] * 2),
                    value=int(item['value'])
                )

    with st.sidebar.expander("📌 Float Pool Inputs"):
        st.markdown("**Float Pool deployment assumptions.**")
        for item in editable_values:
            if 'Open Days per Month (C17)' in f"{item['description']} ({item['cell']})" or 'Average Days per provider per Month (B27)' in f"{item['description']} ({item['cell']})" or 'Providers Onboarded per Month (B26)' in f"{item['description']} ({item['cell']})":
                label = f"{item['description']} ({item['cell']})"
                input_values[label] = st.slider(
                    label,
                    min_value=0,
                    max_value=int(item['value'] * 2),
                    value=int(item['value'])
                )

    with st.sidebar.expander("📌 VISTA Locums Inputs"):
        st.markdown("**Locums usage assumptions.**")
        for item in editable_values:
            if 'Open Days per Month (D17)' in f"{item['description']} ({item['cell']})" or 'Hospitalist' in item['description']:
                label = f"{item['description']} ({item['cell']})"
                input_values[label] = st.slider(
                    label,
                    min_value=0,
                    max_value=int(item['value'] * 2),
                    value=int(item['value'])
                )

    months = list(range(1, 25))

    permanent_onboard_rate = input_values.get('Providers Onboarded per Month (B21)', 0)
    permanent_days_per_provider = input_values.get('Average Days per provider per Month (B22)', 0)
    float_pool_onboard_rate = input_values.get('Providers Onboarded per Month (B26)', 0)
    float_pool_open_days = input_values.get('Open Days per Month (C17)', 0)
    float_pool_days_per_provider = input_values.get('Average Days per provider per Month (B27)', 0)
    locum_open_days = input_values.get('Open Days per Month (D17)', 0)
    hospitalist_rate = input_values.get('Hospitalist (B4)', 0)

    permanent_shifts = []
    total_permanent = 0
    for month in months:
        if month >= 4:
            total_permanent += permanent_onboard_rate
        permanent_shifts.append(total_permanent * permanent_days_per_provider)

    float_pool_shifts = []
    total_float_pool = 0
    for month in months:
        if month >= 12:
            total_float_pool += float_pool_onboard_rate
            float_pool_shifts.append(total_float_pool * float_pool_days_per_provider)
        else:
            float_pool_shifts.append(0)

    locum_shifts = []
    for month in months:
        if month >= 4:
            locum_shifts.append(locum_open_days)
        else:
            locum_shifts.append(0)

    shifts_data = {
        'Permanent': permanent_shifts,
        'Float Pool': float_pool_shifts,
        'VISTA Locums': locum_shifts
    }

    cost_data = {
        'Permanent': [shifts * 100 for shifts in permanent_shifts],
        'Float Pool': [shifts * 80 for shifts in float_pool_shifts],
        'VISTA Locums': [shifts * hospitalist_rate for shifts in locum_shifts]
    }

    st.subheader("📈 Shift Coverage Over 24 Months")
    fig1 = go.Figure()
    for cat in ['Permanent', 'Float Pool', 'VISTA Locums']:
        fig1.add_trace(go.Bar(x=months, y=shifts_data[cat], name=cat, hovertemplate=f"%{{x}} Month<br>%{{y}} Shifts"))
    fig1.update_layout(barmode='stack', xaxis_title="Month", yaxis_title="Shifts")
    st.plotly_chart(fig1, use_container_width=True)

    st.subheader("💰 Staffing Costs & Savings Over 24 Months")
    fig2 = go.Figure()
    for cat in ['Permanent', 'Float Pool', 'VISTA Locums']:
        fig2.add_trace(go.Bar(x=months, y=cost_data[cat], name=cat, hovertemplate=f"%{{x}} Month<br>$%{{y:,.0f}} Cost"))
    fig2.update_layout(barmode='stack', xaxis_title="Month", yaxis_title="Cost ($)")
    st.plotly_chart(fig2, use_container_width=True)

    st.success("✅ Graphs updated with ramp logic, corrected grouping, and realistic locum start assumptions.")
else:
    st.info("Please upload an Excel file to get started.")
