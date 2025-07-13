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
        st.markdown("**These inputs control the onboarding rate and average shifts per provider for permanent staff.**")
        for item in editable_values:
            if 'Provider' in item['description'] or 'Days per provider' in item['description']:
                label = f"{item['description']} ({item['cell']})"
                input_values[label] = st.slider(
                    label,
                    min_value=0,
                    max_value=int(item['value'] * 2),
                    value=int(item['value']),
                    help="Adjusts permanent provider staffing growth and shift impact."
                )

    with st.sidebar.expander("📌 Float Pool Inputs"):
        st.markdown("**These inputs impact Float Pool deployment and cost control.**")
        for item in editable_values:
            if 'Open Days per Month (C17)' in f"{item['description']} ({item['cell']})" or 'Float' in item['description']:
                label = f"{item['description']} ({item['cell']})"
                input_values[label] = st.slider(
                    label,
                    min_value=0,
                    max_value=int(item['value'] * 2),
                    value=int(item['value']),
                    help="Drives Float Pool shift allocation and impacts cost model."
                )

    with st.sidebar.expander("📌 VISTA Locums Inputs"):
        st.markdown("**These inputs affect Locum usage assumptions and related costs.**")
        for item in editable_values:
            if 'Open Days per Month (D17)' in f"{item['description']} ({item['cell']})" or 'Hospitalist' in item['description']:
                label = f"{item['description']} ({item['cell']})"
                input_values[label] = st.slider(
                    label,
                    min_value=0,
                    max_value=int(item['value'] * 2),
                    value=int(item['value']),
                    help="Controls Locum staffing volume and cost impact."
                )

    months = list(range(1, 25))

    permanent_rate = input_values.get('Providers Onboarded per Month (B21)', 0) * input_values.get('Average Days per provider per Month (B22)', 0)
    float_pool_days = input_values.get('Open Days per Month (C17)', 0)
    vista_days = input_values.get('Open Days per Month (D17)', 0)
    hospitalist_rate = input_values.get('Hospitalist (B4)', 0)

    shifts_data = {
        'Permanent': [permanent_rate * month for month in months],
        'Float Pool': [float_pool_days * month for month in months],
        'VISTA Locums': [vista_days * month for month in months]
    }

    baseline_cost = 500000

    cost_data = {
        'Permanent': [baseline_cost - shifts_data['Permanent'][month - 1] * 100 for month in months],
        'Float Pool': [baseline_cost - shifts_data['Float Pool'][month - 1] * 80 for month in months],
        'VISTA Locums': [baseline_cost - shifts_data['VISTA Locums'][month - 1] * hospitalist_rate for month in months]
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

    st.success("✅ Graphs updated with current editable inputs.")
else:
    st.info("Please upload an Excel file to get started.")
