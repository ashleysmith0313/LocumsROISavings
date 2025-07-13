
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from openpyxl import load_workbook
from io import BytesIO

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
                            editable_values.append({'description': description, 'value': cell.value})

    st.sidebar.header("🔧 Adjust Model Inputs")
    input_values = {}
    for item in editable_values:
        label = f"{item['description']}"
        input_values[label] = st.sidebar.slider(label, min_value=0, max_value=int(item['value'] * 2), value=int(item['value']))

    # Simulated data - replace with your real logic
    months = list(range(1, 25))
    categories = ['Permanent', 'Float Pool', 'VISTA Locums']

    shifts_data = {
        'Permanent': [input_values.get('Providers Onboarded per Month', 0) * month for month in months],
        'Float Pool': [input_values.get('Average Days per provider per Month', 0) * month for month in months],
        'VISTA Locums': [input_values.get('Hospitalist', 0) * month for month in months]
    }

    baseline_cost = 500000  # Example baseline for savings calculation

    cost_data = {
        'Permanent': [baseline_cost - shifts_data['Permanent'][month - 1] * 100 for month in months],
        'Float Pool': [baseline_cost - shifts_data['Float Pool'][month - 1] * 80 for month in months],
        'VISTA Locums': [baseline_cost - shifts_data['VISTA Locums'][month - 1] * 200 for month in months]
    }

    st.subheader("📈 Shift Coverage Over 24 Months")
    fig1 = go.Figure()
    for cat in categories:
        fig1.add_trace(go.Bar(x=months, y=shifts_data[cat], name=cat))
    fig1.update_layout(barmode='stack', xaxis_title="Month", yaxis_title="Shifts")

    st.plotly_chart(fig1, use_container_width=True)

    st.subheader("💰 Staffing Costs & Savings Over 24 Months")
    fig2 = go.Figure()
    for cat in categories:
        fig2.add_trace(go.Bar(x=months, y=cost_data[cat], name=cat))

    # Adding savings annotation per month (example with baseline - stacked total)
    for month in months:
        total_cost = sum(cost_data[cat][month - 1] for cat in categories)
        savings = baseline_cost - total_cost
        fig2.add_annotation(x=month, y=0, text=f"Save: ${savings:,.0f}", showarrow=False, yshift=-15, font=dict(size=10))

    fig2.update_layout(barmode='stack', xaxis_title="Month", yaxis_title="Cost ($)")
    st.plotly_chart(fig2, use_container_width=True)

    st.success("✅ Graphs updated with current editable inputs.")
else:
    st.info("Please upload an Excel file to get started.")
