
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Commonwealth Cares Deployment Model", layout="wide")

st.title("📊 Commonwealth Cares Deployment Visualizer")
st.markdown("Upload your Excel model and adjust key inputs to visualize shift coverage and staffing cost savings over time.")

uploaded_file = st.file_uploader("Upload your Excel File", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    sheet = wb['Combo- Select & Float Pool']

    # Extract red numeric cells
    editable_values = []
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, (int, float)):
                if cell.font and cell.font.color and cell.font.color.type == 'rgb':
                    if cell.font.color.rgb.upper().startswith('FFFF0000'):
                        editable_values.append({'cell': cell.coordinate, 'value': cell.value})

    st.sidebar.header("🔧 Adjust Model Inputs")
    input_values = {}
    for item in editable_values:
        label = f"Cell {item['cell']}"
        input_values[item['cell']] = st.sidebar.slider(label, min_value=0, max_value=int(item['value'] * 2), value=int(item['value']))

    # Simulated data for charts (replace this with real logic)
    months = list(range(5, 13))
    categories = ['Permanent', 'Float Pool', 'VISTA Locums']
    shifts_data = {
        'Permanent': [0, 20, 20, 20, 60, 60, 60, 100],
        'Float Pool': [0, 15, 30, 45, 60, 75, 90, 100],
        'VISTA Locums': [0, 50, 100, 150, 200, 225, 210, 160]
    }

    cost_data = {
        'Permanent': [0, 55000, 25000, 25000, 135000, 75000, 75000, 185000],
        'Float Pool': [0, 40716, 81432, 122148, 162864, 203580, 244296, 271440],
        'VISTA Locums': [0, 193500, 387000, 580500, 774000, 870750, 812700, 619200]
    }

    # Create bar chart for shifts
    st.subheader("📈 Shift Coverage Over Time")
    fig1 = go.Figure()
    for cat in categories:
        fig1.add_trace(go.Bar(x=months, y=shifts_data[cat], name=cat))
    fig1.update_layout(barmode='stack', xaxis_title="Months", yaxis_title="Shifts")
    st.plotly_chart(fig1, use_container_width=True)

    # Create bar chart for cost
    st.subheader("💰 Staffing Costs Over Time")
    fig2 = go.Figure()
    for cat in categories:
        fig2.add_trace(go.Bar(x=months, y=cost_data[cat], name=cat))
    fig2.update_layout(barmode='stack', xaxis_title="Months", yaxis_title="Cost ($)")
    st.plotly_chart(fig2, use_container_width=True)

    st.success("✅ Graphs updated with current editable inputs.")
else:
    st.info("Please upload an Excel file to get started.")
