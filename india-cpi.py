import pandas as pd
from datetime import datetime
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import streamlit as st
import io
import msoffcrypto
import numpy as np
import re

pd.set_option('future.no_silent_downcasting', True)
pd.set_option('display.max_columns', None)

st.set_page_config(
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS to reduce font size in multi-select box
custom_css = """
<style>
.css-1okebmr-indicatorSeparator, .css-1wy0on6, .css-1hb7zxy-IndicatorsContainer, .css-1n7v3ny-option {
    font-size: 50% !important;
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# Hide Streamlit style and buttons
hide_st_style = '''
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
'''
st.markdown(hide_st_style, unsafe_allow_html=True)

# Load file function
@st.cache_data
def loadfile():
    password = st.secrets["db_password"]
    excel_content = io.BytesIO()
    with open("cpi_streamlit.xlsx", 'rb') as f:
        excel = msoffcrypto.OfficeFile(f)
        excel.load_key(password)
        excel.decrypt(excel_content)
    
    # Loading data from excel file
    df = pd.read_excel(excel_content, sheet_name="Sheet1")
    return df

# Function to get description order and append weights
def get_description_order(sector_type, df):
    order_dict = {
        "Rural": [
            "A) General Index - Rural",
            "A.1) Food and beverages - Rural",
            "A.1.1) Cereals and products - Rural",
            "A.1.2) Meat and fish - Rural",
            "A.1.3) Egg - Rural",
            "A.1.4) Milk and products - Rural",
            "A.1.5) Oils and fats - Rural",
            "A.1.6) Fruits - Rural",
            "A.1.7) Vegetables - Rural",
            "A.1.8) Pulses and products - Rural",
            "A.1.9) Sugar and confectionery - Rural",
            "A.1.10) Spices - Rural",
            "A.1.11) Non-alcoholic beverages - Rural",
            "A.1.12) Prepared meals, snacks, sweets etc. - Rural",
            "A.2) Pan, tobacco and intoxicants - Rural",
            "A.3) Clothing and footwear - Rural",
            "A.3.1) Clothing - Rural",
            "A.3.2) Footwear - Rural",
            "A.4) Housing - Rural",
            "A.5) Fuel and light - Rural",
            "A.6) Miscellaneous - Rural",
            "A.6.1) Household goods and services - Rural",
            "A.6.2) Health - Rural",
            "A.6.3) Transport and communication - Rural",
            "A.6.4) Recreation and amusement - Rural",
            "A.6.5) Education - Rural",
            "A.6.6) Personal Care and Effects - Rural",
            "B) Consumer Food Price Index - Rural"
        ],
        "Urban": [
            "A) General Index - Urban",
            "A.1) Food and beverages - Urban",
            "A.1.1) Cereals and products - Urban",
            "A.1.2) Meat and fish - Urban",
            "A.1.3) Egg - Urban",
            "A.1.4) Milk and products - Urban",
            "A.1.5) Oils and fats - Urban",
            "A.1.6) Fruits - Urban",
            "A.1.7) Vegetables - Urban",
            "A.1.8) Pulses and products - Urban",
            "A.1.9) Sugar and confectionery - Urban",
            "A.1.10) Spices - Urban",
            "A.1.11) Non-alcoholic beverages - Urban",
            "A.1.12) Prepared meals, snacks, sweets etc. - Urban",
            "A.2) Pan, tobacco and intoxicants - Urban",
            "A.3) Clothing and footwear - Urban",
            "A.3.1) Clothing - Urban",
            "A.3.2) Footwear - Urban",
            "A.4) Housing - Urban",
            "A.5) Fuel and light - Urban",
            "A.6) Miscellaneous - Urban",
            "A.6.1) Household goods and services - Urban",
            "A.6.2) Health - Urban",
            "A.6.3) Transport and communication - Urban",
            "A.6.4) Recreation and amusement - Urban",
            "A.6.5) Education - Urban",
            "A.6.6) Personal Care and Effects - Urban",
            "B) Consumer Food Price Index - Urban"
        ],
        "Combined": [
            "A) General Index - Combined",
            "A.1) Food and beverages - Combined",
            "A.1.1) Cereals and products - Combined",
            "A.1.2) Meat and fish - Combined",
            "A.1.3) Egg - Combined",
            "A.1.4) Milk and products - Combined",
            "A.1.5) Oils and fats - Combined",
            "A.1.6) Fruits - Combined",
            "A.1.7) Vegetables - Combined",
            "A.1.8) Pulses and products - Combined",
            "A.1.9) Sugar and confectionery - Combined",
            "A.1.10) Spices - Combined",
            "A.1.11) Non-alcoholic beverages - Combined",
            "A.1.12) Prepared meals, snacks, sweets etc. - Combined",
            "A.2) Pan, tobacco and intoxicants - Combined",
            "A.3) Clothing and footwear - Combined",
            "A.3.1) Clothing - Combined",
            "A.3.2) Footwear - Combined",
            "A.4) Housing - Combined",
            "A.5) Fuel and light - Combined",
            "A.6) Miscellaneous - Combined",
            "A.6.1) Household goods and services - Combined",
            "A.6.2) Health - Combined",
            "A.6.3) Transport and communication - Combined",
            "A.6.4) Recreation and amusement - Combined",
            "A.6.5) Education - Combined",
            "A.6.6) Personal Care and Effects - Combined",
            "B) Consumer Food Price Index - Combined"
        ]
    }
    
    # Add weights to the descriptions
    df['Description'] = df.apply(lambda row: f"{row['Description']} ({row['Weight']:.2f})", axis=1)
    
    order_list = order_dict.get(sector_type, [])
    new_order_list = []
    for item in order_list:
        matches = df[df['Description'].str.contains(re.escape(item.split(' - ')[0]))]
        if not matches.empty:
            weight = matches['Weight'].values[0]
            new_order_list.append(f"{item} ({weight:.2f})")
        else:
            new_order_list.append(f"{item} (NaN)")
    
    return new_order_list

# Main Program Starts Here
df = loadfile()

# Ensuring the Date column is of datetime type
df['Date'] = pd.to_datetime(df['Date'])

# Sorting dataframe by Date to ensure proper animation sequence
df = df.sort_values(by='Date')

# Convert Date column to string without time
df['Date_str'] = df['Date'].dt.strftime('%d-%m-%Y')

df["Value"] = df["Value"].replace("-", np.nan, regex=True)

# Format the Value column to two decimal places and keep it as a float
df['Value'] = df['Value'].astype(float).round(2)

# Create a column to hold the value information along with the year
df['Text'] = df.apply(lambda row: f"<b>{row['Value']:.2f} ({row['Date_str'][-4:]})</b>", axis=1)

metric_types = ["Index", "Inflation"]
sector_types = ["All", "Rural", "Urban", "Combined"]

selected_metric_type = st.sidebar.selectbox("Select Metric Type", metric_types)

# Filter dataframe based on selected metric type
df_filtered = df[df['ValueType'] == selected_metric_type].copy()

df_filtered = df_filtered.replace("", np.nan).dropna()

selected_sector_type = st.sidebar.selectbox("Select Sector Type", sector_types)

# Prepare options for the multiselect based on sector type selection
if selected_sector_type == "All":
    description_options = df_filtered['Description'].unique().tolist()
    selected_description = st.sidebar.multiselect("Select Description to Display", description_options)
else:
    description_options = df_filtered[df_filtered['Description'].str.contains(re.escape(selected_sector_type))]['Description'].unique().tolist()
    selected_description = st
