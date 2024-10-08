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

st.markdown("""
    <style>
        .stMultiSelect [data-baseweb=select] span{
            max-width: 250px;
            font-size: 0.7rem;
        }
    </style>
    """, unsafe_allow_html=True)

# st.multiselect("pick a long string", 
#     [i*25 for i in "ABCDEWXYZ"])

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

# Create a column to hold the value information without the year
# df['Text'] = df.apply(lambda row: f"<b>{row['Value']:.1f}</b>", axis=1)
# Create a column to hold the value information along with weights
# df['Text'] = df.apply(lambda row: f"<b>{row['Value']:.1f}</b> <span style='font-size:70%'> (w {row['Weight']:.2f})</span>", axis=1)


metric_types = ["Index", "Inflation"]
sector_types = ["All", "Rural", "Urban", "Combined"]

selected_metric_type = st.sidebar.selectbox("Select Metric Type", metric_types)


def format_text(row, metric_type):
    value = f"<b>{row['Value']:.1f}</b>"
    if metric_type == "Inflation":
        value += "%"
    return f"{value} <span style='font-size:70%'> (w {row['Weight']:.2f})</span>"

# Create a column to hold the value information along with weights
df['Text'] = df.apply(lambda row: format_text(row, selected_metric_type), axis=1)

# Filter dataframe based on selected metric type
df_filtered = df[df['ValueType'] == selected_metric_type].copy()

df_filtered = df_filtered.replace("", np.nan).dropna()

# Define main categories for each sector type
main_categories = [
    "A) General Index - Rural", "A.1) Food and beverages - Rural", 
    "A.2) Pan, tobacco and intoxicants - Rural", 
    "A.3) Clothing and footwear - Rural", "A.4) Housing - Rural", 
    "A.5) Fuel and light - Rural", "A.6) Miscellaneous - Rural",
    "A) General Index - Urban", "A.1) Food and beverages - Urban", 
    "A.2) Pan, tobacco and intoxicants - Urban", 
    "A.3) Clothing and footwear - Urban", "A.4) Housing - Urban", 
    "A.5) Fuel and light - Urban", "A.6) Miscellaneous - Urban",
    "A) General Index - Combined", "A.1) Food and beverages - Combined", 
    "A.2) Pan, tobacco and intoxicants - Combined", 
    "A.3) Clothing and footwear - Combined", "A.4) Housing - Combined", 
    "A.5) Fuel and light - Combined", "A.6) Miscellaneous - Combined",
    "B) Consumer Food Price Index - Rural",
    "B) Consumer Food Price Index - Urban",
    "B) Consumer Food Price Index - Combined",
]

# Additional filter for Main Cat, Sub Cat, or Both
category_types = ["Both", "Main Cat", "Sub Cat"]
selected_category_type = st.sidebar.selectbox("Select Category Type", category_types)

if selected_category_type == "Main Cat":
    df_filtered = df_filtered[df_filtered['Description'].apply(lambda x: any(main in x for main in main_categories) or "General Index" in x or "Housing" in x
        or "Pan, tobacco and intoxicants" in x or "Fuel and light" in x)]
elif selected_category_type == "Sub Cat":
    df_filtered = df_filtered[df_filtered['Description'].apply(lambda x: not any(main in x for main in main_categories) or "General Index" in x or "Housing" in x
        or "Pan, tobacco and intoxicants" in x or "Fuel and light" in x)]

selected_sector_type = st.sidebar.selectbox("Select Sector Type", sector_types)

# Initialize session state for selected descriptions (new lines)
# if 'selected_description' not in st.session_state:
#     st.session_state.selected_description = []

# #New Lines added to take care of selection in between so that menue does not reset (begin)
# # Prepare options for the multiselect based on sector type selection
# if selected_sector_type == "All":
#     description_options = df_filtered['Description'].unique().tolist()
# else:
#     description_options = df_filtered[df_filtered['Description'].str.contains(re.escape(selected_sector_type))]['Description'].unique().tolist()

# # Check if the current selected descriptions are valid for the new sector type
# valid_selected_description = [desc for desc in st.session_state.selected_description if desc in description_options]

# # If the selected descriptions are not valid for the new sector type, reset the multiselect
# if len(valid_selected_description) != len(st.session_state.selected_description):
#     st.session_state.selected_description = []

# # Update the multiselect box with valid options
# selected_description = st.sidebar.multiselect("Select Description to Display", description_options, default=st.session_state.selected_description)

# # Update session state with the new selections
# st.session_state.selected_description = selected_description

#New Lines added to take care of selection in between so that menue does not reset (end)

# Incase there is a bug in the above code then replace this block below with the above
# Prepare options for the multiselect based on sector type selection
if selected_sector_type == "All":
    description_options = df_filtered['Description'].unique().tolist()
    selected_description = st.sidebar.multiselect("Select Description to Display", description_options)
else:
    description_options = df_filtered[df_filtered['Description'].str.contains(re.escape(selected_sector_type))]['Description'].unique().tolist()
    selected_description = st.sidebar.multiselect("Select Description to Display", description_options, default=description_options)


# Filter dataframe based on selected main description
if selected_description:
    df_filtered = df_filtered[df_filtered['Description'].isin(selected_description)]

# Calculate the overall min and max values for the 'Value' column in the entire dataset
overall_min_value = df_filtered['Value'].min()
overall_max_value = df_filtered['Value'].max()

# Ensure the order of descriptions does not change when 'All' is selected
if selected_sector_type != "All":
    # Reorder the Description column based on the selected sector type
    description_order = get_description_order(selected_sector_type, df_filtered)
    if description_order:
        df_filtered['Description'] = pd.Categorical(df_filtered['Description'], categories=description_order, ordered=True)
        df_filtered = df_filtered.sort_values('Description')  # Sort the dataframe by Description to ensure the order is maintained
else:
    # Preserve the order of selected descriptions
    selected_description_order = selected_description
    df_filtered['Description'] = pd.Categorical(df_filtered['Description'], categories=selected_description_order, ordered=True)
    df_filtered = df_filtered.sort_values('Description')  # Sort the dataframe by Description to ensure the order is maintained


# Check if there is any data left after filtering
if selected_sector_type == "All" and not selected_description:
    st.write("Please select at least one description to display the data.")
elif df_filtered.empty:
    st.write("No data available for the selected filters.")
else:
    # Create the 'Weighted Average' column
    df_filtered['Weighted Average'] = df_filtered['Value'] * df_filtered['Weight'] / 100
    min_weighted_avg = df_filtered['Weighted Average'].min()
    max_weighted_avg = df_filtered['Weighted Average'].max()
    
    # Manually set the date range in the sidebar
    unique_dates = df_filtered['Date'].dt.date.unique()
    unique_dates = sorted(unique_dates)  # Ensure dates are sorted
    date_index = st.slider("Slider for Selecting Date Index", min_value=0, max_value=len(unique_dates) - 1, value=0)
    # date_index = st.slider("", min_value=0, max_value=len(unique_dates) - 1, value=0)
    selected_date = unique_dates[date_index]
    
    df_filtered_date = df_filtered[df_filtered['Date'].dt.date == selected_date]

    fig = make_subplots(rows=1, cols=2, shared_yaxes=True, column_widths=[0.75, 0.2], horizontal_spacing=0.01)  # Minimal horizontal spacing

    # Create scatter plot
    scatter_fig = px.scatter(df_filtered_date, x="Value", y="Description", color="Description", size_max=20, text="Text")
    scatter_fig.update_traces(marker=dict(size=20))  # Adjust the size value as needed
    scatter_fig.update_traces(marker=dict(line=dict(width=1, color='black')), textposition='middle right', textfont=dict(family='Arial', size=15, color='black', weight='bold'))
    scatter_fig.update_layout(showlegend=False, xaxis_title="Value of " + selected_metric_type)

    # df_filtered_date = df_filtered_date.dropna() #Debug 16th July 2024

    # Map colors from scatter plot to bar plot
    color_map = {desc: trace.marker.color for desc, trace in zip(df_filtered_date['Description'], scatter_fig.data)}

    # Create bar plot
    bar_fig = px.bar(df_filtered_date, x="Weighted Average", y="Description", orientation='h', text_auto='.2f')
    bar_fig.update_traces(textposition='outside', textfont=dict(size=15, family='Arial', color='black', weight='bold'))
    bar_fig.update_traces(marker=dict(line=dict(width=2, color='black')))
    bar_fig.update_traces(marker_color=[color_map[desc] for desc in df_filtered_date['Description']])
    bar_fig.update_layout(showlegend=False, xaxis_title="Weighted Average", yaxis=dict(showticklabels=False))


    # Update the y-axis tick labels to be bold
    fig.update_yaxes(tickfont=dict(size=15, family='Arial', color='black', weight='bold'), row=1, col=1)

    for trace in scatter_fig.data:
        fig.add_trace(trace, row=1, col=1)

    for trace in bar_fig.data:
        fig.add_trace(trace, row=1, col=2)

    # Create a reversed list of categories (descriptions)
    categories_reversed = df_filtered_date['Description'].tolist()[::-1]

    # Reverse the order of the y-axis for both the scatter plot and the bar plot
    fig.update_yaxes(categoryorder='array', categoryarray=categories_reversed, row=1, col=1)
    fig.update_yaxes(categoryorder='array', categoryarray=categories_reversed, row=1, col=2)

   # Update the layout for the combined figure
    fig.update_xaxes(row=1, col=1, range=[overall_min_value, overall_max_value * 1.05], fixedrange=True, showline=True, linewidth=1.5, linecolor='grey', mirror=True, showgrid=True, gridcolor='lightgrey')
    fig.update_yaxes(row=1, col=1, tickfont=dict(size=15),fixedrange=True, showline=True, linewidth=1.5, linecolor='grey', mirror=True, showgrid=True, gridcolor='lightgrey')

    if selected_metric_type == "Inflation":
        fig.update_xaxes(row=1, col=2, range=[min_weighted_avg*3, max_weighted_avg * 1.4],fixedrange=True, showline=True, linewidth=1.5, linecolor='grey', mirror=True, showgrid=True, gridcolor='lightgrey')
    else:
         fig.update_xaxes(row=1, col=2, range=[0, max_weighted_avg * 1.4],fixedrange=True, showline=True, linewidth=1.5, linecolor='grey', mirror=True, showgrid=True, gridcolor='lightgrey')

    fig.update_yaxes(row=1, col=2, fixedrange=True, showline=True, linewidth=1.5, linecolor='grey', mirror=True, showgrid=True, gridcolor='lightgrey')
    
    fig.update_layout(height=700, width=1200, margin=dict(l=5, r=10, t=0, b=0, pad=0), showlegend=False, yaxis=dict(automargin=True))

     # Update the layout for the combined figure with x-axis labels
    fig.update_xaxes(title_text="CPI " + selected_metric_type, row=1, col=1, title_font=dict(size=15, family='Arial', color='black', weight='bold'))
    fig.update_xaxes(title_text="Weight Adjusted Values", row=1, col=2, title_font=dict(size=15, family='Arial', color='black', weight='bold'))


    # Display the date with month on top along with the title
    # Create the styled title
    styled_category_type = f"<span style='color:red; font-weight:bold;'>{selected_category_type}</span>"
    styled_sector_type = f"<span style='color:blue; font-weight:bold;'>{selected_sector_type}</span>"
    styled_metric_type = f"<span style='color:brown; font-weight:bold;'>{selected_metric_type}</span>"
    styled_month = f"<span style='color:green; font-weight:bold;'>{selected_date.strftime('%b %Y')}</span>"
    title = f"Consumer Price {styled_category_type} {styled_sector_type} {styled_metric_type} Data For Month - {styled_month}"

    # Display the date with month on top along with the title
    st.markdown(f"<h1 style='font-size:30px; margin-top: -20px;'>{title}</h1>", unsafe_allow_html=True)

    st.plotly_chart(fig, use_container_width=True)
