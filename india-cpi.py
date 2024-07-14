import pandas as pd
from datetime import datetime
import plotly.express as px
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
    selected_description = st.sidebar.multiselect("Select Description to Display", description_options, default=description_options)

# Filter dataframe based on selected main description
if selected_description:
    df_filtered = df_filtered[df_filtered['Description'].isin(selected_description)]

# Reorder the Description column based on the selected sector type
description_order = get_description_order(selected_sector_type, df_filtered)
if description_order:
    df_filtered['Description'] = pd.Categorical(df_filtered['Description'], categories=description_order, ordered=True)
    df_filtered = df_filtered.sort_values('Description')  # Sort the dataframe by Description to ensure the order is maintained

# Ensure the frames are sorted correctly
df_filtered['Weighted Average'] = df_filtered['Weighted Average'] / 100  # Scale down by 100
df_filtered['Weighted Average'] = df_filtered['Weighted Average'].round(2)  # Round to 2 decimal places

# Check if there is any data left after filtering
if selected_sector_type == "All" and not selected_description:
    st.write("Please select at least one description to display the data.")
elif df_filtered.empty:
    st.write("No data available for the selected filters.")
else:
    with st.spinner('Rendering the display, please wait...'):
        # Calculate min and max values for the dotted lines
        min_value = df_filtered['Value'].min()
        max_value = df_filtered['Value'].max()

        # Ensure Date_str is ordered correctly
        df_filtered['Date_str'] = pd.Categorical(df_filtered['Date_str'], ordered=True, categories=sorted(df_filtered['Date_str'].unique(), key=lambda x: datetime.strptime(x, '%d-%m-%Y')))

        # Calculate the range for the x-axis
        range_min = min_value - abs(min_value) * 0.30
        range_max = max_value + abs(max_value) * 0.15

        # Plotly animation setup
        fig = px.scatter(df_filtered, x="Value", y="Description", animation_frame="Date_str", animation_group="Description",
                         color="Description", range_x=[range_min, range_max],
                         title="", size="Weight", size_max=20, text="Text")

        # Customize text position to the right of the dots
        fig.update_traces(textposition='middle right', textfont=dict(size=16))

        # Add black outlines to the dots
        fig.update_traces(marker=dict(line=dict(width=2, color='black')))

        # Customize y-axis labels font size and make them bold
        fig.update_yaxes(tickfont=dict(size=15, color='black', family='Arial', weight='bold'))

        # Remove y-axis labels and variable labels
        fig.update_yaxes(showticklabels=True)

        # Draw a black line on the y-axis
        fig.add_shape(type='line', x0=0, x1=0, y0=0, y1=1, line=dict(color='black', width=1), xref='x', yref='paper')

        # Remove legend on the right side
        fig.update_layout(showlegend=False)

        # Add dotted lines for min and max values
        fig.add_shape(
            type="line",
            x0=min_value, y0=0, x1=min_value, y1=1,
            xref='x', yref='paper',
            line=dict(color="blue", width=2, dash="dot")
        )
        fig.add_shape(
            type="line",
            x0=max_value, y0=0, x1=max_value, y1=1,
            xref='x', yref='paper',
            line=dict(color="red", width=2, dash="dot")
        )

        # Adjust the layout
        fig.update_layout(
            xaxis_title="Value of "+selected_metric_type,
            yaxis_title="",
            width=1200,
            height=950,  # Adjust the height to make the plot more visible
            margin=dict(l=0, r=10, t=120, b=40, pad=0),  # Add margins to make the plot more readable and closer to the left
            sliders=[{
                'steps': [
                    {
                        'args': [
                            [date_str],
                            {
                                'frame': {'duration': 300, 'redraw': True},
                                'mode': 'immediate',
                                'transition': {'duration': 300}
                            }
                        ],
                        'label': date_str,
                        'method': 'animate'
                    }
                    for date_str in sorted(df_filtered['Date_str'].unique(), key=lambda x: datetime.strptime(x, '%d-%m-%Y'))
                ],
                'x': 0.1,
                'xanchor': 'left',
                'y': 0,
                'yanchor': 'top'
            }]
        )

        # Add initial annotation for the date
        initial_date_annotation = {
            'x': 0,
            'y': 1.15,  # Move the date annotation closer to the top of the chart
            'xref': 'paper',
            'yref':'paper',
            'text': f'<span style="color:red;font-size:30px"><b>Date: {df_filtered["Date_str"].iloc[0]}</b></span>',
            'showarrow': False,
            'font': {
                'size': 20
            }
        }
        fig.update_layout(annotations=[initial_date_annotation])

        # Custom callback to update the date annotation dynamically
        def update_annotations(date_str):
            return [go.layout.Annotation(
                x=0,
                y=1.15,
                xref='paper',
                yref='paper',
                text=f'<span style="color:red;font-size:30px"><b>Date: {date_str}</b></span>',
                showarrow=False,
                font=dict(size=30)
            )]

        # Customize y-axis labels font size
        fig.update_yaxes(tickfont=dict(size=15))

        # Update annotation with each frame
        for frame in fig.frames:
            date_str = frame.name
            frame['layout'].update(annotations=update_annotations(date_str))

        # Ensure the frames are sorted correctly
        fig.frames = sorted(fig.frames, key=lambda frame: datetime.strptime(frame.name, '%d-%m-%Y'))

        # Combine scatter plot and horizontal bar chart in the same layout
        combined_fig = make_subplots(rows=1, cols=2, column_widths=[0.7, 0.3], horizontal_spacing=0.02)

        for trace in fig['data']:
            combined_fig.add_trace(trace, row=1, col=1)

        for frame in fig.frames:
            combined_fig.add_trace(go.Bar(
                x=df_filtered[df_filtered['Date_str'] == frame.name]['Weighted Average'],
                y=df_filtered[df_filtered['Date_str'] == frame.name]['Description'],
                orientation='h',
                text=df_filtered[df_filtered['Date_str'] == frame.name]['Weighted Average'].apply(lambda x: f'{x:.2f}'),
                textposition='outside',
                showlegend=False
            ), row=1, col=2)

        combined_fig.update_yaxes(showticklabels=False, row=1, col=2)
        combined_fig.update_xaxes(title_text="Weighted Average", row=1, col=2)

        combined_fig.update_layout(
            title_text="Index and Weighted Average Over Time",
            height=900,
            width=1400,
            showlegend=False,
            sliders=fig.layout.sliders,
            updatemenus=[{
                'type': 'buttons',
                'showactive': False,
                'buttons': [
                    {
                        'label': 'Play',
                        'method': 'animate',
                        'args': [None, {
                            'frame': {'duration': 500, 'redraw': True},
                            'fromcurrent': True,
                            'transition': {'duration': 300, 'easing': 'linear'}
                        }]
                    },
                    {
                        'label': 'Pause',import pandas as pd
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
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

# Calculate weighted average
df['Weighted Average'] = df['Value'] * df['Weight'] / 100  # Assuming 'Value' is the index and 'Weight' is in percentage
df['Weighted Average'] = df['Weighted Average'].round(2)

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
    selected_description = st.sidebar.multiselect("Select Description to Display", description_options, default=description_options)

# Filter dataframe based on selected main description
if selected_description:
    df_filtered = df_filtered[df_filtered['Description'].isin(selected_description)]

# Reorder the Description column based on the selected sector type
description_order = get_description_order(selected_sector_type, df_filtered)
if description_order:
    df_filtered['Description'] = pd.Categorical(df_filtered['Description'], categories=description_order, ordered=True)
    df_filtered = df_filtered.sort_values('Description')  # Sort the dataframe by Description to ensure the order is maintained

# Check if there is any data left after filtering
if selected_sector_type == "All" and not selected_description:
    st.write("Please select at least one description to display the data.")
elif df_filtered.empty:
    st.write("No data available for the selected filters.")
else:
    with st.spinner('Rendering the display, please wait...'):
        # Calculate min and max values for the dotted lines
        min_value = df_filtered['Value'].min()
        max_value = df_filtered['Value'].max()

        # Ensure Date_str is ordered correctly
        df_filtered['Date_str'] = pd.Categorical(df_filtered['Date_str'], ordered=True, categories=sorted(df_filtered['Date_str'].unique(), key=lambda x: datetime.strptime(x, '%d-%m-%Y')))

        # Calculate the range for the x-axis
        range_min = min_value - abs(min_value) * 0.30
        range_max = max_value + abs(max_value) * 0.15

        # Plotly animation setup
        fig = px.scatter(df_filtered, x="Value", y="Description", animation_frame="Date_str", animation_group="Description",
                         color="Description", range_x=[range_min, range_max],
                         title="", size="Weight", size_max=20, text="Text")

        # Customize text position to the right of the dots
        fig.update_traces(textposition='middle right', textfont=dict(size=16))

        # Add black outlines to the dots
        fig.update_traces(marker=dict(line=dict(width=2, color='black')))

        # Customize y-axis labels font size and make them bold
        fig.update_yaxes(tickfont=dict(size=15, color='black', family='Arial', weight='bold'))

        # Remove y-axis labels and variable labels
        fig.update_yaxes(showticklabels=True)

        # Draw a black line on the y-axis
        fig.add_shape(type='line', x0=0, x1=0, y0=0, y1=1, line=dict(color='black', width=1), xref='x', yref='paper')

        # Remove legend on the right side
        fig.update_layout(showlegend=False)

        # Add dotted lines for min and max values
        fig.add_shape(
            type="line",
            x0=min_value, y0=0, x1=min_value, y1=1,
            xref='x', yref='paper',
            line=dict(color="blue", width=2, dash="dot")
        )
        fig.add_shape(
            type="line",
            x0=max_value, y0=0, x1=max_value, y1=1,
            xref='x', yref='paper',
            line=dict(color="red", width=2, dash="dot")
        )

        # Adjust the layout
        fig.update_layout(
            xaxis_title="Value of "+selected_metric_type,
            yaxis_title="",
            width=1200,
            height=950,  # Adjust the height to make the plot more visible
            margin=dict(l=0, r=10, t=120, b=40, pad=0),  # Add margins to make the plot more readable and closer to the left
            sliders=[{
                'steps': [
                    {
                        'args': [
                            [date_str],
                            {
                                'frame': {'duration': 300, 'redraw': True},
                                'mode': 'immediate',
                                'transition': {'duration': 300}
                            }
                        ],
                        'label': date_str,
                        'method': 'animate'
                    }
                    for date_str in sorted(df_filtered['Date_str'].unique(), key=lambda x: datetime.strptime(x, '%d-%m-%Y'))
                ],
                'x': 0.1,
                'xanchor': 'left',
                'y': 0,
                'yanchor': 'top'
            }]
        )

        # Add initial annotation for the date
        initial_date_annotation = {
            'x': 0,
            'y': 1.15,  # Move the date annotation closer to the top of the chart
            'xref': 'paper',
            'yref':'paper',
            'text': f'<span style="color:red;font-size:30px"><b>Date: {df_filtered["Date_str"].iloc[0]}</b></span>',
            'showarrow': False,
            'font': {
                'size': 20
            }
        }
        fig.update_layout(annotations=[initial_date_annotation])

        # Custom callback to update the date annotation dynamically
        def update_annotations(date_str):
            return [go.layout.Annotation(
                x=0,
                y=1.15,
                xref='paper',
                yref='paper',
                text=f'<span style="color:red;font-size:30px"><b>Date: {date_str}</b></span>',
                showarrow=False,
                font=dict(size=30)
            )]

        # Customize y-axis labels font size
        fig.update_yaxes(tickfont=dict(size=15))

        # Update annotation with each frame
        for frame in fig.frames:
            date_str = frame.name
            frame['layout'].update(annotations=update_annotations(date_str))

        # Ensure the frames are sorted correctly
        fig.frames = sorted(fig.frames, key=lambda frame: datetime.strptime(frame.name, '%d-%m-%Y'))

        # Combine scatter plot and horizontal bar chart in the same layout
        combined_fig = make_subplots(rows=1, cols=2, column_widths=[0.7, 0.3], horizontal_spacing=0.02)

        for trace in fig['data']:
            combined_fig.add_trace(trace, row=1, col=1)

        for frame in fig.frames:
            combined_fig.add_trace(go.Bar(
                x=df_filtered[df_filtered['Date_str'] == frame.name]['Weighted Average'],
                y=df_filtered[df_filtered['Date_str'] == frame.name]['Description'],
                orientation='h',
                text=df_filtered[df_filtered['Date_str'] == frame.name]['Weighted Average'].apply(lambda x: f'{x:.2f}'),
                textposition='outside',
                showlegend=False
            ), row=1, col=2)

        combined_fig.update_yaxes(showticklabels=False, row=1, col=2)
        combined_fig.update_xaxes(title_text="Weighted Average", row=1, col=2)

        combined_fig.update_layout(
            title_text="Index and Weighted Average Over Time",
            height=900,
            width=1400,
            showlegend=False,
            sliders=fig.layout.sliders,
            updatemenus=[{
                'type': 'buttons',
                'showactive': False,
                'buttons': [
                    {
                        'label': 'Play',
                        'method': 'animate',
                        'args': [None, {
                            'frame': {'duration': 500, 'redraw': True},
                            'fromcurrent': True,
                            'transition': {'duration': 300, 'easing': 'linear'}
                        }]
                    },
                    {
                        'label': 'Pause',
                        'method': 'animate',
                        'args': [[None], {
                            'frame': {'duration': 0, 'redraw': False},
                            'mode': 'immediate',
                            'transition': {'duration': 0}
                        }]
                    }
                ]
            }]
        )

        combined_fig.update_layout(annotations=[initial_date_annotation])

        with st.container():
            st.plotly_chart(combined_fig, use_container_width=True)

                        'method': 'animate',
                        'args': [[None], {
                            'frame': {'duration': 0, 'redraw': False},
                            'mode': 'immediate',
                            'transition': {'duration': 0}
                        }]
                    }
                ]
            }]
        )

        combined_fig.update_layout(annotations=[initial_date_annotation])

        with st.container():
            st.plotly_chart(combined_fig, use_container_width=True)
