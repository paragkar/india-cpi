import pandas as pd
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import io
import msoffcrypto

metric_type = ["Index", "Inflation", "WA"]

pd.set_option('display.max_columns', None)

st.set_page_config(
	# page_title="Your App Title",
	layout="wide",
	initial_sidebar_state="expanded"
)


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
@st.cache_resource
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

# Main Program Starts Here
df = loadfile()

# Ensuring the Date column is of datetime type
df['Date'] = pd.to_datetime(df['Date'])

# Sorting dataframe by Date to ensure proper animation sequence
df = df.sort_values(by='Date')

# Convert Date column to string without time
df['Date_str'] = df['Date'].dt.strftime('31st Mar %Y')

# Format the Value column to two decimal places and keep it as a float
df['Value'] = df['Value'].astype(float).round(2)

# # Create a column to hold the value information along with the year
df['Text'] = df.apply(lambda row: f"<b>{row['Value']:.2f} ({row['Date_str'][-4:]})</b>", axis=1)

selected_metric_type = st.sidebar.selectbox("Select Metric Type", metric_type)

# Check if any metric types are selected
if selected_metric_type:
	# Further filter dataframe based on selected metrics
	df = df[df['Metric'].str.contains(selected_metric_type)]

df = df.replace("", np.nan).dropna()

# Sidebar for main category selection
# selected_main_cat = st.sidebar.selectbox("Select MainCat", df['MainCat'].unique())

# Sidebar for sub category selection
# selected_sub_cat = st.sidebar.selectbox("Select SubCat", df['SubCat'].unique())

# Define the order for main category
main_cat_order = ['General Index', 'Food and beverages', 
'Pan, tobacco and intoxicants','Clothing and footwear', 'Housing', 'Fuel and light', 'Miscellaneous']
 

# state_order = ["Revenue Deficit","Gross Fiscal Deficit", "Primary Deficit",
# 			   "Primary Revenue Deficit", "Conventional Deficit", "Aggregrate Disburse", "Revenue Receipt",
# 			   "Tax Receipts", "Non Tax Receipts", "Aggregrate Receipts"]

# # Update title based on selection
# if selected_type == "Center":
# 	title_text = "SELECT FISCAL INDICATORS OF THE CENTER GOVT (% of gdp)"
# 	metric_order = center_order
# else:
# 	title_text = "SELECT FISCAL INDICATORS OF THE STATE GOVTs (% of gdp)"
# 	metric_order = state_order

# st.markdown(f"<h1 style='font-size:25px; margin-top: -60px;'>{title_text.title()}</h1>", unsafe_allow_html=True)

# filtered_main_cat_df = df[df['MainCat'] == selected_main_cat]


# # Set the main cat order for the y-axis
# df['MainCat'] = pd.Categorical(df['MainCat'], categories=main_cat_order, ordered=True)
# df = df.sort_values('MainCat')

selected_main_cat = st.sidebar.multiselect("Select Main Categories to Display", df['MainCat'].unique(), default=list(df['MainCat'].unique()))

# Check if any main categories are selected
if selected_main_cat:
	# Further filter dataframe based on selected metrics
	filtered_df = df[df['MainCat'].isin(selected_main_cat)]

selected_sub_cat = st.sidebar.multiselect("Select Sub Categories to Display", filtered_df['SubCat'].unique(), default=list(filtered_df['SubCat'].unique()))


# Check if any sub categories are selected
if selected_sub_cat:
	# Further filter dataframe based on selected metrics
	filtered_df = filtered_df[filtered_df['SubCat'].isin(selected_sub_cat)]



st.write(filtered_df)

# 	# Calculate min and max values for the dotted lines
# 	min_value = filtered_df['Value'].min()
# 	max_value = filtered_df['Value'].max()

# 	# Ensure Date_str is ordered correctly
# 	filtered_df['Date_str'] = pd.Categorical(filtered_df['Date_str'], ordered=True, categories=sorted(filtered_df['Date_str'].unique(), key=lambda x: datetime.strptime(x, '31st Mar %Y')))

# 	# Calculate the range for the x-axis
# 	range_min = min_value - abs(min_value) * 0.30
# 	range_max = max_value + abs(max_value) * 0.15

# 	# Plotly animation setup
# 	fig = px.scatter(filtered_df, x="Value", y="Metric", animation_frame="Date_str", animation_group="Metric",
# 					 color="Metric", range_x=[range_min, range_max],
# 					 title="", size_max=24, text="Text")

# 	# Customize text position to the right of the dots
# 	fig.update_traces(textposition='middle right', textfont=dict(size=16))

# 	# Add black outlines to the dots
# 	fig.update_traces(marker=dict(size=20, line=dict(width=2, color='black')))

# 	# Customize y-axis labels font size and make them bold
# 	fig.update_yaxes(tickfont=dict(size=20, color='black', family='Arial', weight='bold'))


# 	# Remove y-axis labels and variable labels
# 	fig.update_yaxes(showticklabels=True)
# 	fig.update_traces(marker=dict(size=24))

# 	# Draw a black line on the y-axis
# 	fig.add_shape(type='line', x0=0, x1=0, y0=0, y1=1, line=dict(color='black', width=1), xref='x', yref='paper')

# 	# Remove legend on the right side
# 	fig.update_layout(showlegend=False)

# 	# Add black border to the chart
# 	# fig.update_xaxes(fixedrange=True, showline=True, linewidth=1.2, linecolor='black', mirror=True)
# 	# fig.update_yaxes(fixedrange=True, showline=True, linewidth=1.2, linecolor='black', mirror=True)

# 	# Add dotted lines for min and max values
# 	fig.add_shape(
# 		type="line",
# 		x0=min_value, y0=0, x1=min_value, y1=1,
# 		xref='x', yref='paper',
# 		line=dict(color="blue", width=2, dash="dot")
# 	)
# 	fig.add_shape(
# 		type="line",
# 		x0=max_value, y0=0, x1=max_value, y1=1,
# 		xref='x', yref='paper',
# 		line=dict(color="red", width=2, dash="dot")
# 	)

# 	# Adjust the layout
# 	fig.update_layout(
# 		xaxis_title="Value as Percentage of GDP",
# 		yaxis_title="",
# 		width =1200,
# 		height=900,  # Adjust the height to make the plot more visible
# 		margin=dict(l=0, r=10, t=120, b=40, pad=0),  # Add margins to make the plot more readable and closer to the left
# 		sliders=[{
# 			'steps': [
# 				{
# 					'args': [
# 						[date_str],
# 						{
# 							'frame': {'duration': 300, 'redraw': True},
# 							'mode': 'immediate',
# 							'transition': {'duration': 300}
# 						}
# 					],
# 					'label': date_str,
# 					'method': 'animate'
# 				}
# 				for date_str in sorted(filtered_df['Date_str'].unique(), key=lambda x: datetime.strptime(x, '31st Mar %Y'))
# 			],
# 			'x': 0.1,
# 			'xanchor': 'left',
# 			'y': 0,
# 			'yanchor': 'top'
# 		}]
# 	)

# 	# Add initial annotation for the date
# 	initial_date_annotation = {
# 		'x': 0,
# 		'y': 1.15,  # Move the date annotation closer to the top of the chart
# 		'xref': 'paper',
# 		'yref': 'paper',
# 		'text': f'<span style="color:red;font-size:30px"><b>Date: {filtered_df["Date_str"].iloc[0]}</b></span>',
# 		'showarrow': False,
# 		'font': {
# 			'size': 20
# 		}
# 	}
# 	fig.update_layout(annotations=[initial_date_annotation])

# 	# Custom callback to update the date annotation dynamically
# 	def update_annotations(date_str):
# 		return [go.layout.Annotation(
# 			x=0,
# 			y=1.15,
# 			xref='paper',
# 			yref='paper',
# 			text=f'<span style="color:red;font-size:30px"><b>Date: {date_str}</b></span>',
# 			showarrow=False,
# 			font=dict(size=30)
# 		)]

# 	# Customize y-axis labels font size
# 	fig.update_yaxes(tickfont=dict(size=20))

# 	# Update annotation with each frame
# 	for frame in fig.frames:
# 		date_str = frame.name
# 		frame['layout'].update(annotations=update_annotations(date_str))

# 	# Ensure the frames are sorted correctly
# 	fig.frames = sorted(fig.frames, key=lambda frame: datetime.strptime(frame.name, '31st Mar %Y'))

# 	# Custom callback to update the date annotation dynamically
# 	fig.update_layout(
# 		updatemenus=[{
# 			'type': 'buttons',
# 			'showactive': False,
# 			'buttons': [
# 				{
# 					'label': 'Play',
# 					'method': 'animate',
# 					'args': [None, {
# 						'frame': {'duration': 500, 'redraw': True},
# 						'fromcurrent': True,
# 						'transition': {'duration': 300, 'easing': 'linear'}
# 					}]
# 				},
# 				{
# 					'label': 'Pause',
# 					'method': 'animate',
# 					'args': [[None], {
# 						'frame': {'duration': 0, 'redraw': False},
# 						'mode': 'immediate',
# 						'transition': {'duration': 0}
# 					}]
# 				}
# 			]
# 		}]
# 	)

# 	# Use Streamlit's container to fit the chart properly
# 	with st.container():
# 		st.plotly_chart(fig, use_container_width=True)
# else:
# 	st.write("Please select at least one metric to display the chart.")
