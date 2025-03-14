import subprocess
import sys

# Function to install missing packages
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Try to import packages, and install them if not found
try:
    import streamlit as st
except ImportError:
    install("streamlit")
    import streamlit as st

try:
    import pandas as pd
except ImportError:
    install("pandas")
    import pandas as pd

try:
    import matplotlib.pyplot as plt
except ImportError:
    install("matplotlib")
    import matplotlib.pyplot as plt

try:
    from datetime import datetime
except ImportError:
    install("datetime")  # Note: `datetime` is built into Python, so this is just a fallback
    from datetime import datetime

try:
    import altair as alt
except ImportError:
    install("altair")
    import altair as alt

try:
    from streamlit_echarts import JsCode
except ImportError:
    install("streamlit-echarts")
    from streamlit_echarts import JsCode

try:
    from streamlit_echarts import st_echarts
except ImportError:
    install("streamlit-echarts")
    from streamlit_echarts import st_echarts

import plotly.express as px


# Set wide layout for Streamlit
st.set_page_config(layout="wide")

def get_data_for_chart(pivot_df):
    series_list = []
    color_list = [
    '#FF4500',  # OrangeRed (Vibrant Orange)
    '#FF6347',  # Tomato (Bright Red)
    '#32CD32',  # LimeGreen (Bright Green)
    '#1E90FF',  # DodgerBlue (Bright Blue)
    '#FFD700',  # Gold (Vibrant Yellow)
    '#D2691E',  # Chocolate (Rich Brown)
    '#8A2BE2',  # BlueViolet (Purple)
    '#FF1493',  # DeepPink (Hot Pink)
    '#00FA9A',  # MediumSpringGreen (Aqua Green)
    '#BA55D3',  # MediumOrchid (Bright Purple)
    '#FF4500',  # OrangeRed (Vibrant Orange)
    '#FF6347',  # Tomato (Bright Red)
    '#32CD32',  # LimeGreen (Bright Green)
    '#1E90FF',  # DodgerBlue (Bright Blue)
    '#FFD700',  # Gold (Vibrant Yellow)
    '#D2691E',  # Chocolate (Rich Brown)
    '#8A2BE2',  # BlueViolet (Purple)
    '#FF1493',  # DeepPink (Hot Pink)
    '#00FA9A',  # MediumSpringGreen (Aqua Green)
    '#BA55D3'  # MediumOrchid (Bright Purple)
]

    for num, col in enumerate(pivot_df.columns):
        if col != 'Process' and col != 'Month':
            dict = {
                            'name': f'Month {col}',
                            'type': 'bar',
                            'data': list(pivot_df[col]),
                            "barWidth": "20%",
                            "barGap": "40%",
                            'itemStyle': {
                                'color': color_list[num]
                            }
                        }
            
            series_list.append(dict)

    legend = {
                'data': [col for col in pivot_df.columns if col != 'Process' and col != 'Month'],
                'top': 'bottom'
            }
    
    option = {
            'title': {
                'text': 'Issues by Process and Month',
                'subtext': 'Clustered Bar Chart',
                'left': 'center',
                'textStyle': {
                    'color': 'white',
                    'fontSize': 20
                }
            },
            'tooltip': {
                'trigger': 'axis',
                'axisPointer': {
                    'type': 'shadow'
                }
            },
            'legend': legend,
            'xAxis': {
                'type': 'value',
                'name': 'Issue Count',
                'nameLocation': 'middle'
            },
            'yAxis': {
                'type': 'category',
                'data': list(pivot_df["Process"]),
                'name': 'Process',
                'interval': 5
            },
            'series': series_list
        }
    
    return option

def remove_POs(df_shift_all):
    pattern = "PO#|PO #|INC #|INC#|INC00|received ticket|Received ticket"
    df_filtered_bad = df_shift_all[df_shift_all['Issue'].str.contains(pattern, regex=True)]
    df_filtered_good = df_shift_all[~df_shift_all['Issue'].str.contains(pattern, regex=True)]

    return df_filtered_good, df_filtered_bad

def get_file_as_dataframe(filename):
    df = pd.read_excel(filename, sheet_name="End Of Shift Report")

    df.columns = df.iloc[0, :]
    df_process = df.iloc[1:, :]

    # Set Date/Month with previous information
    df_process["Date/Month"] = df_process["Date/Month"].fillna(method='ffill')

    df_process.columns = ["Date/Month","Pending Action","Shift1_Process","Shift1_Issue","Shift1_Action Taken","Shift2_Process","Shift2_Issue","Shift2_Action Taken","Shift3_Process","Shift3_Issue","Shift3_Action Taken","NaN"]

    df_process = df_process[["Date/Month","Shift1_Process","Shift1_Issue","Shift1_Action Taken","Shift2_Process","Shift2_Issue","Shift2_Action Taken","Shift3_Process","Shift3_Issue","Shift3_Action Taken"]]

    return df_process

def seperate_shift_df(df_process):
    df_shift1 = df_process[["Date/Month","Shift1_Process","Shift1_Issue","Shift1_Action Taken"]]
    df_shift1 = df_shift1.dropna(subset = ["Shift1_Process"])
    df_shift1.columns = ["Date/Month","Process","Issue","Action Taken"]
    df_shift1["Date/Month"] = pd.to_datetime(df_shift1["Date/Month"])

    df_shift2 = df_process[["Date/Month","Shift2_Process","Shift2_Issue","Shift2_Action Taken"]]
    df_shift2 = df_shift2.dropna(subset = ["Shift2_Process"])
    df_shift2.columns = ["Date/Month","Process","Issue","Action Taken"]
    df_shift2["Date/Month"] = pd.to_datetime(df_shift2["Date/Month"])

    df_shift3 = df_process[["Date/Month","Shift3_Process","Shift3_Issue","Shift3_Action Taken"]]
    df_shift3 = df_shift3.dropna(subset = ["Shift3_Process"])
    df_shift3.columns = ["Date/Month","Process","Issue","Action Taken"]
    df_shift3["Date/Month"] = pd.to_datetime(df_shift3["Date/Month"])

    df_shift_all = pd.concat([df_shift1, df_shift2, df_shift3], axis = 0)
    df_shift_all["Date/Month"] = pd.to_datetime(df_shift_all["Date/Month"])
    df_shift_all = df_shift_all.dropna()
    df_shift_all_good, df_shift_all_bad = remove_POs(df_shift_all)

    return df_shift1, df_shift2, df_shift3, df_shift_all_good, df_shift_all_bad

# Streamlit app title
st.title("End Of Shift Report Analysis")

# File uploader widget to upload a CSV file
uploaded_file = st.file_uploader("Upload your End Of Shift Report file", type=["xlsx"])


if uploaded_file is not None:
    # Read the uploaded CSV file into a pandas DataFrame    
    process_df = get_file_as_dataframe(uploaded_file)

    df_shift1, df_shift2, df_shift3, df_shift_all, df_shift_all_bad = seperate_shift_df(process_df)

    st.write(f"Removed {len(df_shift_all_bad)} Tickets from calculations, remaining: {len(df_shift_all)} issues")
    st.dataframe(df_shift_all_bad, use_container_width=True)

    df_shift_all["Date/Month"] = pd.to_datetime(df_shift_all["Date/Month"], errors='coerce')
    df_shift_all["Month-Year"] = df_shift_all["Date/Month"].dt.strftime('%Y-%m')
    
    comparing_months = df_shift_all.groupby(["Month-Year", "Process"]).count()
    comparing_months_final = comparing_months.reset_index()[["Month-Year", "Process", "Issue"]]

    pivot_df = comparing_months_final.pivot_table(index='Process', columns="Month-Year", values='Issue', aggfunc='sum', fill_value=0)

    pivot_df_final = pivot_df.reset_index()

    if 'Process' in df_shift_all.columns:

        start_date, end_date = st.date_input(
            "Choose a date range", 
            [df_shift_all["Date/Month"].min(), df_shift_all["Date/Month"].max()],
            min_value = df_shift_all["Date/Month"].min(),
            max_value= df_shift_all["Date/Month"].max()
        )

        col1, col2 = st.columns([0.5, 1.5])

        start_date_str = start_date.strftime("%Y-%m-%d")  # Correct format string
        end_date_str = end_date.strftime("%Y-%m-%d")

        process_counts = df_shift_all[(df_shift_all["Date/Month"] >= pd.to_datetime(start_date)) & 
    (df_shift_all["Date/Month"] <= pd.to_datetime(end_date))]

        process_counts_to_display = process_counts.groupby(by=["Process"]).size().reset_index(name='ProcessCount')

        # Display the counts in the app
        with col1:
            st.write("Count of each unique Process:")
            st.dataframe(process_counts_to_display[["Process", "ProcessCount"]], use_container_width=True)  # Display the DataFrame with renamed columns
            total_issues = sum(list(process_counts_to_display["ProcessCount"]))
            st.write(f"Total Issues: {total_issues}")
        # Plot the count of 'Process' values as a bar chart
        with col2: 
            st.write(f"Visualizing the count of each Process, {start_date_str} - {end_date_str}:")

            option = {
            "tooltip": {
                "trigger": 'axis',
                "axisPointer": {      
                "type": 'shadow'      
                }
            },
            "xAxis": {
                "type": 'category',
                "data": list(process_counts_to_display.Process),
                "axisLabel": {
                "rotate": 90 
                }
            },
            "yAxis": {
                "type": 'value'
            },
            "series": [
                {
                "data": list(process_counts_to_display.ProcessCount),
                "type": 'bar',
                'itemStyle': {
                'color': 'red'  # Set the color of the line to red
            }
                }
            ]}

            
            clicked_label = st_echarts(option,
            height = "500px",
            events = {"click": "function(params) {return params.name}"})

    else:
        st.error("'Process' column not found in the uploaded data")

    if clicked_label != None:
        clicked_process = process_counts[process_counts["Process"] == clicked_label][["Date/Month","Process","Issue","Action Taken"]]
    else:
        clicked_process = process_counts[["Date/Month","Process","Issue","Action Taken"]]

    st.write("____")

    st.write(f"{clicked_label}")

    process_counts_to_display = clicked_process[["Date/Month"]].groupby(by=["Date/Month"]).size().reset_index(name='ProcessCount')

    # Get the start and end date of each month in 'Date/Month'
    start_of_month = clicked_process['Date/Month'].dt.to_period('M').dt.start_time
    end_of_month = clicked_process['Date/Month'].dt.to_period('M').dt.end_time
    
    # Generate a list of all the unique months in the data
    months = pd.date_range(start=start_of_month.min(), end=end_of_month.max(), freq='M')
    
    # Create an empty DataFrame to hold all date ranges for the months
    all_month_dates = pd.DataFrame()
    
    # For each month, create a date range from the start to the end of the month
    for month in months:
        month_start = month.replace(day=1)
        month_end = month.replace(day=pd.Timestamp(month).days_in_month)
        month_range = pd.date_range(month_start, month_end, freq='D')
        all_month_dates = pd.concat([all_month_dates, pd.DataFrame(month_range, columns=['Date/Month'])], ignore_index=True)
    
    # Merge with the original counts DataFrame
    result = pd.merge(all_month_dates, process_counts_to_display, on='Date/Month', how='left')
    
    # Fill NaN values in 'ProcessCount' with 0
    result['ProcessCount'].fillna(0, inplace=True)

    option = {
                "tooltip": {
                    "trigger": 'axis',
                    "axisPointer": {      
                    "type": 'shadow'      
                    }
                },
        'xAxis': {
            'type': 'category',
            'data': list(result["Date/Month"].dt.strftime('%Y-%m-%dT%H:%M:%S'))
        },
        'yAxis': {
            'type': 'value'
        },
        'series': [
            {
                'data': list(result["ProcessCount"]),
                'type': 'line',
                'itemStyle': {
                'color': 'red'  # Set the color of the line to red
            }
            }
        ]
    }

    
    secondary_clicked_label = st_echarts(option,
        height = "300px",
        events = {"click": "function(params) {return params.name}"})

    if secondary_clicked_label != None:
        seconday_clicked_process = process_counts[(process_counts["Date/Month"] == secondary_clicked_label) & (process_counts["Process"] == clicked_label)][["Date/Month","Process","Issue","Action Taken"]].sort_values(by = "Date/Month")
        st.dataframe(seconday_clicked_process, use_container_width = True)
    else:
        seconday_clicked_process = process_counts[(process_counts["Process"] == clicked_label)][["Date/Month","Process","Issue","Action Taken"]].sort_values(by = "Date/Month")
        st.dataframe(seconday_clicked_process, use_container_width = True)

    compare_month_col, compare_month_col2 = st.columns([1.5, 1])

    with compare_month_col:
        compare_month_option = get_data_for_chart(pivot_df_final)
        st_echarts(compare_month_option,
                height = "1200px",
                events = {"click": "function(params) {return params.name}"})
        
    with compare_month_col2:
        process_counts_pivot = process_counts.groupby(by = ["Process"]).size().reset_index(name='ProcessCount')

        fig = px.pie(
            process_counts_pivot, 
            names="Process", 
            values="ProcessCount", 
            title="Process Issues", 
            color_discrete_sequence=px.colors.qualitative.Set3,  # Use a predefined color set
            hover_data={"ProcessCount": True}  # Show values on hover
        )

        # Set dark theme for Streamlit background
        fig.update_layout(
            paper_bgcolor="#0E1117",  # Dark background for Streamlit
            font=dict(color="white"),  # Set text color to white
            width=900,  # Set width
            height=700,  # Set height
            title_x=0.5  # Center the title
        )

        # Display the chart in Streamlit
        st.plotly_chart(fig, use_container_width=True)

        
    

