import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import os

# Try to import openpyxl, and if it fails, show a user-friendly message
try:
    import openpyxl
except ImportError:
    st.error("The openpyxl library is not installed. Please install it to read Excel files.")
    st.stop()

# Set Streamlit page configuration
st.set_page_config(page_title="ABC Company - Winnipeg Office Room Occupancy", layout="wide")

# Apply custom CSS styles
st.markdown("""
    <style>
    /* Make the title sticky */
    .css-18e3th9 {
        position: -webkit-sticky;
        position: sticky;
        top: 0;
        background-color: white;
        z-index: 100;
        padding-top: 0;
    }
    /* Add rounded corners and shadow */
    .stApp {
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    /* Customize font and colors */
    html, body, [class*="css"]  {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    /* Collapsible sidebar */
    .css-1lcbmhc.e1fqkh3o2 {
        width: 18rem;
        transition: width 0.3s;
    }
    .css-1lcbmhc.e1fqkh3o2:hover {
        width: 21rem;
    }
    </style>
    """, unsafe_allow_html=True)

# Function to convert DataFrame to CSV for download
def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# Modify the load_data function to handle potential file reading errors
@st.cache_data
def load_data(folder_path):
    all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    li = []

    for filename in all_files:
        try:
            df = pd.read_excel(filename, engine='openpyxl')
            li.append(df)
        except Exception as e:
            st.warning(f"Error reading file {filename}: {str(e)}")

    if not li:
        st.error("No valid Excel files could be read. Please check your data files.")
        st.stop()

    df = pd.concat(li, axis=0, ignore_index=True)

    df['Local Date'] = pd.to_datetime(df['Local Date'], format='%Y-%m-%d', errors='coerce')
    df['Local Time'] = pd.to_datetime(df['Local Time'], format='%H:%M:%S', errors='coerce').dt.time

    df['timestamp'] = pd.to_datetime(df['Local Date'].dt.strftime('%Y-%m-%d') + ' ' + df['Local Time'].astype(str), errors='coerce')

    if df['timestamp'].isnull().any():
        st.warning("⚠️ Some timestamps couldn't be parsed and are set to 'NaT'. Please check your Excel files for consistency.")

    df.set_index('timestamp', inplace=True)
    df['Local Date'] = df['Local Date'].dt.date

    if df['Local Date'].isnull().any():
        st.warning("⚠️ Some 'Local Date' entries couldn't be parsed and are set to 'NaT'. Please check your Excel files for consistency.")

    df['Week Start'] = df['Local Date'] - pd.to_timedelta(pd.to_datetime(df['Local Date']).dt.dayofweek, unit='d')
    df['Week Start'] = pd.to_datetime(df['Week Start']).dt.date

    # Ensure 'People Presence' is binary
    df['People Presence'] = df['People Presence'].astype(int)

    return df

# Function to get unique floors
def get_unique_floors(df):
    return sorted(df['Floor Name'].dropna().unique())

# Function to create combined heatmap for all rooms
def create_combined_heatmap(all_room_data, x_label):
    fig = go.Figure()

    # Add heatmap trace
    fig.add_trace(go.Heatmap(
        z=all_room_data.T.values,
        x=all_room_data.index,
        y=all_room_data.columns,
        colorscale='YlOrRd',
        showscale=False,
        hoverongaps=False,
        hovertemplate='Room: %{y}<br>Time: %{x}<br>Occupied: %{z}<extra></extra>'
    ))

    # Update layout
    fig.update_layout(
        title='Combined Room Occupancy',
        xaxis_title=x_label,
        yaxis_title='Rooms',
        height=max(400, 50 * len(all_room_data.columns) + 100),
        xaxis=dict(
            tickmode='array',
            tickvals=all_room_data.index,
            ticktext=[f"{h:02d}:00" for h in range(9, 18)] if x_label == "Time of Day" else ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
            tickangle=45
        ),
        yaxis=dict(autorange="reversed")
    )

    return fig

# Load the data from the 'Room Occupancy' folder
df = load_data('Room Occupancy')

# Streamlit app
st.title("🏢 ABC Company - Winnipeg Office Room Occupancy")

# Create tabs
tab1, tab2 = st.tabs(["Daily Trends", "Weekly Trends"])

# Sidebar for filters
st.sidebar.header("🏢 Floor and Room Selection")

# Get unique floors
floors = get_unique_floors(df)

# Use session state to store selected floors and rooms
if 'selected_floors' not in st.session_state:
    st.session_state.selected_floors = []
if 'selected_rooms' not in st.session_state:
    st.session_state.selected_rooms = []

# Floor selection
selected_floors = st.sidebar.multiselect("Select Floors:", floors, default=st.session_state.selected_floors)

# Update selected floors in session state
st.session_state.selected_floors = selected_floors

# Get unique room names ('Space Name' column) based on selected floors
if selected_floors:
    available_rooms = df[df['Floor Name'].isin(selected_floors)]['Space Name'].dropna().unique()
else:
    available_rooms = df['Space Name'].dropna().unique()

# Button to select all rooms on chosen floors
if st.sidebar.button("Select All Rooms on Chosen Floors"):
    st.session_state.selected_rooms = available_rooms.tolist()

# Room selection
selected_rooms = st.sidebar.multiselect(
    "Select rooms manually:",
    options=available_rooms,
    default=st.session_state.selected_rooms
)

# Update selected rooms in session state
st.session_state.selected_rooms = selected_rooms

# Layout selection
st.sidebar.markdown("### 🖥️ Layout Option")
layout_option = st.sidebar.radio("Select Layout:", ["Focus", "Analyse"], index=0)

# Subheader for daily filters
st.sidebar.subheader("📅 Daily Filters")

# Get available years, months, and days
available_dates = df['Local Date'].dropna().unique()
years = sorted(pd.to_datetime(available_dates).year.unique())
selected_year = st.sidebar.selectbox("Select Year:", years)

# Filter months based on selected year
months = sorted(pd.to_datetime(available_dates)[pd.to_datetime(available_dates).year == selected_year].month.unique())
month_names = [pd.Timestamp(month=month, day=1, year=selected_year).strftime('%B') for month in months]
month_dict = dict(zip(month_names, months))
selected_month_name = st.sidebar.selectbox("Select Month:", month_names)
selected_month = month_dict[selected_month_name]

# Filter days based on selected year and month
days = sorted(pd.to_datetime(available_dates)[
    (pd.to_datetime(available_dates).year == selected_year) &
    (pd.to_datetime(available_dates).month == selected_month)
].day.unique())
selected_day = st.sidebar.selectbox("Select Day:", days)
# Subheader for weekly filters
st.sidebar.subheader("📅 Weekly Filters")

# Get available weeks (as dates)
available_weeks = sorted(df['Week Start'].dropna().unique())
selected_week_start = st.sidebar.selectbox("Select Week Starting (Monday):", available_weeks)

# Office hours note
st.sidebar.markdown("⏰ **Note:** Office hours are defined as 9 AM to 5 PM.")

# Set office hours
start_time = pd.Timestamp("09:00").time()
end_time = pd.Timestamp("17:00").time()

# Function to generate download link with unique key
def get_download_link(df_utilization, title, filename, key):
    csv = convert_df_to_csv(df_utilization)
    return st.sidebar.download_button(
        label=title,
        data=csv,
        file_name=filename,
        mime='text/csv',
        key=key
    )

# Color mapping for rooms
qualitative_colors = px.colors.qualitative.Plotly
color_map = {room: qualitative_colors[i % len(qualitative_colors)] for i, room in enumerate(selected_rooms)}

# Daily Trends Tab
with tab1:
    st.markdown("<h3 style='color: #4CAF50;'>Daily Dashboard</h3>", unsafe_allow_html=True)

    # Create selected date
    selected_date = pd.Timestamp(year=selected_year, month=selected_month, day=selected_day).date()

    if not selected_rooms:
        st.warning("Please select at least one room to view occupancy data.")
    elif selected_date not in available_dates:
        st.warning(f"No data available for {selected_date}")
    else:
        # Filter the DataFrame based on user input for daily trends
        daily_filtered_dfs = {}
        daily_capacities = {}
        for room in selected_rooms:
            filtered_df = df[
                (df['Space Name'] == room) &
                (df['Local Date'] == selected_date)
            ]
            hourly_occupancy = filtered_df.resample('H')['People Presence'].max()
            hourly_occupancy = hourly_occupancy.between_time(start_time.strftime('%H:%M'),
                                                             end_time.strftime('%H:%M'))
            hourly_occupancy = hourly_occupancy.reindex(pd.date_range(
                start=pd.Timestamp.combine(selected_date, start_time),
                end=pd.Timestamp.combine(selected_date, end_time),
                freq='H'
            ), fill_value=0)
            daily_filtered_dfs[room] = hourly_occupancy
            capacity_data = df.loc[df['Space Name'] == room, 'Space Capacity']
            if not capacity_data.empty:
                daily_capacities[room] = capacity_data.values[0]
            else:
                st.warning(f"No capacity data available for {room}. Setting capacity to 0.")
                daily_capacities[room] = 0

        # Safe DataFrame creation and visualization
        if daily_filtered_dfs:
            if any(not df.empty for df in daily_filtered_dfs.values()):
                try:
                    # Create DataFrame with explicit index
                    combined_daily_data = pd.DataFrame(daily_filtered_dfs).fillna(0)
                    if not combined_daily_data.empty:
                        combined_fig = create_combined_heatmap(combined_daily_data, "Time of Day")
                        st.plotly_chart(combined_fig, use_container_width=True)

                        # Calculate and display utilization
                        avg_utilization = {}
                        utilization_records = {}
                        for room, hourly_occupancy in daily_filtered_dfs.items():
                            if not hourly_occupancy.empty:
                                utilization = (hourly_occupancy.sum() / len(hourly_occupancy)) * 100
                                avg_utilization[room] = utilization
                                utilization_records[room] = {'Room': room, 'Usage (%)': utilization}

                        # Display utilization text and download button
                        if avg_utilization:
                            utilization_text = "<div style='border: 2px solid #4CAF50; padding: 10px; border-radius: 10px; background-color: #f9f9f9;'>"
                            utilization_text += "<h3 style='color: #4CAF50;'>Average Daily Utilization</h3>"
                            for room, utilization in avg_utilization.items():
                                utilization_text += f"<p style='font-size: 18px; font-weight: bold; color: #333;'>{room}: {utilization:.2f}% utilized</p>"
                            utilization_text += "</div>"
                            st.markdown(utilization_text, unsafe_allow_html=True)

                            if utilization_records:
                                df_utilization = pd.DataFrame(utilization_records.values())
                                get_download_link(
                                    df_utilization,
                                    title="📄 Download Daily Utilization Data",
                                    filename="average_daily_utilization.csv",
                                    key='download_daily'
                                )
                    else:
                        st.warning("No data available for visualization after processing.")
                except ValueError as e:
                    st.warning("Unable to create visualizations. Please check if data is available for the selected criteria.")
                    st.error(f"Error details: {str(e)}")
            else:
                st.warning("No occupancy data found for the selected criteria.")
        else:
            st.warning("No data available for the selected filters.")
            # Weekly Trends Tab
with tab2:
    st.markdown("<h3 style='color: #4CAF50;'>Weekly Dashboard</h3>", unsafe_allow_html=True)
    st.write("This section displays weekly room occupancy trends.")

    # Ensure 'Week Start' is a date object
    df['Week Start'] = pd.to_datetime(df['Week Start']).dt.date

    # Convert 'selected_week_start' to date object
    selected_week_start_date = pd.to_datetime(selected_week_start).date()

    # Define week end date as Friday (Monday to Friday)
    week_end_date = selected_week_start_date + pd.Timedelta(days=4)

    if not selected_rooms:
        st.warning("Please select at least one room to view occupancy data.")
    elif selected_week_start_date not in available_weeks:
        st.warning(f"No data available for the week starting {selected_week_start_date}")
    else:
        # Create date range for the week
        date_range = pd.date_range(start=selected_week_start_date, end=week_end_date, freq='D')
        
        # Initialize DataFrame with date range as index
        combined_weekly_data = pd.DataFrame(index=date_range)
        
        # Fill data for each room
        for room in selected_rooms:
            filtered_df = df[
                (df['Space Name'] == room) &
                (df['Local Date'] >= selected_week_start_date) &
                (df['Local Date'] <= week_end_date)
            ]
            if not filtered_df.empty:
                daily_occupancy = filtered_df.groupby('Local Date')['People Presence'].max()
                combined_weekly_data[room] = daily_occupancy

        # Continue with visualization if data exists
        if not combined_weekly_data.empty:
            try:
                combined_fig = create_combined_heatmap(combined_weekly_data.fillna(0), "Day of Week")
                st.plotly_chart(combined_fig, use_container_width=True)

                # Calculate weekly utilization
                avg_utilization_weekly = {}
                utilization_records_weekly = {}
                
                for room in selected_rooms:
                    if room in combined_weekly_data.columns:
                        room_data = combined_weekly_data[room].fillna(0)
                        utilization = (room_data.sum() / len(room_data)) * 100
                        avg_utilization_weekly[room] = utilization
                        utilization_records_weekly[room] = {'Room': room, 'Usage (%)': utilization}

                # Display weekly utilization
                if avg_utilization_weekly:
                    utilization_text_weekly = "<div style='border: 2px solid #FF5722; padding: 10px; border-radius: 10px; background-color: #fff7f0;'>"
                    utilization_text_weekly += "<h3 style='color: #FF5722;'>Average Weekly Utilization</h3>"
                    
                    for room, utilization in avg_utilization_weekly.items():
                        utilization_text_weekly += f"<p style='font-size: 18px; font-weight: bold; color: #333;'>{room}: {utilization:.2f}% utilized</p>"
                    
                    utilization_text_weekly += "</div>"
                    st.markdown(utilization_text_weekly, unsafe_allow_html=True)

                    # Add download button for weekly utilization data
                    if utilization_records_weekly:
                        df_utilization_weekly = pd.DataFrame(utilization_records_weekly.values())
                        get_download_link(
                            df_utilization_weekly,
                            title="📄 Download Weekly Utilization Data",
                            filename="average_weekly_utilization.csv",
                            key='download_weekly'
                        )

            except ValueError as e:
                st.warning("Unable to create visualizations. Please check if data is available for the selected criteria.")
                st.error(f"Error details: {str(e)}")
        else:
            st.warning("No data available for the selected week and rooms.")

# Style updates for the utilization boxes
st.markdown("""
<style>
div[data-testid="stMarkdownContainer"] {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    font-size: 16px;
}
</style>
""", unsafe_allow_html=True)