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

# Function to create heatmap for individual rooms
def create_room_heatmap(room_data, room_name, x_label):
    fig = go.Figure()

    # Add heatmap trace
    fig.add_trace(go.Heatmap(
        z=[room_data.values],
        x=room_data.index,
        y=[room_name],
        colorscale='YlOrRd',
        showscale=False,
        hoverongaps=False,
        hovertemplate='Time: %{x}<br>Occupied: %{z}<extra></extra>'
    ))

    # Update layout
    fig.update_layout(
        title=f'{room_name} Occupancy',
        xaxis_title=x_label,
        height=200,
        xaxis=dict(
            tickmode='array',
            tickvals=room_data.index,
            ticktext=[f"{h:02d}:00" for h in range(9, 18)] if x_label == "Time of Day" else ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
            tickangle=45
        ),
        yaxis=dict(showticklabels=False)
    )

    return fig

# Function to create combined heatmap for all rooms
def create_combined_heatmap(all_room_data, x_label):
    fig = go.Figure()

    # Add heatmap trace
    fig.add_trace(go.Heatmap(
        z=all_room_data.T.values,
        x=all_room_data.index,
        y=all_room_data.columns,
        colorscale='YlOrRd',
        showscale=True,
        hoverongaps=False,
        hovertemplate='Room: %{y}<br>Time: %{x}<br>Occupied: %{z}<extra></extra>'
    ))

    # Update layout
    fig.update_layout(
        title='Combined Room Occupancy',
        xaxis_title=x_label,
        yaxis_title='Rooms',
        height=max(400, 50 * len(all_room_data.columns) + 100),  # Adjust height based on number of rooms
        xaxis=dict(
            tickmode='array',
            tickvals=all_room_data.index,
            ticktext=[f"{h:02d}:00" for h in range(9, 18)] if x_label == "Time of Day" else ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
            tickangle=45
        ),
        yaxis=dict(autorange="reversed"),  # To match the traditional heatmap orientation
        coloraxis_colorbar=dict(
            title='Occupancy',
            tickvals=[0, 1],
            ticktext=['Not Occupied', 'Occupied']
        )
    )

    return fig

# Determine number of columns based on layout option
def get_num_columns():
    if layout_option == "Analyse":
        return min(4, len(selected_rooms)) if len(selected_rooms) > 1 else 1
    else:
        return 1

# Daily Trends Tab
with tab1:
    st.markdown("<h3 style='color: #4CAF50;'>Daily Dashboard</h3>", unsafe_allow_html=True)

    if not selected_rooms:
        st.warning("Please select at least one room to view occupancy data.")
    elif selected_date not in available_dates:
        st.warning(f"No data available for {selected_date}")
    else:
        # Filter the DataFrame based on user input for daily trends
        daily_filtered_dfs = {}
        daily_capacities = {}
        
        # ... your existing filtering code ...

        # Safe DataFrame creation and visualization
        if daily_filtered_dfs:
            if any(not df.empty for df in daily_filtered_dfs.values()):
                try:
                    combined_daily_data = pd.DataFrame(daily_filtered_dfs)
                    combined_fig = create_combined_heatmap(combined_daily_data, "Time of Day")
                    st.plotly_chart(combined_fig, use_container_width=True)
                    
                    # Individual room heatmaps
                    for room, hourly_occupancy in daily_filtered_dfs.items():
                        if not hourly_occupancy.empty:
                            fig = create_room_heatmap(hourly_occupancy, room, "Time of Day")
                            st.plotly_chart(fig, use_container_width=True)
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

    # Check if selected week has data
    if selected_week_start_date not in available_weeks:
        st.warning(f"No data available for the week starting {selected_week_start_date}")
    else:
        # Filter the DataFrame based on the selected week
        weekly_filtered_dfs = {}
        weekly_capacities = {}
        for room in selected_rooms:
            filtered_df = df[
                (df['Space Name'] == room) &
                (df['Local Date'] >= selected_week_start_date) &
                (df['Local Date'] <= week_end_date)
            ]

            if filtered_df.empty:
                st.warning(f"No data available for {room} in the selected week.")
                continue

            filtered_df = filtered_df.between_time(start_time.strftime('%H:%M'),
                                                   end_time.strftime('%H:%M'))

            if filtered_df.empty:
                st.warning(f"No data available for {room} during office hours.")
                continue

            daily_occupancy = filtered_df.groupby(filtered_df.index.date)['People Presence'].max()
            daily_occupancy = daily_occupancy.to_frame()

            date_range = pd.date_range(start=selected_week_start_date, end=week_end_date, freq='D').date
            daily_occupancy = daily_occupancy.reindex(date_range, fill_value=0)

            if daily_occupancy.empty:
                st.warning(f"No occupancy data available for {room} after processing.")
                continue

            weekly_filtered_dfs[room] = daily_occupancy

            capacity_data = df.loc[df['Space Name'] == room, 'Space Capacity']
            if not capacity_data.empty:
                weekly_capacities[room] = capacity_data.values[0]
            else:
                st.warning(f"No capacity data available for {room}. Setting capacity to 0.")
                weekly_capacities[room] = 0

        # Proceed with plotting and analysis if there's data
        if weekly_filtered_dfs:
            # Create a date range for the selected week (Monday to Friday)
            date_range = pd.date_range(
                start=selected_week_start_date,
                end=week_end_date,
                freq='D'
            )
            date_labels = date_range.strftime('%A')  # Day names

            # Initialize variables
            avg_utilization = {}
            utilization_records_weekly = {}

            # Calculate utilization
            for room, daily_occupancy in weekly_filtered_dfs.items():
                if not daily_occupancy.empty:
                    utilization = (daily_occupancy['People Presence'].sum() / len(daily_occupancy)) * 100
                    avg_utilization[room] = utilization
                    utilization_records_weekly[room] = {'Room': room, 'Usage (%)': utilization}

            # Create combined heatmap and display it at the top
            combined_weekly_data = pd.DataFrame(weekly_filtered_dfs)
            combined_fig = create_combined_heatmap(combined_weekly_data, "Day of Week")
            st.plotly_chart(combined_fig, use_container_width=True)

            # Create individual room heatmaps
            for room, daily_occupancy in weekly_filtered_dfs.items():
                fig = create_room_heatmap(daily_occupancy, room, "Day of Week")
                st.plotly_chart(fig, use_container_width=True)

            # Display average weekly utilization
            utilization_text_weekly = "<div style='border: 2px solid #FF5722; padding: 10px; border-radius: 10px; background-color: #fff7f0;'>"
            utilization_text_weekly += "<h3 style='color: #FF5722;'>Average Weekly Utilization</h3>"

            for room, utilization in avg_utilization.items():
                utilization_text_weekly += f"<p style='font-size: 18px; font-weight: bold; color: #333;'>{room}: {utilization:.2f}% utilized</p>"

            utilization_text_weekly += "</div>"
            st.markdown(utilization_text_weekly, unsafe_allow_html=True)

            # Add download button for weekly utilization data with distinct label and key
            if utilization_records_weekly:
                df_utilization_weekly = pd.DataFrame(utilization_records_weekly.values())
                get_download_link(
                    df_utilization_weekly,
                    title="📄 Download Weekly Utilization Data",
                    filename="average_weekly_utilization.csv",
                    key='download_weekly'
                )

            # Style updates for the average weekly utilization box
            st.markdown("""
            <style>
            div[data-testid="stMarkdownContainer"] {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                font-size: 16px;
            }
            </style>
            """, unsafe_allow_html=True)
        else:
            st.warning("🚫 No data available after processing. Please adjust your filters.")
