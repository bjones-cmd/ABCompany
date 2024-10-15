import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import os

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

# Load your data
@st.cache_data
def load_data(folder_path):
    all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    li = []

    for filename in all_files:
        df = pd.read_excel(filename)
        li.append(df)

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

    return df

# Load the data from the 'Room Occupancy' folder
df = load_data('Room Occupancy')

# Streamlit app
st.title("🏢 ABC Company - Winnipeg Office Room Occupancy")

# Create tabs
tab1, tab2 = st.tabs(["Daily Trends", "Weekly Trends"])

# Sidebar for filters
st.sidebar.header("🚪 Room Selection")

# Get unique room names ('Space Name' column)
rooms = df['Space Name'].dropna().unique()
selected_rooms = st.sidebar.multiselect("Select Rooms:", rooms)

# Chart type selection
st.sidebar.markdown("### 📊 Chart Type Selector")
chart_type = st.sidebar.radio("Select Chart Type:", ["Bar", "Line", "Area", "Scatter"])

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

# Function to create individual plots
def create_individual_plots(filtered_dfs, time_labels, y_limit_upper, plot_type, x_label):
    plots = []
    for room, filtered_df in filtered_dfs.items():
        if not filtered_df.empty:
            # Reset index and rename columns for plotting
            filtered_df = filtered_df.reset_index()
            # Rename 'timestamp' to 'Time' if it exists
            if 'timestamp' in filtered_df.columns:
                filtered_df.rename(columns={'timestamp': 'Time'}, inplace=True)
            else:
                filtered_df.rename(columns={'index': 'Time'}, inplace=True)

            # Ensure 'Time' is in string format '%H:%M' or '%A'
            if isinstance(filtered_df['Time'].iloc[0], pd.Timestamp):
                if x_label == "Time":
                    filtered_df['Time'] = filtered_df['Time'].dt.strftime('%H:%M')
                else:
                    filtered_df['Time'] = filtered_df['Time'].dt.strftime('%A')
            else:
                # Handle cases where 'Time' might already be a string but ensure proper formatting
                filtered_df['Time'] = pd.to_datetime(filtered_df['Time'], errors='coerce').dt.strftime('%H:%M')

            filtered_df = filtered_df.dropna(subset=['Time'])

            color = color_map[room]
            if plot_type == "Bar":
                fig = px.bar(filtered_df, x='Time', y='Peak People Count',
                             title=f"{room}",
                             labels={"Time": x_label, "Peak People Count": "Number of People"},
                             template='plotly_white',
                             color_discrete_sequence=[color])
            elif plot_type == "Line":
                fig = px.line(filtered_df, x='Time', y='Peak People Count',
                              title=f"{room}",
                              labels={"Time": x_label, "Peak People Count": "Number of People"},
                              markers=True, template='plotly_white',
                              color_discrete_sequence=[color])
            elif plot_type == "Area":
                fig = px.area(filtered_df, x='Time', y='Peak People Count',
                              title=f"{room}",
                              labels={"Time": x_label, "Peak People Count": "Number of People"},
                              template='plotly_white',
                              color_discrete_sequence=[color])
            elif plot_type == "Scatter":
                fig = px.scatter(filtered_df, x='Time', y='Peak People Count',
                                 title=f"{room}",
                                 labels={"Time": x_label, "Peak People Count": "Number of People"},
                                 template='plotly_white',
                                 color_discrete_sequence=[color])
            else:
                continue  # Skip if an unknown plot_type is passed

            fig.update_layout(
                xaxis=dict(tickmode='array', tickvals=time_labels, ticktext=time_labels),
                yaxis=dict(range=[0, y_limit_upper]),
                title_x=0.5,
                hovermode='x unified'
            )
            fig.update_xaxes(tickangle=45)
            plots.append(fig)
    return plots

# Determine number of columns based on layout option
def get_num_columns():
    if layout_option == "Analyse":
        return min(4, len(selected_rooms)) if len(selected_rooms) > 1 else 1
    else:
        return 1

# Daily Trends Tab
with tab1:
    st.markdown("<h3 style='color: #4CAF50;'>Daily Dashboard</h3>", unsafe_allow_html=True)

    # Create selected date
    selected_date = pd.Timestamp(year=selected_year, month=selected_month, day=selected_day).date()

    # Check if selected date is available
    if selected_date not in available_dates:
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
            # Resample and process data
            hourly_occupancy = filtered_df.resample('H').max()
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

        # Check if any filtered_df is empty
        if all(filtered_df.empty for filtered_df in daily_filtered_dfs.values()):
            st.warning("🚫 No data available for the selected filters.")
        else:
            # Create a common set of timestamps for the selected date and time range
            time_range = pd.date_range(
                start=pd.Timestamp.combine(selected_date, start_time),
                end=pd.Timestamp.combine(selected_date, end_time),
                freq='H'
            )
            time_labels = time_range.strftime('%H:%M')

            # Initialize variables
            max_users = 0
            avg_utilization = {}
            utilization_records = {}

            # Calculate the maximum number of users and average utilization
            for room, hourly_occupancy in daily_filtered_dfs.items():
                if not hourly_occupancy.empty:
                    room_max = hourly_occupancy['Peak People Count'].max()
                    max_users = max(max_users, room_max)
                    avg_users = hourly_occupancy['Peak People Count'].mean()
                    capacity = daily_capacities[room]
                    if capacity > 0:
                        utilization = (avg_users / capacity) * 100
                        avg_utilization[room] = utilization
                        utilization_records[room] = {'Room': room, 'Average Utilization (%)': utilization}
                    else:
                        avg_utilization[room] = 0
                        st.warning(f"Capacity for {room} is zero. Utilization set to 0%.")

            # Round up the y-axis limit to the nearest ten
            y_limit_upper = (int(max_users) // 10 + 1) * 10

            # Create individual plots
            plots = create_individual_plots(daily_filtered_dfs, time_labels, y_limit_upper, chart_type, "Time")

            # Display plots based on layout option
            num_cols = get_num_columns()
            if num_cols > 1:
                cols = st.columns(num_cols)
                for idx, fig in enumerate(plots):
                    with cols[idx % num_cols]:
                        st.plotly_chart(fig, use_container_width=True)
            else:
                for fig in plots:
                    st.plotly_chart(fig, use_container_width=True)

            # Create combined Plotly graph
            combined_fig = go.Figure()

            for room, hourly_occupancy in daily_filtered_dfs.items():
                hourly_occupancy = hourly_occupancy.reset_index()
                hourly_occupancy.rename(columns={'index': 'Time'}, inplace=True)

                # Ensure 'Time' is formatted correctly
                hourly_occupancy['Time'] = hourly_occupancy['Time'].dt.strftime('%H:%M')

                color = color_map[room]
                if chart_type == "Bar":
                    combined_fig.add_trace(go.Bar(
                        x=hourly_occupancy['Time'],
                        y=hourly_occupancy['Peak People Count'],
                        name=room,
                        marker_color=color,
                        opacity=0.6
                    ))
                elif chart_type == "Line":
                    combined_fig.add_trace(go.Scatter(
                        x=hourly_occupancy['Time'],
                        y=hourly_occupancy['Peak People Count'],
                        mode='lines+markers',
                        name=room,
                        marker=dict(color=color),
                        line=dict(width=2)
                    ))
                elif chart_type == "Area":
                    combined_fig.add_trace(go.Scatter(
                        x=hourly_occupancy['Time'],
                        y=hourly_occupancy['Peak People Count'],
                        mode='lines',
                        name=room,
                        fill='tozeroy',
                        marker=dict(color=color),
                        line=dict(width=2)
                    ))
                elif chart_type == "Scatter":
                    combined_fig.add_trace(go.Scatter(
                        x=hourly_occupancy['Time'],
                        y=hourly_occupancy['Peak People Count'],
                        mode='markers',
                        name=room,
                        marker=dict(color=color, size=8)
                    ))

            combined_fig.update_layout(
                title="Combined Room Occupancy",
                xaxis_title="Time",
                yaxis_title="Number of People",
                xaxis=dict(tickmode='array', tickvals=time_labels, ticktext=time_labels),
                barmode='group' if chart_type == "Bar" else None,
                yaxis=dict(range=[0, y_limit_upper]),
                legend_title="Rooms",
                template='plotly_white',
                hovermode='x unified'
            )
            combined_fig.update_xaxes(tickangle=45)
            st.plotly_chart(combined_fig, use_container_width=True)

            # Display average daily utilization
            utilization_text = "<div style='border: 2px solid #4CAF50; padding: 10px; border-radius: 10px; background-color: #f9f9f9;'>"
            utilization_text += "<h3 style='color: #4CAF50;'>Average Daily Utilization</h3>"

            for room, utilization in avg_utilization.items():
                utilization_text += f"<p style='font-size: 18px; font-weight: bold; color: #333;'>{room}: {utilization:.2f}% utilized</p>"

            utilization_text += "</div>"
            st.markdown(utilization_text, unsafe_allow_html=True)

            # Add download button for utilization data with distinct labels and keys
            if utilization_records:
                df_utilization = pd.DataFrame(utilization_records.values())
                get_download_link(
                    df_utilization,
                    title="📄 Download Daily Utilization Data",
                    filename="average_daily_utilization.csv",
                    key='download_daily'
                )

            # Style updates for the average daily utilization box
            st.markdown("""
            <style>
            div[data-testid="stMarkdownContainer"] {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                font-size: 16px;
            }
            </style>
            """, unsafe_allow_html=True)

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

            daily_occupancy = filtered_df.groupby(filtered_df.index.date)['Peak People Count'].max()
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
            max_users = 0
            avg_utilization = {}
            utilization_records_weekly = {}

            # Calculate the maximum number of users and average utilization
            for room, daily_occupancy in weekly_filtered_dfs.items():
                if not daily_occupancy.empty:
                    room_max = daily_occupancy['Peak People Count'].max()
                    max_users = max(max_users, room_max)
                    avg_users = daily_occupancy['Peak People Count'].mean()
                    capacity = weekly_capacities[room]
                    if capacity > 0:
                        utilization = (avg_users / capacity) * 100
                        avg_utilization[room] = utilization
                        utilization_records_weekly[room] = {'Room': room, 'Average Utilization (%)': utilization}
                    else:
                        avg_utilization[room] = 0
                        st.warning(f"Capacity for {room} is zero. Utilization set to 0%.")

            # Round up the y-axis limit to the nearest ten
            y_limit_upper = (int(max_users) // 10 + 1) * 10

            # Create individual plots
            plots_weekly = create_individual_plots(weekly_filtered_dfs, date_labels, y_limit_upper, chart_type, "Day")

            # Display plots based on layout option
            num_cols = get_num_columns()
            if num_cols > 1:
                cols = st.columns(num_cols)
                for idx, fig in enumerate(plots_weekly):
                    with cols[idx % num_cols]:
                        st.plotly_chart(fig, use_container_width=True)
            else:
                for fig in plots_weekly:
                    st.plotly_chart(fig, use_container_width=True)

            # Create combined Plotly graph
            combined_fig_weekly = go.Figure()

            for room, daily_occupancy in weekly_filtered_dfs.items():
                daily_occupancy = daily_occupancy.reset_index()
                daily_occupancy.rename(columns={'index': 'Day'}, inplace=True)

                # Ensure 'Day' is formatted correctly
                daily_occupancy['Day'] = pd.to_datetime(daily_occupancy['Day']).dt.strftime('%A')

                color = color_map[room]
                if chart_type == "Bar":
                    combined_fig_weekly.add_trace(go.Bar(
                        x=daily_occupancy['Day'],
                        y=daily_occupancy['Peak People Count'],
                        name=room,
                        marker_color=color,
                        opacity=0.6
                    ))
                elif chart_type == "Line":
                    combined_fig_weekly.add_trace(go.Scatter(
                        x=daily_occupancy['Day'],
                        y=daily_occupancy['Peak People Count'],
                        mode='lines+markers',
                        name=room,
                        marker=dict(color=color),
                        line=dict(width=2)
                    ))
                elif chart_type == "Area":
                    combined_fig_weekly.add_trace(go.Scatter(
                        x=daily_occupancy['Day'],
                        y=daily_occupancy['Peak People Count'],
                        mode='lines',
                        name=room,
                        fill='tozeroy',
                        marker=dict(color=color),
                        line=dict(width=2)
                    ))
                elif chart_type == "Scatter":
                    combined_fig_weekly.add_trace(go.Scatter(
                        x=daily_occupancy['Day'],
                        y=daily_occupancy['Peak People Count'],
                        mode='markers',
                        name=room,
                        marker=dict(color=color, size=8)
                    ))

            combined_fig_weekly.update_layout(
                title="Combined Weekly Room Occupancy",
                xaxis_title="Day",
                yaxis_title="Number of People",
                xaxis=dict(tickmode='array', tickvals=date_labels, ticktext=date_labels),
                barmode='group' if chart_type == "Bar" else None,
                yaxis=dict(range=[0, y_limit_upper]),
                legend_title="Rooms",
                template='plotly_white',
                hovermode='x unified'
            )
            combined_fig_weekly.update_xaxes(tickangle=45)
            st.plotly_chart(combined_fig_weekly, use_container_width=True)

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