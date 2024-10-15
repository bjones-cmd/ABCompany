import os
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Set Streamlit page configuration
st.set_page_config(page_title="Room Occupancy Dashboard", layout="wide")

# Apply custom CSS styles
st.markdown("""
    <style>
    .css-18e3th9 {
        position: -webkit-sticky;
        position: sticky;
        top: 0;
        background-color: white;
        z-index: 100;
        padding-top: 0;
    }
    .stApp {
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    html, body, [class*="css"]  {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .css-1lcbmhc.e1fqkh3o2 {
        width: 18rem;
        transition: width 0.3s;
    }
    .css-1lcbmhc.e1fqkh3o2:hover {
        width: 21rem;
    }
    </style>
    """, unsafe_allow_html=True)

# Function to load all Excel files from a directory
@st.cache_data
def load_all_data(directory):
    all_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.xlsx')]
    df_list = [pd.read_excel(file) for file in all_files]
    df = pd.concat(df_list, ignore_index=True)
    return df

# Load data from the Room Occupancy folder please
df = load_all_data('Room Occupancy')

# Ensure the data is in the correct format
# Assuming the CSVs have columns like 'Date', 'Hour', 'Minute', 'Floor', 'Room Name', 'Occupancy Count'
df['Hour'] = pd.to_numeric(df['Hour'], errors='coerce').astype('Int64')
df['Minute'] = pd.to_numeric(df['Minute'], errors='coerce').astype('Int64')
df['timestamp'] = pd.to_datetime(
    df['Date'].astype(str) + ' ' +
    df['Hour'].astype(str).str.zfill(2) + ':' +
    df['Minute'].astype(str).str.zfill(2),
    format='%Y-%m-%d %H:%M',
    errors='coerce'
)
df.set_index('timestamp', inplace=True)
df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d', errors='coerce').dt.date
df['Week Start'] = df['Date'] - pd.to_timedelta(pd.to_datetime(df['Date']).dt.dayofweek, unit='d')
df['Week Start'] = pd.to_datetime(df['Week Start']).dt.date

# Streamlit app
st.title("Room Occupancy Dashboard")

# Create tabs
tab1, tab2 = st.tabs(["Daily Trends", "Weekly Trends"])

# Sidebar for filters
st.sidebar.header("Filter Options")

# Get unique floors and rooms
floors = df['Floor'].dropna().unique()
rooms = df['Room Name'].dropna().unique()
selected_floors = st.sidebar.multiselect("Select Floors:", floors)
selected_rooms = st.sidebar.multiselect("Select Rooms:", rooms)

# Chart type selection
st.sidebar.markdown("### Chart Type Selector")
chart_type = st.sidebar.radio("Select Chart Type:", ["Bar", "Line", "Area", "Scatter"])

# Layout selection
st.sidebar.markdown("### Layout Option")
layout_option = st.sidebar.radio("Select Layout:", ["Focus", "Analyse"], index=0)

# Daily Trends Tab
with tab1:
    st.markdown("<h3 style='color: #4CAF50;'>Daily Dashboard</h3>", unsafe_allow_html=True)
    # Implement daily trends logic similar to app40.py

# Weekly Trends Tab
with tab2:
    st.markdown("<h3 style='color: #4CAF50;'>Weekly Dashboard</h3>", unsafe_allow_html=True)
    # Implement weekly trends logic similar to app40.py

