import streamlit as st
import pandas as pd
import os
import glob
import tkinter as tk
from tkinter import filedialog

# Set page config
st.set_page_config(
    page_title="Google Trends Data Processor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }
    .sub-header {
        font-size: 1.5rem;
        color: #ff7f0e;
        margin: 1rem 0;
        border-bottom: 2px solid #ff7f0e;
        padding-bottom: 0.5rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        margin: 1rem 0;
    }
    .error-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        margin: 1rem 0;
    }
    .stButton > button {
        background-color: #1f77b4;
        color: white;
        border-radius: 0.5rem;
        border: none;
        padding: 0.5rem 1rem;
        font-weight: bold;
        transition: all 0.3s;
    }
    .stButton > button:hover {
        background-color: #0d5a8a;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'base_path' not in st.session_state:
    st.session_state.base_path = ""
if 'output_path' not in st.session_state:
    st.session_state.output_path = ""
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = {}

# Main title
st.markdown('<h1 class="main-header">📊 Google Trends Data Processor</h1>', unsafe_allow_html=True)

# Folder selection functions
def select_folder(title, key):
    """Function to select folder using tkinter"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    root.wm_attributes('-topmost', 1)  # Bring to front
    folder_path = filedialog.askdirectory(title=title)
    root.destroy()
    return folder_path

# Sidebar for folder selection
with st.sidebar:
    st.markdown('<h2 class="sub-header">🗂️ Folder Selection</h2>', unsafe_allow_html=True)
    
    # Upload folder selection
    col1, col2 = st.columns([3, 1])
    with col1:
        st.text_input("Google Trends Data Folder:", value=st.session_state.base_path, key="base_path_display", disabled=True)
    with col2:
        if st.button("📁 Browse", key="select_base"):
            selected_path = select_folder("Select Google Trends Data Folder", "base_path")
            if selected_path:
                st.session_state.base_path = selected_path
                st.rerun()
    
    # Output folder selection
    col1, col2 = st.columns([3, 1])
    with col1:
        st.text_input("Output Folder:", value=st.session_state.output_path, key="output_path_display", disabled=True)
    with col2:
        if st.button("💾 Browse", key="select_output"):
            selected_path = select_folder("Select Output Folder", "output_path")
            if selected_path:
                st.session_state.output_path = selected_path
                st.rerun()
    
    st.markdown("---")
    
    # Process buttons
    st.markdown('<h3 class="sub-header">⚙️ Process Data</h3>', unsafe_allow_html=True)
    
    process_timeline = st.button("🕒 Process Timeline Data", use_container_width=True)
    process_geomap = st.button("🗺️ Process GeoMap Data", use_container_width=True)
    
    st.markdown("---")
    
    # Download buttons
    st.markdown('<h3 class="sub-header">💾 Download Files</h3>', unsafe_allow_html=True)

# Helper functions
def process_timeline_data(base_path):
    """Process timeline CSV files"""
    try:
        results = {}
        
        # Process Web folder
        web_file = os.path.join(base_path, "Web", "multiTimeline.csv")
        if os.path.exists(web_file):
            df_web = pd.read_csv(web_file, skiprows=2)
            df_web.columns = [col.replace(": (Sri Lanka)", "") for col in df_web.columns]
            df_web['Platform'] = "Web"
            results['web'] = df_web
        
        # Process YouTube folder
        youtube_file = os.path.join(base_path, "Youtube", "multiTimeline.csv")
        if os.path.exists(youtube_file):
            df_youtube = pd.read_csv(youtube_file, skiprows=2)
            df_youtube.columns = [col.replace(": (Sri Lanka)", "") for col in df_youtube.columns]
            df_youtube['Platform'] = "Youtube"
            results['youtube'] = df_youtube
        
        # Merge data
        if 'web' in results and 'youtube' in results:
            df_merged = pd.concat([results['web'], results['youtube']], ignore_index=True)
            results['merged'] = df_merged
        
        return results, None
    except Exception as e:
        return None, str(e)

def process_geomap_files(folder_name, platform_name, base_path):
    """Process geoMap CSV files"""
    folder_path = os.path.join(base_path, folder_name)
    geomap_files = glob.glob(os.path.join(folder_path, "geoMap*.csv"))
    
    df_city = None
    df_region = None
    
    for file_path in geomap_files:
        df = pd.read_csv(file_path, skiprows=2)
        first_column = df.columns[0]
        
        if first_column == "City":
            df['Breakdown'] = "City"
            df_city = df.copy()
        elif first_column == "Region":
            df['Breakdown'] = "Region"
            df_region = df.copy()
    
    # Clean headers and add platform for both dataframes
    if df_city is not None:
        df_city.columns = [col.split(':')[0] for col in df_city.columns]
        df_city['Platform'] = platform_name
    
    if df_region is not None:
        df_region.columns = [col.split(':')[0] for col in df_region.columns]
        df_region['Platform'] = platform_name
    
    return df_city, df_region

def process_geomap_data(base_path):
    """Process geomap CSV files"""
    try:
        results = {}
        
        # Process Web folder
        web_city, web_region = process_geomap_files("Web", "Web", base_path)
        
        # Process YouTube folder
        youtube_city, youtube_region = process_geomap_files("Youtube", "Youtube", base_path)
        
        # Merge city data
        if web_city is not None and youtube_city is not None:
            df_city_merged = pd.concat([web_city, youtube_city], ignore_index=True)
            results['city'] = df_city_merged
        
        # Merge region data
        if web_region is not None and youtube_region is not None:
            df_region_merged = pd.concat([web_region, youtube_region], ignore_index=True)
            results['region'] = df_region_merged
        
        return results, None
    except Exception as e:
        return None, str(e)

# Main content area
# Check if folders are selected
if not st.session_state.base_path or not st.session_state.output_path:
    st.markdown('''
    <div class="warning-box">
        <h3>⚠️ Please select folders first</h3>
        <p>Use the sidebar to select your Google Trends data folder and output folder before processing.</p>
    </div>
    ''', unsafe_allow_html=True)
else:
    st.markdown(f'<div class="success-box"><strong>📂 Data Folder:</strong> {st.session_state.base_path}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="success-box"><strong>💾 Output Folder:</strong> {st.session_state.output_path}</div>', unsafe_allow_html=True)

# Process timeline data
if process_timeline and st.session_state.base_path and st.session_state.output_path:
    with st.spinner("Processing timeline data..."):
        timeline_results, timeline_error = process_timeline_data(st.session_state.base_path)
        
        if timeline_error:
            st.error(f"Error processing timeline data: {timeline_error}")
        elif timeline_results and 'merged' in timeline_results:
            st.session_state.processed_data['timeline'] = timeline_results['merged']
            st.success("✅ Timeline data processed successfully!")
            
            # Display summary
            df_merged = timeline_results['merged']
            st.write("**Data Summary:**")
            st.write(f"- Total rows: {df_merged.shape[0]}")
            st.write(f"- Columns: {df_merged.shape[1]}")
            st.write("- Platform distribution:")
            st.write(df_merged['Platform'].value_counts())
            
            # Show preview
            st.write("**Data Preview:**")
            st.dataframe(df_merged.head(), use_container_width=True)

# Process geomap data
if process_geomap and st.session_state.base_path and st.session_state.output_path:
    with st.spinner("Processing geomap data..."):
        geomap_results, geomap_error = process_geomap_data(st.session_state.base_path)
        
        if geomap_error:
            st.error(f"Error processing geomap data: {geomap_error}")
        elif geomap_results:
            st.session_state.processed_data.update(geomap_results)
            st.success("✅ GeoMap data processed successfully!")
            
            # Display summary for city data
            if 'city' in geomap_results:
                st.write("**City Data Summary:**")
                df_city = geomap_results['city']
                st.write(f"- Total rows: {df_city.shape[0]}")
                st.write(f"- Columns: {df_city.shape[1]}")
                st.write("- Platform distribution:")
                st.write(df_city['Platform'].value_counts())
                st.dataframe(df_city.head(), use_container_width=True)
            
            # Display summary for region data
            if 'region' in geomap_results:
                st.write("**Region Data Summary:**")
                df_region = geomap_results['region']
                st.write(f"- Total rows: {df_region.shape[0]}")
                st.write(f"- Columns: {df_region.shape[1]}")
                st.write("- Platform distribution:")
                st.write(df_region['Platform'].value_counts())
                st.dataframe(df_region.head(), use_container_width=True)

# Save files to output folder section in sidebar
with st.sidebar:
    if st.session_state.processed_data and st.session_state.output_path:
        st.markdown("---")
        
        def save_files_to_folder():
            """Save all processed files to the output folder"""
            try:
                saved_files = []
                
                # Create output directory if it doesn't exist
                os.makedirs(st.session_state.output_path, exist_ok=True)
                
                if 'timeline' in st.session_state.processed_data:
                    timeline_path = os.path.join(st.session_state.output_path, "gt_timeline.csv")
                    st.session_state.processed_data['timeline'].to_csv(timeline_path, index=False)
                    saved_files.append("gt_timeline.csv")
                
                if 'city' in st.session_state.processed_data:
                    city_path = os.path.join(st.session_state.output_path, "gt_geomap_city.csv")
                    st.session_state.processed_data['city'].to_csv(city_path, index=False)
                    saved_files.append("gt_geomap_city.csv")
                
                if 'region' in st.session_state.processed_data:
                    region_path = os.path.join(st.session_state.output_path, "gt_geomap_region.csv")
                    st.session_state.processed_data['region'].to_csv(region_path, index=False)
                    saved_files.append("gt_geomap_region.csv")
                
                return saved_files, None
            except Exception as e:
                return None, str(e)
        
        if st.button("💾 Save Files to Output Folder", use_container_width=True, type="primary"):
            with st.spinner("Saving files..."):
                saved_files, save_error = save_files_to_folder()
                
                if save_error:
                    st.error(f"Error saving files: {save_error}")
                elif saved_files:
                    st.success(f"✅ Successfully saved {len(saved_files)} files!")
                    for file in saved_files:
                        st.write(f"✓ {file}")
                    st.info(f"📂 Files saved to: {st.session_state.output_path}")


# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #666; margin-top: 2rem;'>
        <p>Created by @djslash9 | 2025</p>
    </div>
    """, 
    unsafe_allow_html=True
)