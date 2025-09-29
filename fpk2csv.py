import os
import sys
import asyncio
# Set event loop policy for Windows before any other imports
if sys.platform == "win32":
    if hasattr(asyncio, "WindowsSelectorEventLoopPolicy"):
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
# Create and set event loop
try:
    loop = asyncio.get_event_loop()
except RuntimeError:
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
# Suppress warnings
import warnings
warnings.filterwarnings('ignore')
# Configure logging before other imports
import logging
logging.getLogger('nltk').setLevel(logging.ERROR)
logging.getLogger('streamlit').setLevel(logging.ERROR)
# Fix for PyTorch/Streamlit compatibility - must be done before Streamlit import
def patch_streamlit_watcher():
    """Patch Streamlit to avoid PyTorch class inspection errors"""
    try:
        import streamlit.watcher.local_sources_watcher as watcher
        
        original_get_module_paths = watcher.get_module_paths
        
        def safe_get_module_paths(module):
            """Safe version that skips problematic modules"""
            try:
                # Skip torch.classes and other problematic modules
                if hasattr(module, '__name__'):
                    module_name = str(module.__name__)
                    if any(problematic in module_name for problematic in ['torch.classes', 'torch._classes']):
                        return []
                
                # Skip modules that might cause attribute errors
                if not hasattr(module, '__path__'):
                    return []
                    
                return original_get_module_paths(module)
            except Exception:
                return []
        
        watcher.get_module_paths = safe_get_module_paths
        return True
    except Exception:
        return False
# Apply the patch before importing Streamlit
patch_streamlit_watcher()
# Now import Streamlit and other packages
import streamlit as st
import pandas as pd
import glob
import re
from datetime import datetime
from pathlib import Path
import time
# Import tkinter for file dialogs
import tkinter as tk
from tkinter import filedialog, messagebox

# Fix for NLTK - check and download only if needed
try:
    import nltk
    # Check if data is already downloaded without triggering download
    nltk_data_path = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'nltk_data')
    
    # Only download if not present
    required_packages = ['punkt', 'stopwords', 'wordnet', 'omw-1.4', 'vader_lexicon']
    for package in required_packages:
        package_path = os.path.join(nltk_data_path, package)
        if not os.path.exists(package_path):
            nltk.download(package, quiet=True)
except Exception:
    # Silently continue if NLTK issues occur
    pass

class FileDialogManager:
    """Manager for tkinter file dialogs"""
    
    def __init__(self):
        self.root = None
        self._init_root()
    
    def _init_root(self):
        """Initialize the tkinter root window (hidden)"""
        try:
            if self.root is None:
                self.root = tk.Tk()
                self.root.withdraw()  # Hide the main window
                self.root.attributes('-topmost', True)  # Make dialogs appear on top
        except Exception as e:
            st.error(f"Error initializing file dialog: {e}")
    
    def select_input_folder(self, initial_dir=None):
        """Open folder selection dialog for input folder"""
        try:
            self._init_root()
            folder_path = filedialog.askdirectory(
                title="Select FPK Input Folder",
                initialdir=initial_dir
            )
            return folder_path if folder_path else None
        except Exception as e:
            st.error(f"Error selecting input folder: {e}")
            return None
    
    def select_output_folder(self, initial_dir=None):
        """Open folder selection dialog for output folder"""
        try:
            self._init_root()
            folder_path = filedialog.askdirectory(
                title="Select Output Folder",
                initialdir=initial_dir
            )
            return folder_path if folder_path else None
        except Exception as e:
            st.error(f"Error selecting output folder: {e}")
            return None
    
    def cleanup(self):
        """Clean up tkinter root window"""
        try:
            if self.root:
                self.root.destroy()
                self.root = None
        except:
            pass

# Page configuration
st.set_page_config(
    page_title="FPK File Processor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .path-display {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 5px;
        border-left: 4px solid #1f77b4;
        font-family: monospace;
        font-size: 0.9em;
        word-break: break-all;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        border: 1px solid #e9ecef;
    }
    .metric-value {
        font-size: 2rem;
        font-weight: bold;
        color: #1f77b4;
    }
    .metric-label {
        font-size: 1rem;
        color: #6c757d;
    }
</style>
""", unsafe_allow_html=True)

def validate_date_folder_name(folder_name):
    """
    Validate if folder name matches date format
    Returns datetime object if valid, None otherwise
    """
    date_patterns = [
        r'^\d{4}-\d{2}-\d{2}$',  # YYYY-MM-DD
        r'^\d{4}\.\d{2}\.\d{2}$',  # YYYY.MM.DD
        r'^\d{4}_\d{2}_\d{2}$',  # YYYY_MM_DD
        r'^\d{8}$'  # YYYYMMDD
    ]
    
    for pattern in date_patterns:
        if re.match(pattern, folder_name):
            try:
                if pattern == r'^\d{8}$':  # YYYYMMDD
                    return datetime.strptime(folder_name, '%Y%m%d')
                elif pattern == r'^\d{4}\.\d{2}\.\d{2}$':  # YYYY.MM.DD
                    return datetime.strptime(folder_name, '%Y.%m.%d')
                elif pattern == r'^\d{4}_\d{2}_\d{2}$':  # YYYY_MM_DD
                    return datetime.strptime(folder_name, '%Y_%m_%d')
                else:  # YYYY-MM-DD
                    return datetime.strptime(folder_name, '%Y-%m-%d')
            except ValueError:
                continue
    return None

def count_files_in_folder(folder_path):
    """Count Excel files in folder structure"""
    if not os.path.exists(folder_path):
        return 0
    
    count = 0
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(('.xlsx', '.xls')):
                count += 1
    return count

def scan_fpk_structure(input_folder):
    """Scan and analyze the FPK folder structure"""
    if not os.path.exists(input_folder):
        return None
    
    date_folders = []
    excel_files = []
    total_files = 0
    
    for item in os.listdir(input_folder):
        item_path = os.path.join(input_folder, item)
        if os.path.isdir(item_path):
            Date = validate_date_folder_name(item)
            if Date:
                # Count Excel files in this date folder
                xlsx_files = glob.glob(os.path.join(item_path, "*.xlsx"))
                xls_files = glob.glob(os.path.join(item_path, "*.xls"))
                file_count = len(xlsx_files) + len(xls_files)
                total_files += file_count
                
                date_folders.append({
                    'name': item,
                    'path': item_path,
                    'date': Date,
                    'file_count': file_count
                })
                
                for file_path in xlsx_files + xls_files:
                    excel_files.append({
                        'path': file_path,
                        'date': Date,
                        'folder_name': item
                    })
    
    return {
        'date_folders': date_folders,
        'excel_files': excel_files,
        'total_files': total_files,
        'total_date_folders': len(date_folders)
    }

def sanitize_sheet_name(sheet_name):
    """
    Sanitize sheet name for use in folder and file names
    Remove dots and other special characters
    """
    # Replace dots, ellipsis, and other special characters with underscores
    sanitized = re.sub(r'[\.…]', '_', sheet_name)
    # Replace other problematic characters
    sanitized = re.sub(r'[\/\\\:\*\?\"\<\>\|]', '_', sanitized)
    # Remove multiple consecutive underscores
    sanitized = re.sub(r'_{2,}', '_', sanitized)
    # Remove leading/trailing underscores and spaces
    sanitized = sanitized.strip(' _')
    # If empty after sanitization, use a default name
    if not sanitized:
        sanitized = "Sheet"
    return sanitized

def process_excel_file(file_path, Date, base_output_dir):
    """
    Process a single Excel file and save as CSV
    """
    try:
        # Read all sheets from the Excel file
        excel_file = pd.ExcelFile(file_path)
        processed_files = []
        
        for sheet_name in excel_file.sheet_names:
            try:
                # Read the sheet with no header initially
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                # Skip if file is empty or has very few rows
                if df.empty or len(df) < 2:
                    processed_files.append({
                        'status': 'skipped',
                        'file_path': '',
                        'reason': 'Sheet is empty or has insufficient data',
                        'sheet_name': sheet_name
                    })
                    continue
                
                # Delete first column (A) if it exists and is mostly empty
                if len(df.columns) > 0:
                    first_col_empty_ratio = df.iloc[:, 0].isna().sum() / len(df)
                    if first_col_empty_ratio > 0.8:  # If 80%+ empty, remove it
                        df = df.drop(df.columns[0], axis=1)
                
                # Use row 5 as headers (index 4 since we start from 0)
                header_row = 4
                
                # If header row is beyond available rows, adjust
                if header_row >= len(df):
                    header_row = min(4, len(df) - 1) if len(df) > 1 else 0
                
                # Set row 5 as headers
                if len(df) > header_row:
                    # Use the header row as column names
                    df.columns = df.iloc[header_row]
                    
                    # Remove rows before the header row + 1 (keep only data rows)
                    if header_row + 1 < len(df):
                        df = df.iloc[header_row + 1:].reset_index(drop=True)
                    else:
                        df = pd.DataFrame(columns=df.columns)  # Empty but with correct columns
                
                # Remove completely empty rows
                df = df.dropna(how='all')
                
                # Skip if no data remains
                if df.empty:
                    processed_files.append({
                        'status': 'skipped',
                        'file_path': '',
                        'reason': 'No data rows after processing',
                        'sheet_name': sheet_name
                    })
                    continue
                
                # Add Date column only (no file_name column)
                if 'Date' not in df.columns:
                    df.insert(0, 'Date', Date.strftime('%Y-%m-%d'))
                
                # Create output directory for this sheet using sanitized name
                sanitized_sheet_name = sanitize_sheet_name(sheet_name)
                output_dir = os.path.join(base_output_dir, sanitized_sheet_name)
                os.makedirs(output_dir, exist_ok=True)
                
                # Create output filename using sanitized sheet name
                output_filename = f"{sanitized_sheet_name} {Date.strftime('%Y%m%d')}.csv"
                output_path = os.path.join(output_dir, output_filename)
                
                # Skip if file already exists
                if os.path.exists(output_path):
                    processed_files.append({
                        'status': 'skipped',
                        'file_path': output_path,
                        'reason': 'File already exists',
                        'sheet_name': sheet_name
                    })
                    continue
                
                # Save as CSV
                df.to_csv(output_path, index=False)
                
                processed_files.append({
                    'status': 'success',
                    'file_path': output_path,
                    'sheet_name': sheet_name,
                    'rows_processed': len(df)
                })
                
            except Exception as sheet_error:
                processed_files.append({
                    'status': 'error',
                    'file_path': file_path,
                    'reason': f"Sheet processing error: {str(sheet_error)}",
                    'sheet_name': sheet_name
                })
                continue
        
        return processed_files
        
    except Exception as e:
        return [{
            'status': 'error',
            'file_path': file_path,
            'reason': str(e),
            'sheet_name': 'unknown'
        }]

def load_file_paths():
    """Load file paths from a configuration file"""
    config_file = "file_paths.config"
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r') as f:
                lines = f.readlines()
                paths = {}
                for line in lines:
                    if '=' in line:
                        key, value = line.strip().split('=', 1)
                        paths[key] = value
                return paths
        except:
            return {}
    return {}

def save_file_paths(input_path, output_path):
    """Save file paths to a configuration file"""
    config_file = "file_paths.config"
    try:
        with open(config_file, 'w') as f:
            f.write(f"input_folder={input_path}\n")
            f.write(f"output_folder={output_path}\n")
        return True
    except:
        return False

def main():
    """Main Streamlit application"""
    
    st.markdown('<div class="main-header">📊 FPK File Processor</div>', unsafe_allow_html=True)
    
    # Initialize file dialog manager
    if 'file_dialog' not in st.session_state:
        st.session_state.file_dialog = FileDialogManager()
    
    # Sidebar for folder selection
    st.sidebar.title("Folder Selection")
    st.sidebar.markdown("Select input and output folders using file dialogs or manual entry")
    
    # Load saved file paths
    saved_paths = load_file_paths()
    
    # Folder selection with tkinter dialogs
    st.sidebar.markdown("### Input Folder")
    col1, col2 = st.sidebar.columns([3, 1])
    
    with col1:
        input_folder = st.text_input(
            "FPK Folder Path:",
            value=st.session_state.get('input_folder', saved_paths.get('input_folder', '')),
            placeholder="C:/path/to/your/FPK/folder",
            key="input_path",
            label_visibility="collapsed"
        )
    
    with col2:
        if st.button("📁 Browse", key="browse_input", use_container_width=True):
            selected_folder = st.session_state.file_dialog.select_input_folder(
                initial_dir=st.session_state.get('input_folder', None)
            )
            if selected_folder:
                st.session_state.input_folder = selected_folder
                st.rerun()
    
    st.sidebar.markdown("### Output Folder")
    col1, col2 = st.sidebar.columns([3, 1])
    
    with col1:
        output_folder = st.text_input(
            "Output Folder Path:",
            value=st.session_state.get('output_folder', saved_paths.get('output_folder', '')),
            placeholder="C:/path/to/output/folder",
            key="output_path",
            label_visibility="collapsed"
        )
    
    with col2:
        if st.button("📁 Browse", key="browse_output", use_container_width=True):
            selected_folder = st.session_state.file_dialog.select_output_folder(
                initial_dir=st.session_state.get('output_folder', None)
            )
            if selected_folder:
                st.session_state.output_folder = selected_folder
                st.rerun()
    
    # Load File Path Button
    st.sidebar.markdown("### Path Management")
    col1, col2 = st.sidebar.columns(2)
    
    with col1:
        if st.button("💾 Save Paths", use_container_width=True):
            if input_folder and output_folder:
                if save_file_paths(input_folder, output_folder):
                    st.sidebar.success("✅ Paths saved successfully!")
                else:
                    st.sidebar.error("❌ Failed to save paths")
            else:
                st.sidebar.warning("⚠️ Please enter both paths first")
    
    with col2:
        if st.button("📂 Load Paths", use_container_width=True):
            saved_paths = load_file_paths()
            if saved_paths:
                st.session_state.input_folder = saved_paths.get('input_folder', '')
                st.session_state.output_folder = saved_paths.get('output_folder', '')
                st.sidebar.success("✅ Paths loaded successfully!")
                st.rerun()
            else:
                st.sidebar.warning("⚠️ No saved paths found")
    
    # Store in session state
    if input_folder:
        st.session_state.input_folder = input_folder
    if output_folder:
        st.session_state.output_folder = output_folder
    
    # Display selected folders
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Selected Folders")
    
    if st.session_state.get('input_folder'):
        st.sidebar.markdown("**Input:**")
        st.sidebar.markdown(f'<div class="path-display">{st.session_state.input_folder}</div>', unsafe_allow_html=True)
    
    if st.session_state.get('output_folder'):
        st.sidebar.markdown("**Output:**")
        st.sidebar.markdown(f'<div class="path-display">{st.session_state.output_folder}</div>', unsafe_allow_html=True)
    
    # Quick actions
    st.sidebar.markdown("### Quick Actions")
    if st.sidebar.button("🔄 Refresh Folder Scan", use_container_width=True):
        st.rerun()
    
    # Main content area
    if not st.session_state.get('input_folder'):
        st.info("👈 Please select an input folder using the file browser or enter the path manually")
        st.markdown("""
        ### Expected Folder Structure:
        ```
        FPK/
        ├── 2024-01-15/
        │   └── datafile.xlsx
        ├── 2024-01-16/
        │   └── datafile.xlsx
        └── 2024-01-17/
            └── datafile.xlsx
        ```
        
        ### How to use:
        1. Click **📁 Browse** next to "Input Folder" to select your FPK folder
        2. Click **📁 Browse** next to "Output Folder" to select where to save processed files
        3. Click **💾 Save Paths** to remember your selections
        4. Click **🚀 Start Processing** to begin conversion
        """)
        return
    
    # Check if input folder exists
    if not os.path.exists(st.session_state.input_folder):
        st.error(f"❌ Input folder does not exist: {st.session_state.input_folder}")
        st.info("Please select a valid folder using the file browser")
        return
    
    if not st.session_state.get('output_folder'):
        st.info("👈 Please select an output folder using the file browser or enter the path manually")
        return
    
    # Scan folder structure
    with st.spinner("Scanning folder structure..."):
        scan_results = scan_fpk_structure(st.session_state.input_folder)
    
    if not scan_results:
        st.error("❌ No valid FPK structure found!")
        st.markdown("""
        ### Required Folder Structure:
        - Root folder containing date-named subfolders
        - Date format: `YYYY-MM-DD`, `YYYY.MM.DD`, `YYYYMMDD`, etc.
        - Each date folder should contain Excel files (.xlsx or .xls)
        """)
        return
    
    if scan_results['total_files'] == 0:
        st.warning("⚠️ No Excel files found in the FPK structure!")
        return
    
    # Display folder statistics
    st.markdown("### 📊 Folder Statistics")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{scan_results['total_date_folders']}</div>
            <div class="metric-label">📁 Date Folders</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{scan_results['total_files']}</div>
            <div class="metric-label">📄 Excel Files</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        output_file_count = count_files_in_folder(st.session_state.output_folder)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{output_file_count}</div>
            <div class="metric-label">💾 Existing Output Files</div>
        </div>
        """, unsafe_allow_html=True)
    
    # Display date folders
    st.markdown("### 📅 Date Folders Found")
    
    if scan_results['date_folders']:
        date_df = pd.DataFrame([
            {
                'Folder Name': folder['name'],
                'Date': folder['date'].strftime('%Y-%m-%d'),
                'Files': folder['file_count'],
                'Path': folder['path']
            }
            for folder in scan_results['date_folders']
        ])
        st.dataframe(date_df, use_container_width=True, hide_index=True)
    else:
        st.warning("No valid date folders found!")
    
    # Show Sheet Names from first Excel file
    if scan_results['excel_files']:
        st.markdown("### 📋 Sheet Names Found")
        sample_file = scan_results['excel_files'][0]
        try:
            excel_file = pd.ExcelFile(sample_file['path'])
            sheet_names = excel_file.sheet_names
            
            st.write(f"**Sheets in** `{os.path.basename(sample_file['path'])}`:")
            for i, sheet_name in enumerate(sheet_names, 1):
                sanitized_name = sanitize_sheet_name(sheet_name)
                st.write(f"{i}. **{sheet_name}** → `{sanitized_name}`")
                
        except Exception as e:
            st.error(f"Error reading sheet names: {e}")
    
    # Processing options
    st.markdown("### ⚙️ Processing Options")
    
    col1, col2 = st.columns(2)
    
    with col1:
        skip_existing = st.checkbox("Skip existing files", value=True)
        show_detailed_progress = st.checkbox("Show detailed progress", value=True)
    
    with col2:
        debug_mode = st.checkbox("Debug mode (show sheet structure)", value=False)
        create_sample_structure = st.checkbox("Create sample structure for testing", value=False)
    
    # Create sample structure if requested
    if create_sample_structure:
        if st.button("Create Sample FPK Structure"):
            sample_dir = os.path.join(st.session_state.input_folder, "sample_fpk")
            os.makedirs(sample_dir, exist_ok=True)
            
            sample_dates = ['2024-01-15', '2024-01-16', '2024-01-17']
            sample_data = {
                'Unnamed: 0': ['', '', '', '', 'Row0'],
                'Col1': ['', '', '', 'Header1', 'Data1'],
                'Col2': ['', '', '', 'Header2', 'Data2'],
                'Col3': ['', '', '', 'Header3', 'Data3']
            }
            
            for date_str in sample_dates:
                date_dir = os.path.join(sample_dir, date_str)
                os.makedirs(date_dir, exist_ok=True)
                
                df = pd.DataFrame(sample_data)
                file_path = os.path.join(date_dir, f"sample_data_{date_str}.xlsx")
                
                # Create Excel file with multiple sheets
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False)
                    df.to_excel(writer, sheet_name='Data', index=False)
                    df.to_excel(writer, sheet_name='Report', index=False)
            
            st.success(f"✅ Sample structure created at: {sample_dir}")
            st.session_state.input_folder = sample_dir
            st.rerun()
    
    # Process button
    st.markdown("---")
    if st.button("🚀 Start Processing", type="primary", use_container_width=True):
        if not os.path.exists(st.session_state.output_folder):
            os.makedirs(st.session_state.output_folder, exist_ok=True)
            st.success(f"✅ Created output directory: {st.session_state.output_folder}")
        
        # Save paths when processing starts
        if input_folder and output_folder:
            save_file_paths(input_folder, output_folder)
        
        # Initialize progress
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Process files
        all_results = []
        created_folders = set()
        file_counter = {}
        
        total_files = len(scan_results['excel_files'])
        
        for i, file_info in enumerate(scan_results['excel_files']):
            file_path = file_info['path']
            Date = file_info['date']
            
            # Update progress
            progress = (i + 1) / total_files
            progress_bar.progress(progress)
            status_text.text(f"Processing {i+1}/{total_files}: {os.path.basename(file_path)}")
            
            results = process_excel_file(file_path, Date, st.session_state.output_folder)
            
            for result in results:
                result['source_file'] = file_path
                result['date_folder'] = file_info['folder_name']
                all_results.append(result)
                
                # Track created folders
                if result['status'] == 'success':
                    folder_path = os.path.dirname(result['file_path'])
                    created_folders.add(folder_path)
                    
                    # Track file names for duplicates
                    filename = os.path.basename(result['file_path'])
                    if filename in file_counter:
                        file_counter[filename].append(result['file_path'])
                    else:
                        file_counter[filename] = [result['file_path']]
            
            # Small delay to show progress
            time.sleep(0.1)
        
        # Complete progress
        progress_bar.progress(1.0)
        status_text.text("✅ Processing complete!")
        
        # Generate report
        st.markdown("## 📋 Processing Report")
        
        # Statistics
        success_count = len([r for r in all_results if r['status'] == 'success'])
        skipped_count = len([r for r in all_results if r['status'] == 'skipped'])
        error_count = len([r for r in all_results if r['status'] == 'error'])
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("✅ Success", success_count)
        with col2:
            st.metric("⏭️ Skipped", skipped_count)
        with col3:
            st.metric("❌ Errors", error_count)
        
        # Show detailed results by sheet
        if debug_mode:
            with st.expander("🔍 Detailed Sheet Processing Results"):
                sheet_results = {}
                for result in all_results:
                    sheet_name = result['sheet_name']
                    if sheet_name not in sheet_results:
                        sheet_results[sheet_name] = []
                    sheet_results[sheet_name].append(result)
                
                for sheet_name, results in sheet_results.items():
                    st.subheader(f"Sheet: {sheet_name}")
                    success = len([r for r in results if r['status'] == 'success'])
                    errors = len([r for r in results if r['status'] == 'error'])
                    skipped = len([r for r in results if r['status'] == 'skipped'])
                    
                    st.write(f"Success: {success}, Errors: {errors}, Skipped: {skipped}")
                    
                    for result in results:
                        if result['status'] == 'error':
                            st.error(f"Error: {result['reason']}")
                        elif result['status'] == 'skipped':
                            st.warning(f"Skipped: {result['reason']}")
        
        # Show created folders
        if created_folders:
            with st.expander(f"📁 Created Folders ({len(created_folders)})"):
                for folder in sorted(created_folders):
                    st.code(f"📂 {os.path.relpath(folder, st.session_state.output_folder)}")
        
        # Show success messages
        if success_count > 0 and show_detailed_progress:
            with st.expander("🎉 Successfully Created Files"):
                for result in all_results:
                    if result['status'] == 'success':
                        rel_path = os.path.relpath(result['file_path'], st.session_state.output_folder)
                        rows_info = f" ({result['rows_processed']} rows)" if 'rows_processed' in result else ""
                        st.success(f"✅ {rel_path}{rows_info} (Sheet: {result['sheet_name']})")
        
        # Show errors
        if error_count > 0:
            with st.expander("❌ Errors Encountered"):
                for result in all_results:
                    if result['status'] == 'error':
                        st.error(f"❌ {os.path.basename(result['source_file'])} - {result['sheet_name']}: {result['reason']}")
        
        # Show skipped files
        if skipped_count > 0 and show_detailed_progress:
            with st.expander("⏭️ Skipped Files"):
                for result in all_results:
                    if result['status'] == 'skipped':
                        if result.get('file_path'):
                            rel_path = os.path.relpath(result['file_path'], st.session_state.output_folder)
                            st.warning(f"⏭️ {rel_path} - {result['reason']}")
                        else:
                            st.warning(f"⏭️ {result['sheet_name']} - {result['reason']}")

# Run the app
if __name__ == "__main__":
    # Initialize session state
    if 'input_folder' not in st.session_state:
        st.session_state.input_folder = ''
    if 'output_folder' not in st.session_state:
        st.session_state.output_folder = ''
    
    try:
        main()
    finally:
        # Clean up tkinter resources
        if 'file_dialog' in st.session_state:
            st.session_state.file_dialog.cleanup()