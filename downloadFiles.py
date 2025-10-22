import streamlit as st
import pandas as pd
import requests
import os
from urllib.parse import urlparse
import time
from pathlib import Path
from datetime import datetime

def download_file(url, download_folder, filename=None, progress_bar=None):
    """
    Download a file from a URL with progress tracking
    """
    try:
        # Create download folder if it doesn't exist
        Path(download_folder).mkdir(parents=True, exist_ok=True)
        
        # Get filename from URL if not provided
        if not filename:
            parsed_url = urlparse(url)
            filename = os.path.basename(parsed_url.path)
            if not filename or filename == '/':
                filename = f"downloaded_file_{int(time.time())}"
        
        # Clean filename
        filename = "".join(c for c in filename if c.isalnum() or c in "._- ")
        
        # Full path for the file
        file_path = os.path.join(download_folder, filename)
        username='nahimana2023'
        password='CDRReports2025'
        # Download the file with progress
        response = requests.get(url, stream=True, timeout=30,auth=(username,password))
        response.raise_for_status()
        
        total_size = int(response.headers.get('content-length', 0))
        downloaded_size = 0
        
        with open(file_path, 'wb') as file:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    file.write(chunk)
                    downloaded_size += len(chunk)
                    if progress_bar and total_size > 0:
                        progress = downloaded_size / total_size
                        progress_bar.progress(progress)
        
        return file_path, True, "Download completed"
        
    except requests.exceptions.RequestException as e:
        return None, False, f"Network error: {str(e)}"
    except Exception as e:
        return None, False, f"Error: {str(e)}"

def main():
    st.set_page_config(
        page_title="Excel URL Downloader",
        page_icon="📥",
        layout="wide"
    )
    
    st.title("📥 Excel URL Downloader")
    st.markdown("Upload an Excel file containing URLs and download all files to your chosen location.")
    
    # Initialize session state for folder path
    if 'download_folder' not in st.session_state:
        st.session_state.download_folder = "./downloads"
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("Configuration")
        
        # File upload
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls'],
            help="Upload an Excel file with URLs in one column"
        )
        
        # Download location selection
        st.subheader("Download Location")
        
        # Method 1: Manual path input
        download_folder = st.text_input(
            "Enter Download Folder Path",
            value=st.session_state.download_folder,
            placeholder="./downloads or /path/to/your/folder",
            help="Enter full path to where you want to save files"
        )
        
        if download_folder != st.session_state.download_folder:
            st.session_state.download_folder = download_folder
        
        # Quick select buttons for common locations
        st.write("Quick select paths:")
        quick_col1, quick_col2 = st.columns(2)
        
        with quick_col1:
            if st.button("🖥️ Desktop", use_container_width=True):
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                st.session_state.download_folder = desktop_path
                st.rerun()
            
            if st.button("📁 Documents", use_container_width=True):
                documents_path = os.path.join(os.path.expanduser("~"), "Documents")
                st.session_state.download_folder = documents_path
                st.rerun()
        
        with quick_col2:
            if st.button("📂 Downloads", use_container_width=True):
                downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
                st.session_state.download_folder = downloads_path
                st.rerun()
            
            if st.button("🗂️ Current Dir", use_container_width=True):
                current_dir = os.getcwd()
                st.session_state.download_folder = current_dir
                st.rerun()
        
        # Show folder info
        if st.session_state.download_folder:
            if os.path.exists(st.session_state.download_folder):
                try:
                    folder_size = sum(
                        os.path.getsize(os.path.join(dirpath, filename))
                        for dirpath, dirnames, filenames in os.walk(st.session_state.download_folder)
                        for filename in filenames
                    )
                    st.info(f"✅ Folder exists\n**Size:** {folder_size:,} bytes")
                except:
                    st.info("✅ Folder exists")
            else:
                st.warning("⚠️ Folder will be created on download")
        
        # Advanced options
        st.subheader("Advanced Options")
        
        url_column = st.text_input(
            "URL Column Name",
            value="URL",
            help="Name of the column containing URLs"
        )
        
        filename_column = st.text_input(
            "Filename Column Name (Optional)",
            value="",
            help="Name of the column containing custom filenames"
        )
        
        sheet_name = st.text_input(
            "Sheet Name",
            value="0",
            help="Sheet name or index (0 for first sheet)"
        )
        
        delay = st.slider(
            "Delay between downloads (seconds)",
            min_value=0.0,
            max_value=5.0,
            value=1.0,
            step=0.5,
        )
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if uploaded_file is not None:
            try:
                # Read the Excel file
                if sheet_name.isdigit():
                    sheet_name = int(sheet_name)
                
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                
                st.success(f"✅ Excel file loaded successfully!")
                st.write(f"**File:** {uploaded_file.name}")
                st.write(f"**Shape:** {df.shape[0]} rows × {df.shape[1]} columns")
                
                # Show current download folder
                if st.session_state.download_folder:
                    st.info(f"**Download folder:** `{st.session_state.download_folder}`")
                
                # Show dataframe preview
                with st.expander("Preview Data", expanded=True):
                    st.dataframe(df.head(10), use_container_width=True)
                
                # Check if URL column exists
                if url_column not in df.columns:
                    st.error(f"❌ Column '{url_column}' not found in the Excel file.")
                    st.write("Available columns:", list(df.columns))
                else:
                    # Filter valid URLs
                    valid_urls = df[df[url_column].notna() & (df[url_column].astype(str).str.strip() != '')]
                    total_urls = len(valid_urls)
                    
                    st.write(f"**Valid URLs found:** {total_urls}")
                    
                    if total_urls == 0:
                        st.warning("No valid URLs found in the specified column.")
                    else:
                        # Download progress
                        st.subheader("Download Progress")
                        
                        # Check if folder is selected
                        if not st.session_state.download_folder:
                            st.error("❌ Please select a download folder first!")
                        else:
                            if st.button("🚀 Start Download", type="primary", use_container_width=True):
                                # Initialize progress tracking
                                progress_text = st.empty()
                                overall_progress = st.progress(0)
                                status_text = st.empty()
                                
                                successful_downloads = 0
                                failed_downloads = []
                                download_log = []
                                
                                for index, row in valid_urls.iterrows():
                                    url = str(row[url_column]).strip()
                                    
                                    # Get custom filename if specified
                                    custom_filename = None
                                    if filename_column and filename_column.strip() and filename_column in df.columns and pd.notna(row[filename_column]):
                                        custom_filename = str(row[filename_column])
                                    
                                    # Update progress
                                    progress_text.write(f"Downloading {index + 1}/{total_urls}: {url}")
                                    overall_progress.progress((index) / total_urls)
                                    
                                    # Create individual progress bar for current download
                                    file_progress = st.progress(0)
                                    
                                    # Download the file
                                    file_path, success, message = download_file(
                                        url, 
                                        st.session_state.download_folder, 
                                        custom_filename,
                                        file_progress
                                    )
                                    
                                    if success:
                                        successful_downloads += 1
                                        file_size = os.path.getsize(file_path) if file_path and os.path.exists(file_path) else 0
                                        status_text.success(f"✅ {os.path.basename(file_path)} ({file_size:,} bytes)")
                                        download_log.append({
                                            "url": url,
                                            "filename": os.path.basename(file_path),
                                            "status": "Success",
                                            "size": file_size
                                        })
                                    else:
                                        failed_downloads.append({"url": url, "error": message})
                                        status_text.error(f"❌ Failed: {url}")
                                        download_log.append({
                                            "url": url,
                                            "filename": custom_filename or "N/A",
                                            "status": "Failed",
                                            "error": message
                                        })
                                    
                                    # Clear individual progress bar
                                    file_progress.empty()
                                    
                                    # Delay between downloads
                                    if delay > 0 and index < total_urls - 1:
                                        time.sleep(delay)
                                
                                # Final progress update
                                overall_progress.progress(1.0)
                                progress_text.empty()
                                
                                # Show summary
                                st.subheader("📊 Download Summary")
                                
                                col_success, col_failed, col_total = st.columns(3)
                                
                                with col_success:
                                    st.metric("Successful", successful_downloads, delta=f"+{successful_downloads}")
                                
                                with col_failed:
                                    st.metric("Failed", len(failed_downloads), delta=f"+{len(failed_downloads)}", delta_color="inverse")
                                
                                with col_total:
                                    st.metric("Total", total_urls)
                                
                                # Show download log
                                with st.expander("📋 Download Log", expanded=True):
                                    log_df = pd.DataFrame(download_log)
                                    st.dataframe(log_df, use_container_width=True)
                                
                                st.success(f"📁 Files downloaded to: `{st.session_state.download_folder}`")
            
            except Exception as e:
                st.error(f"Error reading Excel file: {str(e)}")
        
        else:
            # Instructions when no file is uploaded
            st.info("👈 Please upload an Excel file to get started")
            
            with st.expander("📋 Expected Excel Format"):
                example_data = {
                    'URL': [
                        'https://example.com/file1.pdf',
                        'https://example.com/file2.jpg',
                        'https://example.com/file3.zip'
                    ],
                    'Filename': [
                        'document1.pdf',
                        'image2.jpg',
                        'archive3.zip'
                    ]
                }
                example_df = pd.DataFrame(example_data)
                st.dataframe(example_df, use_container_width=True)
    
    with col2:
        st.subheader("ℹ️ Instructions")
        
        st.markdown("""
        1. **Upload Excel File**
        2. **Enter Folder Path** or use quick buttons
        3. **Configure Columns**
        4. **Start Download**
        
        ### 📁 Folder Path Examples
        - **Windows**: `C:/Users/YourName/Downloads`
        - **Mac**: `/Users/YourName/Downloads`
        - **Linux**: `/home/yourname/Downloads`
        - **Relative**: `./downloads`
        """)
        
        # System info
        st.subheader("System Info")
        st.write(f"Current time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        if uploaded_file:
            st.write(f"File size: {uploaded_file.size:,} bytes")

if __name__ == "__main__":
    main()