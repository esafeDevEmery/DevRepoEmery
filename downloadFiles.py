import streamlit as st
import pandas as pd
import requests
import os
from urllib.parse import urlparse
import time
import io
import base64
from datetime import datetime
import zipfile

def download_file_to_memory(url):
    """
    Download a file to memory and return its content and metadata
    """
    try:
        username='nahimana2023'
        password='CDRReports2025'
        response = requests.get(url, stream=True, timeout=30,auth=(username,password))
        response.raise_for_status()
        
        # Get filename from URL
        parsed_url = urlparse(url)
        filename = os.path.basename(parsed_url.path)
        if not filename or filename == '/':
            filename = f"downloaded_file_{int(time.time())}"
        
        # Clean filename
        filename = "".join(c for c in filename if c.isalnum() or c in "._- ")
        
        # Read content into memory
        content = response.content
        file_size = len(content)
        
        return content, filename, file_size, True, "Download completed"
        
    except requests.exceptions.RequestException as e:
        return None, None, 0, False, f"Network error: {str(e)}"
    except Exception as e:
        return None, None, 0, False, f"Error: {str(e)}"

def create_download_link(content, filename, file_type="application/octet-stream"):
    """
    Create a download link for a file in memory
    """
    b64 = base64.b64encode(content).decode()
    href = f'<a href="data:{file_type};base64,{b64}" download="{filename}">Download {filename}</a>'
    return href

def create_zip_download_link(files_data, zip_filename="downloaded_files.zip"):
    """
    Create a ZIP file containing multiple files and provide download link
    """
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file_data in files_data:
            if file_data['success']:
                zip_file.writestr(file_data['filename'], file_data['content'])
    
    zip_buffer.seek(0)
    zip_content = zip_buffer.getvalue()
    
    b64 = base64.b64encode(zip_content).decode()
    href = f'<a href="data:application/zip;base64,{b64}" download="{zip_filename}">📦 Download All Files as ZIP</a>'
    return href

def get_file_type(filename):
    """
    Get MIME type based on file extension
    """
    extension = filename.split('.')[-1].lower()
    file_types = {
        'pdf': 'application/pdf',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'png': 'image/png',
        'gif': 'image/gif',
        'txt': 'text/plain',
        'csv': 'text/csv',
        'zip': 'application/zip',
        'doc': 'application/msword',
        'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'xls': 'application/vnd.ms-excel',
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }
    return file_types.get(extension, 'application/octet-stream')

def main():
    st.set_page_config(
        page_title="Excel URL Downloader",
        page_icon="📥",
        layout="wide"
    )
    
    st.title("📥 Excel URL Downloader")
    st.markdown("Upload an Excel file containing URLs and download files directly to your browser.")
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("Configuration")
        
        # File upload
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls'],
            help="Upload an Excel file with URLs in one column"
        )
        
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
        
        download_mode = st.radio(
            "Download Mode",
            ["Individual Files", "ZIP Archive"],
            help="Download files individually or as a single ZIP file"
        )
        
        delay = st.slider(
            "Delay between downloads (seconds)",
            min_value=0.0,
            max_value=5.0,
            value=1.0,
            step=0.5,
            help="Delay between processing URLs"
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
                        
                        if st.button("🚀 Process URLs", type="primary", use_container_width=True):
                            # Initialize progress tracking
                            progress_text = st.empty()
                            overall_progress = st.progress(0)
                            status_text = st.empty()
                            
                            successful_downloads = []
                            failed_downloads = []
                            all_downloaded_files = []
                            
                            for index, row in valid_urls.iterrows():
                                url = str(row[url_column]).strip()
                                
                                # Get custom filename if specified
                                custom_filename = None
                                if filename_column and filename_column.strip() and filename_column in df.columns and pd.notna(row[filename_column]):
                                    custom_filename = str(row[filename_column])
                                
                                # Update progress
                                progress_text.write(f"Processing {index + 1}/{total_urls}: {url}")
                                overall_progress.progress((index) / total_urls)
                                
                                # Download the file to memory
                                content, auto_filename, file_size, success, message = download_file_to_memory(url)
                                
                                # Use custom filename if provided, otherwise use auto-generated
                                filename = custom_filename or auto_filename
                                
                                if success:
                                    file_data = {
                                        'content': content,
                                        'filename': filename,
                                        'size': file_size,
                                        'url': url,
                                        'file_type': get_file_type(filename)
                                    }
                                    successful_downloads.append(file_data)
                                    all_downloaded_files.append({
                                        'content': content,
                                        'filename': filename,
                                        'success': True
                                    })
                                    status_text.success(f"✅ Downloaded: {filename} ({file_size:,} bytes)")
                                else:
                                    failed_downloads.append({"url": url, "error": message})
                                    all_downloaded_files.append({
                                        'filename': filename or "Unknown",
                                        'success': False
                                    })
                                    status_text.error(f"❌ Failed: {url} - {message}")
                                
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
                                st.metric("Successful", len(successful_downloads), delta=f"+{len(successful_downloads)}")
                            
                            with col_failed:
                                st.metric("Failed", len(failed_downloads), delta=f"+{len(failed_downloads)}", delta_color="inverse")
                            
                            with col_total:
                                st.metric("Total", total_urls)
                            
                            # Download options based on mode
                            st.subheader("📥 Download Files")
                            
                            if download_mode == "Individual Files":
                                # Show individual download links
                                if successful_downloads:
                                    st.write("### Individual File Downloads")
                                    for file_data in successful_downloads:
                                        download_link = create_download_link(
                                            file_data['content'],
                                            file_data['filename'],
                                            file_data['file_type']
                                        )
                                        st.markdown(download_link, unsafe_allow_html=True)
                                        st.caption(f"Size: {file_data['size']:,} bytes | URL: {file_data['url']}")
                                        st.write("---")
                                
                            else:  # ZIP Archive mode
                                # Show ZIP download link
                                if successful_downloads:
                                    st.write("### Download All Files as ZIP")
                                    zip_link = create_zip_download_link(all_downloaded_files)
                                    st.markdown(zip_link, unsafe_allow_html=True)
                                    st.caption(f"Contains {len(successful_downloads)} files, total size: {sum(f['size'] for f in successful_downloads):,} bytes")
                                
                                # Also show individual links in expander
                                if successful_downloads:
                                    with st.expander("📄 Individual File Links (for selective download)"):
                                        for file_data in successful_downloads:
                                            download_link = create_download_link(
                                                file_data['content'],
                                                file_data['filename'],
                                                file_data['file_type']
                                            )
                                            st.markdown(download_link, unsafe_allow_html=True)
                                            st.caption(f"Size: {file_data['size']:,} bytes")
                                            st.write("---")
                            
                            # Show failed downloads if any
                            if failed_downloads:
                                with st.expander("❌ Failed Downloads", expanded=False):
                                    for failed in failed_downloads:
                                        st.write(f"**URL:** {failed['url']}")
                                        st.write(f"**Error:** {failed['error']}")
                                        st.write("---")
                            
                            # File preview for common file types
                            if successful_downloads:
                                with st.expander("👀 File Previews", expanded=False):
                                    for file_data in successful_downloads[:5]:  # Limit to first 5 files
                                        filename = file_data['filename'].lower()
                                        content = file_data['content']
                                        
                                        st.write(f"**{file_data['filename']}** ({file_data['size']:,} bytes)")
                                        
                                        if filename.endswith(('.png', '.jpg', '.jpeg', '.gif')):
                                            # Display images
                                            st.image(content, caption=file_data['filename'], use_column_width=True)
                                        elif filename.endswith('.pdf'):
                                            # Show PDF info
                                            st.info("PDF file - download to view")
                                        elif filename.endswith(('.txt', '.csv')):
                                            # Show text content preview
                                            try:
                                                text_content = content.decode('utf-8')
                                                st.text_area("Content preview", text_content[:1000] + ("..." if len(text_content) > 1000 else ""), height=150)
                                            except:
                                                st.info("Binary file - download to view")
                                        else:
                                            st.info("Download file to view content")
                                        st.write("---")
            
            except Exception as e:
                st.error(f"Error reading Excel file: {str(e)}")
        
        else:
            # Instructions when no file is uploaded
            st.info("👈 Please upload an Excel file to get started")
            
            # Example format
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
                st.caption("Note: The 'Filename' column is optional. If not provided, filenames will be extracted from URLs.")
    
    with col2:
        st.subheader("ℹ️ Instructions")
        
        st.markdown("""
        1. **Upload Excel File**: Click 'Browse files' to upload
        2. **Configure Settings**: Set column names and options
        3. **Process URLs**: Click the button to download files to memory
        4. **Download Files**: Click links to save files to your device
        
        ### 📥 Download Modes
        - **Individual Files**: Download each file separately
        - **ZIP Archive**: All files in one ZIP download
        
        ### 🔒 Security & Privacy
        - Files are processed in memory
        - No files stored on server
        - Secure HTTPS connections
        - Your data is not persisted
        
        ### ⚡ Supported Files
        - Images (PNG, JPG, GIF)
        - Documents (PDF, DOC, XLS)
        - Archives (ZIP)
        - Text files (TXT, CSV)
        - And many more!
        """)
        
        # System info
        st.subheader("System Info")
        st.write(f"Current time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        if uploaded_file:
            st.write(f"File size: {uploaded_file.size:,} bytes")

if __name__ == "__main__":
    main()