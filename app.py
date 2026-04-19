import io
import os
import glob
import zipfile

import streamlit as st
import pandas as pd
from invoice_processor import process_single_file, compute_summary, build_zip

# --- Page Config ---
st.set_page_config(
    page_title="Invoice Categorizer",
    page_icon="📄",
    layout="centered",
)

# --- Custom CSS ---
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    .main-header {
        text-align: center;
        padding: 0.5rem 0 1.5rem 0;
    }
    .main-header h1 {
        color: #1E293B;
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 0.25rem;
    }
    .main-header p {
        color: #64748B;
        font-size: 1.05rem;
        margin-top: 0;
    }

    div[data-testid="stMetric"] {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 0.5rem;
        padding: 0.75rem 1rem;
    }

    div.stDownloadButton > button {
        background-color: #2563EB !important;
        color: white !important;
        font-size: 1.05rem;
        padding: 0.75rem 2rem;
        border-radius: 0.5rem;
        width: 100%;
        border: none !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #1D4ED8 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- Header ---
st.markdown("""
<div class="main-header">
    <h1>Invoice Categorizer</h1>
    <p>Upload PDF invoices to automatically sort them by customer name and date</p>
</div>
""", unsafe_allow_html=True)

# --- Init Session State ---
for key in ['processing_complete', 'results', 'summary', 'zip_bytes', 'files_map']:
    if key not in st.session_state:
        st.session_state[key] = None

# --- Results Section ---
if st.session_state.processing_complete:
    summary = st.session_state.summary

    # Metrics
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total", f"{summary['total']:,}")
    c2.metric("Originals", f"{summary['ok_count']:,}")
    c3.metric("Duplicates", f"{summary['duplicate_count']:,}")
    c4.metric("Errors", f"{summary['error_count']:,}")

    st.divider()

    # Customer breakdown
    if summary['customer_counts']:
        st.subheader("Customer Breakdown")
        cust_df = pd.DataFrame(
            list(summary['customer_counts'].items()),
            columns=["Customer", "Invoices"]
        )
        st.dataframe(cust_df, use_container_width=True, hide_index=True, height=min(400, 35 * len(cust_df) + 38))

    # Date breakdown
    if summary['year_month_counts']:
        st.subheader("Date Breakdown")
        date_df = pd.DataFrame(
            list(summary['year_month_counts'].items()),
            columns=["Year / Month", "Invoices"]
        )
        st.dataframe(date_df, use_container_width=True, hide_index=True)

    # Errors
    if summary['error_count'] > 0:
        with st.expander(f"Errors ({summary['error_count']} files)"):
            st.dataframe(pd.DataFrame(summary['errors']), use_container_width=True, hide_index=True)

    # Output
    if summary['ok_count'] == 0:
        st.warning("No original invoices found. All files were either duplicates or could not be processed.")
    else:
        st.divider()
        st.subheader("Get Results")
        dl_tab, gdrive_tab = st.tabs(["Download ZIP", "Upload to Google Drive"])

        with dl_tab:
            zip_size = len(st.session_state.zip_bytes)
            if zip_size > 1024 * 1024:
                size_str = f"{zip_size / (1024 * 1024):.1f} MB"
            else:
                size_str = f"{zip_size / 1024:.0f} KB"

            st.download_button(
                label=f"Download Categorized Invoices ({size_str})",
                data=st.session_state.zip_bytes,
                file_name="categorized_invoices.zip",
                mime="application/zip",
                use_container_width=True,
            )

        with gdrive_tab:
            import json
            from gdrive_uploader import upload_to_drive

            # Check if credentials are saved in secrets
            has_saved = False
            saved_creds = None
            saved_folder = ""
            try:
                saved_creds = dict(st.secrets["gdrive"]["credentials"])
                saved_folder = st.secrets["gdrive"].get("folder_id", "")
                has_saved = True
            except Exception:
                pass

            if has_saved:
                st.success("Google Drive is connected.")
                if st.button("Upload to Google Drive", type="primary", use_container_width=True, key="btn_gdrive_up"):
                    ok_results = [r for r in st.session_state.results if r['status'] == 'ok']
                    progress = st.progress(0, text="Uploading to Google Drive...")
                    status = st.empty()

                    def up_progress(current, total, filename):
                        progress.progress(current / total, text=f"Uploading {current:,} / {total:,}")
                        status.caption(f"Uploading: {filename}")

                    try:
                        result = upload_to_drive(
                            credentials_json=saved_creds,
                            parent_folder_id=saved_folder,
                            ok_results=ok_results,
                            uploaded_files=st.session_state.get('files_map', {}),
                            progress_callback=up_progress,
                        )
                        progress.progress(1.0, text="Upload complete!")
                        status.empty()
                        st.success(f"Uploaded **{result['uploaded']:,}** invoices to Google Drive.")
                    except Exception as e:
                        st.error(f"Upload failed: {e}")
            else:
                st.info("**One-time setup** — after this, it's one click every time.")
                creds_file = st.file_uploader("Service Account JSON Key", type=["json"], key="gdrive_creds")
                folder_id = st.text_input("Google Drive Folder ID", placeholder="from the folder URL after /folders/")

                if st.button("Save & Upload", type="primary", disabled=(creds_file is None or not folder_id.strip()), use_container_width=True, key="btn_gdrive_up"):
                    try:
                        creds_json = json.loads(creds_file.read())
                    except Exception:
                        st.error("Invalid JSON key file.")
                        st.stop()

                    # Generate secrets config for permanent saving
                    secrets_toml = f'[gdrive]\nfolder_id = "{folder_id.strip()}"\n\n[gdrive.credentials]\n'
                    for k, v in creds_json.items():
                        secrets_toml += f'{k} = """{v}"""\n' if isinstance(v, str) else f"{k} = {json.dumps(v)}\n"

                    with st.expander("Save this so you never have to do it again", expanded=True):
                        st.markdown("Go to **Streamlit Cloud** > your app > **Settings** > **Secrets** and paste:")
                        st.code(secrets_toml, language="toml")

                    # Upload now
                    ok_results = [r for r in st.session_state.results if r['status'] == 'ok']
                    progress = st.progress(0, text="Uploading to Google Drive...")
                    status = st.empty()

                    def up_progress_first(current, total, filename):
                        progress.progress(current / total, text=f"Uploading {current:,} / {total:,}")
                        status.caption(f"Uploading: {filename}")

                    try:
                        result = upload_to_drive(
                            credentials_json=creds_json,
                            parent_folder_id=folder_id.strip(),
                            ok_results=ok_results,
                            uploaded_files=st.session_state.get('files_map', {}),
                            progress_callback=up_progress_first,
                        )
                        progress.progress(1.0, text="Upload complete!")
                        status.empty()
                        st.success(f"Uploaded **{result['uploaded']:,}** invoices to Google Drive.")
                    except Exception as e:
                        st.error(f"Upload failed: {e}")

    # Reset
    st.markdown("")
    if st.button("Process New Batch", use_container_width=True):
        for key in ['processing_complete', 'results', 'summary', 'zip_bytes', 'files_map']:
            st.session_state[key] = None
        st.rerun()

else:
    # --- Options ---
    skip_duplicates = st.toggle("Skip duplicate invoices", value=True,
                                help="Invoices marked as '*** DUPLICATE ***' will be skipped. Turn off to include all files.")

    # --- Input Tabs ---
    is_cloud = os.environ.get("STREAMLIT_SHARING_MODE") or os.environ.get("HOSTNAME", "").startswith("streamlit")

    if is_cloud:
        tab_upload, tab_gdrive, tab_onedrive = st.tabs(["Upload Files", "Google Drive", "OneDrive"])
        tab_folder = None
    else:
        tab_upload, tab_folder, tab_gdrive, tab_onedrive = st.tabs(
            ["Upload Files", "Local Folder", "Google Drive", "OneDrive"]
        )

    files_map = None

    # --- Upload Files ---
    with tab_upload:
        uploaded_files = st.file_uploader(
            "Upload invoice PDFs or a ZIP file",
            type=["pdf", "zip"],
            accept_multiple_files=True,
            help="Select PDF files, or upload a ZIP containing PDFs. Max: 1 GB.",
        )

        if uploaded_files:
            total_size = sum(f.size for f in uploaded_files)
            if total_size > 1024 * 1024:
                size_str = f"{total_size / (1024 * 1024):.1f} MB"
            else:
                size_str = f"{total_size / 1024:.0f} KB"
            zip_count = sum(1 for f in uploaded_files if f.name.lower().endswith('.zip'))
            pdf_count = len(uploaded_files) - zip_count
            parts = []
            if pdf_count:
                parts.append(f"{pdf_count:,} PDF{'s' if pdf_count != 1 else ''}")
            if zip_count:
                parts.append(f"{zip_count:,} ZIP{'s' if zip_count != 1 else ''}")
            st.info(f"**{' + '.join(parts)}** selected ({size_str})")

        if st.button("Categorize Invoices", type="primary", disabled=(not uploaded_files), use_container_width=True, key="btn_upload"):
            files_map = {}
            for f in uploaded_files:
                if f.name.lower().endswith('.zip'):
                    try:
                        with zipfile.ZipFile(io.BytesIO(f.read())) as zf:
                            for name in zf.namelist():
                                if name.lower().endswith('.pdf') and not name.startswith('__MACOSX'):
                                    basename = os.path.basename(name)
                                    if basename:
                                        files_map[basename] = zf.read(name)
                    except zipfile.BadZipFile:
                        st.error(f"Could not read ZIP file: {f.name}")
                else:
                    files_map[f.name] = f.read()

    # --- Local Folder ---
    if tab_folder:
        with tab_folder:
            st.caption("Drag a folder from Finder/Explorer into the box below, or paste the path.")
            folder_path = st.text_input(
                "Folder path",
                placeholder="C:\\Users\\James\\Invoices  or  /Users/james/Invoices",
                label_visibility="collapsed",
            )

            pdf_files = []
            if folder_path:
                folder_path = folder_path.strip().strip('"').strip("'")
                if os.path.isdir(folder_path):
                    pdf_files = sorted(set(
                        glob.glob(os.path.join(folder_path, '**', '*.PDF'), recursive=True)
                        + glob.glob(os.path.join(folder_path, '**', '*.pdf'), recursive=True)
                    ))
                    if pdf_files:
                        total_size = sum(os.path.getsize(f) for f in pdf_files)
                        if total_size > 1024 * 1024:
                            size_str = f"{total_size / (1024 * 1024):.1f} MB"
                        else:
                            size_str = f"{total_size / 1024:.0f} KB"
                        st.info(f"**{len(pdf_files):,}** PDF files found ({size_str})")
                    else:
                        st.warning("No PDF files found in this folder.")
                else:
                    st.error("Folder not found. Please check the path.")

            if st.button("Categorize Invoices", type="primary", disabled=(not pdf_files), use_container_width=True, key="btn_folder"):
                files_map = {}
                for f in pdf_files:
                    with open(f, 'rb') as fh:
                        files_map[os.path.basename(f)] = fh.read()

    # --- Google Drive ---
    with tab_gdrive:
        gd_link = st.text_input(
            "Google Drive Folder Link",
            placeholder="https://drive.google.com/drive/folders/...",
            help="Paste the shared folder link. Must be set to 'Anyone with the link'.",
        )

        if st.button("Fetch & Categorize", type="primary", disabled=(not gd_link.strip() if gd_link else True), use_container_width=True, key="btn_gdrive"):
            from cloud_sources import list_and_download_gdrive_public

            progress = st.progress(0, text="Downloading from Google Drive...")
            status = st.empty()

            def gd_progress(current, total, filename):
                progress.progress(current / total, text=f"Downloading {current:,} / {total:,}")
                status.caption(f"Downloading: {filename}")

            try:
                files_map = list_and_download_gdrive_public(gd_link.strip(), gd_progress)
                progress.progress(1.0, text=f"Downloaded {len(files_map):,} PDFs")
                status.empty()
                if not files_map:
                    st.warning("No PDF files found in the folder.")
                    files_map = None
            except Exception as e:
                st.error(f"Google Drive error: {e}")
                files_map = None

    # --- OneDrive ---
    with tab_onedrive:
        od_link = st.text_input(
            "OneDrive Sharing Link",
            placeholder="https://1drv.ms/f/...  (from Share → Copy link, not the address bar)",
            help="Use Share → Anyone with the link can view → Copy link. The URL must start with https://1drv.ms/ "
            "or be a onedrive.live.com link that contains redeem=. Address-bar links (?id=…) will not work.",
        )

        if st.button("Fetch & Categorize", type="primary", disabled=(not od_link.strip() if od_link else True), use_container_width=True, key="btn_onedrive"):
            from cloud_sources import list_and_download_onedrive_link

            progress = st.progress(0, text="Downloading from OneDrive...")
            status = st.empty()

            def od_progress(current, total, filename):
                progress.progress(current / total, text=f"Downloading {current:,} / {total:,}")
                status.caption(f"Downloading: {filename}")

            try:
                files_map = list_and_download_onedrive_link(od_link.strip(), od_progress)
                progress.progress(1.0, text=f"Downloaded {len(files_map):,} PDFs")
                status.empty()
                if not files_map:
                    st.warning("No PDF files found. Make sure the link is a public sharing link.")
                    files_map = None
            except Exception as e:
                st.error(f"OneDrive error: {e}")
                files_map = None

    # --- Process ---
    if files_map:
        progress = st.progress(0, text="Starting...")
        status = st.empty()

        results = []
        total = len(files_map)

        for i, (filename, pdf_bytes) in enumerate(files_map.items()):
            status.caption(f"Processing: {filename}")
            result = process_single_file(pdf_bytes, filename, skip_duplicates=skip_duplicates)
            results.append(result)
            progress.progress((i + 1) / total, text=f"{i + 1:,} / {total:,}")

        progress.progress(1.0, text="Complete!")
        status.empty()

        summary = compute_summary(results)
        ok_results = [r for r in results if r['status'] == 'ok']
        zip_bytes = build_zip(ok_results, files_map) if ok_results else b''

        st.session_state.results = results
        st.session_state.summary = summary
        st.session_state.zip_bytes = zip_bytes
        st.session_state.files_map = files_map
        st.session_state.processing_complete = True

        st.rerun()
