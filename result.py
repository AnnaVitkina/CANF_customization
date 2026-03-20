import os
import sys
import shutil
import gradio as gr
from pathlib import Path
from typing import Optional

# ---------------------------------------------------------------------------
# Project root detection (for Colab exec/notebook when __file__ or cwd is wrong):
# CANF_PROJECT_ROOT, __file__, common clone paths, scan /content/*, cwd-relative
# repo names, then walk parents. A directory qualifies if it has shipment_input.py
# or vocabulary.py (this repo).
# ---------------------------------------------------------------------------

_PROJECT_MARKERS = (
    "shipment_input.py",
    "vocabulary.py",
)


def _is_project_dir(p: Path) -> bool:
    if not p.is_dir():
        return False
    return any((p / name).is_file() for name in _PROJECT_MARKERS)


def _known_repo_paths_colab() -> list:
    """Typical Google Colab clone locations for this repo (underscore or hyphen)."""
    extra = []
    for path in ("/content/CANF_customization", "/content/CANF-customization"):
        p = Path(path)
        if _is_project_dir(p):
            extra.append(p.resolve())
    cwd = Path.cwd()
    for rel in (Path("CANF_customization"), Path("CANF-customization")):
        p = (cwd / rel).resolve()
        if _is_project_dir(p):
            extra.append(p)
    return extra


def _colab_clone_candidates() -> list:
    """Likely repo locations when cwd is /content but code lives in /content/<repo>."""
    extra = []
    for p in _known_repo_paths_colab():
        if p not in extra:
            extra.append(p)
    content = Path("/content")
    if content.is_dir():
        try:
            for child in sorted(content.iterdir()):
                if child.is_dir() and _is_project_dir(child):
                    if child.resolve() not in extra:
                        extra.append(child.resolve())
        except OSError:
            pass
    cwd = Path.cwd()
    for folder_name in ("CANF_customization", "CANF-customization", "Apple CANF customization"):
        p = (cwd / folder_name).resolve()
        if _is_project_dir(p) and p not in extra:
            extra.append(p)
    return extra


def get_project_root() -> Optional[Path]:
    """
    Find the folder that contains this project's modules (see _PROJECT_MARKERS).
    Set env CANF_PROJECT_ROOT if auto-detection fails.
    """
    candidates = []
    env_root = os.environ.get("CANF_PROJECT_ROOT", "").strip()
    if env_root:
        candidates.append(Path(env_root).resolve())
    try:
        candidates.append(Path(__file__).resolve().parent)
    except NameError:
        pass
    # Prefer known Colab clone paths, then scan all /content/* subdirs that look like this repo
    _seen_norm = {c.resolve() for c in candidates if hasattr(c, "resolve")}
    for p in _colab_clone_candidates():
        try:
            pr = p.resolve()
        except OSError:
            pr = p
        if pr not in _seen_norm:
            _seen_norm.add(pr)
            candidates.insert(0, p)
    cwd = Path.cwd().resolve()
    candidates.append(cwd)
    # Walk up from cwd (user may launch from a subfolder)
    p = cwd
    for _ in range(8):
        candidates.append(p)
        if p.parent == p:
            break
        p = p.parent
    seen = set()
    for cand in candidates:
        try:
            c = cand.resolve()
        except OSError:
            continue
        if c in seen:
            continue
        seen.add(c)
        if _is_project_dir(c):
            return c
    return None


def ensure_project_on_syspath() -> Optional[str]:
    """Insert project root at front of sys.path so `import shipment_input` always works."""
    root = get_project_root()
    if root is None:
        return None
    s = str(root)
    if s not in sys.path:
        sys.path.insert(0, s)
    return s


def setup_python_path():
    """Setup Python path to include the project directory for imports."""
    try:
        added = ensure_project_on_syspath()
        if added:
            print(f"📁 Added project root to Python path: {added}")
        else:
            # Last resort: cwd
            cwd = os.getcwd()
            if cwd and cwd not in sys.path:
                sys.path.insert(0, cwd)
                print(f"📁 Added cwd to Python path (shipment_input.py not found): {cwd}")
    except Exception as e:
        print(f"⚠️ Warning: Could not set up Python path: {e}")


setup_python_path()

def run_full_workflow_gradio(rate_card_file, etof_file, mismatch_report_files=None, 
                             ignore_rate_card_columns=None):
    """
    Main workflow for use in Gradio.
    Accepts uploaded files and user input; returns downloadable files and status messages.
    
    Workflow:
    1. Save uploaded files to input/ folder
    2. Process ETOF file (shipment_input.py)
    3. Process Rate Card file (rate_card_input.py) 
    4. Run vocabulary mapping (vocabulary.py) -> creates vocabulary_mapping.json and Filtered_Rate_Card_with_Conditions.json
    5. Run matching (matching.py) -> creates Matched_Shipments_with.json
    6. Run formatting (formatting.py) -> creates Matched_Shipments_formatted.json and .xlsx
    7. Save final results to output/ folder
    """
    status_messages = []
    errors = []
    warnings = []
    
    def log_status(msg, level="info"):
        """Log status messages with different levels"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_msg = f"[{timestamp}] {msg}"
        status_messages.append(formatted_msg)
        
        if level == "error":
            errors.append(msg)
        elif level == "warning":
            warnings.append(msg)
        
        print(formatted_msg)
    
    # Handle file input (Gradio may give strings or tempfile paths)
    def _handle_upload(uploaded, allow_multiple=False):
        if uploaded is None:
            return None if not allow_multiple else []
        if isinstance(uploaded, list):
            if not allow_multiple:
                return _handle_upload(uploaded[0] if uploaded else None, allow_multiple=False)
            result = []
            for item in uploaded:
                if item is None:
                    continue
                if hasattr(item, "name"):
                    result.append(item.name)
                elif isinstance(item, str):
                    result.append(item)
            return result if result else []
        if hasattr(uploaded, "name"):
            return uploaded.name
        if isinstance(uploaded, str):
            return uploaded
        return None if not allow_multiple else []
    
    # Convert all filepaths to correct types
    rate_card_path = _handle_upload(rate_card_file)
    etof_path = _handle_upload(etof_file)
    mismatch_report_path = _handle_upload(mismatch_report_files, allow_multiple=True)
    
    # Validate required fields
    if not etof_path:
        error_msg = "❌ Error: ETOF File is required."
        log_status(error_msg, "error")
        return None, error_msg
    
    if not rate_card_path:
        error_msg = "❌ Error: Rate Card File is required."
        log_status(error_msg, "error")
        return None, error_msg
    
    log_status("✅ Validation passed. Starting workflow...", "info")
    
    # Resolve project root (folder with shipment_input.py) so imports and folders stay consistent
    ensure_project_on_syspath()
    project_path = get_project_root()
    if project_path:
        script_dir = str(project_path)
    else:
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            script_dir = os.getcwd()
    log_status(f"📁 Working project directory: {script_dir}", "info")

    # Create output and input directories
    input_dir = os.path.join(script_dir, "input")
    output_dir = os.path.join(script_dir, "output")
    partly_df_dir = os.path.join(script_dir, "partly_df")
    
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(partly_df_dir, exist_ok=True)
    
    log_status(f"📁 Input folder: {input_dir}", "info")
    log_status(f"📁 Output folder: {output_dir}", "info")
    log_status(f"📁 Intermediate files folder: {partly_df_dir}", "info")
    
    # Copy uploaded files to input directory
    rate_card_filename = None
    etof_filename = None
    mismatch_report_filenames = []
    
    # Copy rate card file
    if rate_card_path:
        rate_card_filename = os.path.basename(rate_card_path)
        input_rc_path = os.path.join(input_dir, rate_card_filename)
        shutil.copy2(rate_card_path, input_rc_path)
        log_status(f"✓ Rate Card file saved: {rate_card_filename}", "info")
        if not os.path.exists(input_rc_path):
            error_msg = f"❌ Error: Failed to copy rate card file to {input_rc_path}"
            log_status(error_msg, "error")
            return None, error_msg
    
    # Copy ETOF file
    if etof_path:
        etof_filename = os.path.basename(etof_path)
        input_etof_path = os.path.join(input_dir, etof_filename)
        shutil.copy2(etof_path, input_etof_path)
        log_status(f"✓ ETOF file saved: {etof_filename}", "info")
        if not os.path.exists(input_etof_path):
            error_msg = f"❌ Error: Failed to copy ETOF file to {input_etof_path}"
            log_status(error_msg, "error")
            return None, error_msg
    
    # Copy mismatch report files (if provided)
    if mismatch_report_path:
        mismatch_files_list = mismatch_report_path if isinstance(mismatch_report_path, list) else [mismatch_report_path]
        for idx, mismatch_file_path in enumerate(mismatch_files_list):
            if mismatch_file_path:
                mismatch_filename = os.path.basename(mismatch_file_path)
                input_mismatch_path = os.path.join(input_dir, mismatch_filename)
                shutil.copy2(mismatch_file_path, input_mismatch_path)
                mismatch_report_filenames.append(mismatch_filename)
                log_status(f"✓ Mismatch Report file saved: {mismatch_filename}", "info")
    
    # Change to project directory so relative paths (input/, partly_df/) work
    original_cwd = os.getcwd()
    try:
        os.chdir(script_dir)
        
        # --- STEP 1: Configure ETOF Enrichment (if mismatch reports provided) ---
        if mismatch_report_filenames:
            try:
                from shipment_input import configure_enrichment
                mismatch_paths = mismatch_report_filenames if len(mismatch_report_filenames) > 1 else mismatch_report_filenames[0]
                configure_enrichment(mismatch_report_paths=mismatch_paths)
                log_status(f"✓ Enrichment configured with {len(mismatch_report_filenames)} mismatch report(s)", "info")
            except Exception as e:
                log_status(f"⚠️ Warning: Could not configure enrichment: {str(e)}", "warning")
        
        # --- STEP 2: Process ETOF File (shipment_input.py) ---
        try:
            from shipment_input import process_etof_file
            if etof_filename:
                log_status(f"📄 Processing ETOF file: {etof_filename}", "info")
                # process_etof_file prepends "input/" internally; pass filename only
                etof_df, etof_columns = process_etof_file(etof_filename)
                log_status(f"✓ ETOF processed: {etof_df.shape[0]} rows, {len(etof_columns)} columns", "info")
        except Exception as e:
            error_msg = f"❌ Error processing ETOF file: {str(e)}"
            log_status(error_msg, "error")
            import traceback
            log_status(traceback.format_exc(), "error")
            return None, "\n".join(status_messages)
        
        # --- STEP 3: Process Rate Card File (rate_card_input.py) ---
        try:
            from rate_card_input import process_rate_card, process_business_rules
            if rate_card_filename:
                log_status(f"📄 Processing Rate Card file: {rate_card_filename}", "info")
                # process_rate_card prepends "input/" internally; pass filename only
                rate_card_df, rate_card_columns, rate_card_conditions = process_rate_card(rate_card_filename)
                log_status(f"✓ Rate Card processed: {rate_card_df.shape[0]} rows, {len(rate_card_columns)} columns, {len(rate_card_conditions)} conditions", "info")
        except Exception as e:
            error_msg = f"❌ Error processing Rate Card file: {str(e)}"
            log_status(error_msg, "error")
            import traceback
            log_status(traceback.format_exc(), "error")
            return None, "\n".join(status_messages)
        
        # --- STEP 4: Vocabulary Mapping (vocabulary.py) ---
        try:
            from vocabulary import map_and_rename_columns
            
            # Parse ignore_rate_card_columns from comma-separated string to list
            ignore_columns_list = None
            if ignore_rate_card_columns and ignore_rate_card_columns.strip():
                ignore_columns_list = [col.strip() for col in ignore_rate_card_columns.split(',') if col.strip()]
                log_status(f"ℹ️ Ignoring rate card columns: {', '.join(ignore_columns_list)}", "info")
            
            log_status(f"🔤 Processing Vocabulary Mapping...", "info")
            log_status(f"   This step maps rate card columns to ETOF columns and creates vocabulary_mapping.json", "info")
            log_status(f"   Ignore Rate Card Columns: {ignore_columns_list if ignore_columns_list else 'None'}", "info")
            
            # vocabulary.map_and_rename_columns joins with input/; pass filenames only
            vocab_result = map_and_rename_columns(
                rate_card_file_path=rate_card_filename,
                etof_file_path=etof_filename,
                output_txt_path=os.path.join("partly_df", "column_mapping_results.txt"),
                ignore_rate_card_columns=ignore_columns_list
            )
            
            if vocab_result is None:
                error_msg = "❌ Error: Vocabulary mapping returned None"
                log_status(error_msg, "error")
                return None, "\n".join(status_messages)
            
            etof_renamed, _, _ = vocab_result
            
            if etof_renamed is not None and not etof_renamed.empty:
                log_status(f"✓ Vocabulary mapping completed: {etof_renamed.shape[0]} rows", "info")
                log_status(f"   Created: partly_df/vocabulary_mapping.json", "info")
                log_status(f"   Created: partly_df/Filtered_Rate_Card_with_Conditions.json", "info")
            else:
                log_status(f"⚠️ Warning: Vocabulary mapping completed but no data available", "warning")
                
        except Exception as e:
            error_msg = f"❌ Error in vocabulary mapping: {str(e)}"
            log_status(error_msg, "error")
            import traceback
            log_status(traceback.format_exc(), "error")
            return None, "\n".join(status_messages)
        
        # --- STEP 5: Matching (matching.py) ---
        try:
            from matching import run_matching_from_json
            
            log_status(f"🔍 Running Matching Process...", "info")
            log_status(f"   This step compares each shipment to rate card lanes and finds best matches", "info")
            
            matching_result = run_matching_from_json(
                rate_card_json_path=os.path.join("partly_df", "Filtered_Rate_Card_with_Conditions.json"),
                vocabulary_json_path=os.path.join("partly_df", "vocabulary_mapping.json"),
                output_dir="partly_df"
            )
            
            if matching_result and matching_result[0]:
                log_status(f"✓ Matching completed successfully", "info")
                log_status(f"   Created: partly_df/Matched_Shipments_with.json", "info")
            else:
                log_status(f"⚠️ Warning: Matching process did not produce output", "warning")
                
        except Exception as e:
            error_msg = f"❌ Error in matching: {str(e)}"
            log_status(error_msg, "error")
            import traceback
            log_status(traceback.format_exc(), "error")
            return None, "\n".join(status_messages)
        
        # --- STEP 6: Formatting (formatting.py) ---
        try:
            from formatting import run_formatting
            
            log_status(f"📝 Running Formatting Process...", "info")
            log_status(f"   This step adds 'Possible Best Match' column and reformats comments", "info")
            
            formatting_result = run_formatting(
                input_json_path=os.path.join("partly_df", "Matched_Shipments_with.json"),
                output_json_path=os.path.join("partly_df", "Matched_Shipments_formatted.json"),
                output_xlsx_path=os.path.join("partly_df", "Matched_Shipments_formatted.xlsx")
            )
            
            if formatting_result and formatting_result[0]:
                log_status(f"✓ Formatting completed successfully", "info")
                log_status(f"   Created: partly_df/Matched_Shipments_formatted.json", "info")
                log_status(f"   Created: partly_df/Matched_Shipments_formatted.xlsx", "info")
            else:
                log_status(f"⚠️ Warning: Formatting process did not produce output", "warning")
                
        except Exception as e:
            error_msg = f"❌ Error in formatting: {str(e)}"
            log_status(error_msg, "error")
            import traceback
            log_status(traceback.format_exc(), "error")
            return None, "\n".join(status_messages)
        
        # --- STEP 7: Copy final results to output folder ---
        final_json_path = os.path.join(output_dir, "Matched_Shipments_formatted.json")
        final_xlsx_path = os.path.join(output_dir, "Matched_Shipments_formatted.xlsx")
        
        try:
            source_json = os.path.join("partly_df", "Matched_Shipments_formatted.json")
            source_xlsx = os.path.join("partly_df", "Matched_Shipments_formatted.xlsx")
            
            if os.path.exists(source_json):
                shutil.copy2(source_json, final_json_path)
                log_status(f"✓ Final JSON saved to output folder: {final_json_path}", "info")
            
            if os.path.exists(source_xlsx):
                shutil.copy2(source_xlsx, final_xlsx_path)
                log_status(f"✓ Final Excel saved to output folder: {final_xlsx_path}", "info")
                final_file_path = final_xlsx_path
            else:
                final_file_path = final_json_path if os.path.exists(final_json_path) else None
                
        except Exception as e:
            log_status(f"⚠️ Warning: Could not copy final results: {str(e)}", "warning")
            final_file_path = None
        
    finally:
        os.chdir(original_cwd)
    
    # Prepare status summary
    status_summary = []
    status_summary.append("=" * 60)
    status_summary.append("WORKFLOW SUMMARY")
    status_summary.append("=" * 60)
    status_summary.append("")
    
    if final_file_path and os.path.exists(final_file_path):
        status_summary.append(f"✅ SUCCESS: Output file created")
        status_summary.append(f"   Location: {final_file_path}")
    else:
        status_summary.append(f"❌ Workflow did not complete successfully")
    
    status_summary.append("")
    
    if errors:
        status_summary.append(f"❌ ERRORS ({len(errors)}):")
        for i, error in enumerate(errors[:5], 1):
            status_summary.append(f"  {i}. {error}")
        if len(errors) > 5:
            status_summary.append(f"  ... and {len(errors) - 5} more errors")
        status_summary.append("")
    
    if warnings:
        status_summary.append(f"⚠️  WARNINGS ({len(warnings)}):")
        for i, warning in enumerate(warnings[:5], 1):
            status_summary.append(f"  {i}. {warning}")
        if len(warnings) > 5:
            status_summary.append(f"  ... and {len(warnings) - 5} more warnings")
        status_summary.append("")
    
    # Add key status messages
    key_messages = [msg for msg in status_messages if any(keyword in msg for keyword in 
                    ['✓', '❌', '⚠️', 'Error', 'Warning', 'SUCCESS', 'completed', 'failed'])]
    
    if key_messages:
        status_summary.append("Key Steps:")
        status_summary.append("-" * 60)
        status_summary.extend(key_messages[-15:])
    
    status_text = "\n".join(status_summary)
    return (final_file_path, status_text) if final_file_path and os.path.exists(final_file_path) else (None, status_text)


# ---- Gradio UI definition ----
with gr.Blocks(title="CANF Analyzer", theme=gr.themes.Soft()) as demo:
    gr.Markdown("# 📊 CANF Analyzer")
    gr.Markdown("### Process and match shipment data with rate card lanes")
    
    with gr.Accordion("📖 Instructions & Information", open=False):
        gr.Markdown("""
        ## How to Use This Workflow

        ### Google Colab (recommended)
        Do **not** `chdir` to `/content` before loading this app — stay inside the cloned repo, or set `CANF_PROJECT_ROOT`.
        ```text
        !git clone https://github.com/YOUR_ORG/CANF_customization.git  # or pull if already cloned
        %cd /content/CANF_customization
        !pip install -q gradio pandas openpyxl nest_asyncio
        !python result.py
        ```
        If you use `exec(open(...).read())` from `/content`, the app still tries to auto-detect `/content/CANF_customization`.
        Optional: `os.environ["CANF_PROJECT_ROOT"] = "/content/CANF_customization"` before `exec`.
        
        ### Step 1: Upload Required Files
        - **Rate Card File** (Required): Excel file containing rate card data (.xlsx)
        - **ETOF File** (Required): Excel file containing ETOF shipment data (.xlsx)
        
        ### Step 2: Upload Optional Files
        - **Mismatch Report File(s)** (Optional): Excel file(s) for ETOF enrichment
          - You can upload multiple mismatch report files
        
        ### Step 3: Configure Advanced Options (Optional)
        - **Ignore Rate Card Columns**: Enter comma-separated column names to exclude from processing
          - Example: `Column1, Column2, Column3`
          - **How it works**: These columns will be excluded from vocabulary mapping and matching processes.
            They are removed from the rate card dataframe before column mapping occurs, so they won't be
            matched to ETOF columns or used in the matching logic. This is useful for excluding columns
            that are not relevant for matching (e.g., internal notes, metadata columns, etc.)
        
        ### Step 4: Run Workflow
        - Click "Run Analyzer" button
        - Wait for processing to complete
        - Check the Status/Errors section for any issues
        - Download the Result files from the output folder
        
        ## Workflow Steps
        1. **File Processing**: Uploaded files are saved to `input/` folder
        2. **ETOF Processing**: ETOF file is processed (with optional enrichment from mismatch reports)
        3. **Rate Card Processing**: Rate card file is processed and business rules are extracted
        4. **Vocabulary Mapping**: Columns are mapped and renamed to standard names
           - Creates `partly_df/vocabulary_mapping.json`
           - Creates `partly_df/Filtered_Rate_Card_with_Conditions.json`
        5. **Matching**: Shipments are matched with rate card entries
           - Creates `partly_df/Matched_Shipments_with.json`
        6. **Formatting**: Adds "Possible Best Match" column and reformats comments
           - Creates `partly_df/Matched_Shipments_formatted.json`
           - Creates `partly_df/Matched_Shipments_formatted.xlsx`
        7. **Output**: Final results are copied to `output/` folder
        
        ## Output Files
        - **Matched_Shipments_formatted.json**: JSON file with all matched shipments and formatted comments
        - **Matched_Shipments_formatted.xlsx**: Excel file with the same data
        
        ## Troubleshooting
        - **Errors are shown in red** in the Status/Errors section
        - **Warnings are shown in yellow** - these may not prevent completion
        - Check that all required files are uploaded
        - Verify file formats are correct (.xlsx)
        - Ensure Rate Card and ETOF files have the expected structure
        """)
    
    gr.Markdown("---")
    gr.Markdown("### 📁 File Upload")
    gr.Markdown("**Required:** Rate Card File and ETOF File  |  **Optional:** Mismatch Report File(s)")
    
    with gr.Row():
        rate_card_input = gr.File(label="Rate Card File (.xlsx) *Required", file_types=[".xlsx", ".xls"])
        etof_input = gr.File(label="ETOF File (.xlsx) *Required", file_types=[".xlsx", ".xls"])
    
    with gr.Row():
        mismatch_report_input = gr.File(
            label="Mismatch Report File(s) (.xlsx) *Optional - for ETOF enrichment",
            file_types=[".xlsx", ".xls"],
            file_count="multiple"
        )
    
    # Ignore rate card columns input
    ignore_rate_card_columns_input = gr.Textbox(
        label="Ignore Rate Card Columns (Optional)",
        placeholder="Enter column names separated by commas (e.g., Column1, Column2, Column3)",
        info="Rate card columns to exclude from processing. Separate multiple columns with commas. These columns will be removed from the rate card before vocabulary mapping and matching."
    )
    
    gr.Markdown("---")
    launch_button = gr.Button("🚀 Run Analyzer", variant="primary", size="lg")
    
    with gr.Row():
        out = gr.File(label="📥 Result Files (Download Final Output)")
        status_output = gr.Textbox(
            label="📋 Status & Errors",
            lines=20,
            max_lines=30,
            interactive=False,
            placeholder="Workflow status and error messages will appear here..."
        )
    
    def launch_workflow(rate_card_file, etof_file, mismatch_report_files, ignore_rate_card_columns):
        try:
            result_file, status_text = run_full_workflow_gradio(
                rate_card_file=rate_card_file,
                etof_file=etof_file,
                mismatch_report_files=mismatch_report_files,
                ignore_rate_card_columns=ignore_rate_card_columns
            )
            return result_file, status_text
        except Exception as e:
            import traceback
            error_details = f"❌ CRITICAL ERROR:\n{str(e)}\n\nTraceback:\n{traceback.format_exc()}"
            return None, error_details
    
    launch_button.click(
        launch_workflow,
        inputs=[
            rate_card_input, etof_input, mismatch_report_input, ignore_rate_card_columns_input
        ],
        outputs=[out, status_output]
    )

if __name__ == "__main__":
    import sys
    
    # Create input, output, and partly_df folders when program starts
    ensure_project_on_syspath()
    _root = get_project_root()
    if _root:
        script_dir = str(_root)
    else:
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            script_dir = os.getcwd()
    
    input_dir = os.path.join(script_dir, "input")
    output_dir = os.path.join(script_dir, "output")
    partly_df_dir = os.path.join(script_dir, "partly_df")
    
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(partly_df_dir, exist_ok=True)
    
    print(f"📁 Created input folder: {input_dir}")
    print(f"📁 Created output folder: {output_dir}")
    print(f"📁 Created intermediate files folder: {partly_df_dir}")
    
    # Check if running in Colab
    in_colab = 'google.colab' in sys.modules
    
    if in_colab:
        print("🚀 Launching Gradio interface for Google Colab...")
        demo.launch(server_name="0.0.0.0", share=False, debug=False, show_error=True)
    else:
        print("🚀 Launching Gradio interface locally...")
        print(f"💡 Input files will be saved to: {input_dir}")
        print(f"💡 Output files will be saved to: {output_dir}")
        print(f"💡 Intermediate files will be saved to: {partly_df_dir}")
        demo.launch(server_name="127.0.0.1", share=False)
