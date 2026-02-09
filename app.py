"""
Enterprise Data Quality Rule Engine - Main Application
Modular architecture with specialized annexures for Uniqueness, Completeness, and Standardization
"""

import streamlit as st
import os
from pathlib import Path

# Import all modules
from modules.config import AppConfig
from modules.file_loader import FileLoaderService
from modules.rulebook_builder import RulebookBuilderService
from modules.rule_executor import RuleExecutorEngine
from modules.scoring_engine import ScoringService
from modules.report_generator import ExcelReportGenerator
from modules.ui_components import UIComponents
from modules.utils import setup_directories, save_uploaded_file, clean_temp_directory

# =========================
# PAGE CONFIGURATION
# =========================
st.set_page_config(
    page_title=AppConfig.APP_TITLE,
    page_icon=AppConfig.APP_ICON,
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================
# SETUP
# =========================
setup_directories()

# =========================
# MAIN APPLICATION
# =========================
def main():
    """Main application entry point"""
    
    # Render header and sidebar
    UIComponents.render_header()
    UIComponents.render_sidebar()
    
    # Main workflow
    st.markdown("---")
    
    # Step 1: File Upload Section
    st.subheader("📤 Step 1: Upload Files")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**📊 Master Dataset**")
        data_file = st.file_uploader(
            "Upload your data file",
            type=AppConfig.SUPPORTED_DATA_FORMATS,
            key="data_file",
            help="Supported: CSV, Excel, TSV, JSON"
        )
    
    with col2:
        st.markdown("**📋 DQ Rules Configuration**")
        upload_type = st.radio(
            "Upload Type",
            ["Rules Dataset (CSV/Excel)", "JSON Rulebook"],
            horizontal=True
        )
        
        if upload_type == "Rules Dataset (CSV/Excel)":
            rules_file = st.file_uploader(
                "Upload rules file",
                type=AppConfig.SUPPORTED_RULES_FORMATS,
                key="rules_csv",
                help="CSV or Excel with required columns"
            )
        else:
            rules_file = st.file_uploader(
                "Upload JSON rulebook",
                type=["json"],
                key="rules_json",
                help="Pre-built JSON rulebook"
            )
    
    # Help section
    UIComponents.render_file_format_help()
    
    st.markdown("---")
    
    # Step 2: Process Button
    if data_file and rules_file:
        
        if st.button("🚀 Run Data Quality Check", type="primary", use_container_width=True):
            
            try:
                # Clean previous temp files
                clean_temp_directory()
                
                # Progress tracking
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # ==========================================
                # STEP 1: SAVE UPLOADED FILES
                # ==========================================
                status_text.text("📂 Saving uploaded files...")
                progress_bar.progress(5)
                
                data_path = save_uploaded_file(data_file, AppConfig.TEMP_DIR)
                rules_path = save_uploaded_file(rules_file, AppConfig.TEMP_DIR)
                
                # ==========================================
                # STEP 2: LOAD MASTER DATASET
                # ==========================================
                status_text.text("📊 Loading master dataset...")
                progress_bar.progress(15)
                
                loader = FileLoaderService()
                data_df = loader.load_dataframe(data_path)
                all_columns = data_df.columns.tolist()
                
                st.info(f"✅ Loaded {len(data_df):,} records with {len(all_columns)} columns")
                
                # ==========================================
                # STEP 3: BUILD RULEBOOK
                # ==========================================
                status_text.text("🔧 Building rulebook...")
                progress_bar.progress(30)
                
                rulebook_builder = RulebookBuilderService()
                
                if upload_type == "JSON Rulebook":
                    rulebook = rulebook_builder.load_json_rulebook(rules_path)
                else:
                    rules_df = loader.load_dataframe(rules_path)
                    rulebook = rulebook_builder.build_from_rules_dataset(
                        rules_df, 
                        all_columns
                    )
                
                # Display rulebook metrics
                metrics_col1, metrics_col2, metrics_col3, metrics_col4 = st.columns(4)
                metrics_col1.metric("📊 Records", f"{len(data_df):,}")
                metrics_col2.metric("📋 Columns", len(all_columns))
                metrics_col3.metric("⚙️ Rules Loaded", len(rulebook.get("rules", [])))
                
                # Count unique columns (including combinations)
                unique_targets = set()
                for rule in rulebook["rules"]:
                    if "column" in rule:
                        unique_targets.add(rule["column"])
                    elif "columns" in rule:
                        unique_targets.add(" + ".join(rule["columns"]))
                
                metrics_col4.metric("🎯 Unique Targets", len(unique_targets))
                
                # ==========================================
                # STEP 4: EXECUTE RULES
                # ==========================================
                status_text.text("✅ Executing validation rules...")
                progress_bar.progress(50)
                
                executor = RuleExecutorEngine(data_df, rulebook)
                results_df = executor.execute_all_rules()
                
                # Get combination duplicates for specialized reporting
                combination_duplicates = executor.get_combination_duplicates()
                
                # ==========================================
                # STEP 5: CALCULATE SCORES
                # ==========================================
                status_text.text("📊 Calculating DQ scores...")
                progress_bar.progress(70)
                
                scoring = ScoringService()
                overall_score = scoring.calculate_overall_score(results_df)
                column_scores = scoring.calculate_column_scores(results_df, all_columns)
                dimension_scores = scoring.calculate_dimension_scores(results_df)
                
                # ==========================================
                # STEP 6: GENERATE REPORTS
                # ==========================================
                status_text.text("💾 Generating Excel report with specialized annexures...")
                progress_bar.progress(85)
                
                report_gen = ExcelReportGenerator(
                    results_df=results_df,
                    rulebook=rulebook,
                    all_columns=all_columns,
                    column_scores=column_scores,
                    overall_score=overall_score,
                    dimension_scores=dimension_scores,
                    duplicate_combinations=combination_duplicates
                )
                
                output_path = report_gen.generate_report(AppConfig.OUTPUT_DIR)
                rulebook_path = report_gen.save_rulebook_json(AppConfig.OUTPUT_DIR, rulebook)
                
                progress_bar.progress(100)
                status_text.text("✅ Processing complete!")
                
                st.success("✅ Data Quality Check Completed Successfully!")
                
                st.markdown("---")
                
                # ==========================================
                # DISPLAY RESULTS
                # ==========================================
                UIComponents.render_results_dashboard(
                    overall_score=overall_score,
                    results_df=results_df,
                    column_scores=column_scores,
                    dimension_scores=dimension_scores
                )
                
                st.markdown("---")
                
                # ==========================================
                # SPECIALIZED ANNEXURES INFO
                # ==========================================
                st.subheader("📑 Report Includes Specialized Annexures")
                
                info_col1, info_col2, info_col3 = st.columns(3)
                
                with info_col1:
                    st.info("""
                    **🔍 Uniqueness Issues**
                    - Duplicate combinations
                    - All duplicate records
                    - Combination details
                    """)
                
                with info_col2:
                    st.info("""
                    **📝 Completeness Issues**
                    - Missing values
                    - Null/empty fields
                    - Incomplete records
                    """)
                
                with info_col3:
                    st.info("""
                    **📐 Standardization Issues**
                    - Format violations
                    - Pattern mismatches
                    - Non-standard values
                    """)
                
                st.markdown("---")
                
                # ==========================================
                # DOWNLOAD SECTION
                # ==========================================
                UIComponents.render_download_section(
                    output_path=output_path,
                    rulebook_path=rulebook_path,
                    total_annexures=len(all_columns) + 3  # +3 for specialized annexures
                )
                
                st.markdown("---")
                
                # ==========================================
                # DETAILED VIEWS
                # ==========================================
                UIComponents.render_detailed_views(
                    rulebook=rulebook,
                    results_df=results_df,
                    column_scores=column_scores,
                    dimension_scores=dimension_scores
                )
                
                # ==========================================
                # COMBINATION DUPLICATES PREVIEW
                # ==========================================
                if combination_duplicates:
                    st.markdown("---")
                    st.subheader("🔍 Combination Duplicates Preview")
                    
                    for combo_key, dup_groups in combination_duplicates.items():
                        with st.expander(f"**{combo_key}** - {len(dup_groups)} duplicate groups found"):
                            for idx, group in enumerate(dup_groups[:5], 1):  # Show first 5 groups
                                st.write(f"Duplicate Group {idx}: Rows {group}")
                            
                            if len(dup_groups) > 5:
                                st.info(f"... and {len(dup_groups) - 5} more duplicate groups. See full report in Excel.")
                
            except Exception as e:
                st.error(f"❌ Error during processing: {str(e)}")
                UIComponents.render_error_details(e)
    
    else:
        # ==========================================
        # WELCOME SCREEN
        # ==========================================
        UIComponents.render_welcome_screen()
        
        # Additional info about combination rules
        st.markdown("---")
        st.subheader("💡 New Feature: Combination Uniqueness")
        
        st.info("""
        **Check uniqueness across multiple columns!**
        
        In your rules file, specify combinations like:
        - `column_name`: `user_name + address`
        - `rule`: `uniqueness`
        - `dimension`: `Uniqueness`
        - `message`: `Combination should identify a single unique entity`
        
        The system will automatically detect duplicates across the combination and create a specialized annexure.
        """)
    
    # Footer
    st.markdown("---")
    UIComponents.render_footer()


if __name__ == "__main__":
    main()
