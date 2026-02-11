"""
UI Components Module
Streamlit UI components for the DQ Engine
All styling moved to external CSS file
"""

import streamlit as st
import pandas as pd
import traceback
from typing import Dict
from .config import AppConfig


class UIComponents:
    """Streamlit UI components"""
    
    @staticmethod
    def render_header():
        """Render application header"""
        st.markdown(
            f'<h1 class="purple-text">{AppConfig.APP_ICON} {AppConfig.APP_TITLE}</h1>',
            unsafe_allow_html=True
        )
        st.caption(f"Version {AppConfig.VERSION} | Enterprise-Grade Data Quality Validation")
        st.markdown("---")
    
    @staticmethod
    def render_sidebar():
        """Render sidebar with information"""
        with st.sidebar:
            st.header("📋 About")
            st.info(f"""
            **{AppConfig.APP_TITLE}**
            
            - 100% Dynamic Rules
            - Zero Hardcoded Logic
            - Multiple File Formats
            - Comprehensive Reports
            - Column-wise Annexures
            """)
            
            st.markdown("---")
            st.header("📊 Supported Features")
            st.markdown("""
            **Rule Types:**
            - Completeness
            - Uniqueness
            - Validity
            - Standardization
            
            **File Formats:**
            - CSV, Excel, TSV, JSON
            """)
            
            st.markdown("---")
            st.header("🔗 Quick Links")
            st.markdown("""
            - [Documentation](#)
            - [Support](#)
            - [GitHub](#)
            """)
    
    @staticmethod
    def render_file_format_help():
        """Render file format help section"""
        with st.expander("📋 Expected File Formats"):
            st.markdown("""
            ### Rules Dataset Format (CSV/Excel)
            
            **Required Columns:**
            - `column_name` or `column`: Target column name
            - `rule` or `rule_type`: Type of validation
            - `dimension` or `rule_category`: DQ dimension
            - `message`: Validation error message
            
            **Optional Columns:**
            - `expression`: Rule expression (regex, range, etc.)
            - `severity`: HIGH, MEDIUM, or LOW
            
            **Example:**
            ```csv
            column_name,rule,dimension,message,expression
            email,not_null,Completeness,Email is required,
            age,range,Validity,Age must be 0-120,"0,120"
            status,allowed_values,Validity,Invalid status,"Active,Inactive"
            ```
            
            ### JSON Rulebook Format
            ```json
            {
              "rules": [
                {
                  "column": "email",
                  "rule_type": "not_null",
                  "dimension": "Completeness",
                  "message": "Email is required",
                  "expression": null,
                  "severity": "HIGH"
                }
              ]
            }
            ```
            """)
    
    @staticmethod
    def render_results_dashboard(
        overall_score: float,
        results_df: pd.DataFrame,
        column_scores: Dict[str, float],
        dimension_scores: Dict[str, float]
    ):
        """Render results dashboard"""
        st.subheader("📊 Data Quality Dashboard")
        
        # Main metrics
        col1, col2, col3, col4 = st.columns(4)
        
        clean_count = len(results_df[results_df["Count of issues"] == 0])
        issue_count = len(results_df) - clean_count
        pass_columns = sum(1 for s in column_scores.values() if s == 100)
        
        col1.metric("Overall DQ Score", f"{overall_score}%")
        col2.metric("Clean Records", f"{clean_count:,}")
        col3.metric("Records with Issues", f"{issue_count:,}")
        col4.metric("Columns at 100%", f"{pass_columns}/{len(column_scores)}")
        
        # Score interpretation
        if overall_score >= 95:
            st.success("🎉 **Excellent!** Outstanding data quality.")
        elif overall_score >= 80:
            st.info("👍 **Good!** Minor improvements needed.")
        elif overall_score >= 60:
            st.warning("⚠️ **Fair!** Significant improvements required.")
        else:
            st.error("❌ **Poor!** Critical data quality issues detected.")
    
    @staticmethod
    def render_download_section(
        output_path,
        rulebook_path,
        total_annexures: int
    ):
        """Render download section"""
        st.subheader("📥 Download Reports")
        
        col1, col2 = st.columns(2)
        
        with col1:
            with open(output_path, "rb") as f:
                st.download_button(
                    "📊 Download Complete Excel Report",
                    f,
                    file_name=output_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help=f"Includes: DQ Score, Results, Summary, Dimension Analysis, and {total_annexures} Annexures"
                )
        
        with col2:
            with open(rulebook_path, "rb") as f:
                st.download_button(
                    "📋 Download Rulebook JSON",
                    f,
                    file_name=rulebook_path.name,
                    mime="application/json",
                    use_container_width=True,
                    help="Generated rulebook for reference and reuse"
                )
    
    @staticmethod
    def render_detailed_views(
        rulebook: Dict,
        results_df: pd.DataFrame,
        column_scores: Dict[str, float],
        dimension_scores: Dict[str, float]
    ):
        """Render detailed view tabs"""
        st.subheader("🔍 Detailed Analysis")
        
        tab1, tab2, tab3, tab4 = st.tabs([
            "📈 Column Scores",
            "🎯 Dimension Analysis",
            "📋 Rulebook",
            "📊 Results Preview"
        ])
        
        with tab1:
            UIComponents._render_column_scores(column_scores)
        
        with tab2:
            UIComponents._render_dimension_scores(dimension_scores)
        
        with tab3:
            st.json(rulebook)
            st.info(f"Total rules: {len(rulebook.get('rules', []))}")
        
        with tab4:
            UIComponents._render_results_preview(results_df)
    
    @staticmethod
    def _render_column_scores(column_scores: Dict[str, float]):
        """Render column scores table"""
        score_data = []
        for col, score in sorted(column_scores.items(), key=lambda x: x[1]):
            status = "✅ PASSED" if score == 100 else "❌ FAILED"
            score_data.append({
                "Column": col,
                "DQ Score (%)": score,
                "Status": status
            })
        
        score_df = pd.DataFrame(score_data)
        
        def color_status(val):
            if val == "✅ PASSED":
                return 'background-color: #C6EFCE; color: #006100'
            else:
                return 'background-color: #FFC7CE; color: #9C0006'
        
        st.dataframe(
            score_df.style.applymap(color_status, subset=['Status']),
            use_container_width=True,
            hide_index=True
        )
    
    @staticmethod
    def _render_dimension_scores(dimension_scores: Dict[str, float]):
        """Render dimension scores table"""
        if not dimension_scores:
            st.info("No dimension analysis available")
            return
        
        dim_data = []
        for dimension, score in dimension_scores.items():
            status = "✅ PASSED" if score == 100 else "❌ FAILED"
            dim_data.append({
                "Dimension": dimension,
                "DQ Score (%)": score,
                "Status": status
            })
        
        dim_df = pd.DataFrame(dim_data)
        st.dataframe(dim_df, use_container_width=True, hide_index=True)
    
    @staticmethod
    def _render_results_preview(results_df: pd.DataFrame):
        """Render results preview"""
        display_cols = [c for c in results_df.columns if not c.startswith("_")]
        
        # Show issues summary
        issues_df = results_df[results_df["Count of issues"] > 0]
        if len(issues_df) > 0:
            st.write(f"**Total records with issues: {len(issues_df):,}**")
            st.dataframe(issues_df[display_cols].head(100), use_container_width=True)
        else:
            st.success("🎉 No issues found! All records passed validation.")
            st.dataframe(results_df[display_cols].head(20), use_container_width=True)
    
    @staticmethod
    def render_welcome_screen():
        """Render welcome screen when no files uploaded"""
        st.info("👆 Please upload both Master Dataset and Rules Configuration to begin")
        
        st.markdown("---")
        st.subheader("🚀 How It Works")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown(
                '<div class="step-card">'
                '<h3>1️⃣ Upload Files</h3>'
                '<p>Master dataset (CSV/Excel/JSON)</p>'
                '<p>Rules dataset or JSON rulebook</p>'
                '<p>Supports multiple formats</p>'
                '</div>',
                unsafe_allow_html=True
            )
        
        with col2:
            st.markdown(
                '<div class="step-card">'
                '<h3>2️⃣ Generate Rulebook</h3>'
                '<p>Automatic rule mapping</p>'
                '<p>Dynamic validation logic</p>'
                '<p>Zero hardcoded rules</p>'
                '</div>',
                unsafe_allow_html=True
            )
        
        with col3:
            st.markdown(
                '<div class="step-card">'
                '<h3>3️⃣ Get Results</h3>'
                '<p>DQ Score dashboard</p>'
                '<p>Column-wise annexures</p>'
                '<p>Downloadable reports</p>'
                '</div>',
                unsafe_allow_html=True
            )
        
        st.markdown("---")
        
        # Feature highlights
        st.subheader("✨ Key Features")
        
        feature_col1, feature_col2 = st.columns(2)
        
        with feature_col1:
            st.markdown("""
            - ✅ 100% Dynamic Rules
            - ✅ Multiple File Formats
            - ✅ Comprehensive Validation
            - ✅ Column-wise Analysis
            """)
        
        with feature_col2:
            st.markdown("""
            - ✅ DQ Score Calculation
            - ✅ Dimension Analysis
            - ✅ Excel Report with Annexures
            - ✅ JSON Rulebook Export
            """)
    
    @staticmethod
    def render_error_details(error: Exception):
        """Render error details"""
        with st.expander("🔍 Show Detailed Error"):
            st.code(traceback.format_exc())
    
    @staticmethod
    def render_footer():
        """Render application footer"""
        st.markdown(
            '<div class="text-center margin-top-1">'
            '<p class="caption">✔ Enterprise Grade | ✔ 100% Dynamic | ✔ Zero Hardcoded Logic</p>'
            '<p class="caption">Powered by ' + AppConfig.APP_TITLE + ' v' + AppConfig.VERSION + '</p>'
            '</div>',
            unsafe_allow_html=True
        )