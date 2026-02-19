"""
SC Result Processing Application
=================================================
Streamlit app for processing Short Circuit analysis results
Supports HVCB, LVCB, and Bus sheets with automated evaluation
"""

import streamlit as st
import pandas as pd
from typing import Tuple
from io import BytesIO
import traceback


class SCProcessingEngine:
    """
    Short Circuit Result Processing Engine
    Processes HVCB, LVCB, and Bus sheets with rating evaluation logic
    """

    def __init__(self, hvcb_df: pd.DataFrame,
                 lvcb_df: pd.DataFrame,
                 bus_df: pd.DataFrame):

        self.hvcb_df = hvcb_df.copy()
        self.lvcb_df = lvcb_df.copy()
        self.bus_df = bus_df.copy()

    # ==========================================================
    # PUBLIC ENTRY
    # ==========================================================

    def process_all(self) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Process all three sheets and return processed dataframes"""

        processed_hv = self._process_cb(self.hvcb_df, cb_type="HVCB")
        processed_lv = self._process_cb(self.lvcb_df, cb_type="LVCB")
        processed_bus = self._process_bus(self.bus_df)

        return processed_hv, processed_lv, processed_bus

    # ==========================================================
    # COMMON CB PROCESSOR (HVCB + LVCB)
    # ==========================================================

    def _process_cb(self, df: pd.DataFrame, cb_type: str) -> pd.DataFrame:
        """Process Circuit Breaker sheets (HV and LV)"""

        required_cols = [
            "Rated ip",
            "Rated Ib Sym",
            "Ip (Simulated)",
            'I"k (Simulated)'
        ]

        self._validate_columns(df, required_cols, cb_type)

        # Rated I"k = Rated Ib Sym
        df['Rated I"k (Calculated)'] = df["Rated Ib Sym"]

        # Margin Columns
        df["Breaking Margin"] = df['Rated I"k (Calculated)'] - df['I"k (Simulated)']
        df["Bracing Margin"] = df["Rated ip"] - df["Ip (Simulated)"]

        # Utilization %
        df["Breaking Utilization %"] = (
            df['I"k (Simulated)'] / df['Rated I"k (Calculated)'] * 100
        ).round(2)

        df["Bracing Utilization %"] = (
            df["Ip (Simulated)"] / df["Rated ip"] * 100
        ).round(2)

        # Status Evaluation
        df["Device Duty Status"] = df.apply(
            lambda row: self._evaluate(row['Rated I"k (Calculated)'],
                                       row['I"k (Simulated)']),
            axis=1
        )

        df["Bracing Status"] = df.apply(
            lambda row: self._evaluate(row["Rated ip"],
                                       row["Ip (Simulated)"]),
            axis=1
        )

        df["Equipment Type"] = cb_type

        return df

    # ==========================================================
    # BUS PROCESSOR
    # ==========================================================

    def _process_bus(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process Bus sheet with special Type='Other' logic"""

        required_cols = [
            "Nominal kV",
            "Rated Peak",
            "Ip (Simulated)",
            'I"k (Simulated)',
            "Type"
        ]

        self._validate_columns(df, required_cols, "Bus")

        df['Rated I"k (Calculated)'] = df.apply(
            lambda row: self._calculate_bus_rated_ik(row),
            axis=1
        )

        # Margin Columns
        df["Breaking Margin"] = (
            df['Rated I"k (Calculated)'] - df['I"k (Simulated)']
        )

        df["Bracing Margin"] = (
            df["Rated Peak"] - df["Ip (Simulated)"]
        )

        # Utilization % (handle None values)
        df["Breaking Utilization %"] = df.apply(
            lambda row: round((row['I"k (Simulated)'] / row['Rated I"k (Calculated)']) * 100, 2)
            if pd.notna(row['Rated I"k (Calculated)']) and row['Rated I"k (Calculated)'] != 0
            else None,
            axis=1
        )

        df["Bracing Utilization %"] = df.apply(
            lambda row: round((row["Ip (Simulated)"] / row["Rated Peak"]) * 100, 2)
            if pd.notna(row["Rated Peak"]) and row["Rated Peak"] != 0
            else None,
            axis=1
        )

        # Status Evaluation
        df["Device Duty Status"] = df.apply(
            lambda row: self._evaluate_bus_breaking(row),
            axis=1
        )

        df["Bracing Status"] = df.apply(
            lambda row: self._evaluate(row["Rated Peak"],
                                       row["Ip (Simulated)"]),
            axis=1
        )

        df["Equipment Type"] = "Bus"

        return df

    # ==========================================================
    # BUS RATED I"k LOGIC
    # ==========================================================

    def _calculate_bus_rated_ik(self, row):
        """Calculate Bus Rated I"k based on voltage level and rated peak"""

        # SPECIAL RULE: Type="Other" with no Rated Peak
        if pd.isna(row["Rated Peak"]) and row["Type"] == "Other":
            return None

        nominal_kv = row["Nominal kV"]
        rated_peak = row["Rated Peak"]

        # HV Bus Rule (> 1 kV)
        if nominal_kv > 1:
            return round(rated_peak / 2.5, 3)

        # LV Bus Mapping
        lv_mapping = {
            143: 65,
            110: 50,
            75.6: 36,
            52.5: 25,
            32: 16,
            20: 10
        }

        if rated_peak in lv_mapping:
            return lv_mapping[rated_peak]

        # Fallback for unmapped LV values
        return round(rated_peak / 2.5, 3)

    # ==========================================================
    # BUS SPECIAL DUTY EVALUATION
    # ==========================================================

    def _evaluate_bus_breaking(self, row):
        """Evaluate bus breaking status with special *PASS rule"""

        # SPECIAL RULE: Type="Other" with no Rated Peak gets *PASS
        if pd.isna(row["Rated Peak"]) and row["Type"] == "Other":
            return "*PASS"

        return self._evaluate(
            row['Rated I"k (Calculated)'],
            row['I"k (Simulated)']
        )

    # ==========================================================
    # GENERIC EVALUATION FUNCTION
    # ==========================================================

    @staticmethod
    def _evaluate(rated, simulated):
        """Generic PASS/FAIL evaluation"""

        if pd.isna(rated) or pd.isna(simulated):
            return "DATA ERROR"

        return "PASS" if rated >= simulated else "FAIL"

    # ==========================================================
    # VALIDATION
    # ==========================================================

    @staticmethod
    def _validate_columns(df, required_cols, sheet_name):
        """Validate required columns exist in dataframe"""

        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            raise ValueError(
                f"Missing columns in {sheet_name} sheet: {', '.join(missing_cols)}"
            )


# ==========================================================
# STREAMLIT APPLICATION
# ==========================================================

def create_excel_download(hvcb_df, lvcb_df, bus_df) -> BytesIO:
    """Create Excel file with multiple sheets for download"""
    
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        hvcb_df.to_excel(writer, sheet_name='HVCB_Processed', index=False)
        lvcb_df.to_excel(writer, sheet_name='LVCB_Processed', index=False)
        bus_df.to_excel(writer, sheet_name='Bus_Processed', index=False)
    
    output.seek(0)
    return output


def apply_status_styling(df: pd.DataFrame) -> pd.DataFrame:
    """Apply color coding to status columns"""
    
    def highlight_status(val):
        if val == "PASS":
            return 'background-color: #d4edda; color: #155724'
        elif val == "FAIL":
            return 'background-color: #f8d7da; color: #721c24'
        elif val == "*PASS":
            return 'background-color: #fff3cd; color: #856404'
        elif val == "DATA ERROR":
            return 'background-color: #f5c6cb; color: #721c24'
        return ''
    
    # Apply styling to status columns if they exist
    styled_df = df.style
    
    if "Device Duty Status" in df.columns:
        styled_df = styled_df.applymap(
            highlight_status, 
            subset=["Device Duty Status"]
        )
    
    if "Bracing Status" in df.columns:
        styled_df = styled_df.applymap(
            highlight_status, 
            subset=["Bracing Status"]
        )
    
    return styled_df


def display_summary_metrics(hvcb_df, lvcb_df, bus_df):
    """Display summary statistics for processed results"""
    
    st.subheader("📊 Processing Summary")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("HVCB Devices", len(hvcb_df))
        hvcb_fails = (hvcb_df["Device Duty Status"] == "FAIL").sum() + \
                     (hvcb_df["Bracing Status"] == "FAIL").sum()
        st.metric("HVCB Failures", hvcb_fails)
    
    with col2:
        st.metric("LVCB Devices", len(lvcb_df))
        lvcb_fails = (lvcb_df["Device Duty Status"] == "FAIL").sum() + \
                     (lvcb_df["Bracing Status"] == "FAIL").sum()
        st.metric("LVCB Failures", lvcb_fails)
    
    with col3:
        st.metric("Bus Elements", len(bus_df))
        bus_fails = (bus_df["Device Duty Status"] == "FAIL").sum() + \
                    (bus_df["Bracing Status"] == "FAIL").sum()
        st.metric("Bus Failures", bus_fails)


def main():
    """Main Streamlit application"""
    
    # Page Configuration
    st.set_page_config(
        page_title="SC Result Processor",
        page_icon="⚡",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Header
    st.title("⚡ Short Circuit Result Processing Engine")
    st.markdown("""
    Upload your raw Short Circuit analysis Excel file containing **HVCB**, **LVCB**, and **Bus** sheets.
    The app will automatically:
    - Calculate rated I"k values
    - Evaluate device duty and bracing status (PASS/FAIL/*PASS)
    - Calculate margins and utilization percentages
    - Generate downloadable processed Excel file
    """)
    
    st.divider()
    
    # Sidebar
    with st.sidebar:
        st.header("📂 File Upload")
        uploaded_file = st.file_uploader(
            "Upload SC Results Excel File",
            type=['xlsx', 'xls'],
            help="Excel file must contain sheets: HVCB, LVCB, Bus"
        )
        
        st.divider()
        
        st.header("ℹ️ Required Columns")
        
        with st.expander("HVCB / LVCB Sheets"):
            st.markdown("""
            - Rated ip
            - Rated Ib Sym
            - Ip (Simulated)
            - I"k (Simulated)
            """)
        
        with st.expander("Bus Sheet"):
            st.markdown("""
            - Nominal kV
            - Rated Peak
            - Ip (Simulated)
            - I"k (Simulated)
            - Type
            """)
        
        st.divider()
        
        st.header("📋 Legend")
        st.markdown("""
        - **PASS**: Rated ≥ Simulated ✅
        - **FAIL**: Rated < Simulated ❌
        - ***PASS**: Special case (Bus Type='Other') ⚠️
        """)
    
    # Main Processing Area
    if uploaded_file is not None:
        try:
            # Read Excel file
            with st.spinner("Reading Excel file..."):
                excel_file = pd.ExcelFile(uploaded_file)
                available_sheets = excel_file.sheet_names
                
                st.success(f"✅ File loaded successfully! Found sheets: {', '.join(available_sheets)}")
            
            # Check for required sheets
            required_sheets = ["HVCB", "LVCB", "Bus"]
            missing_sheets = [sheet for sheet in required_sheets if sheet not in available_sheets]
            
            if missing_sheets:
                st.error(f"❌ Missing required sheets: {', '.join(missing_sheets)}")
                st.stop()
            
            # Load dataframes
            hvcb_raw = pd.read_excel(uploaded_file, sheet_name="HVCB")
            lvcb_raw = pd.read_excel(uploaded_file, sheet_name="LVCB")
            bus_raw = pd.read_excel(uploaded_file, sheet_name="Bus")
            
            st.info(f"📋 Loaded: {len(hvcb_raw)} HVCB, {len(lvcb_raw)} LVCB, {len(bus_raw)} Bus records")
            
            # Process data
            with st.spinner("Processing Short Circuit results..."):
                engine = SCProcessingEngine(hvcb_raw, lvcb_raw, bus_raw)
                hvcb_processed, lvcb_processed, bus_processed = engine.process_all()
            
            st.success("✅ Processing completed successfully!")
            
            # Display summary metrics
            display_summary_metrics(hvcb_processed, lvcb_processed, bus_processed)
            
            st.divider()
            
            # Tabbed display of results
            tab1, tab2, tab3 = st.tabs(["🔌 HVCB Results", "🔌 LVCB Results", "🚌 Bus Results"])
            
            with tab1:
                st.subheader("High Voltage Circuit Breaker Results")
                st.dataframe(
                    apply_status_styling(hvcb_processed),
                    use_container_width=True,
                    height=400
                )
                
                # Show failures only
                hvcb_failures = hvcb_processed[
                    (hvcb_processed["Device Duty Status"] == "FAIL") | 
                    (hvcb_processed["Bracing Status"] == "FAIL")
                ]
                if not hvcb_failures.empty:
                    with st.expander(f"⚠️ Show {len(hvcb_failures)} HVCB Failures"):
                        st.dataframe(hvcb_failures, use_container_width=True)
            
            with tab2:
                st.subheader("Low Voltage Circuit Breaker Results")
                st.dataframe(
                    apply_status_styling(lvcb_processed),
                    use_container_width=True,
                    height=400
                )
                
                # Show failures only
                lvcb_failures = lvcb_processed[
                    (lvcb_processed["Device Duty Status"] == "FAIL") | 
                    (lvcb_processed["Bracing Status"] == "FAIL")
                ]
                if not lvcb_failures.empty:
                    with st.expander(f"⚠️ Show {len(lvcb_failures)} LVCB Failures"):
                        st.dataframe(lvcb_failures, use_container_width=True)
            
            with tab3:
                st.subheader("Bus Results")
                st.dataframe(
                    apply_status_styling(bus_processed),
                    use_container_width=True,
                    height=400
                )
                
                # Show failures only
                bus_failures = bus_processed[
                    (bus_processed["Device Duty Status"] == "FAIL") | 
                    (bus_processed["Bracing Status"] == "FAIL")
                ]
                if not bus_failures.empty:
                    with st.expander(f"⚠️ Show {len(bus_failures)} Bus Failures"):
                        st.dataframe(bus_failures, use_container_width=True)
            
            st.divider()
            
            # Download section
            st.subheader("💾 Download Processed Results")
            
            col1, col2 = st.columns([3, 1])
            
            with col1:
                excel_data = create_excel_download(
                    hvcb_processed, 
                    lvcb_processed, 
                    bus_processed
                )
                
                st.download_button(
                    label="📥 Download Processed Excel File",
                    data=excel_data,
                    file_name="SC_Results_Processed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                if st.button("🔄 Process New File", use_container_width=True):
                    st.rerun()
        
        except Exception as e:
            st.error("❌ An error occurred during processing")
            st.error(f"Error details: {str(e)}")
            
            with st.expander("🐛 Show Full Error Traceback"):
                st.code(traceback.format_exc())
    
    else:
        # Welcome screen
        st.info("👆 Upload an Excel file from the sidebar to begin processing")
        
        st.subheader("🎯 Features")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **Circuit Breaker Processing:**
            - Automatic Rated I"k calculation
            - Breaking and bracing margin analysis
            - Device utilization percentage
            - PASS/FAIL status evaluation
            """)
        
        with col2:
            st.markdown("""
            **Bus Processing:**
            - HV/LV voltage-based logic
            - Special Type='Other' handling
            - Peak current evaluation
            - *PASS status for special cases
            """)
        
        st.divider()
        
        st.subheader("📖 How to Use")
        st.markdown("""
        1. **Prepare your Excel file** with three sheets: HVCB, LVCB, Bus
        2. **Upload the file** using the sidebar file uploader
        3. **Review results** in the tabbed interface
        4. **Check failures** using the expandable sections
        5. **Download processed file** with all calculations and evaluations
        """)


if __name__ == "__main__":
    main()
