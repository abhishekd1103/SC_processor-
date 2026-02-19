"""
SC Result Processing Application - OPTIMIZED
=================================================
Lightweight Streamlit app for Short Circuit analysis
Optimized for fast deployment and minimal dependencies
"""

import streamlit as st
import pandas as pd
from typing import Tuple
from io import BytesIO


class SCProcessingEngine:
    """Short Circuit Result Processing Engine"""

    def __init__(self, hvcb_df: pd.DataFrame, lvcb_df: pd.DataFrame, bus_df: pd.DataFrame):
        self.hvcb_df = hvcb_df.copy()
        self.lvcb_df = lvcb_df.copy()
        self.bus_df = bus_df.copy()

    def process_all(self) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Process all three sheets"""
        processed_hv = self._process_cb(self.hvcb_df, cb_type="HVCB")
        processed_lv = self._process_cb(self.lvcb_df, cb_type="LVCB")
        processed_bus = self._process_bus(self.bus_df)
        return processed_hv, processed_lv, processed_bus

    def _process_cb(self, df: pd.DataFrame, cb_type: str) -> pd.DataFrame:
        """Process Circuit Breaker sheets"""
        required_cols = ["Rated ip", "Rated Ib Sym", "Ip (Simulated)", 'I"k (Simulated)']
        self._validate_columns(df, required_cols, cb_type)

        df['Rated I"k (Calculated)'] = df["Rated Ib Sym"]
        df["Breaking Margin"] = df['Rated I"k (Calculated)'] - df['I"k (Simulated)']
        df["Bracing Margin"] = df["Rated ip"] - df["Ip (Simulated)"]
        
        df["Breaking Utilization %"] = (df['I"k (Simulated)'] / df['Rated I"k (Calculated)'] * 100).round(2)
        df["Bracing Utilization %"] = (df["Ip (Simulated)"] / df["Rated ip"] * 100).round(2)
        
        df["Device Duty Status"] = df.apply(lambda row: self._evaluate(row['Rated I"k (Calculated)'], row['I"k (Simulated)']), axis=1)
        df["Bracing Status"] = df.apply(lambda row: self._evaluate(row["Rated ip"], row["Ip (Simulated)"]), axis=1)
        df["Equipment Type"] = cb_type

        return df

    def _process_bus(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process Bus sheet"""
        required_cols = ["Nominal kV", "Rated Peak", "Ip (Simulated)", 'I"k (Simulated)', "Type"]
        self._validate_columns(df, required_cols, "Bus")

        df['Rated I"k (Calculated)'] = df.apply(lambda row: self._calculate_bus_rated_ik(row), axis=1)
        df["Breaking Margin"] = df['Rated I"k (Calculated)'] - df['I"k (Simulated)']
        df["Bracing Margin"] = df["Rated Peak"] - df["Ip (Simulated)"]
        
        df["Breaking Utilization %"] = df.apply(
            lambda row: round((row['I"k (Simulated)'] / row['Rated I"k (Calculated)']) * 100, 2)
            if pd.notna(row['Rated I"k (Calculated)']) and row['Rated I"k (Calculated)'] != 0 else None, axis=1
        )
        
        df["Bracing Utilization %"] = df.apply(
            lambda row: round((row["Ip (Simulated)"] / row["Rated Peak"]) * 100, 2)
            if pd.notna(row["Rated Peak"]) and row["Rated Peak"] != 0 else None, axis=1
        )
        
        df["Device Duty Status"] = df.apply(lambda row: self._evaluate_bus_breaking(row), axis=1)
        df["Bracing Status"] = df.apply(lambda row: self._evaluate(row["Rated Peak"], row["Ip (Simulated)"]), axis=1)
        df["Equipment Type"] = "Bus"

        return df

    def _calculate_bus_rated_ik(self, row):
        """Calculate Bus Rated I"k"""
        if pd.isna(row["Rated Peak"]) and row["Type"] == "Other":
            return None

        nominal_kv = row["Nominal kV"]
        rated_peak = row["Rated Peak"]

        if nominal_kv > 1:
            return round(rated_peak / 2.5, 3)

        lv_mapping = {143: 65, 110: 50, 75.6: 36, 52.5: 25, 32: 16, 20: 10}
        return lv_mapping.get(rated_peak, round(rated_peak / 2.5, 3))

    def _evaluate_bus_breaking(self, row):
        """Evaluate bus breaking status"""
        if pd.isna(row["Rated Peak"]) and row["Type"] == "Other":
            return "*PASS"
        return self._evaluate(row['Rated I"k (Calculated)'], row['I"k (Simulated)'])

    @staticmethod
    def _evaluate(rated, simulated):
        """Generic PASS/FAIL evaluation"""
        if pd.isna(rated) or pd.isna(simulated):
            return "DATA ERROR"
        return "PASS" if rated >= simulated else "FAIL"

    @staticmethod
    def _validate_columns(df, required_cols, sheet_name):
        """Validate required columns"""
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            raise ValueError(f"Missing columns in {sheet_name}: {', '.join(missing)}")


@st.cache_data
def create_excel_download(hvcb_df, lvcb_df, bus_df) -> bytes:
    """Create Excel file - CACHED for performance"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        hvcb_df.to_excel(writer, sheet_name='HVCB_Processed', index=False)
        lvcb_df.to_excel(writer, sheet_name='LVCB_Processed', index=False)
        bus_df.to_excel(writer, sheet_name='Bus_Processed', index=False)
    return output.getvalue()


def style_dataframe(df: pd.DataFrame):
    """Lightweight styling for dataframes"""
    def color_status(val):
        colors = {
            "PASS": "background-color: #d4edda",
            "FAIL": "background-color: #f8d7da",
            "*PASS": "background-color: #fff3cd",
            "DATA ERROR": "background-color: #f5c6cb"
        }
        return colors.get(val, "")
    
    styled = df.style
    if "Device Duty Status" in df.columns:
        styled = styled.applymap(color_status, subset=["Device Duty Status"])
    if "Bracing Status" in df.columns:
        styled = styled.applymap(color_status, subset=["Bracing Status"])
    return styled


def main():
    """Main Application"""
    
    st.set_page_config(page_title="SC Processor", page_icon="⚡", layout="wide")
    
    st.title("⚡ Short Circuit Result Processor")
    st.markdown("Upload Excel file with **HVCB**, **LVCB**, and **Bus** sheets for automated processing")
    
    # Sidebar
    with st.sidebar:
        st.header("📂 Upload File")
        uploaded_file = st.file_uploader("SC Results Excel", type=['xlsx', 'xls'])
        
        st.divider()
        st.markdown("**Legend:**\n- ✅ PASS\n- ❌ FAIL\n- ⚠️ *PASS")
    
    if uploaded_file:
        try:
            # Load data
            excel_file = pd.ExcelFile(uploaded_file)
            sheets = excel_file.sheet_names
            
            required = ["HVCB", "LVCB", "Bus"]
            missing = [s for s in required if s not in sheets]
            
            if missing:
                st.error(f"❌ Missing sheets: {', '.join(missing)}")
                return
            
            hvcb_raw = pd.read_excel(uploaded_file, sheet_name="HVCB")
            lvcb_raw = pd.read_excel(uploaded_file, sheet_name="LVCB")
            bus_raw = pd.read_excel(uploaded_file, sheet_name="Bus")
            
            # Process
            with st.spinner("Processing..."):
                engine = SCProcessingEngine(hvcb_raw, lvcb_raw, bus_raw)
                hvcb_processed, lvcb_processed, bus_processed = engine.process_all()
            
            st.success("✅ Processing completed!")
            
            # Metrics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("HVCB", len(hvcb_processed))
                fails = ((hvcb_processed["Device Duty Status"] == "FAIL") | 
                        (hvcb_processed["Bracing Status"] == "FAIL")).sum()
                st.metric("Failures", fails, delta_color="inverse")
            
            with col2:
                st.metric("LVCB", len(lvcb_processed))
                fails = ((lvcb_processed["Device Duty Status"] == "FAIL") | 
                        (lvcb_processed["Bracing Status"] == "FAIL")).sum()
                st.metric("Failures", fails, delta_color="inverse")
            
            with col3:
                st.metric("Bus", len(bus_processed))
                fails = ((bus_processed["Device Duty Status"] == "FAIL") | 
                        (bus_processed["Bracing Status"] == "FAIL")).sum()
                st.metric("Failures", fails, delta_color="inverse")
            
            # Results
            tab1, tab2, tab3 = st.tabs(["HVCB", "LVCB", "Bus"])
            
            with tab1:
                st.dataframe(style_dataframe(hvcb_processed), use_container_width=True, height=400)
            
            with tab2:
                st.dataframe(style_dataframe(lvcb_processed), use_container_width=True, height=400)
            
            with tab3:
                st.dataframe(style_dataframe(bus_processed), use_container_width=True, height=400)
            
            # Download
            st.divider()
            excel_data = create_excel_download(hvcb_processed, lvcb_processed, bus_processed)
            st.download_button(
                "📥 Download Processed Excel",
                data=excel_data,
                file_name="SC_Results_Processed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            with st.expander("Details"):
                st.exception(e)
    
    else:
        st.info("👆 Upload file to start")


if __name__ == "__main__":
    main()
