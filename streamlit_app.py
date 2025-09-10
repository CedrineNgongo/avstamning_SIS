import streamlit as st
import tempfile
from pathlib import Path

# Byt modulnamn här om din fil heter något annat än avstamning_SIS.py
import avstamning_SIS as avm

st.set_page_config(page_title="Avstämning – SIS", page_icon="🧮", layout="centered")

st.title("Avstämning (SIS)")
st.write("Ladda upp **kontoutdrag (Bank)** och **bokföringslista** (CSV/XLSX/XLS).")

col1, col2 = st.columns(2)
with col1:
    bank_file = st.file_uploader("Kontoutdrag (Bank)", type=["csv", "xlsx", "xls"], key="bank")
with col2:
    bokf_file = st.file_uploader("Bokföringslista", type=["csv", "xlsx", "xls"], key="bokf")

run = st.button("Kör avstämning", type="primary", disabled=not (bank_file and bokf_file))

if run:
    try:
        with st.spinner("Bearbetar…"):
            # Spara upp till tempfiler så avstamning_SIS kan läsa via filväg
            with tempfile.TemporaryDirectory() as td:
                bank_path = Path(td) / f"bank_{bank_file.name}"
                bokf_path  = Path(td) / f"bokf_{bokf_file.name}"
                bank_path.write_bytes(bank_file.read())
                bokf_path.write_bytes(bokf_file.read())

                xlsx_bytes = avm.build_output_excel_bytes(str(bank_path), str(bokf_path))

        st.success("Klar! Ladda ned resultatet nedan.")
        st.download_button(
            label="Ladda ned Excel",
            data=xlsx_bytes,
            file_name="output_avstamning.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Något gick fel: {e}")
