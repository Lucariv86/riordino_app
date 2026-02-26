import io
from datetime import datetime

import streamlit as st

try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

from reorder_engine import (
    ReorderConfig,
    parse_input_excel_fixed_columns,
    compute_reorders,
    export_to_excel,
    build_order_csv_bytes,
)

st.set_page_config(page_title="Riordini Magazzino", layout="wide")
st.title("Tool Riordini Magazzino (colonne fisse)")

uploaded = st.file_uploader(
    "Carica Excel (layout fisso: A=Marca, D=Codice, F=Descrizione, G=Gruppo, H=Scar.AC, I=UPA, L=Scar.AP, M=Giacenza)",
    type=["xlsx", "xls"],
)

col1, col2, col3 = st.columns(3)
with col1:
    coverage_days = st.slider("Copertura target (giorni)", min_value=30, max_value=180, value=30, step=5)
with col2:
    target_value = st.number_input("Target valore ordine (€) - opzionale", min_value=0.0, value=0.0, step=100.0)
with col3:
    three_dec_style = st.checkbox('Giacenza stile "-1,000" = -1', value=True)

generate_cosso_csv = st.checkbox("Genera CSV ordine Cosso (BLU/COR/UNI)", value=True)
generate_aspl_csv = st.checkbox("Genera CSV ordine ASPL", value=True)

run = st.button("Calcola riordino", type="primary")

if run:
    if not uploaded:
        st.error("Carica un file Excel prima di calcolare.")
        st.stop()

    config = ReorderConfig(
        coverage_days=int(coverage_days),
        target_value_eur=(float(target_value) if target_value and target_value > 0 else None),
        giacenza_three_decimals_style=bool(three_dec_style),
    )

    st.write("Nome file:", uploaded.name)

    try:
        df = parse_input_excel_fixed_columns(uploaded, config)
    except Exception as e:
        st.exception(e)
        st.stop()

    st.subheader("Anteprima input (parser colonne fisse)")
    st.dataframe(df.head(30), use_container_width=True)

    df_riordino, df_scartati, summary, warnings = compute_reorders(df, config)

    st.subheader("Risultato riordino")
    st.write(summary)
    if warnings:
        st.warning(" | ".join(warnings))

    st.dataframe(df_riordino, use_container_width=True, height=520)

    # Export in memoria per download
    output = io.BytesIO()
    import tempfile, os

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        export_to_excel(df_riordino, df_scartati, summary, warnings, tmp_path)
        with open(tmp_path, "rb") as f:
            output.write(f.read())
        output.seek(0)
    finally:
        try:
            os.remove(tmp_path)
        except Exception:
            pass

    st.download_button(
        label="Scarica Excel riordino",
        data=output,
        file_name="riordino_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    now = datetime.now()
    if ZoneInfo is not None:
        try:
            now = datetime.now(ZoneInfo("Europe/Rome"))
        except Exception:
            now = datetime.now()

    if generate_cosso_csv:
        cosso_csv_bytes = build_order_csv_bytes(
            df_riordino=df_riordino,
            brands=["BLU", "COR", "UNI"],
            code_col="CODICE ARTICOLO",
            qty_col="qty_to_order",
        )
        cosso_filename = f"ordine_cosso_{now.strftime('%Y%m%d')}.csv"
        st.download_button(
            label="Scarica CSV ordine Cosso",
            data=cosso_csv_bytes,
            file_name=cosso_filename,
            mime="text/csv",
        )

    if generate_aspl_csv:
        aspl_csv_bytes = build_order_csv_bytes(
            df_riordino=df_riordino,
            brands=["ASPL"],
            code_col="CODICE ARTICOLO",
            qty_col="qty_to_order",
        )
        aspl_filename = f"ordine_ASPL_{now.strftime('%Y%m%d_%H%M%S')}.csv"
        st.download_button(
            label="Scarica CSV ordine ASPL",
            data=aspl_csv_bytes,
            file_name=aspl_filename,
            mime="text/csv",
        )
