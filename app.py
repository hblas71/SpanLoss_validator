import streamlit as st
import os
import sys
import tempfile
import pandas as pd

# Añadir directorio actual al path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scripts.analyze_fiber_pair import run_pipeline
from scripts.generate_docx import build_docx_report

st.set_page_config(page_title="Nokia OTDR Fiber Characterization", page_icon="📶", layout="wide")

st.title("📶 Caracterización Técnica de Fibras Ópticas (OTDR)")
st.caption("Norma ITU-T G.652.D | Análisis en 2 partes: Recorrido Completo y Salida de Central")

uploaded_file = st.file_uploader(
    "Arrastra la plantilla de vanos (.xlsm / .xlsx) o paquete ZIP con trazas .sor:",
    type=["xlsm", "xlsx", "zip", "sor"]
)

if uploaded_file is not None:
    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = os.path.join(temp_dir, uploaded_file.name)
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        with st.spinner("Procesando trazas reflectométricas y analizando eventos..."):
            out_dir = os.path.join(temp_dir, "output")
            os.makedirs(out_dir, exist_ok=True)
            results = run_pipeline([input_path], output_dir=out_dir)

        global_status = results.get("global_status", "NO APTO")
        excel_info = results.get("excel_info", {})
        nodo_a = excel_info.get("nodo_a", "NodoA")
        nodo_b = excel_info.get("nodo_b", "NodoB")

        st.subheader(f"Vano: {nodo_a} ↔ {nodo_b}")

        if global_status == "APTO":
            st.success("### 🟢 DICTAMEN GLOBAL: VANO APTO")
        else:
            st.error("### 🔴 DICTAMEN GLOBAL: VANO NO APTO (NO OK)")
            st.markdown("**Motivos de no conformidad / Puntos NOK:**")
            for reas in results.get("verdict_reasons", []):
                st.write(f"- ❌ {reas}")

        tab1, tab2, tab3 = st.tabs(["📊 Resumen Ejecutivo", "🔍 Auditoría de Eventos", "📥 Descarga de Informes"])

        with tab1:
            rows = []
            delta_loss = results.get("pair_symmetry", {}).get("delta_loss_db", 0.0)
            for t in results.get("traces_full_span", []):
                co = next((c for c in results.get("traces_central_exit", []) if c["direction"] == t["direction"]), {})
                rows.append({
                    "Hilo": t["direction"],
                    "Longitud (km)": f"{t['span_length_km']:.3f}",
                    "Pérdida Vano (dB)": f"{t['total_loss_db']:.2f}",
                    "Desbalance Par": f"{delta_loss:.2f} dB",
                    "Pérdida Central (27-150m)": f"{co.get('loss_db', 0.0):.2f} dB",
                    "Eventos NOK": len(t.get("out_of_spec_events", [])),
                    "Dictamen": t["status"]
                })
            df_exec = pd.DataFrame(rows)
            st.dataframe(df_exec, use_container_width=True)

        with tab2:
            st.markdown("#### Detalle de Eventos Fuera de Norma por Hilo:")
            for t in results.get("traces_full_span", []):
                st.write(f"**Hilo {t['direction']} ({t['filename']}):**")
                nok_events = t.get("out_of_spec_events", [])
                if not nok_events:
                    st.write("✔️ Todos los eventos están dentro de norma.")
                else:
                    for ev in nok_events:
                        e = ev["event"]
                        st.write(f"- PK **{e['distance_km']:.3f} km** ({e['distance_m']:.1f} m): {ev['reason']}")

        with tab3:
            st.markdown("#### Descarga de Archivos Generados:")
            c1, c2 = st.columns(2)

            docx_name = f"Informe_Caracterizacion_{nodo_a}-{nodo_b}.docx"
            docx_file = os.path.join(out_dir, docx_name)
            build_docx_report(results, docx_file)

            if os.path.exists(docx_file):
                with open(docx_file, "rb") as f_d:
                    c1.download_button(
                        label="📄 Descargar Informe Word (.docx)",
                        data=f_d,
                        file_name=docx_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            zip_path = excel_info.get("zip_path")
            if zip_path and os.path.exists(zip_path):
                with open(zip_path, "rb") as f_z:
                    c2.download_button(
                        label="📦 Descargar Trazas .sor (ZIP)",
                        data=f_z,
                        file_name=os.path.basename(zip_path),
                        mime="application/zip"
                    )
