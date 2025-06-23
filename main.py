
import streamlit as st
import pandas as pd
from datetime import datetime
from fpdf import FPDF
import io
import os

st.set_page_config(page_title="SOP compare", layout="wide")

st.title("🗂️ Compara SOP si pipeline")
st.write("Descarcă din Asgard și încarcă două fișiere Excel care conțin sheet-urile 'PIPELINE' și 'SOP'. Unul vechi și unul actual (nou). Scriptul va detecta automat diferențele.")

file_old = st.file_uploader("Fișier Excel - Vechi", type=["xlsx"], key="old")
file_new = st.file_uploader("Fișier Excel - Nou", type=["xlsx"], key="new")

def read_clean_excel(file, sheet_name):
    raw_df = pd.read_excel(file, sheet_name=sheet_name, header=None)
    header_row_idx = None
    for idx, row in raw_df.iterrows():
        if "Asgard ID" in row.values:
            header_row_idx = idx
            break
    if header_row_idx is None:
        raise ValueError(f"Nu s-a găsit coloana 'Asgard ID' în sheetul '{sheet_name}'.")
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row_idx)
    df.dropna(axis=1, how='all', inplace=True)
    df.dropna(axis=0, how='all', inplace=True)
    return df

def load_data(file):
    return {
        "PIPELINE": read_clean_excel(file, "PIPELINE"),
        "SOP": read_clean_excel(file, "SOP")
    }

def compare_data(old_df, new_df, id_col):
    compare_cols = ["City", "Format", "Typology", "Estimated Opening Date"]
    old_ids = set(old_df[id_col])
    new_ids = set(new_df[id_col])

    added = new_df[~new_df[id_col].isin(old_ids)][[id_col] + compare_cols]
    removed = old_df[~old_df[id_col].isin(new_ids)][[id_col] + compare_cols]

    modified = []
    common_ids = old_ids & new_ids
    for id_ in common_ids:
        old_row = old_df[old_df[id_col] == id_].iloc[0]
        new_row = new_df[new_df[id_col] == id_].iloc[0]
        diff_lines = []

        row_data = {
            "Asgard ID": id_,
            "City": new_row.get("City", "")
        }
        for col in compare_cols:
            if col not in old_df.columns or col not in new_df.columns:
                continue
            old_val = old_row[col]
            new_val = new_row[col]

            if col == "Estimated Opening Date":
                try:
                    old_val_fmt = pd.to_datetime(old_val).strftime("%d.%m")
                    new_val_fmt = pd.to_datetime(new_val).strftime("%d.%m")
                except:
                    old_val_fmt = str(old_val)
                    new_val_fmt = str(new_val)
                if old_val_fmt != new_val_fmt:
                    diff_lines.append(f"{col}: {old_val_fmt} -> {new_val_fmt}")
            else:
                if pd.isna(old_val) and pd.isna(new_val):
                    continue
                if old_val != new_val:
                    diff_lines.append(f"{col}: {old_val} -> {new_val}")

        if diff_lines:
            row_data["Diferențe"] = "\n".join(diff_lines)
            modified.append(row_data)

    return added, modified, removed

if file_old and file_new:
    try:
        st.header("📄 Informații generale")

        st.markdown(f"- **Raport generat la:** `{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}`")
        st.markdown(f"- **Fișier vechi:** `{file_old.name}`")
        st.markdown(f"- **Fișier nou:** `{file_new.name}`")

        data_old = load_data(file_old)
        data_new = load_data(file_new)
        lung_pipeline_vechi = len(data_old['PIPELINE'])
        lung_pip_vechi_minusdoi = lung_pipeline_vechi - 2
        lung_sop_vechi = len(data_old['SOP'])
        lung_sop_vechi_minusdoi = lung_sop_vechi - 2
        lung_pipeline_noi = len(data_new['PIPELINE'])
        lung_pip_noi_minusdoi = lung_pipeline_noi - 2
        lung_sop_noi = len(data_new['SOP'])
        lung_sop_noi_minusdoi = lung_sop_noi - 2

        st.markdown(f"- **Nr. magazine în PIPELINE vechi: {lung_pip_vechi_minusdoi} --> noi: {lung_pip_noi_minusdoi}**")
        st.markdown(f"- **Nr. magazine în SOP vechi: {lung_sop_vechi_minusdoi} --> noi: {lung_sop_noi_minusdoi}**")

        st.header("📊 Rezultatele comparației")

        added_pipeline, modified_pipeline, removed_pipeline = compare_data(
            data_old["PIPELINE"], data_new["PIPELINE"], "Asgard ID"
        )

        # 1. Magazine nou apărute în PIPELINE
        st.subheader("1. Magazine nou apărute în PIPELINE")
        st.write(added_pipeline)

        # 2. Magazine din PIPELINE la care s-a modificat ceva
        st.subheader("2. Magazine din PIPELINE la care s-a modificat ceva")
        if modified_pipeline:
            df_mod = pd.DataFrame(modified_pipeline)
            cols = [col for col in ["Asgard ID", "City", "Diferențe"] if col in df_mod.columns]
            st.write(df_mod[cols])
        else:
            st.write("Nu există Magazine modificate în PIPELINE.")

        removed_ids = set(removed_pipeline["Asgard ID"])
        sop_new_ids = set(data_new["SOP"]["Asgard ID"])
        removed_not_in_sop = removed_pipeline[~removed_pipeline["Asgard ID"].isin(sop_new_ids)]
        removed_in_sop = removed_pipeline[removed_pipeline["Asgard ID"].isin(sop_new_ids)]

        # 3. Magazine scoase din PIPELINE care nu au apărut în SOP
        st.subheader("3. Magazine scoase din PIPELINE care nu au apărut în SOP")
        st.write(removed_not_in_sop)

        # 4. Magazine mutate din PIPELINE în SOP (contract semnat)
        st.subheader("4. Magazine mutate din PIPELINE în SOP (contract semnat)")
        st.write(removed_in_sop)

        # 5. Magazine apărute în SOP care nu erau în SOP vechi și nici în PIPELINE vechi
            # Seturi cu ID-urile existente anterior
        sop_vechi_ids = set(data_old["SOP"]["Asgard ID"])
        pipeline_vechi_ids = set(data_old["PIPELINE"]["Asgard ID"])

            # SOP nou
        sop_nou_df = data_new["SOP"]

            # Filtrare: doar ID-urile noi, care nu existau în niciuna dintre listele vechi
        sop_added_filtered = sop_nou_df[
            ~sop_nou_df["Asgard ID"].isin(sop_vechi_ids.union(pipeline_vechi_ids))
        ]

            # Afișare
        st.subheader("5. Magazine apărute în SOP care nu erau în SOP vechi și nici în PIPELINE vechi")
        st.write(sop_added_filtered)

        # 6. Magazine din SOP la care s-a modificat ceva
        _, sop_modified, _ = compare_data(data_old["SOP"], data_new["SOP"], "Asgard ID")
        st.subheader("6. Magazine din SOP la care s-a modificat ceva")
        if sop_modified:
            st.write(pd.DataFrame(sop_modified)[["Asgard ID", "Diferențe"]])
        else:
            st.write("Nu există Magazine modificate în SOP.")

        # 7. Magazine care au fost scoase din SOP (probabil deschise)
        _, _, sop_removed = compare_data(data_old["SOP"], data_new["SOP"], "Asgard ID")
        st.subheader("7. Magazine care au fost scoase din SOP (probabil deschise)")
        st.write(sop_removed)

        # ========================
        # de aici e partea de PDF
        # ========================

        st.header("📥 Export Raport în PDF")
        if st.button("📄 Descarcă raportul ca PDF"):
            pdf = FPDF(format='A4')
            pdf.add_page()
            pdf.add_font('DejaVu', '', 'DejaVuSans.ttf', uni=True)
            pdf.set_font('DejaVu', '', 10)

            def write_title(title):
                pdf.set_font("DejaVu", size=12)
                pdf.cell(0, 10, title, ln=True)
                pdf.set_font("DejaVu", size=10)

            def write_dataframe(df):
                for _, row in df.iterrows():
                    line = ', '.join([f"{col}: {row[col]}" for col in df.columns])
                    pdf.multi_cell(0, 5, line)
                    pdf.ln(1)

            write_title("Informații generale")
            pdf.multi_cell(0, 5, f"Raport generat la: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            pdf.multi_cell(0, 5, f"Fișier vechi: {file_old.name}")
            pdf.multi_cell(0, 5, f"Fișier nou: {file_new.name}")
            pdf.multi_cell(0, 5, f"Nr. magazine în PIPELINE vechi: {lung_pip_vechi_minusdoi} --> noi: {lung_pip_noi_minusdoi}")
            pdf.multi_cell(0, 5, f"Nr. magazine în SOP vechi: {lung_sop_vechi_minusdoi} --> noi: {lung_sop_noi_minusdoi}")
            pdf.ln(5)

            added_pipeline, modified_pipeline, removed_pipeline = compare_data(data_old["PIPELINE"], data_new["PIPELINE"], "Asgard ID")
            removed_ids = set(removed_pipeline["Asgard ID"])
            sop_new_ids = set(data_new["SOP"]["Asgard ID"])
            removed_not_in_sop = removed_pipeline[~removed_pipeline["Asgard ID"].isin(sop_new_ids)]
            removed_in_sop = removed_pipeline[removed_pipeline["Asgard ID"].isin(sop_new_ids)]
            sop_added, _, _ = compare_data(data_old["SOP"], data_new["SOP"], "Asgard ID")
            _, sop_modified, _ = compare_data(data_old["SOP"], data_new["SOP"], "Asgard ID")
            _, _, sop_removed = compare_data(data_old["SOP"], data_new["SOP"], "Asgard ID")

            sections = [
                ("1. Magazine nou apărute în PIPELINE", added_pipeline),
                ("2. Magazine din PIPELINE la care s-a modificat ceva",
                    pd.DataFrame(modified_pipeline)[["Asgard ID", "City", "Diferențe"]]
                    if modified_pipeline else pd.DataFrame()),
                ("3. Magazine scoase din PIPELINE care nu au apărut în SOP", removed_not_in_sop),
                ("4. Magazine mutate din PIPELINE în SOP (contract semnat)", removed_in_sop),
                ("5. Magazine apărute în SOP care nu erau în SOP vechi", sop_added),
                ("5. Magazine apărute în SOP care nu erau în SOP vechi și nici în PIPELINE vechi", sop_added_filtered),
                ("6. Magazine din SOP la care s-a modificat ceva", pd.DataFrame(sop_modified)[["Asgard ID", "Diferențe"]]) if sop_modified else ("6. Magazine din SOP la care s-au modificat parametri", pd.DataFrame()),
                ("7. Magazine care au fost scoase din SOP (probabil deschise)", sop_removed),
            ]

            for title, df in sections:
                write_title(title)
                if df.empty:
                    pdf.multi_cell(0, 5, "Nu există înregistrări.")
                else:
                    write_dataframe(df)
                pdf.ln(3)

            pdf_buffer = io.BytesIO()
            pdf_bytes = pdf.output(dest='S').encode('latin1')
            st.download_button(
                label="📥 Descarcă PDF",
                data=pdf_bytes,
                file_name=f"Raport comparatie SOP {datetime.now().strftime('%Y-%m-%d %H:%M')}.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Eroare la procesare: {e}")
