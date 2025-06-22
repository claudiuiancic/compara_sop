
import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Comparare Proiecte Excel", layout="wide")

st.title("🗂️ Comparare Proiecte între două versiuni Excel")
st.write("Încarcă două fișiere Excel care conțin sheet-urile 'pipeline' și 'SOP'. Scriptul va detecta automat diferențele.")

file_old = st.file_uploader("Fișier Excel - Vechi", type=["xlsx"], key="old")
file_new = st.file_uploader("Fișier Excel - Nou", type=["xlsx"], key="new")

def read_clean_excel(file, sheet_name):
    # Citește toate datele din sheet
    raw_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

    # Găsește rândul care conține "Asgard ID" (headerul real)
    header_row_idx = None
    for idx, row in raw_df.iterrows():
        if "Asgard ID" in row.values:
            header_row_idx = idx
            break

    if header_row_idx is None:
        raise ValueError(f"Nu s-a găsit coloana 'Asgard ID' în sheetul '{sheet_name}'.")

    # Folosește acel rând ca header și reîncarcă DataFrame-ul curățat
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row_idx)

    # Elimină coloanele complet goale
    df.dropna(axis=1, how='all', inplace=True)
    # Elimină rândurile complet goale
    df.dropna(axis=0, how='all', inplace=True)

    return df

def load_data(file):
    # Încarcă ambele sheet-uri curate
    return {
        "pipeline": read_clean_excel(file, "pipeline"),
        "SOP": read_clean_excel(file, "SOP")
    }

def compare_data(old_df, new_df, id_col):
    # Obține seturile de ID-uri din fiecare versiune
    old_ids = set(old_df[id_col])
    new_ids = set(new_df[id_col])

    # Proiecte noi și eliminate
    added = new_df[~new_df[id_col].isin(old_ids)]
    removed = old_df[~old_df[id_col].isin(new_ids)]

    # Proiecte comune cu modificări
    modified = []
    common_ids = old_ids & new_ids
    for id_ in common_ids:
        old_row = old_df[old_df[id_col] == id_].iloc[0]
        new_row = new_df[new_df[id_col] == id_].iloc[0]
        diff = {}
        for col in old_df.columns:
            if col != id_col and col in new_df.columns:
                if pd.isna(old_row[col]) and pd.isna(new_row[col]):
                    continue
                if old_row[col] != new_row[col]:
                    diff[col] = (old_row[col], new_row[col])
        if diff:
            modified.append({"Asgard ID": id_, "Diferențe": diff})
    return added, modified, removed


if file_old and file_new:
    try:
        # Afișează informații generale despre fișiere
        st.header("📄 Informații generale")

        st.markdown(f"- **Fișier vechi:** `{file_old.name}`")
        st.markdown(f"- **Fișier nou:** `{file_new.name}`")

        # Încarcă datele
        data_old = load_data(file_old)
        data_new = load_data(file_new)

        # Statistici generale
        st.markdown(f"- **Nr. proiecte în pipeline vechi:** {len(data_old['pipeline'])}")
        st.markdown(f"- **Nr. proiecte în pipeline nou:** {len(data_new['pipeline'])}")
        st.markdown(f"- **Nr. proiecte în SOP vechi:** {len(data_old['SOP'])}")
        st.markdown(f"- **Nr. proiecte în SOP nou:** {len(data_new['SOP'])}")
        st.markdown(f"- **Raport generat la:** `{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}`")


        # Încarcă datele
        data_old = load_data(file_old)
        data_new = load_data(file_new)

        st.header("📊 Rezultatele comparației")

        # 1. Proiecte noi în pipeline
        st.subheader("1. Proiecte nou apărute în pipeline")
        added_pipeline, modified_pipeline, removed_pipeline = compare_data(
            data_old["pipeline"], data_new["pipeline"], "Asgard ID")
        st.write(added_pipeline)

        # 2. Proiecte modificate în pipeline
        st.subheader("2. Proiecte din pipeline care au suferit modificări")
        st.write(pd.DataFrame(modified_pipeline))

        # 3. Proiecte scoase din pipeline care nu au apărut în SOP
        removed_ids = set(removed_pipeline["Asgard ID"])
        sop_new_ids = set(data_new["SOP"]["Asgard ID"])
        removed_not_in_sop = removed_pipeline[~removed_pipeline["Asgard ID"].isin(sop_new_ids)]
        st.subheader("3. Proiecte scoase din pipeline care nu au apărut în SOP")
        st.write(removed_not_in_sop)

        # 4. Proiecte scoase din pipeline care au apărut în SOP
        removed_in_sop = removed_pipeline[removed_pipeline["Asgard ID"].isin(sop_new_ids)]
        st.subheader("4. Proiecte scoase din pipeline care au apărut în SOP")
        st.write(removed_in_sop)

        # 5. Proiecte apărute în SOP care nu erau în pipeline
        sop_added, _, _ = compare_data(data_old["pipeline"], data_new["SOP"], "Asgard ID")
        st.subheader("5. Proiecte apărute în SOP care nu erau în pipeline")
        st.write(sop_added)

        # 6. Proiecte din SOP cu modificări
        _, sop_modified, _ = compare_data(data_old["SOP"], data_new["SOP"], "Asgard ID")
        st.subheader("6. Proiecte din SOP la care s-au modificat parametri")
        st.write(pd.DataFrame(sop_modified))

        # 7. Proiecte scoase din SOP
        _, _, sop_removed = compare_data(data_old["SOP"], data_new["SOP"], "Asgard ID")
        st.subheader("7. Proiecte care au fost scoase din lista SOP")
        st.write(sop_removed)

    except Exception as e:
        st.error(f"Eroare la procesare: {e}")

        # Generare PDF la cerere
        from fpdf import FPDF
        import io

        st.header("📥 Export Raport în PDF")

        if st.button("📄 Descarcă raportul ca PDF"):
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.set_font("Arial", size=10)

            def write_title(title):
                pdf.set_font("Arial", style="B", size=12)
                pdf.cell(0, 10, title, ln=True)
                pdf.set_font("Arial", size=10)

            def write_dataframe(df):
                for _, row in df.iterrows():
                    line = ', '.join([f"{col}: {row[col]}" for col in df.columns])
                    pdf.multi_cell(0, 5, line)
                    pdf.ln(1)

            write_title("Informații generale")
            pdf.multi_cell(0, 5, f"Fișier vechi: {file_old.name}")
            pdf.multi_cell(0, 5, f"Fișier nou: {file_new.name}")
            pdf.multi_cell(0, 5, f"Nr. proiecte în pipeline vechi: {len(data_old['pipeline'])}")
            pdf.multi_cell(0, 5, f"Nr. proiecte în pipeline nou: {len(data_new['pipeline'])}")
            pdf.multi_cell(0, 5, f"Nr. proiecte în SOP vechi: {len(data_old['SOP'])}")
            pdf.multi_cell(0, 5, f"Nr. proiecte în SOP nou: {len(data_new['SOP'])}")
            pdf.multi_cell(0, 5, f"Raport generat la: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            pdf.ln(5)

            sections = [
                ("1. Proiecte nou apărute în pipeline", added_pipeline),
                ("2. Proiecte modificate în pipeline", pd.DataFrame(modified_pipeline)),
                ("3. Proiecte scoase din pipeline care nu au apărut în SOP", removed_not_in_sop),
                ("4. Proiecte scoase din pipeline care au apărut în SOP", removed_in_sop),
                ("5. Proiecte apărute în SOP care nu erau în pipeline", sop_added),
                ("6. Proiecte din SOP la care s-au modificat parametri", pd.DataFrame(sop_modified)),
                ("7. Proiecte care au fost scoase din lista SOP", sop_removed),
            ]

            for title, df in sections:
                write_title(title)
                if df.empty:
                    pdf.multi_cell(0, 5, "Nu există înregistrări.")
                else:
                    write_dataframe(df)
                pdf.ln(3)

            # Export PDF în memorie
            pdf_buffer = io.BytesIO()
            pdf.output(pdf_buffer)
            st.download_button(
                label="📥 Descarcă PDF",
                data=pdf_buffer.getvalue(),
                file_name="raport_proiecte.pdf",
                mime="application/pdf"
            )
