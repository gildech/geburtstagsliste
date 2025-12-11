import io

import pandas as pd
import numpy as np
import streamlit as st


def prepare_dataframe(df_raw: pd.DataFrame, target_year: int = 2026) -> pd.DataFrame:
    """
    Wendet die gleiche Logik an wie im Notebook:
    - Datum parsen und nach Monat/Tag sortieren
    - Spalten umbenennen / bereinigen
    - Alter im Zieljahr berechnen
    """
    df = df_raw.copy()

    # Datum parsen (wie im Notebook, dayfirst=True)
    if "Geburtsdatum" in df.columns:
        df["Geburtsdatum"] = pd.to_datetime(
            df["Geburtsdatum"], errors="coerce", dayfirst=True
        )

        # Nach Monat/Tag sortieren (Jahr ignorieren für Sortierung)
        df["_geburtstag_sort"] = df["Geburtsdatum"].dt.strftime("%m-%d")
        df = df.sort_values("_geburtstag_sort").reset_index(drop=True)
        df = df.drop(columns=["_geburtstag_sort"])

        # Alter im Zieljahr
        df[f"Alter {target_year}"] = target_year - df["Geburtsdatum"].dt.year

    # Unnötige Spalten entfernen (falls vorhanden)
    df = df.drop(columns=["Kontakte", "Anredeart"], errors="ignore")

    # Spalten wie im Notebook umbenennen / harmonisieren
    rename_map = {
        "Strasse (Korr.)": "Strasse",
        "PLZ (Korr.)": "PLZ",
        "Ort (Korr.)": "Ort",
        # Korrespondenzsprache-Harmonisierung
        "Korresp.sprache": "Korrespondenzsprache",
        "Korresp. Sprache": "Korrespondenzsprache",
        "Korrespondenz Sprache": "Korrespondenzsprache",
    }
    existing_rename = {k: v for k, v in rename_map.items() if k in df.columns}
    if existing_rename:
        df = df.rename(columns=existing_rename)

    return df


def build_geburtstagsliste_excel(
    df: pd.DataFrame,
    target_year: int = 2026,
    include_no_date_sheet: bool = True,
) -> tuple[bytes, dict]:
    """
    Erzeugt eine Excel-Datei im Speicher (Bytes) mit:
    - einem Tabellenblatt pro Monat
    - runden Geburtstagen grün markiert
    - Ehrenmitglieder/-präsidenten/Prinzenrolle gelb hervorgehoben
    Gibt zusätzlich einfache Statistik zurück.
    """
    df = df.copy()

    # Daten mit und ohne Geburtsdatum trennen
    df_with_date = pd.DataFrame()
    df_no_date = pd.DataFrame()
    if "Geburtsdatum" in df.columns:
        df_no_date = df[df["Geburtsdatum"].isna()].copy()
        df_with_date = df[df["Geburtsdatum"].notna()].copy()

        # Zusätzliche Hilfsspalten (nur für Zeilen mit Geburtsdatum)
        df_with_date["Monat"] = df_with_date["Geburtsdatum"].dt.month
        df_with_date["Tag"] = df_with_date["Geburtsdatum"].dt.day

    monat_namen = {
        1: "Januar",
        2: "Februar",
        3: "März",
        4: "April",
        5: "Mai",
        6: "Juni",
        7: "Juli",
        8: "August",
        9: "September",
        10: "Oktober",
        11: "November",
        12: "Dezember",
    }

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Monatsblätter für alle Kontakte mit Geburtsdatum
        monate_mit_daten = False
        if (
            not df_with_date.empty
            and "Monat" in df_with_date.columns
            and df_with_date["Monat"].notna().any()
        ):
            for monat in range(1, 13):
                df_monat = df_with_date[df_with_date["Monat"] == monat].copy()
                if df_monat.empty:
                    continue
                monate_mit_daten = True

                # Nach Tag sortieren
                if "Tag" in df_monat.columns:
                    df_monat = df_monat.sort_values("Tag")

                # Anzeigeformat für Geburtsdatum als 'dd.mm.yyyy'
                if "Geburtsdatum" in df_monat.columns:
                    df_monat["Geburtsdatum (TT.MM.JJJJ)"] = df_monat[
                        "Geburtsdatum"
                    ].dt.strftime("%d.%m.%Y")

                # Mitgliedschaft-Spalte finden
                mitgliedschaft_col = None
                for mgl_col in ["Mitgliedschaft", "mitgliedschaft", "Mitglied", "mitglied"]:
                    if mgl_col in df_monat.columns:
                        mitgliedschaft_col = mgl_col
                        break

                # Korrespondenzsprache-Spalte robust finden (inkl. abgekürzter Varianten)
                corr_lang_col = None
                for col in df_monat.columns:
                    col_norm = str(col).strip().lower().replace(" ", "")
                    if col_norm in ("korrespondenzsprache", "korrespsprache"):
                        corr_lang_col = col
                        break
                    if ("korresp" in col_norm or "korrespondenz" in col_norm) and "sprache" in col_norm:
                        corr_lang_col = col
                        break

                # Spaltenauswahl wie im Notebook
                desired_cols = [
                    "Vorname",
                    "Nachname",
                    "Geburtsdatum (TT.MM.JJJJ)",
                    "Firma",
                    "Strasse",
                    "PLZ",
                    "Ort",
                    f"Alter {target_year}",
                ]
                if mitgliedschaft_col:
                    desired_cols.append(mitgliedschaft_col)
                if corr_lang_col:
                    # Korrespondenzsprache direkt nach Ort
                    desired_cols.insert(7, corr_lang_col)

                sheet_cols = [c for c in desired_cols if c in df_monat.columns]
                # Falls Korrespondenzspalte nicht in desired_cols war, aber existiert, trotzdem anhängen
                if corr_lang_col and corr_lang_col not in sheet_cols:
                    sheet_cols.append(corr_lang_col)
                # Fallback: alle Spalten, falls keine der gewünschten da ist
                if not sheet_cols:
                    sheet_cols = list(df_monat.columns)

                export_df = df_monat[sheet_cols]

                # In Excel-Sheet schreiben
                sheet_name = monat_namen.get(monat, f"Monat_{monat}")
                export_df.to_excel(writer, sheet_name=sheet_name, index=False)

                workbook = writer.book
                worksheet = writer.sheets[sheet_name]

                # Formate
                date_format_german = workbook.add_format({"num_format": "dd.mm.yyyy"})
                green_highlight_format = workbook.add_format(
                    {"bg_color": "#C6EFCE", "font_color": "#006100"}
                )
                yellow_highlight_format = workbook.add_format(
                    {"bg_color": "#FFF2CC", "font_color": "#7F6000"}
                )
                bold_format = workbook.add_format({"bold": True})

                # Header fett
                worksheet.set_row(0, None, bold_format)
                # Spaltenbreite
                worksheet.set_column(0, len(sheet_cols) - 1, 18)

                # Runde Geburtstage grün hervorheben
                alter_col_name = f"Alter {target_year}"
                if alter_col_name in sheet_cols:
                    alter_col_idx = sheet_cols.index(alter_col_name)
                    col_letter = chr(ord("A") + alter_col_idx)
                    row_end = len(export_df)
                    rng = f"{col_letter}2:{col_letter}{row_end + 1}"
                    worksheet.conditional_format(
                        rng,
                        {
                            "type": "formula",
                            "criteria": f"=AND(MOD({col_letter}2,10)=0,NOT(ISBLANK({col_letter}2)))",
                            "format": green_highlight_format,
                        },
                    )

                # Datumsformat
                if "Geburtsdatum (TT.MM.JJJJ)" in sheet_cols:
                    geb_de_col_idx = sheet_cols.index("Geburtsdatum (TT.MM.JJJJ)")
                    worksheet.set_column(
                        geb_de_col_idx, geb_de_col_idx, 18, date_format_german
                    )

                # Ehrenmitglied / Ehrenpräsident / Prinzenrolle gelb hervorheben
                if mitgliedschaft_col and mitgliedschaft_col in sheet_cols:
                    mg_col_idx = sheet_cols.index(mitgliedschaft_col)
                    mg_col_letter = chr(ord("A") + mg_col_idx)
                    row_end = len(export_df)
                    for rolle in ["Ehrenmitglied", "Ehrenpräsident", "Prinzenrolle"]:
                        worksheet.conditional_format(
                            f"A2:{chr(ord('A') + len(sheet_cols) - 1)}{row_end + 1}",
                            {
                                "type": "formula",
                                "criteria": f'=ISNUMBER(SEARCH("{rolle}", ${mg_col_letter}2))',
                                "format": yellow_highlight_format,
                            },
                        )

        # Zusätzliches Sheet für Kontakte ohne Geburtsdatum (nach Dezember)
        if include_no_date_sheet and not df_no_date.empty:
            sheet_name = "Ohne_Geburtsdatum"
            # Möglichst ähnliche Spaltenreihenfolge
            desired_cols_no_date = [
                "Vorname",
                "Nachname",
                "Firma",
                "Strasse",
                "PLZ",
                "Ort",
                "Mitgliedschaft",
            ]

            # Korrespondenzsprache-Spalte auch hier robust erkennen
            corr_lang_col_no_date = None
            for col in df_no_date.columns:
                col_norm = str(col).strip().lower().replace(" ", "")
                if col_norm in ("korrespondenzsprache", "korrespsprache"):
                    corr_lang_col_no_date = col
                    break
                if ("korresp" in col_norm or "korrespondenz" in col_norm) and "sprache" in col_norm:
                    corr_lang_col_no_date = col
                    break

            if corr_lang_col_no_date:
                # nach Ort einfügen
                desired_cols_no_date.insert(6, corr_lang_col_no_date)

            sheet_cols_no_date = [
                c for c in desired_cols_no_date if c in df_no_date.columns
            ]
            # Falls andere Spalten existieren, hinten anhängen
            for c in df_no_date.columns:
                if c not in sheet_cols_no_date:
                    sheet_cols_no_date.append(c)

            # Fallback: mindestens alle Spalten, falls Liste am Ende leer ist
            if not sheet_cols_no_date:
                sheet_cols_no_date = list(df_no_date.columns)

            export_no_date = df_no_date[sheet_cols_no_date]
            export_no_date.to_excel(writer, sheet_name=sheet_name, index=False)

            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            yellow_highlight_format = workbook.add_format(
                {"bg_color": "#FFF2CC", "font_color": "#7F6000"}
            )
            bold_format = workbook.add_format({"bold": True})

            worksheet.set_row(0, None, bold_format)
            worksheet.set_column(0, len(sheet_cols_no_date) - 1, 18)

            # Auch hier Ehrenrollen hervorheben, falls Mitgliedschaft vorhanden
            if "Mitgliedschaft" in sheet_cols_no_date:
                mg_col_idx = sheet_cols_no_date.index("Mitgliedschaft")
                mg_col_letter = chr(ord("A") + mg_col_idx)
                row_end = len(export_no_date)
                for rolle in ["Ehrenmitglied", "Ehrenpräsident", "Prinzenrolle"]:
                    worksheet.conditional_format(
                        f"A2:{chr(ord('A') + len(sheet_cols_no_date) - 1)}{row_end + 1}",
                        {
                            "type": "formula",
                            "criteria": f'=ISNUMBER(SEARCH("{rolle}", ${mg_col_letter}2))',
                            "format": yellow_highlight_format,
                        },
                    )

        # Falls keine Monatsdaten vorhanden, mindestens das Ohne_Geburtsdatum-Sheet
        if df_with_date.empty and (not include_no_date_sheet or df_no_date.empty):
            # Keine Daten, trotzdem ein leeres Sheet anlegen
            empty_df = pd.DataFrame({"Info": ["Keine Daten vorhanden"]})
            empty_df.to_excel(writer, sheet_name="Keine_Daten", index=False)

    # Statistik (global, nicht pro Monat)
    stats = {}
    if "Vorname" in df.columns:
        stats["anzahl_namen"] = int(df["Vorname"].notna().sum())
    if "Geburtsdatum" in df.columns:
        stats["anzahl_geburtsdaten"] = int(df["Geburtsdatum"].notna().sum())
        stats["anzahl_fehlende_geburtsdaten"] = int(df["Geburtsdatum"].isna().sum())

    output.seek(0)
    return output.getvalue(), stats


def main():
    st.set_page_config(page_title="Geburtstagsliste 2026", layout="centered")

    # Moderner heller Notion-Style / responsive Cards
    st.markdown(
        """
<style>
/* Hintergrund */
.stApp {
    background: linear-gradient(180deg, #f9fafb 0%, #f3f4f6 100%);
    color: #0f172a;
}

/* Hauptcontainer als helle Card */
.block-container {
    padding-top: 2.6rem; /* Abstand unterhalb des Streamlit-Header-Balkens */
    padding-bottom: 2.2rem;
    max-width: 1080px;
    background-color: #ffffff;
    border-radius: 18px;
    box-shadow: 0 18px 45px rgba(15, 23, 42, 0.12);
    border: 1px solid #e5e7eb;
}

/* Texte */
h1, h2, h3, h4, h5, h6, p, li, span, label {
    color: #0f172a !important;
}
p, li, span, label {
    font-size: 16px;
    line-height: 1.6;
}

/* Buttons */
div.stButton > button {
    border-radius: 12px;
    padding: 0.65rem 1.5rem;
    border: 1px solid #e2e8f0;
    background: #ffffff;
    color: #0f172a;
    font-weight: 600;
    letter-spacing: 0.01em;
    box-shadow: 0 10px 22px rgba(15, 23, 42, 0.06);
    transition: all 0.18s ease;
}
div.stButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 14px 30px rgba(15, 23, 42, 0.12);
}

/* Download-Button */
[data-testid="baseButton-secondary"] {
    border-radius: 12px !important;
    border: 1px solid #e2e8f0 !important;
}

/* Info / Alert Boxen */
.stAlert {
    background-color: #f8fafc;
    border-radius: 14px;
    border: 1px solid #e2e8f0;
}

/* Tabelle */
[data-testid="stDataFrame"] {
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid #e5e7eb;
}

/* Upload Dropzone */
[data-testid="stFileUploadDropzone"] {
    border: 2px dashed #cbd5e1;
    border-radius: 16px;
    background: #f8fafc;
}

/* Mobile Responsiveness */
@media (max-width: 768px) {
    .block-container {
        padding: 1rem;
        border-radius: 12px;
    }
    .stMarkdown h2, .stMarkdown h3 {
        font-size: 1.2rem;
    }
    div.stButton > button {
        width: 100%;
    }
}
</style>
""",
        unsafe_allow_html=True,
    )

    # Logo dezent zentriert
    logo_cols = st.columns([1, 1, 1])
    with logo_cols[1]:
        try:
            st.image("Gilde-Brandlogo-Petrol.jpg", width=220)
        except Exception:
            st.write(" ")

    # Einstellungen ohne Sidebar (responsive Controls)
    st.markdown("### Einstellungen")
    ctrl_cols = st.columns([1, 1])
    with ctrl_cols[0]:
        target_year = st.number_input(
            "Zieljahr für Geburtstagsliste",
            min_value=1900,
            max_value=2100,
            value=2026,
            step=1,
        )
    with ctrl_cols[1]:
        include_no_date_sheet = st.checkbox(
            "Blatt für Einträge ohne Geburtsdatum hinzufügen",
            value=True,
        )

    st.markdown("## 🎂 Geburtstagsliste Generator")
    st.caption(
        "Automatischer Excel-Export aus Fairgate – mit runden Geburtstagen und Ehrenmitgliedern hervorgehoben."
    )

    st.markdown(
        """
**Schritt 1 – Export aus Fairgate**

1. Öffne [`mein.fairgate.ch`](https://mein.fairgate.ch/gilde/backend/contact/list).
2. Unter **„Gespeicherte Filter“** wähle **„Geburtstags-Exportliste“**.
3. Klicke oben rechts auf die **drei Balken (Menü)** und dann auf **„Exportieren“**.
4. Unter **„Gespeicherte Spalteneinstellungen“** wähle **„Geburtstagstabelle“** und lade die Excel-Datei herunter.

**Schritt 2 – Geburtstagsliste erzeugen**

Lade hier die eben aus Fairgate exportierte Excel-Datei hoch.
"""
    )

    uploaded_file = st.file_uploader(
        "Excel-Datei auswählen (.xlsx)", type=["xlsx"], accept_multiple_files=False
    )

    if not uploaded_file:
        st.info("Bitte eine Excel-Datei per Drag & Drop hierher ziehen oder auswählen.")
        return

    # Datei einlesen
    try:
        df_raw = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Fehler beim Einlesen der Datei: {e}")
        return

    st.subheader("Hochgeladene Datei")

    # Spaltenübersicht per Toggle-Button
    if "show_columns" not in st.session_state:
        st.session_state.show_columns = False

    cols_top = st.columns([2, 1])
    with cols_top[0]:
        st.caption("Vorschau der ersten Zeilen aus dem Export.")
    with cols_top[1]:
        if st.button(
            "Spalten verbergen" if st.session_state.show_columns else "Spalten einblenden"
        ):
            st.session_state.show_columns = not st.session_state.show_columns

    if st.session_state.show_columns:
        with st.expander("Spalten der hochgeladenen Datei", expanded=True):
            st.write(list(df_raw.columns))

    st.dataframe(df_raw.head(30), use_container_width=True)

    df_prepared = prepare_dataframe(df_raw.copy(), target_year=target_year)

    with st.expander(
        f"Vorschau der aufbereiteten Daten (für {target_year})", expanded=False
    ):
        st.dataframe(df_prepared.head(30), use_container_width=True)

    if st.button(f"Geburtstagsliste {target_year} erstellen"):
        with st.spinner("Erzeuge Excel-Datei ..."):
            excel_bytes, stats = build_geburtstagsliste_excel(
                df_prepared,
                target_year=target_year,
                include_no_date_sheet=include_no_date_sheet,
            )

        st.success(f"Geburtstagsliste_{target_year}.xlsx wurde erfolgreich erzeugt.")

        # Statistik anzeigen
        col1, col2, col3 = st.columns(3)
        if "anzahl_namen" in stats:
            col1.metric("Anzahl Namen", stats["anzahl_namen"])
        if "anzahl_geburtsdaten" in stats:
            col2.metric("Geburtsdaten vorhanden", stats["anzahl_geburtsdaten"])
        if "anzahl_fehlende_geburtsdaten" in stats:
            col3.metric("Fehlende Geburtsdaten", stats["anzahl_fehlende_geburtsdaten"])

        st.download_button(
            label=f"📥 Geburtstagsliste_{target_year}.xlsx herunterladen",
            data=excel_bytes,
            file_name=f"Geburtstagsliste_{target_year}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )


if __name__ == "__main__":
    main()
