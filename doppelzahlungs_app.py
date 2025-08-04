import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from io import BytesIO

st.set_page_config(page_title="Doppelzahlungsprüfung", layout="wide")
st.title("🔍 Doppelzahlungsprüfung – ARE Projekt")

# Datei-Upload
uploaded_file = st.file_uploader("📁 Excel-Datei hochladen", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)

    # Validierung der Spalten
    required_columns = ["BelegID", "RechnungNr", "Saldo", "BelegDatum", "BuText", "AusgleichsID", "KtoNr", "BelegNr", "PersKto"]
    if not all(col in df.columns for col in required_columns):
        st.error("❌ Die Datei enthält nicht alle benötigten Spalten.")
        st.stop()

    # Statistik: Grunddaten
    st.subheader("📊 Grundstatistiken")
    st.write(f"**Gesamtanzahl Zeilen:** {len(df)}")
    
    df["Saldo"] = pd.to_numeric(df["Saldo"], errors="coerce").abs()
    df["AusgleichsID"] = df["AusgleichsID"].fillna("").astype(str)
    df["Mit_AusgleichsID"] = df["AusgleichsID"].apply(lambda x: 1 if x.strip() != "" else 0)
    df["Ohne_AusgleichsID"] = df["AusgleichsID"].apply(lambda x: 1 if x.strip() == "" else 0)

    st.write(f"✅ Zeilen mit AusgleichsID: {df['Mit_AusgleichsID'].sum()}")
    st.write(f"⚠️ Zeilen ohne AusgleichsID: {df['Ohne_AusgleichsID'].sum()}")

    # Filter auf KtoNr = 4200 und SHKz = H
    df_4200 = df[(df["KtoNr"] == 4200) & (df["SHKz"] != "S")].copy()

    st.write(f"🔎 Gefilterte Zeilen (KtoNr=4200 & SHKz≠S): {len(df_4200)}")

    # Analysen
    df_4200["Analyse_1"] = df_4200.duplicated(subset=["RechnungNr", "Saldo", "PersKto"], keep=False).astype(int)
    df_4200["Analyse_2"] = df_4200.duplicated(subset=["AusgleichsID"], keep=False).astype(int)
    df_4200 = df_4200[df_4200["Analyse_1"] == 1]
    df_4200 = df_4200[df_4200["Analyse_2"] == 0]
    df_4200 = df_4200[~df_4200["BelegNr"].astype(str).str.startswith("50-")]

    st.write(f"🔁 Übrig nach Analyse & Filter: {len(df_4200)}")

    # BelegNr-Prefix Diagramm
    df_4200["BelegNr_prefix"] = df_4200["BelegNr"].astype(str).str[:3]
    prefix_counts = df_4200["BelegNr_prefix"].value_counts().sort_index()

    st.subheader("📌 BelegNr-Prefix Histogramm")
    fig, ax = plt.subplots(figsize=(2, 1))
    bars = ax.bar(prefix_counts.index, prefix_counts.values)
    for bar in bars:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, height + 0.5, int(height), ha='center')
    ax.set_title("Buchungen je BelegNr-Präfix")
    st.pyplot(fig)

    # Wordcloud BuText
    st.subheader("💬 Wordcloud aus Buchungstexten")
    bu_text = " ".join(df["BuText"].dropna().astype(str))
    wordcloud = WordCloud(width=800, height=400, background_color='white').generate(bu_text)
    fig_wc, ax_wc = plt.subplots(figsize=(2, 1))
    ax_wc.imshow(wordcloud, interpolation='bilinear')
    ax_wc.axis("off")
    st.pyplot(fig_wc)

    # Erkennung möglicher Gegenbuchungen (Merge)
    gewuenschte_kontonummern = [193003, 193303, 193308, 193401, 193302, 1933]
    df_merge = df[df["KtoNr"].isin(gewuenschte_kontonummern)].copy()
    df_merge["BelegID"] = pd.to_numeric(df_merge["BelegID"], errors="coerce").astype("Int64")
    df_4200["AusgleichsID"] = pd.to_numeric(df_4200["AusgleichsID"], errors="coerce").astype("Int64")

    df_result = df_4200.merge(
        df_merge[["BelegID", "KtoNr"]],
        left_on="AusgleichsID",
        right_on="BelegID",
        how="left"
    )

    df_result["Identifiziert"] = df_result["KtoNr_y"].apply(lambda x: "x" if x in gewuenschte_kontonummern else "")
    anzahl_identifiziert = (df_result["Identifiziert"] == "x").sum()
    st.success(f"✅ Mögliche Gegenbuchungen (KtoNr): {anzahl_identifiziert}")

    # Export als Excel
    st.subheader("⬇️ Ergebnis herunterladen")
    excel_buffer = BytesIO()
    df_result.to_excel(excel_buffer, index=False)
    st.download_button("📥 Ergebnis als Excel herunterladen", data=excel_buffer.getvalue(), file_name="df_4200_Gefiltert.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


