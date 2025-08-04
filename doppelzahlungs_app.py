import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from wordcloud import WordCloud

# Titel und Einführung
st.set_page_config(layout="wide")
st.title("Doppelzahlungsanalyse")
st.markdown("Diese Anwendung dient zur **systematischen Identifikation von potentiellen Doppelzahlungen** in der Buchhaltung.")

# Daten laden
file_path = "Callies_Fibu_Buchungen_neu_Jahr 2022-2023.xlsx"
df = pd.read_excel(file_path)

# Relevante Spalten prüfen
required_columns = ["BelegID", "RechnungNr", "Saldo", "BelegDatum", "BuText", "AusgleichsID", "KtoNr", "BelegNr", "PersKto"]
if not all(col in df.columns for col in required_columns):
    st.error("Eine oder mehrere erforderliche Spalten fehlen in der Excel-Datei.")
    st.stop()

# Grundbereinigung
df["AusgleichsID"] = df["AusgleichsID"].fillna("").astype(str)
df["Saldo"] = pd.to_numeric(df["Saldo"], errors="coerce").abs()
df["KtoNr"] = pd.to_numeric(df["KtoNr"], errors="coerce")
df["BelegNr"] = df["BelegNr"].astype(str)
df["BelegNr_prefix"] = df["BelegNr"].str[:3]

# Filter: nur KtoNr = 4200 & SHKz ≠ S
df_4200 = df[(df["KtoNr"] == 4200) & (df["SHKz"] != "S")].copy()

# Analysen: Duplikate nach RechnungNr, Saldo, PersKto
df_4200["Analyse_1"] = df_4200.duplicated(subset=["RechnungNr", "Saldo", "PersKto"], keep=False).astype(int)
df_4200["Analyse_2"] = df_4200.duplicated(subset=["AusgleichsID"], keep=False).astype(int)

# Filter anwenden
df_4200 = df_4200[(df_4200["Analyse_1"] == 1) & (df_4200["Analyse_2"] == 0)]
df_4200 = df_4200[~df_4200["BelegNr"].str.startswith("50-")]

# Merge: Buchungen mit Zielkonten über AusgleichsID
zielkonten = [193003, 193303, 193308, 193401, 193302, 1933]
df_merge = df[df["KtoNr"].isin(zielkonten)].copy()

df_merge["BelegID"] = pd.to_numeric(df_merge["BelegID"], errors="coerce").astype("Int64")
df_4200["AusgleichsID"] = pd.to_numeric(df_4200["AusgleichsID"], errors="coerce").astype("Int64")

df_4200 = df_4200.merge(
    df_merge[["BelegID", "KtoNr"]],
    left_on="AusgleichsID",
    right_on="BelegID",
    how="left"
)

# Kennzeichnung
df_4200["Identifiziert"] = df_4200["KtoNr_y"].apply(lambda x: "x" if x in zielkonten else "")

# Ausgabe: Statistiken
st.header("📊 Statistiken")
col1, col2, col3 = st.columns(3)
col1.metric("Gesamtzeilen (Excel)", len(df))
col2.metric("Gefiltert KtoNr=4200", len(df_4200))
col3.metric("Identifizierte Doppelzahlungen", df_4200["Identifiziert"].eq("x").sum())

# Grafiken: nebeneinander
st.header("📈 Visualisierungen")
col1, col2 = st.columns(2)

with col1:
    st.subheader("BelegNr-Präfixe (Kto 4200)")
    fig1, ax1 = plt.subplots(figsize=(5, 2))
    prefix_counts = df_4200["BelegNr_prefix"].value_counts().sort_index()
    bars = ax1.bar(prefix_counts.index, prefix_counts.values)
    for bar in bars:
        yval = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width() / 2, yval + 0.1, int(yval), ha='center', va='bottom')
    ax1.set_xlabel("Präfix")
    ax1.set_ylabel("Anzahl")
    st.pyplot(fig1)

with col2:
    st.subheader("Wordcloud: Buchungstexte")
    bu_text = " ".join(df["BuText"].dropna().astype(str).tolist())
    wordcloud = WordCloud(width=400, height=200, background_color='white').generate(bu_text)
    fig2, ax2 = plt.subplots(figsize=(5, 2))
    ax2.imshow(wordcloud, interpolation='bilinear')
    ax2.axis("off")
    st.pyplot(fig2)

# Exportoption
st.header("⬇️ Download Ergebnis")
@st.cache_data
def convert_df(df):
    return df.to_excel(index=False, engine='openpyxl')

excel_bytes = convert_df(df_4200)
st.download_button(
    label="Download als Excel-Datei",
    data=excel_bytes,
    file_name="doppelzahlungen_analyse.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
