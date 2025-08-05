import streamlit as st
import pandas as pd
import io
import unicodedata
from datetime import datetime
import matplotlib.pyplot as plt

st.set_page_config(page_title="ACX Analyzer", layout="wide")
st.title("📞 ACX Analyzer – porównanie baz (do 50 plików)")

uploaded_files = st.file_uploader("📤 Wgraj pliki ACX (.xlsx)", type=["xlsx"], accept_multiple_files=True)

def normalize_text(text):
    return unicodedata.normalize("NFKD", str(text).lower()).encode("ascii", errors="ignore").decode("utf-8")

if uploaded_files:
    all_data = []
    for file in uploaded_files[:50]:
        df = pd.read_excel(file)
        df["Baza"] = file.name.replace(".xlsx", "")
        all_data.append(df)

    df_all = pd.concat(all_data, ignore_index=True)
    df_all["LastCallCode_clean"] = df_all["LastCallCode"].astype(str).apply(normalize_text)
    df_all["Skuteczny"] = df_all["LastCallCode_clean"].str.contains("umow|sukces|magazyn")
    df_all["PonownyKontakt"] = df_all["LastCallCode_clean"].str.contains("ponowny kontakt")
    df_all["Bledny"] = df_all["LastCallCode_clean"].str.contains("bledn|zly|rozlacz")
    df_all["TotalTries"] = df_all["TotalTries"].fillna(0)
    df_all["LastTryTime"] = pd.to_datetime(df_all["LastTryTime"], errors="coerce")
    df_all["ImportCreatedOn"] = pd.to_datetime(df_all["ImportCreatedOn"], errors="coerce")

    summary = df_all.groupby("Baza").agg({
        "Id": "count",
        "TotalTries": "sum",
        "Skuteczny": "sum",
        "PonownyKontakt": "sum",
        "Bledny": "sum",
        "LastTryTime": "max",
        "LastTryUser": lambda x: ", ".join(set(x.dropna().astype(str))),
        "RejectReason": lambda x: ", ".join(x.dropna().astype(str).value_counts().head(3).index),
        "ImportCreatedOn": "min",
        "CampaignRecordPhoneIndex": lambda x: ", ".join(set(x.dropna().astype(str)))
    }).reset_index()

    summary.rename(columns={
        "Baza": "📁 Baza",
        "Id": "📋 Rekordów",
        "TotalTries": "📞 Połączeń",
        "Skuteczny": "✅ Spotkań",
        "PonownyKontakt": "🔁 Ponowny kontakt",
        "Bledny": "❌ Błędnych",
        "LastTryTime": "📅 Ostatni kontakt",
        "LastTryUser": "👤 Konsultanci",
        "RejectReason": "🧱 Top odmowy",
        "ImportCreatedOn": "🕓 Data importu",
        "CampaignRecordPhoneIndex": "🧭 Regiony"
    }, inplace=True)

    summary["💯 L100R"] = round((summary["✅ Spotkań"] / summary["📋 Rekordów"]) * 100, 2)
    summary["📉 CTR"] = round(summary["📞 Połączeń"] / summary["✅ Spotkań"].replace(0, 1), 2)
    summary["🔁 % Ponowny kontakt"] = round((summary["🔁 Ponowny kontakt"] / summary["📋 Rekordów"]) * 100, 2)
    summary["🔁 Śr. prób"] = round(summary["📞 Połączeń"] / summary["📋 Rekordów"].replace(0, 1), 2)
    summary["❌ % błędnych"] = round((summary["❌ Błędnych"] / summary["📋 Rekordów"]) * 100, 2)
    summary["⏳ Śr. czas reakcji (dni)"] = (summary["📅 Ostatni kontakt"] - summary["🕓 Data importu"]).dt.days

    def alert(row):
        if row["💯 L100R"] >= 0.20:
            return "🟢 Baza dobra"
        elif row["💯 L100R"] >= 0.10:
            return "🟡 Do obserwacji"
        else:
            return "🔴 Baza martwa"
    summary["🚨 Alert"] = summary.apply(alert, axis=1)

    ponowne = df_all[df_all["PonownyKontakt"] == True].copy()
    ponowna_analiza = ponowne.groupby("Baza").agg({
        "Id": "count",
        "Skuteczny": "sum",
        "TotalTries": "sum"
    }).reset_index()
    ponowna_analiza.rename(columns={
        "Baza": "📁 Baza",
        "Id": "🔁 Rekordów ponownych",
        "Skuteczny": "✅ Skuteczne",
        "TotalTries": "📞 Połączeń"
    }, inplace=True)
    ponowna_analiza["💯 L100R"] = round((ponowna_analiza["✅ Skuteczne"] / ponowna_analiza["🔁 Rekordów ponownych"]) * 100, 2)
    ponowna_analiza["📉 CTR"] = round(ponowna_analiza["📞 Połączeń"] / ponowna_analiza["✅ Skuteczne"].replace(0, 1), 2)

    st.subheader("📊 Porównanie baz – rozszerzone")
    st.dataframe(summary, use_container_width=True)

    st.subheader("📈 Wykresy skuteczności baz")
    chart_cols = ["💯 L100R", "📉 CTR", "🔁 % Ponowny kontakt", "❌ % błędnych"]
    for col in chart_cols:
        fig, ax = plt.subplots()
        summary_sorted = summary.sort_values(col, ascending=False)
        ax.bar(summary_sorted["📁 Baza"], summary_sorted[col], color="skyblue")
        ax.set_title(col)
        ax.set_ylabel(col)
        ax.set_xticklabels(summary_sorted["📁 Baza"], rotation=90)
        st.pyplot(fig)

    st.subheader("📊 Skuteczność ponownych kontaktów")
    st.dataframe(ponowna_analiza, use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        summary.to_excel(writer, index=False, sheet_name="Porównanie baz")
        ponowna_analiza.to_excel(writer, index=False, sheet_name="Ponowny kontakt")
        ws = writer.sheets["Porównanie baz"]
        ws2 = writer.sheets["Ponowny kontakt"]
        bold = writer.book.add_format({'bold': True})
        for i, col in enumerate(summary.columns): ws.set_column(i, i, max(15, len(str(col)) + 2))
        for i, col in enumerate(ponowna_analiza.columns): ws2.set_column(i, i, max(15, len(str(col)) + 2))
        ws.write(len(summary)+2, 0, "📌 LEGENDA", bold)
        ws2.write(len(ponowna_analiza)+2, 0, "📌 LEGENDA", bold)

    st.download_button("⬇️ Pobierz raport Excel", data=buffer.getvalue(), file_name="Raport_Porownanie_Baz_ACX.xlsx", mime="application/vnd.ms-excel")
