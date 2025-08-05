import streamlit as st
import pandas as pd
import io
import unicodedata
from datetime import datetime

st.set_page_config(page_title="ACX Analyzer", layout="wide")
st.title("📞 ACX Analyzer – porównanie baz (do 50 plików)")

uploaded_files = st.file_uploader("📤 Wgraj pliki ACX (.xlsx)", type=["xlsx"], accept_multiple_files=True)

def normalize_text(text):
    return unicodedata.normalize("NFKD", str(text).lower()).encode("ascii", errors="ignore").decode("utf-8")

if uploaded_files:
    import xlsxwriter

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
        "CampaignRecordPhoneIndex": lambda x: ", ".join(set(x.dropna().astype(str))) if set(x.dropna().astype(str)) != {"Nr_telefonu"} else ""
    }).reset_index()

    summary.rename(columns={
        "Baza": "📁 Baza",
        "Id": "📋 Rekordów",
        "TotalTries": "📞 Połączeń",
        "Skuteczny": "✅ Spotkań",
        "PonownyKontakt": "🔁 Ponowny kontakt",
        "Bledny": "❌ Rekordy z błędem",
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
    summary["❌ % błędnych"] = round((summary["❌ Rekordy z błędem"] / summary["📋 Rekordów"]) * 100, 2)
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
    st.subheader("📊 Skuteczność ponownych kontaktów")
    st.dataframe(ponowna_analiza, use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        summary.to_excel(writer, index=False, sheet_name="Porównanie baz")
        ponowna_analiza.to_excel(writer, index=False, sheet_name="Ponowny kontakt")

        wb = writer.book
        ws_summary = writer.sheets["Porównanie baz"]
        ws_ponowny = writer.sheets["Ponowny kontakt"]

        for i, col in enumerate(summary.columns):
            ws_summary.set_column(i, i, 20)
        for i, col in enumerate(ponowna_analiza.columns):
            ws_ponowny.set_column(i, i, 20)

        # Wstaw wykresy
        chart_sheet = wb.add_worksheet("Wykresy")
        metrics = ["💯 L100R", "📉 CTR", "🔁 % Ponowny kontakt", "❌ % błędnych"]
        for i, metric in enumerate(metrics):
            chart = wb.add_chart({'type': 'column'})
            chart.add_series({
                'name': metric,
                'categories': ['Porównanie baz', 1, 0, len(summary), 0],
                'values':     ['Porównanie baz', 1, summary.columns.get_loc(metric), len(summary), summary.columns.get_loc(metric)],
            })
            chart.set_title({'name': metric})
            chart.set_x_axis({'name': 'Baza'})
            chart.set_y_axis({'name': metric})
            chart.set_size({'width': 1440, 'height': 480})  # A–T + 23 wiersze
            chart_sheet.insert_chart(i * 25, 0, chart)

        # Legenda dla obu arkuszy
        legend = [
            ("📋 Rekordów", "Liczba rekordów w bazie"),
            ("📞 Połączeń", "Łączna liczba prób kontaktu"),
            ("✅ Spotkań", "Ile razy zakończono sukcesem"),
            ("❌ Rekordy z błędem", "Błędny numer, rozłączenie"),
            ("💯 L100R", "Spotkania na 100 rekordów"),
            ("📉 CTR", "Połączenia / spotkania"),
            ("🔁 % Ponowny kontakt", "Odsetek ponownych prób"),
            ("🔁 Śr. prób", "Średnia liczba prób / rekord"),
            ("⏳ Śr. czas reakcji", "Dni od importu do kontaktu"),
            ("🧱 Top odmowy", "3 najczęstsze powody odmowy")
        ]

        for start_row, ws in [(len(summary) + 4, ws_summary), (len(ponowna_analiza) + 4, ws_ponowny)]:
            ws.write(start_row, 0, "📌 LEGENDA METRYK")
            for idx, (label, desc) in enumerate(legend, start=start_row + 1):
                ws.write(idx, 0, label)
                ws.write(idx, 1, desc)

    st.download_button("⬇️ Pobierz raport Excel", data=buffer.getvalue(), file_name="Raport_Porownanie_Baz_ACX.xlsx", mime="application/vnd.ms-excel")
