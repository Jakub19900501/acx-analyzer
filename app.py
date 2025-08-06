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

def klasyfikuj_alert_ctr(ctr_val):
    if ctr_val is None:
        return "Brak danych"
    elif ctr_val < 150:
        return "🟣 Genialna"
    elif ctr_val < 300:
        return "🟢 Bardzo dobra"
    elif ctr_val < 500:
        return "🟡 Solidna"
    elif ctr_val < 700:
        return "🟠 Przeciętna"
    elif ctr_val < 1000:
        return "🔴 Słaba"
    else:
        return "⚫ Martwa"

def generuj_wniosek_ctr_roe(ctr_val, roe_val, umowienia):
    if ctr_val is None or umowienia == 0:
        return "❌ Brak umówień – baza prawdopodobnie martwa."
    elif ctr_val >= 1000:
        return "⚠️ CTR ≥ 1000 – baza wypalona. Zalecane wycofanie lub filtrowanie."
    elif roe_val > 5:
        return "✅ ROE > 5% – baza bardzo efektywna. Warto kontynuować."
    elif ctr_val < 300:
        return "👍 CTR < 300 – kaloryczna baza, szybkie efekty."
    else:
        return ""

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
        "ImportCreatedOn": "min"
    }).reset_index()

    summary.rename(columns={
        "Baza": "📁 Baza",
        "Id": "📋 Rekordów",
        "TotalTries": "📞 Połączeń",
        "Skuteczny": "✅ Spotkań",
        "PonownyKontakt": "🔁 Ponowny kontakt",
        "Bledny": "❌ Rekordy z błędem",
        "LastTryTime": "📅 Ostatni kontakt",
        "ImportCreatedOn": "🕓 Data importu"
    }, inplace=True)

    summary["💯 L100R"] = round((summary["✅ Spotkań"] / summary["📋 Rekordów"]) * 100, 2)
    summary["📉 CTR"] = round(summary["📞 Połączeń"] / summary["✅ Spotkań"].replace(0, 1), 2)
    summary["ROE (%)"] = round((summary["✅ Spotkań"] / summary["📞 Połączeń"]) * 100, 2)
    summary["🔁 % Ponowny kontakt"] = round((summary["🔁 Ponowny kontakt"] / summary["📋 Rekordów"]) * 100, 2)
    summary["🔁 Śr. prób"] = round(summary["📞 Połączeń"] / summary["📋 Rekordów"].replace(0, 1), 2)
    summary["⏳ Śr. czas reakcji (dni)"] = (summary["📅 Ostatni kontakt"] - summary["🕓 Data importu"]).dt.days
    summary["🚨 Alert CTR"] = summary["📉 CTR"].apply(klasyfikuj_alert_ctr)
    summary["📝 Wniosek"] = summary.apply(
        lambda row: generuj_wniosek_ctr_roe(row["📉 CTR"], row["ROE (%)"], row["✅ Spotkań"]), axis=1
    )

    summary = summary.sort_values("📉 CTR")

    metryki_kolejnosc = [
        "📁 Baza", "💯 L100R", "📉 CTR", "ROE (%)", "🔁 % Ponowny kontakt", "🔁 Śr. prób",
        "📋 Rekordów", "📞 Połączeń", "✅ Spotkań", "🔁 Ponowny kontakt",
        "❌ Rekordy z błędem", "📅 Ostatni kontakt", "🕓 Data importu",
        "⏳ Śr. czas reakcji (dni)", "🚨 Alert CTR", "📝 Wniosek"
    ]
    summary = summary[[col for col in metryki_kolejnosc if col in summary.columns]]

    # 🔁 ANALIZA PONOWNYCH KONTAKTÓW
    ponowne = df_all[df_all["TotalTries"] > 1].copy()
    ponowne["Skuteczne"] = ponowne["LastCallCode_clean"].str.contains("umowienie")
    ponowne_umowienia = ponowne[ponowne["Skuteczne"] == True]

    ponowny_raport = []
    for baza in ponowne_umowienia["Baza"].unique():
        baza_df = ponowne_umowienia[ponowne_umowienia["Baza"] == baza]
        rozklad = baza_df["TotalTries"].value_counts().sort_index()
        suma = len(baza_df)
        sr = baza_df["TotalTries"].mean()
        med = baza_df["TotalTries"].median()
        rozklad_str = ", ".join([f"{int(k)}: {v}" for k, v in rozklad.items()])
        ponowny_raport.append({
            "📁 Baza": baza,
            "🔁 Rekordów ponownych": len(ponowne[ponowne["Baza"] == baza]),
            "✅ Umówienia": suma,
            "📈 Śr. próba": round(sr, 2),
            "🎯 Mediana": round(med, 2),
            "📊 Rozkład prób": rozklad_str
        })

    ponowna_analiza = pd.DataFrame(ponowny_raport)

    # 📊 WYŚWIETLENIE
    st.subheader("📊 Porównanie baz – rozszerzone")
    st.dataframe(summary, use_container_width=True)
    st.subheader("📊 Skuteczność ponownych kontaktów")
    st.dataframe(ponowna_analiza, use_container_width=True)

    # 📥 EXPORT XLSX
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        summary.to_excel(writer, index=False, sheet_name="Porównanie baz")
        ponowna_analiza.to_excel(writer, index=False, sheet_name="Ponowny kontakt")

        wb = writer.book
        ws_summary = writer.sheets["Porównanie baz"]
        ws_ponowny = writer.sheets["Ponowny kontakt"]

        ws_summary.freeze_panes(1, 5)
        ws_ponowny.freeze_panes(1, 0)

        for i, col in enumerate(summary.columns):
            ws_summary.set_column(i, i, 22)
        for i, col in enumerate(ponowna_analiza.columns):
            ws_ponowny.set_column(i, i, 28)

        chart_sheet = wb.add_worksheet("Wykresy")
        metrics = ["💯 L100R", "📉 CTR", "ROE (%)", "🔁 % Ponowny kontakt", "🔁 Śr. prób"]
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
            chart.set_size({'width': 1440, 'height': 480})
            chart_sheet.insert_chart(i * 25, 0, chart)

        legenda = [
            ("💯 L100R", "Spotkania na 100 rekordów"),
            ("📉 CTR", "Połączenia / spotkania"),
            ("ROE (%)", "Efektywność: % umówień / prób"),
            ("🔁 % Ponowny kontakt", "Odsetek ponownych prób"),
            ("🔁 Śr. prób", "Średnia prób per rekord"),
            ("📋 Rekordów", "Liczba rekordów"),
            ("📞 Połączeń", "Łączna liczba prób kontaktu"),
            ("✅ Spotkań", "Zakończone sukcesem"),
            ("❌ Rekordy z błędem", "Rozłączone / błędny numer"),
            ("📅 Ostatni kontakt", "Data ostatniego kontaktu"),
            ("⏳ Śr. czas reakcji (dni)", "Import → Kontakt"),
            ("🚨 Alert CTR", "Ocena bazy wg CTR"),
            ("📝 Wniosek", "Ocena i zalecenie na podstawie CTR/ROE")
        ]

        for ws, start in [
            (ws_summary, len(summary) + 4),
            (ws_ponowny, len(ponowna_analiza) + 4),
            (chart_sheet, 102)
        ]:
            ws.write(start, 0, "📌 LEGENDA METRYK")
            for idx, (label, desc) in enumerate(legenda, start + 1):
                ws.write(idx, 0, label)
                ws.write(idx, 1, desc)

    st.download_button(
        "⬇️ Pobierz raport Excel",
        data=buffer.getvalue(),
        file_name="Raport_Porownanie_Baz_ACX.xlsx",
        mime="application/vnd.ms-excel"
    )
