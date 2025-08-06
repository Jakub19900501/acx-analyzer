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
    summary["🔁 % Ponowny kontakt"] = round((summary["🔁 Ponowny kontakt"] / summary["📋 Rekordów"]) * 100, 2)
    summary["🔁 Śr. prób"] = round(summary["📞 Połączeń"] / summary["📋 Rekordów"].replace(0, 1), 2)
    summary["⏳ Śr. czas reakcji (dni)"] = (summary["📅 Ostatni kontakt"] - summary["🕓 Data importu"]).dt.days

    def alert(row):
        l100r = row["💯 L100R"]
        if l100r >= 1.0:
            return "🟣 Baza genialna"
        elif l100r >= 0.57:
            return "🟢 Baza bardzo dobra"
        elif l100r >= 0.32:
            return "🟡 Baza solidna"
        elif l100r >= 0.23:
            return "🟠 Baza przeciętna"
        elif l100r >= 0.10:
            return "🔴 Baza słaba"
        else:
            return "⚫ Baza martwa"

    summary["🚨 Alert"] = summary.apply(alert, axis=1)

    alert_order = [
        "🟣 Baza genialna",
        "🟢 Baza bardzo dobra",
        "🟡 Baza solidna",
        "🟠 Baza przeciętna",
        "🔴 Baza słaba",
        "⚫ Baza martwa"
    ]
    summary["🚨 Alert"] = pd.Categorical(summary["🚨 Alert"], categories=alert_order, ordered=True)
    summary = summary.sort_values("🚨 Alert")

    metryki_kolejnosc = [
        "📁 Baza", "💯 L100R", "📉 CTR", "🔁 % Ponowny kontakt", "🔁 Śr. prób",
        "📋 Rekordów", "📞 Połączeń", "✅ Spotkań", "🔁 Ponowny kontakt",
        "❌ Rekordy z błędem", "📅 Ostatni kontakt", "🕓 Data importu",
        "⏳ Śr. czas reakcji (dni)", "🚨 Alert"
    ]
    summary = summary[[col for col in metryki_kolejnosc if col in summary.columns]]

    # ✅ NOWA ANALIZA PONOWNYCH KONTAKTÓW PER BAZA
    ponowne = df_all[df_all["TotalTries"] > 1].copy()
    ponowne["LastCallCode_clean"] = ponowne["LastCallCode"].astype(str).apply(normalize_text)
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

        ws_summary.freeze_panes(1, 5)
        ws_ponowny.freeze_panes(1, 0)

        for i, col in enumerate(summary.columns):
            ws_summary.set_column(i, i, 22)
        for i, col in enumerate(ponowna_analiza.columns):
            ws_ponowny.set_column(i, i, 28)

        # Dodanie wykresów i legendy
        chart_sheet = wb.add_worksheet("Wykresy")
        metrics = ["💯 L100R", "📉 CTR", "🔁 % Ponowny kontakt", "🔁 Śr. prób"]
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
            ("🔁 % Ponowny kontakt", "Odsetek ponownych prób"),
            ("🔁 Śr. prób", "Średnia prób per rekord"),
            ("📋 Rekordów", "Liczba rekordów"),
            ("📞 Połączeń", "Łączna liczba prób kontaktu"),
            ("✅ Spotkań", "Zakończone sukcesem"),
            ("❌ Rekordy z błędem", "Rozłączone / błędny numer"),
            ("📅 Ostatni kontakt", "Data ostatniego kontaktu"),
            ("⏳ Śr. czas reakcji (dni)", "Import → Kontakt"),
            ("🚨 Alert", "Ocena bazy wg L100R (🟣 do ⚫)")
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

            ws.write(start, 3, "🚨 Alert — jakość bazy wg L100R:")
            ws.write(start + 1, 3, "🟣 ≥ 1.00");      ws.write(start + 1, 4, "Baza genialna")
            ws.write(start + 2, 3, "🟢 0.57–0.99");   ws.write(start + 2, 4, "Baza bardzo dobra")
            ws.write(start + 3, 3, "🟡 0.32–0.56");   ws.write(start + 3, 4, "Baza solidna")
            ws.write(start + 4, 3, "🟠 0.23–0.31");   ws.write(start + 4, 4, "Baza przeciętna")
            ws.write(start + 5, 3, "🔴 0.10–0.22");   ws.write(start + 5, 4, "Baza słaba")
            ws.write(start + 6, 3, "⚫ < 0.10");       ws.write(start + 6, 4, "Baza martwa")

    st.download_button(
        "⬇️ Pobierz raport Excel",
        data=buffer.getvalue(),
        file_name="Raport_Porownanie_Baz_ACX.xlsx",
        mime="application/vnd.ms-excel"
    )
