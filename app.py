import streamlit as st
import pandas as pd
import io
import unicodedata

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

    # Normalizacja i przygotowanie danych
    df_all["LastCallCode_clean"] = df_all["LastCallCode"].astype(str).apply(normalize_text)
    df_all["Skuteczny"] = df_all["LastCallCode_clean"].str.contains("umow|sukces|magazyn")
    df_all["PonownyKontakt"] = df_all["LastCallCode_clean"].str.contains("ponowny kontakt")
    df_all["TotalTries"] = df_all["TotalTries"].fillna(0)

    # 📊 Główna tabela porównania baz
    summary = df_all.groupby("Baza").agg({
        "Id": "count",
        "TotalTries": "sum",
        "Skuteczny": "sum",
        "PonownyKontakt": "sum"
    }).reset_index()

    summary.rename(columns={
        "Baza": "📁 Baza",
        "Id": "📋 Rekordów",
        "TotalTries": "📞 Połączeń",
        "Skuteczny": "✅ Spotkań",
        "PonownyKontakt": "🔁 Ponowny kontakt"
    }, inplace=True)

    summary["💯 L100R"] = round((summary["✅ Spotkań"] / summary["📋 Rekordów"]) * 100, 2)
    summary["📉 CTR"] = round(summary["📞 Połączeń"] / summary["✅ Spotkań"].replace(0, 1), 2)
    summary["🔁 % Ponowny kontakt"] = round((summary["🔁 Ponowny kontakt"] / summary["📋 Rekordów"]) * 100, 2)

    # 🧠 ALERTY według L100R
    def alert(row):
        if row["💯 L100R"] >= 0.20:
            return "🟢 Baza dobra"
        elif row["💯 L100R"] >= 0.10:
            return "🟡 Średnia"
        else:
            return "🔴 Baza martwa"

    summary["🚨 Alert"] = summary.apply(alert, axis=1)

    # 📈 Tabela tylko dla „ponownych kontaktów”
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

    # ✅ Wyświetlanie w Streamlit
    st.subheader("📊 Porównanie baz")
    st.dataframe(summary, use_container_width=True)

    st.subheader("📊 Skuteczność ponownych kontaktów")
    st.dataframe(ponowna_analiza, use_container_width=True)

    # 📥 Eksport do Excela
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        summary.to_excel(writer, index=False, sheet_name="Porównanie baz")
        ponowna_analiza.to_excel(writer, index=False, sheet_name="Ponowny kontakt")

        # Formatowanie + legenda w zakładce 1
        ws = writer.sheets["Porównanie baz"]
        for i, col in enumerate(summary.columns):
            ws.set_column(i, i, max(15, len(str(col)) + 2))
        legend = [
            ("📋 Rekordów", "Liczba rekordów w bazie"),
            ("📞 Połączeń", "Suma prób kontaktu (TotalTries)"),
            ("✅ Spotkań", "Spotkania: umówione / magazyn / sukces"),
            ("💯 L100R", "Leady na 100 rekordów"),
            ("📉 CTR", "Połączenia / spotkania"),
            ("🔁 Ponowny kontakt", "Liczba rekordów z kodem 'ponowny kontakt'"),
            ("🔁 % Ponowny kontakt", "Odsetek ponownych kontaktów"),
            ("🚨 Alert", "🟢 ≥ 0.20 (1/500) | 🟡 ≥ 0.10 | 🔴 < 0.10")
        ]
        start = len(summary) + 12
        ws.write(start, 0, "📌 LEGENDA METRYK")
        bold = writer.book.add_format({'bold': True})
        for label, desc in legend:
            start += 1
            ws.write(start, 0, label, bold)
            ws.write(start, 1, desc)

        # Legenda w zakładce 2
        ws2 = writer.sheets["Ponowny kontakt"]
        for i, col in enumerate(ponowna_analiza.columns):
            ws2.set_column(i, i, max(15, len(str(col)) + 2))
        start2 = len(ponowna_analiza) + 12
        ws2.write(start2, 0, "📌 LEGENDA METRYK")
        legend2 = [
            ("🔁 Rekordów ponownych", "Ile rekordów oznaczono jako ponowny kontakt"),
            ("✅ Skuteczne", "Ile z nich zakończyło się spotkaniem"),
            ("💯 L100R", "Skuteczność w % (spotkania/rekordy)"),
            ("📉 CTR", "Połączeń / spotkania")
        ]
        for label, desc in legend2:
            start2 += 1
            ws2.write(start2, 0, label, bold)
            ws2.write(start2, 1, desc)

    st.download_button(
        label="⬇️ Pobierz raport Excel",
        data=buffer.getvalue(),
        file_name="Raport_Porownanie_Baz_ACX.xlsx",
        mime="application/vnd.ms-excel"
    )
