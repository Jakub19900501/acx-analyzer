import streamlit as st
import pandas as pd
import io
import unicodedata
from datetime import datetime

st.set_page_config(page_title="ACX Analyzer", layout="wide")
st.title("📞 ACX Analyzer – Porównanie baz (do 50 plików)")

uploaded_files = st.file_uploader("📤 Wgraj pliki ACX (max 50)", type=["xlsx"], accept_multiple_files=True)

def normalize_text(text):
    return unicodedata.normalize("NFKD", str(text).lower()).encode("ascii", errors="ignore").decode("utf-8")

if uploaded_files:
    all_data = []
    for file in uploaded_files[:50]:
        df = pd.read_excel(file)
        file_name = file.name.replace(".xlsx", "")
        df["Baza"] = file_name
        all_data.append(df)
    
    df_all = pd.concat(all_data, ignore_index=True)

    # Przygotowanie danych
    df_all['LastCallCode_clean'] = df_all['LastCallCode'].astype(str).apply(normalize_text)
    df_all['Skuteczny'] = df_all['LastCallCode_clean'].str.contains("umow|sukces|magazyn")
    df_all['TotalTries'] = df_all['TotalTries'].fillna(0)

    # Agregacja
    summary = df_all.groupby("Baza").agg({
        "Id": "count",
        "TotalTries": "sum",
        "Skuteczny": "sum"
    }).reset_index()

    summary.rename(columns={
        "Baza": "📁 Baza",
        "Id": "📋 Rekordów",
        "TotalTries": "📞 Połączeń",
        "Skuteczny": "✅ Spotkań"
    }, inplace=True)

    summary["📉 CTR"] = round(summary["📞 Połączeń"] / summary["✅ Spotkań"].replace(0, 1), 2)
    summary["💯 L100R"] = round((summary["✅ Spotkań"] / summary["📋 Rekordów"]) * 100, 2)

    # ALERTY – wg Twoich progów
    def alert(row):
        if row["💯 L100R"] <= 0.18:
            return "🔴 Baza martwa"
        elif row["💯 L100R"] >= 5:
            return "🟢 Baza cudowna"
        else:
            return "🟡 Do obserwacji"

    summary["🚨 Alert"] = summary.apply(alert, axis=1)

    st.subheader("📊 Porównanie baz")
    st.dataframe(summary, use_container_width=True)

    # Export Excel z legendą
    st.subheader("📥 Pobierz raport Excel")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        summary.to_excel(writer, index=False, sheet_name="Porównanie baz")
        ws = writer.sheets["Porównanie baz"]
        for i, col in enumerate(summary.columns):
            ws.set_column(i, i, max(15, len(str(col)) + 2))

        # Legenda pod tabelą
        legend = [
            ("📁 Baza", "Nazwa pliku z bazą"),
            ("📋 Rekordów", "Liczba rekordów w bazie (Id)"),
            ("📞 Połączeń", "Suma prób kontaktu (TotalTries)"),
            ("✅ Spotkań", "Rekordy z kodem zawierającym 'umówione', 'sukces', 'magazyn'"),
            ("📉 CTR", "Połączenia / Spotkania – im niższy, tym lepiej"),
            ("💯 L100R", "Spotkania na 100 rekordów – im wyższy, tym lepiej"),
            ("🚨 Alert", "🔴 ≤ 0.18 = martwa baza, 🟢 ≥ 5 = cudowna baza")
        ]
        start_row = len(summary) + 12
        bold = writer.book.add_format({"bold": True})
        ws.write(start_row, 0, "📌 LEGENDA METRYK")
        for label, desc in legend:
            start_row += 1
            ws.write(start_row, 0, label, bold)
            ws.write(start_row, 1, desc)

    st.download_button(
        label="⬇️ Pobierz raport Excel",
        data=buffer.getvalue(),
        file_name="Raport_Porownanie_Baz_ACX.xlsx",
        mime="application/vnd.ms-excel"
    )
