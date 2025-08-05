import streamlit as st
import pandas as pd
import plotly.express as px
import io
import unicodedata
from datetime import datetime

st.set_page_config(page_title="ACX Analyzer V3", layout="wide")
st.title("📞 ACX Analyzer V3 – analiza i porównanie baz")

uploaded_files = st.file_uploader("📤 Wgraj dowolną liczbę plików Excel z ACX", type=["xlsx"], accept_multiple_files=True)

def normalize(text):
    return unicodedata.normalize('NFKD', str(text).lower()).encode('ascii', errors='ignore').decode('utf-8')

def extract_top(values, top_n=3):
    return ', '.join(pd.Series(values).value_counts().head(top_n).index.tolist())

def calc_days_delta(start, end):
    try:
        delta = (end - start).days
        return delta if delta >= 0 else None
    except:
        return None

if uploaded_files:
    summary = []
    city_counts = pd.DataFrame()
    for file in uploaded_files:
        file_name = file.name.replace(".xlsx", "")
        df = pd.read_excel(file)

        df['LastCallCode_clean'] = df['LastCallCode'].astype(str).apply(normalize)
        df['Skuteczny'] = df['LastCallCode_clean'].str.contains("umow|spotkanie|sukces|magazyn")
        df['Bledny'] = df['CloseReason'].astype(str).apply(normalize).str.contains("brak dostepnych telefonow|bledny numer")
        df['Polaczony'] = df['CloseReason'].astype(str).apply(normalize).str.contains("polaczony")
        df['Proby'] = df['Tries'].fillna(0)

        # Czas reakcji
        df['ImportCreatedOn'] = pd.to_datetime(df.get('ImportCreatedOn'), errors='coerce')
        df['ShiftTime'] = pd.to_datetime(df.get('ShiftTime'), errors='coerce')
        df['CzasReakcji'] = df.apply(lambda row: calc_days_delta(row['ImportCreatedOn'], row['ShiftTime']), axis=1)

        total = len(df)
        polaczone = df['Polaczony'].sum()
        skuteczne = df['Skuteczny'].sum()
        bledne = df['Bledny'].sum()
        proby = df['Proby'].mean()
        l100r = round((skuteczne / total) * 100, 2) if total else 0
        ctr = round(polaczone / skuteczne, 2) if skuteczne else float('inf')
        err_rate = round((bledne / total) * 100, 2)
        last_date = df['ShiftTime'].max().strftime('%Y-%m-%d') if not df['ShiftTime'].isna().all() else "brak danych"
        reakcja = round(df['CzasReakcji'].dropna().mean(), 2) if not df['CzasReakcji'].dropna().empty else None

        # TOP odmowy i konsultanci
        odmowy = extract_top(df['LastCallCode_clean'])
        konsultanci = extract_top(df['LastTryUser'].dropna())

        # Regiony
        regiony = extract_top(df['Miejscowosc']) if 'Miejscowosc' in df.columns else ""

        # Alert
        if l100r < 3:
            alert = "🔴 Baza martwa – L100R < 3"
        elif ctr > 12:
            alert = "🟠 CTR wysoki – trudne umawianie"
        elif err_rate > 10:
            alert = "⚠️ Dużo błędnych numerów"
        elif l100r > 12:
            alert = "🟢 Baza kaloryczna"
        else:
            alert = "🟡 Średnia skuteczność"

        summary.append({
            "📁 Baza": file_name,
            "📋 Rekordów": total,
            "📞 Połączeń": polaczone,
            "✅ Spotkań": skuteczne,
            "🔁 Śr. prób": round(proby, 2),
            "❌ % błędnych": err_rate,
            "📉 CTR": ctr,
            "💯 L100R": l100r,
            "📅 Ostatni kontakt": last_date,
            "👤 Konsultanci": konsultanci,
            "🧱 Top odmowy": odmowy,
            "⏳ Śr. czas reakcji (dni)": reakcja,
            "🧭 Regiony": regiony,
            "🚨 Alert": alert
        })

        # Wykres miejscowości
        if 'Miejscowosc' in df.columns:
            count_df = df['Miejscowosc'].value_counts().head(10).reset_index()
            count_df.columns = ['Miasto', 'Liczba kontaktów']
            count_df['Baza'] = file_name
            city_counts = pd.concat([city_counts, count_df], ignore_index=True)

    summary_df = pd.DataFrame(summary)
    st.subheader("📊 Tabela porównawcza baz")
    st.dataframe(summary_df, use_container_width=True)

    st.subheader("📍 Wykres kontaktów wg miejscowości")
    if not city_counts.empty:
        fig = px.bar(city_counts, x="Miasto", y="Liczba kontaktów", color="Baza", barmode="group")
        st.plotly_chart(fig, use_container_width=True)

    # 📥 Raport Excel z legendą i alertami
    st.subheader("📥 Pobierz raport Excel")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, sheet_name='Porównanie baz', index=False)
        worksheet = writer.sheets['Porównanie baz']
        for i, col in enumerate(summary_df.columns):
            worksheet.set_column(i, i, max(15, len(col) + 2))
        start_row = len(summary_df) + 12
        worksheet.write(start_row, 0, "📌 LEGENDA METRYK")
        legend = [
            ("📋 Rekordów", "Liczba wszystkich rekordów w bazie"),
            ("📞 Połączeń", "Liczba rekordów oznaczonych jako połączone"),
            ("✅ Spotkań", "Rekordy oznaczone jako skuteczne (umówione, sukces)"),
            ("🔁 Śr. prób", "Średnia liczba prób kontaktu na rekord"),
            ("❌ % błędnych", "Procent błędnych numerów"),
            ("📉 CTR", "Ile połączeń potrzeba na 1 spotkanie"),
            ("💯 L100R", "Leady na 100 rekordów"),
            ("📅 Ostatni kontakt", "Data ostatniej rozmowy w bazie"),
            ("👤 Konsultanci", "Najaktywniejsi dzwoniący"),
            ("🧱 Top odmowy", "Najczęstsze powody odmowy"),
            ("⏳ Śr. czas reakcji", "Średni czas od importu do kontaktu"),
            ("🧭 Regiony", "Najczęściej występujące miejscowości"),
            ("🚨 Alert", "Szybka ocena skuteczności bazy")
        ]
        bold = writer.book.add_format({'bold': True})
        for label, desc in legend:
            start_row += 1
            worksheet.write(start_row, 0, label, bold)
            worksheet.write(start_row, 1, desc)

    st.download_button("⬇️ Pobierz Excel", data=buffer.getvalue(), file_name="Raport_ACX_V3.xlsx", mime="application/vnd.ms-excel")
