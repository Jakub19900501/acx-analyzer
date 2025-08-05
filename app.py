import streamlit as st
import pandas as pd
import plotly.express as px
import io
import unicodedata

st.set_page_config(page_title="ACX Analyzer V2", layout="wide")
st.title("📞 ACX Analyzer V2 – porównanie baz kontaktowych")

uploaded_files = st.file_uploader("📤 Wgraj dowolną liczbę plików Excel z ACX", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    summary_rows = []
    city_counts = pd.DataFrame()

    for file in uploaded_files:
        file_name = file.name.replace(".xlsx", "")
        df = pd.read_excel(file)

        # ✅ Naprawa: Normalizacja i odporność
        if 'LastCallCode' in df.columns:
            lastcall_clean = df['LastCallCode'].astype(str).str.lower().apply(
                lambda x: unicodedata.normalize('NFKD', x).encode('ascii', errors='ignore').decode('utf-8')
            )
            df['Skuteczny'] = lastcall_clean.str.contains('umowione|potwierdzone')
        else:
            df['Skuteczny'] = False

        df['Błędny numer'] = df['CloseReason'].fillna('').str.lower().str.contains('brak dostępnych telefonów|błędny numer')
        df['Połączony'] = df['CloseReason'].fillna('').str.lower().str.contains('połączony')
        df['Próby'] = df['Tries'].fillna(0)

        total = len(df)
        connected = df['Połączony'].sum()
        leads = df['Skuteczny'].sum()
        bad = df['Błędny numer'].sum()
        tries = df['Próby'].mean()
        l100r = round((leads / total) * 100, 2) if total else 0
        ctr = round(connected / leads, 2) if leads else float('inf')
        error_rate = round((bad / total) * 100, 2)

        summary_rows.append({
            "📁 Baza": file_name,
            "📋 Rekordów": total,
            "📞 Połączeń": connected,
            "✅ Spotkań": leads,
            "🔁 Śr. prób": round(tries, 2),
            "❌ % błędnych": error_rate,
            "📉 CTR": ctr,
            "💯 L100R": l100r
        })

        # Miejscowości (do wykresów)
        if 'Miejscowosc' in df.columns:
            count_df = df['Miejscowosc'].value_counts().head(10).reset_index()
            count_df.columns = ['Miasto', 'Liczba kontaktów']
            count_df['Baza'] = file_name
            city_counts = pd.concat([city_counts, count_df], ignore_index=True)

    summary_df = pd.DataFrame(summary_rows).sort_values(by="💯 L100R", ascending=False)
    st.subheader("📊 Porównanie skuteczności baz")
    st.dataframe(summary_df, use_container_width=True)

    st.subheader("📍 Wykres: TOP 10 miejscowości wg kontaktów (sumarycznie)")
    if not city_counts.empty:
        fig = px.bar(city_counts, x="Miasto", y="Liczba kontaktów", color="Baza", barmode="group")
        st.plotly_chart(fig, use_container_width=True)

    # 📥 Export raportu
    st.subheader("📥 Generuj raport Excel")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, index=False, sheet_name='Porównanie baz')
    st.download_button("⬇️ Pobierz raport Excel", data=buffer.getvalue(), file_name="porownanie_baz.xlsx", mime="application/vnd.ms-excel")

    st.markdown("---")
    st.subheader("📌 Legenda metryk")
    st.markdown("""
    - **CTR** – ile połączeń potrzeba, by umówić 1 spotkanie
    - **L100R** – leady na 100 rekordów
    - **Śr. prób** – średnia liczba prób na rekord
    - **% błędnych** – procent rekordów z błędnym numerem
    """)
