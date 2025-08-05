import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="ACX Analyzer", layout="wide")
st.title("📞 ACX Analyzer – analiza baz kontaktowych")

uploaded_file = st.file_uploader("📤 Wgraj plik Excel z ACX", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Przetwarzanie danych
    df['Skuteczny'] = df['CloseReason'].fillna('').str.lower().str.contains('umówione|potwierdzone')
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
    burnout = round((connected / total) * 100, 2)

    # Metryki
    col1, col2, col3 = st.columns(3)
    col1.metric("📋 Rekordów", total)
    col2.metric("📞 Połączeń", connected)
    col3.metric("✅ Spotkań", leads)

    col1.metric("❌ Błędnych nr", bad)
    col2.metric("🔁 Śr. prób", f"{tries:.2f}")
    col3.metric("💯 L100R", l100r)

    st.markdown("---")

    # Wykresy
    st.subheader("📊 Wykres kontaktów wg miejscowości")
    if 'Miejscowosc' in df.columns:
        top_city = df['Miejscowosc'].value_counts().head(20).reset_index()
        top_city.columns = ['Miasto', 'Liczba kontaktów']
        fig = px.bar(top_city, x='Miasto', y='Liczba kontaktów')
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("📥 Generuj raport Excel")

    # Export Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Dane')
        summary = pd.DataFrame({
            'Metryka': ['Rekordy', 'Połączone', 'Spotkania', 'Błędne nr', 'Śr. prób', 'CTR', 'L100R', 'Wypalenie %'],
            'Wartość': [total, connected, leads, bad, round(tries, 2), ctr, l100r, burnout]
        })
        summary.to_excel(writer, index=False, sheet_name='Podsumowanie')
        writer.save()

        st.download_button("⬇️ Pobierz raport Excel", data=buffer.getvalue(),
                           file_name="raport_acx.xlsx", mime="application/vnd.ms-excel")

    # Legenda
    st.markdown("---")
    st.subheader("📌 Legenda metryk")
    st.markdown("""
    - **CTR** – ile kontaktów potrzeba, by umówić 1 spotkanie
    - **L100R** – liczba leadów na każde 100 rekordów
    - **Wypalenie** – % rekordów już kontaktowanych
    - **Błędne numery** – numery, z którymi nie udało się połączyć
    - **Śr. prób** – średnia liczba prób na rekord
    """)

