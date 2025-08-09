import streamlit as st
import pandas as pd
import io
import re
import unicodedata
from datetime import datetime

st.set_page_config(page_title="ACX Analyzer", layout="wide")
st.title("📞 ACX Analyzer – porównanie baz (do 50 plików)")

uploaded_files = st.file_uploader("📤 Wgraj pliki ACX (.xlsx)", type=["xlsx"], accept_multiple_files=True)

# ---------- NARZĘDZIA ----------

def normalize_text(text) -> str:
    """lower + bez polskich znaków + bez nadmiarowych spacji"""
    s = unicodedata.normalize("NFKD", str(text).strip().lower()).encode("ascii", errors="ignore").decode("utf-8")
    return re.sub(r"\s+", " ", s)

# Ujednolicony wzorzec sukcesu (umówienie/umówienie magazyn/sukces/magazyn)
SUKCES_REGEX = re.compile(r"(umowienie magazyn|umowienie|umow|sukces|magazyn)")

def is_sukces(x: str) -> bool:
    return bool(SUKCES_REGEX.search(normalize_text(x)))

def klasyfikuj_alert_ctr(ctr_val: float) -> str:
    if pd.isna(ctr_val):
        return "Brak danych"
    if ctr_val < 150:
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

def generuj_wniosek_ctr_roe(ctr_val, roe_val, umowienia, proc_wykorzystania):
    # priorytety
    if (umowienia is None or umowienia == 0):
        return "❌ Brak umówień – baza prawdopodobnie martwa."
    if not pd.isna(ctr_val) and ctr_val >= 1000:
        return "⚠️ CTR ≥ 1000 (na wykorzystanych) – baza wypalona. Rozważ wycofanie/filtry."
    if not pd.isna(roe_val) and roe_val > 5:
        return "✅ ROE > 5% – baza bardzo efektywna. Warto kontynuować."
    if not pd.isna(ctr_val) and ctr_val < 300:
        return "👍 CTR < 300 – kaloryczna baza, szybkie efekty."
    # komentarz o niskim wykorzystaniu
    if proc_wykorzystania < 50:
        return f"⏳ W trakcie – wykorzystanie {proc_wykorzystania:.0f}%. Wnioski ostrożne."
    return ""

# ---------- LOGIKA ----------

if uploaded_files:
    import xlsxwriter

    all_data = []
    for file in uploaded_files[:50]:
        df = pd.read_excel(file)
        df["Baza"] = file.name.replace(".xlsx", "")
        all_data.append(df)

    df_all = pd.concat(all_data, ignore_index=True)

    # Bezpieczne kolumny wejściowe (tworzymy jeśli brakuje)
    for col in ["Id","LastCallCode","TotalTries","LastTryTime","ImportCreatedOn","CloseReason",
                "RecordState","EndReason","LastCallReason"]:
        if col not in df_all.columns:
            df_all[col] = None

    # Normalizacje pomocnicze
    df_all["LastCallCode_clean"] = df_all["LastCallCode"].apply(normalize_text)
    df_all["CloseReason_clean"]   = df_all["CloseReason"].apply(normalize_text)
    df_all["RecordState_clean"]   = df_all["RecordState"].apply(normalize_text)
    df_all["EndReason_clean"]     = df_all["EndReason"].apply(normalize_text)
    df_all["LastCallReason_clean"]= df_all["LastCallReason"].apply(normalize_text)

    # Flagi
    df_all["Skuteczny"]       = df_all["LastCallCode"].apply(is_sukces)

    # Ponowny kontakt (pozostaje na LastCallCode – zgodnie z Twoim opisem)
    df_all["PonownyKontakt"]  = df_all["LastCallCode_clean"].str.contains(r"\bponowny kontakt\b")

    # Błędny – teraz WYŁĄCZNIE z CloseReason == "brak dostępnych telefonow"
    df_all["Bledny"]          = df_all["CloseReason_clean"].eq("brak dostepnych telefonow")

    # Niewydzwonione (otwarte), Przełożony, Zamknięte przez system/konsultanta
    df_all["Otwarte"]         = df_all["RecordState_clean"].eq("otwarty")
    df_all["Przelozony"]      = df_all["RecordState_clean"].eq("przelozony")
    df_all["ZamknSystem"]     = df_all["EndReason_clean"].eq("nie udalo sie polaczyc")
    df_all["ZamknKons"]       = (df_all["LastCallReason_clean"].notna() &
                                 (df_all["LastCallReason_clean"] != "") &
                                 (~df_all["LastCallReason_clean"].eq("ponowny kontakt")))

    # Liczby/czasy
    df_all["TotalTries"]      = pd.to_numeric(df_all["TotalTries"], errors="coerce").fillna(0)
    df_all["LastTryTime"]     = pd.to_datetime(df_all["LastTryTime"], errors="coerce")
    df_all["ImportCreatedOn"] = pd.to_datetime(df_all["ImportCreatedOn"], errors="coerce")

    # Agregacje per Baza
    summary = df_all.groupby("Baza").agg({
        "Id": "count",
        "TotalTries": "sum",
        "Skuteczny": "sum",
        "PonownyKontakt": "sum",
        "Bledny": "sum",
        "LastTryTime": "max",
        "ImportCreatedOn": "min",
        "Otwarte": "sum",
        "Przelozony": "sum",
        "ZamknSystem": "sum",
        "ZamknKons": "sum"
    }).reset_index()

    # Nazwy PL
    summary.rename(columns={
        "Baza": "📁 Baza",
        "Id": "📋 Rekordów",
        "TotalTries": "📞 Połączeń",
        "Skuteczny": "✅ Spotkań",
        "PonownyKontakt": "🔁 Ponowny kontakt",
        "Bledny": "❌ Brak tel. (CloseReason)",
        "LastTryTime": "📅 Ostatni kontakt",
        "ImportCreatedOn": "🕓 Data importu",
        "Otwarte": "🟦 Niewydzwonione (otwarte)",
        "Przelozony": "🟧 Przełożony",
        "ZamknSystem": "🤖 Zamkn. system",
        "ZamknKons": "👤 Zamkn. konsultant"
    }, inplace=True)

    # Wykorzystanie bazy
    summary["% Niewykorzystane"] = (summary["🟦 Niewydzwonione (otwarte)"] / summary["📋 Rekordów"] * 100).round(2)
    summary["% Wykorzystane"]    = (100 - summary["% Niewykorzystane"]).round(2)

    # Metryki
    # L100R: spotkania na 100 rekordów (całej bazy)
    summary["💯 L100R"] = (summary["✅ Spotkań"] / summary["📋 Rekordów"] * 100).round(2)

    # CTR i ROE – same z natury bazują tylko na faktycznych próbach/umówieniach (czyli na wykorzystanych rekordach)
    # Zabezpieczenia na 0
    _spotkania = summary["✅ Spotkań"].replace(0, pd.NA)
    _polaczenia = summary["📞 Połączeń"].replace(0, pd.NA)

    summary["📉 CTR"] = (summary["📞 Połączeń"] / _spotkania).round(2)
    summary["ROE (%)"] = (summary["✅ Spotkań"] / _polaczenia * 100).round(2)

    # Średnia prób / % ponownego
    summary["🔁 % Ponowny kontakt"] = (summary["🔁 Ponowny kontakt"] / summary["📋 Rekordów"] * 100).round(2)
    summary["🔁 Śr. prób"] = (summary["📞 Połączeń"] / summary["📋 Rekordów"].replace(0, 1)).round(2)

    # Czas
    summary["⏳ Śr. czas reakcji (dni)"] = (summary["📅 Ostatni kontakt"] - summary["🕓 Data importu"]).dt.days

    # Alert CTR wg wykorzystania (CTR jako metryka pozostaje, ale dodajemy kontekst % wykorzystania)
    summary["🚨 Alert CTR (wykorz.)"] = summary["📉 CTR"].apply(klasyfikuj_alert_ctr)

    # Wniosek (z komentarzem o niskim wykorzystaniu)
    summary["📝 Wniosek"] = summary.apply(
        lambda r: generuj_wniosek_ctr_roe(
            r["📉 CTR"], r["ROE (%)"], r["✅ Spotkań"], r["% Wykorzystane"]
        ),
        axis=1
    )

    # Kolejność kolumn
    metryki_kolejnosc = [
        "📁 Baza",
        "💯 L100R", "📉 CTR", "ROE (%)", "🔁 % Ponowny kontakt", "🔁 Śr. prób",
        "% Wykorzystane", "% Niewykorzystane",
        "📋 Rekordów", "🟦 Niewydzwonione (otwarte)", "🟧 Przełożony",
        "🤖 Zamkn. system", "👤 Zamkn. konsultant",
        "📞 Połączeń", "✅ Spotkań", "🔁 Ponowny kontakt",
        "❌ Brak tel. (CloseReason)",
        "📅 Ostatni kontakt", "🕓 Data importu", "⏳ Śr. czas reakcji (dni)",
        "🚨 Alert CTR (wykorz.)", "📝 Wniosek"
    ]
    summary = summary[[c for c in metryki_kolejnosc if c in summary.columns]]

    # Sortowanie: najpierw po Alert/CTR (rosnąco CTR – lepsze wyżej), przy remisie po % wykorzystane malejąco
    summary = summary.sort_values(by=["📉 CTR", "% Wykorzystane"], ascending=[True, False])

    # ---------- ANALIZA PONOWNYCH KONTAKTÓW ----------
    ponowne = df_all[df_all["TotalTries"] > 1].copy()
    # Sukces jak wyżej – ten sam regex
    ponowne["Skuteczne"] = ponowne["LastCallCode"].apply(is_sukces)
    ponowne_umowienia = ponowne[ponowne["Skuteczne"] == True]

    ponowny_raport = []
    for baza in ponowne["Baza"].unique():
        baza_all = ponowne[ponowne["Baza"] == baza]
        baza_succ = ponowne_umowienia[ponowne_umowienia["Baza"] == baza]

        if len(baza_succ) > 0:
            rozklad = baza_succ["TotalTries"].value_counts().sort_index()
            sr = float(baza_succ["TotalTries"].mean())
            med = float(baza_succ["TotalTries"].median())
            rozklad_str = ", ".join([f"przy {int(k)}. próbie: {v} umówień" for k, v in rozklad.items()])
            umowienia_cnt = len(baza_succ)
        else:
            rozklad_str, sr, med, umowienia_cnt = "", float("nan"), float("nan"), 0

        ponowny_raport.append({
            "📁 Baza": baza,
            "🔁 Rekordów ponownych (>1 próba)": len(baza_all),
            "✅ Umówienia (z ponownych)": umowienia_cnt,
            "📈 Śr. próba umówienia": None if pd.isna(sr) else round(sr, 2),
            "🎯 Mediana próby": None if pd.isna(med) else round(med, 2),
            "📊 Rozkład prób": rozklad_str
        })

    ponowna_analiza = pd.DataFrame(ponowny_raport)

    # ---------- WIDOK ----------
    st.subheader("📊 Porównanie baz – rozszerzone")
    st.dataframe(summary, use_container_width=True)

    st.subheader("📊 Skuteczność ponownych kontaktów")
    st.dataframe(ponowna_analiza, use_container_width=True)

    # ---------- EXPORT XLSX ----------
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        summary.to_excel(writer, index=False, sheet_name="Porównanie baz")
        ponowna_analiza.to_excel(writer, index=False, sheet_name="Ponowny kontakt")

        wb = writer.book
        ws_summary = writer.sheets["Porównanie baz"]
        ws_ponowny = writer.sheets["Ponowny kontakt"]

        ws_summary.freeze_panes(1, 6)  # zamroź do kolumny z % wykorzystane
        ws_ponowny.freeze_panes(1, 0)

        for i, _ in enumerate(summary.columns):
            ws_summary.set_column(i, i, 24)
        for i, _ in enumerate(ponowna_analiza.columns):
            ws_ponowny.set_column(i, i, 34)

        chart_sheet = wb.add_worksheet("Wykresy")
        metrics = ["💯 L100R", "📉 CTR", "ROE (%)", "% Wykorzystane", "% Niewykorzystane", "🔁 % Ponowny kontakt", "🔁 Śr. prób"]
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

        # Legenda
        legenda = [
            ("💯 L100R", "Spotkania na 100 rekordów (całej bazy)"),
            ("📉 CTR", "Połączenia / umówienia (liczone na faktycznie wykorzystanych rekordach)"),
            ("ROE (%)", "% połączeń zakończonych umówieniem"),
            ("% Wykorzystane", "Jaki % rekordów nie jest 'otwarty'"),
            ("% Niewykorzystane", "Jaki % rekordów ma RecordState='otwarty'"),
            ("🔁 % Ponowny kontakt", "Odsetek rekordów z oznaczeniem ponownego kontaktu"),
            ("🔁 Śr. prób", "Średnia prób per rekord"),
            ("📋 Rekordów", "Liczba rekordów w pliku"),
            ("🟦 Niewydzwonione (otwarte)", "Liczba rekordów z RecordState='otwarty'"),
            ("🟧 Przełożony", "Liczba rekordów z RecordState='przełożony'"),
            ("🤖 Zamkn. system", "EndReason='nie udało się połączyć'"),
            ("👤 Zamkn. konsultant", "LastCallReason ≠ 'ponowny kontakt' i niepusty"),
            ("❌ Brak tel. (CloseReason)", "CloseReason='brak dostępnych telefonow'"),
            ("📅 Ostatni kontakt", "Maksymalna data LastTryTime"),
            ("⏳ Śr. czas reakcji (dni)", "Różnica: Ostatni kontakt – Data importu"),
            ("🚨 Alert CTR (wykorz.)", "Klasyfikacja jakości wg CTR"),
            ("📝 Wniosek", "Komentarz na bazie CTR/ROE/% wykorzystania")
        ]

        for ws, start in [
            (ws_summary, len(summary) + 4),
            (ws_ponowny, len(ponowna_analiza) + 4),
            (chart_sheet, 130)
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

