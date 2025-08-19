import streamlit as st
import pandas as pd
import io
import re
import unicodedata
from datetime import datetime

st.set_page_config(page_title="ACX Analyzer", layout="wide")
st.title("📞 ACX Analyzer – porównanie baz (do 50 plików)")

uploaded_files = st.file_uploader("📤 Wgraj pliki ACX (.xlsx)", type=["xlsx"], accept_multiple_files=True)

# ---------- UTIL ----------

def normalize_text(text) -> str:
    s = unicodedata.normalize("NFKD", str(text).strip().lower()).encode("ascii", errors="ignore").decode("utf-8")
    return re.sub(r"\s+", " ", s)

def resolve_col(df: pd.DataFrame, *cands):
    """Zwraca istniejącą kolumnę wg pierwszego trafienia (case-insensitive), wpp tworzy pustą o nazwie pierwszego kandydata."""
    lower = {c.lower(): c for c in df.columns}
    for c in cands:
        if c.lower() in lower:
            return lower[c.lower()]
    new_name = cands[0]
    df[new_name] = None
    return new_name

# Sukces (w tym "umówienie magazyn")
SUKCES_REGEX = re.compile(r"(umowienie magazyn|umowienie|umow|sukces|magazyn)")

# Wzorce
PON_KONTAKT_PAT = re.compile(r"\bponowny kontakt\b")
# Elastyczne: "ponowny kontakt (system)", "ponowny kontakt - system", "systemowy ponowny kontakt", itp.
PON_KONTAKT_SYSTEM_PAT = re.compile(
    r"\bponown\w*\s*konta\w*(?:\s*[-(]\s*)?system\w*\b|\bsystemow\w*\s*ponown\w*\s*konta\w*\b"
)

def is_sukces(x: str) -> bool:
    return bool(SUKCES_REGEX.search(normalize_text(x)))

def klasyfikuj_alert_ctr_with_util(ctr_val: float, util_pct: float) -> str:
    if pd.isna(util_pct):
        return "Brak danych"
    if util_pct < 40:
        return "⏳ Za wczesnie (wykorz. <40%)"
    if pd.isna(ctr_val):
        return "Brak danych"
    if ctr_val < 150:   return "🟣 Genialna"
    if ctr_val < 300:   return "🟢 Bardzo dobra"
    if ctr_val < 500:   return "🟡 Solidna"
    if ctr_val < 700:   return "🟠 Przeciętna"
    if ctr_val < 1000:  return "🔴 Słaba"
    return "⚫ Martwa"

def status_bazy(util_pct: float) -> str:
    if pd.isna(util_pct):
        return "Brak danych"
    if util_pct >= 90:
        return "🔴 Prawie pusta – czas dokupić"
    if util_pct >= 70:
        return "🟡 Na wyczerpaniu"
    return "🟢 OK"

def generuj_wniosek(ctr_val, roe_val, umowienia, util_pct):
    if (umowienia is None or umowienia == 0):
        return "❌ Brak umówień – baza prawdopodobnie martwa."
    if util_pct < 40:
        return f"⏳ W trakcie – wykorzystanie {util_pct:.0f}%. Nie wyciągaj pochopnych wniosków."
    if not pd.isna(ctr_val) and ctr_val >= 1000:
        return "⚠️ CTR ≥ 1000 – baza wypalona. Rozważ wycofanie/filtry."
    if not pd.isna(roe_val) and roe_val > 5:
        return "✅ ROE > 5% – baza bardzo efektywna. Warto kontynuować."
    if not pd.isna(ctr_val) and ctr_val < 300:
        return "👍 CTR < 300 – kaloryczna baza."
    return ""

# ---------- MAIN ----------

if uploaded_files:
    import xlsxwriter

    frames = []
    for file in uploaded_files[:50]:
        df = pd.read_excel(file)
        df["Baza"] = file.name.replace(".xlsx", "")

        # mapowanie nazw
        col_id             = resolve_col(df, "Id","id")
        col_lcc            = resolve_col(df, "LastCallCode","lastcallcode")
        col_lcr            = resolve_col(df, "LastCallReason","lastcallreason","Last Call Reason","last call reason")
        col_tries          = resolve_col(df, "TotalTries","totaltries","Tries","tries")
        col_lasttry        = resolve_col(df, "LastTryTime","lasttrytime")
        col_import         = resolve_col(df, "ImportCreatedOn","importcreatedon")
        col_closereason    = resolve_col(df, "CloseReason","closereason","Close Reason","close reason")
        col_recordstate    = resolve_col(df, "RecordState","recordstate","State","state","Status","status")
        col_endreason      = resolve_col(df, "EndReason","endreason","End Reason","end reason")
        col_disc_reason    = resolve_col(df, "DisconnectedReason","disconnectedreason")
        col_disc_cause     = resolve_col(df, "DisconnectedCause","disconnectedcause")

        # normalizacja
        for c in [col_lcc,col_lcr,col_closereason,col_recordstate,col_endreason,col_disc_reason,col_disc_cause]:
            df[c+"_clean"] = df[c].apply(normalize_text)

        rs_clean  = df[col_recordstate+"_clean"].fillna("")
        lcr_clean = df[col_lcr+"_clean"].fillna("")
        lcc_clean = df[col_lcc+"_clean"].fillna("")
        dcc_clean = df[col_disc_cause+"_clean"].fillna("")

        # --- flagi ---
        df["Skuteczny"] = df[col_lcc].apply(is_sukces)

        # ogólny "Ponowny kontakt"
        pon_from_reason = lcr_clean.str.contains(PON_KONTAKT_PAT, na=False)
        pon_from_code   = lcc_clean.str.contains(PON_KONTAKT_PAT, na=False)
        df["PonownyKontakt"] = pon_from_reason | pon_from_code

        # "Ponowny kontakt (system)" (po reason/code/status – wszystkie warianty)
        pon_sys_from_reason = lcr_clean.str.contains(PON_KONTAKT_SYSTEM_PAT, na=False)
        pon_sys_from_code   = lcc_clean.str.contains(PON_KONTAKT_SYSTEM_PAT, na=False)
        pon_sys_from_state  = rs_clean.str.contains(PON_KONTAKT_SYSTEM_PAT, na=False)
        df["PonownyKontaktSystem"] = pon_sys_from_reason | pon_sys_from_code | pon_sys_from_state

        # "Przełożony" – licz TYLKO czyste warianty RecordState; wyklucz „ponowny kontakt (system)”
        df["Przelozony"] = (
            rs_clean.isin(["przelozony","przeozony","prze lozony"])
            & (~df["PonownyKontaktSystem"])
        )

        # "Otwarte"
        df["Otwarte"] = rs_clean.str.contains(r"\botwart", na=False)

        # "Zamkn. konsultant": jest LastCallReason i ≠ 'ponowny kontakt'
        df["ZamknKons"] = (lcr_clean != "") & (~lcr_clean.str.contains(PON_KONTAKT_PAT, na=False))

        # "Zamkn. system": DOKŁADNIE wg Twojej definicji
        # RecordState = Zamknięty AND DisconnectedCause = "Błędny numer"
        df["ZamknSystem"] = (
            rs_clean.str.contains(r"\bzamkn", na=False)
            & dcc_clean.str.contains(r"\bbledny numer\b", na=False)
        )

        # "Błędny numer" (CloseReason wariant)
        df["Bledny"] = df[col_closereason+"_clean"].fillna("").str.contains(r"brak dostepnych telefon", na=False)

        # liczby/czasy
        df["TotalTries"]      = pd.to_numeric(df[col_tries], errors="coerce").fillna(0)
        df["LastTryTime"]     = pd.to_datetime(df[col_lasttry], errors="coerce")
        df["ImportCreatedOn"] = pd.to_datetime(df[col_import], errors="coerce")

        df.rename(columns={col_id:"Id"}, inplace=True)
        frames.append(df)

    df_all = pd.concat(frames, ignore_index=True)

    # --- agregacja per baza ---
    summary = df_all.groupby("Baza").agg({
        "Id": "count",
        "TotalTries": "sum",
        "Skuteczny": "sum",
        "PonownyKontakt": "sum",
        "PonownyKontaktSystem": "sum",
        "Bledny": "sum",
        "LastTryTime": "max",
        "ImportCreatedOn": "min",
        "Otwarte": "sum",
        "Przelozony": "sum",
        "ZamknSystem": "sum",
        "ZamknKons": "sum"
    }).reset_index()

    summary.rename(columns={
        "Baza": "📁 Baza",
        "Id": "📋 Rekordów",
        "TotalTries": "📞 Połączeń",
        "Skuteczny": "✅ Spotkań",
        "PonownyKontakt": "🔁 Ponowny kontakt",
        "PonownyKontaktSystem": "🔁 Ponowny kontakt (system)",
        "Bledny": "❌ Brak tel. (CloseReason)",
        "LastTryTime": "📅 Ostatni kontakt",
        "ImportCreatedOn": "🕓 Data importu",
        "Otwarte": "🟦 Niewydzwonione (otwarte)",
        "Przelozony": "🟧 Przełożony",
        "ZamknSystem": "🤖 Zamkn. system",
        "ZamknKons": "👤 Zamkn. konsultant"
    }, inplace=True)

    # --- wykorzystanie ---
    niewyk = summary["🟦 Niewydzwonione (otwarte)"] + summary["🔁 Ponowny kontakt"]
    summary["% Niewykorzystane"] = (niewyk / summary["📋 Rekordów"] * 100).round(2)
    summary["% Wykorzystane"]    = (100 - summary["% Niewykorzystane"]).round(2)
    summary["🛒 Status bazy"]    = summary["% Wykorzystane"].apply(status_bazy)

    # --- metryki ---
    summary["💯 L100R"] = (summary["✅ Spotkań"] / summary["📋 Rekordów"] * 100).round(2)
    spotkania_safe  = summary["✅ Spotkań"].replace(0, pd.NA)
    polaczenia_safe = summary["📞 Połączeń"].replace(0, pd.NA)
    summary["📉 CTR"]  = (summary["📞 Połączeń"] / spotkania_safe).round(2)
    summary["ROE (%)"] = (summary["✅ Spotkań"] / polaczenia_safe * 100).round(2)
    summary["🔁 % Ponowny kontakt"] = (summary["🔁 Ponowny kontakt"] / summary["📋 Rekordów"] * 100).round(2)
    summary["🔁 % Ponowny kontakt (system)"] = (summary["🔁 Ponowny kontakt (system)"] / summary["📋 Rekordów"] * 100).round(2)
    summary["🔁 Śr. prób"] = (summary["📞 Połączeń"] / summary["📋 Rekordów"].replace(0,1)).round(2)
    summary["⏳ Śr. czas reakcji (dni)"] = (summary["📅 Ostatni kontakt"] - summary["🕓 Data importu"]).dt.days

    summary["🚨 Alert CTR"] = summary.apply(
        lambda r: klasyfikuj_alert_ctr_with_util(r["📉 CTR"], r["% Wykorzystane"]), axis=1
    )
    summary["📝 Wniosek"] = summary.apply(
        lambda r: generuj_wniosek(r["📉 CTR"], r["ROE (%)"], r["✅ Spotkań"], r["% Wykorzystane"]),
        axis=1
    )

    # kolejność i sort
    order = [
        "📁 Baza",
        "💯 L100R", "📉 CTR", "ROE (%)",
        "% Wykorzystane", "% Niewykorzystane", "🛒 Status bazy",
        "🔁 % Ponowny kontakt", "🔁 % Ponowny kontakt (system)", "🔁 Śr. prób",
        "📋 Rekordów",
        "🟦 Niewydzwonione (otwarte)", "🔁 Ponowny kontakt", "🔁 Ponowny kontakt (system)",
        "🟧 Przełożony", "🤖 Zamkn. system", "👤 Zamkn. konsultant",
        "❌ Brak tel. (CloseReason)",
        "📞 Połączeń", "✅ Spotkań",
        "📅 Ostatni kontakt", "🕓 Data importu", "⏳ Śr. czas reakcji (dni)",
        "🚨 Alert CTR", "📝 Wniosek"
    ]
    summary = summary[[c for c in order if c in summary.columns]]
    summary = summary.sort_values(by=["📉 CTR", "% Wykorzystane"], ascending=[True, False])

    # ---------- PONOWNY KONTAKT (głębiej) ----------
    ponowne = df_all[df_all["TotalTries"] > 1].copy()
    ponowne["Skuteczne"] = df_all.loc[ponowne.index, "Skuteczny"]
    ponowne_um = ponowne[ponowne["Skuteczne"]]

    rows = []
    for baza in ponowne["Baza"].unique():
        b_all = ponowne[ponowne["Baza"] == baza]
        b_ok  = ponowne_um[ponowne_um["Baza"] == baza]

        if len(b_ok) > 0:
            vc = b_ok["TotalTries"].value_counts().sort_index()
            rozklad = ", ".join([f"przy {int(k)}. próbie: {v} umówień" for k, v in vc.items()])
            sr = round(float(b_ok["TotalTries"].mean()), 2)
            med = round(float(b_ok["TotalTries"].median()), 2)
            okcnt = len(b_ok)
        else:
            rozklad, sr, med, okcnt = "", float("nan"), float("nan"), 0

        pon_all = df_all[(df_all["Baza"]==baza) & (df_all["PonownyKontakt"]==True)]
        pon_sys = df_all[(df_all["Baza"]==baza) & (df_all["PonownyKontaktSystem"]==True)]

        rows.append({
            "📁 Baza": baza,
            "🔁 Rekordów ponownych (>1 próba)": len(b_all),
            "✅ Umówienia (z ponownych)": okcnt,
            "📈 Śr. próba umówienia": None if pd.isna(sr) else sr,
            "🎯 Mediana próby": None if pd.isna(med) else med,
            "📊 Rozkład prób": rozklad,
            "🔁 Ponowny kontakt (wszystkie)": len(pon_all),
            "🔁 Ponowny kontakt (system)": len(pon_sys)
        })
    ponowna_analiza = pd.DataFrame(rows)

    # ---------- UI ----------
    st.subheader("📊 Porównanie baz – rozszerzone")
    st.dataframe(summary, use_container_width=True)

    st.subheader("📊 Skuteczność ponownych kontaktów")
    st.dataframe(ponowna_analiza, use_container_width=True)

    # ---------- EXPORT ----------
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        summary.to_excel(writer, index=False, sheet_name="Porównanie baz")
        ponowna_analiza.to_excel(writer, index=False, sheet_name="Ponowny kontakt")

        wb = writer.book
        ws_s = writer.sheets["Porównanie baz"]
        ws_p = writer.sheets["Ponowny kontakt"]

        ws_s.freeze_panes(1, 7)
        ws_p.freeze_panes(1, 0)

        for i,_ in enumerate(summary.columns):
            ws_s.set_column(i, i, 24)
        for i,_ in enumerate(ponowna_analiza.columns):
            ws_p.set_column(i, i, 34)

        chart_sheet = wb.add_worksheet("Wykresy")
        metrics = ["💯 L100R","📉 CTR","ROE (%)","% Wykorzystane","% Niewykorzystane","🔁 % Ponowny kontakt","🔁 % Ponowny kontakt (system)","🔁 Śr. prób"]
        for i, m in enumerate(metrics):
            ch = wb.add_chart({'type':'column'})
            ch.add_series({
                'name': m,
                'categories': ['Porównanie baz', 1, 0, len(summary), 0],
                'values':     ['Porównanie baz', 1, summary.columns.get_loc(m), len(summary), summary.columns.get_loc(m)]
            })
            ch.set_title({'name': m})
            ch.set_x_axis({'name':'Baza'})
            ch.set_y_axis({'name': m})
            ch.set_size({'width': 1440, 'height': 480})
            chart_sheet.insert_chart(i*25, 0, ch)

        legenda = [
            ("💯 L100R","Spotkania na 100 rekordów."),
            ("📉 CTR","Połączenia / umówienia – zawsze z realnych prób/umówień."),
            ("ROE (%)","% połączeń zakończonych umówieniem."),
            ("% Wykorzystane","100 - % niewykorzystanych; niewykorzystane = otwarte + wszystkie ponowne kontakty."),
            ("% Niewykorzystane","Otwarte + wszystkie 'ponowny kontakt'."),
            ("🛒 Status bazy","🟢 <70%, 🟡 70–89%, 🔴 ≥90%."),
            ("🔁 % Ponowny kontakt","Odsetek rekordów z 'ponowny kontakt'."),
            ("🔁 % Ponowny kontakt (system)","Odsetek rekordów z 'ponowny kontakt (system)'."),
            ("🔁 Śr. prób","Średnia prób per rekord."),
            ("🟦 Niewydzwonione (otwarte)","RecordState = otwarty."),
            ("🟧 Przełożony","Tylko RecordState = 'przelozony/przeozony' (bez 'ponowny kontakt (system)')."),
            ("🤖 Zamkn. system","RecordState = 'zamkniety' AND DisconnectedCause = 'bledny numer'."),
            ("👤 Zamkn. konsultant","LastCallReason ustawione i ≠ 'ponowny kontakt'."),
            ("❌ Brak tel. (CloseReason)","CloseReason zawiera 'brak dostepnych telefon'."),
            ("🚨 Alert CTR","Kolor wg CTR tylko gdy wykorzystanie ≥40%; przy <40% → ⏳ Za wczesnie."),
            ("📝 Wniosek","Komentarz wg CTR/ROE/% wykorzystania.")
        ]
        start_rows = [
            (ws_s, len(summary)+4),
            (ws_p, len(ponowna_analiza)+4),
            (chart_sheet, 130)
        ]
        for ws, start in start_rows:
            ws.write(start, 0, "📌 LEGENDA METRYK")
            for idx, (lbl, desc) in enumerate(legenda, start+1):
                ws.write(idx, 0, lbl)
                ws.write(idx, 1, desc)

    st.download_button(
        "⬇️ Pobierz raport Excel",
        data=buffer.getvalue(),
        file_name="Raport_Porownanie_Baz_ACX.xlsx",
        mime="application/vnd.ms-excel"
    )
