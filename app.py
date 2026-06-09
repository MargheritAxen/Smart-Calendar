import streamlit as st
import pandas as pd
import holidays
import gspread
from google.oauth2.service_account import Credentials
from datetime import date
from io import BytesIO

st.set_page_config(
    page_title="Smart Calendar",
    page_icon="📅",
    layout="centered"
)

STATI = ["Ufficio", "Smart", "Assenza"]
PERSONE = ["Margherita", "Roberto"]

NOME_FILE_GOOGLE_SHEETS = "Smart Calendar"

FOGLIO_PRESENZE = "presenze"
FOGLIO_FESTE = "festivita_manuali"
FOGLIO_RIEPILOGO = "riepilogo"

festivita_italiane = holidays.Italy()


@st.cache_resource
def connetti_google_sheets():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scope
    )

    client = gspread.authorize(credentials)

    try:
        spreadsheet = client.open(NOME_FILE_GOOGLE_SHEETS)
    except gspread.SpreadsheetNotFound:
        spreadsheet = client.create(NOME_FILE_GOOGLE_SHEETS)

    return spreadsheet


@st.cache_resource
def inizializza_fogli():
    spreadsheet = connetti_google_sheets()

    fogli_esistenti = {
        ws.title: ws
        for ws in spreadsheet.worksheets()
    }

    def crea_o_prendi(nome_foglio, intestazioni):
        if nome_foglio in fogli_esistenti:
            worksheet = fogli_esistenti[nome_foglio]
        else:
            worksheet = spreadsheet.add_worksheet(
                title=nome_foglio,
                rows=1000,
                cols=len(intestazioni)
            )
            worksheet.update("A1", [intestazioni])

        valori = worksheet.get_all_values()

        if len(valori) == 0:
            worksheet.update("A1", [intestazioni])

        return worksheet

    ws_presenze = crea_o_prendi(
        FOGLIO_PRESENZE,
        ["data", "persona", "stato"]
    )

    ws_feste = crea_o_prendi(
        FOGLIO_FESTE,
        ["data", "descrizione"]
    )

    ws_riepilogo = crea_o_prendi(
        FOGLIO_RIEPILOGO,
        [
            "persona",
            "trimestre",
            "Ufficio",
            "Smart",
            "Giorni lavorati",
            "% Ufficio",
            "% Smart",
            "Esito"
        ]
    )

    try:
        fogli_aggiornati = spreadsheet.worksheets()
        if len(fogli_aggiornati) > 1:
            for ws in fogli_aggiornati:
                if ws.title == "Foglio1":
                    spreadsheet.del_worksheet(ws)
                    break
    except Exception:
        pass

    return spreadsheet, ws_presenze, ws_feste, ws_riepilogo


spreadsheet, ws_presenze, ws_feste, ws_riepilogo = inizializza_fogli()

# ===== EXPORT EXCEL FORMATTATO =====
def genera_excel_formattato(df, df_feste, riepilogo, anno):
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    import calendar

    buffer = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = str(anno)
    ws_riep = wb.create_sheet("Riepilogo")

    codici = {
        "Ufficio": "PRE",
        "Smart": "LAW",
        "Assenza": "ASS"
    }

    persone = PERSONE
    mesi = [
        "GENNAIO", "FEBBRAIO", "MARZO", "APRILE", "MAGGIO", "GIUGNO",
        "LUGLIO", "AGOSTO", "SETTEMBRE", "OTTOBRE", "NOVEMBRE", "DICEMBRE"
    ]
    giorni_it = ["Lu", "Ma", "Me", "Gi", "Ve", "Sa", "Do"]

    fill_header = PatternFill("solid", fgColor="D9EAF7")
    fill_month = PatternFill("solid", fgColor="1F4E78")
    fill_legend = PatternFill("solid", fgColor="E2F0D9")
    fill_weekend = PatternFill("solid", fgColor="E7E6E6")
    fill_ass = PatternFill("solid", fgColor="F4CCCC")
    fill_pre = PatternFill("solid", fgColor="D9EAD3")
    fill_law = PatternFill("solid", fgColor="CFE2F3")
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Dizionario rapido: (data, persona) -> codice
    presenze = df.copy()
    if not presenze.empty:
        presenze["data"] = pd.to_datetime(presenze["data"]).dt.date
        presenze["codice"] = presenze["stato"].map(codici).fillna(presenze["stato"])
        lookup = {
            (r["data"], r["persona"]): r["codice"]
            for _, r in presenze.iterrows()
        }
    else:
        lookup = {}

    # Legenda
    ws["A1"] = "LEGENDA"
    ws["A1"].font = Font(bold=True)
    legenda = [("ASS", "Assenza"), ("PRE", "Lavoro in presenza"), ("LAW", "Smart working")]
    for i, (codice, descrizione) in enumerate(legenda, start=2):
        ws[f"A{i}"] = codice
        ws[f"B{i}"] = descrizione
        ws[f"A{i}"].font = Font(bold=True)
        ws[f"A{i}"].fill = fill_legend
        ws[f"A{i}"].border = border
        ws[f"B{i}"].border = border

    start_cols = [1, 6, 11, 16]  # 4 mesi per riga
    start_rows = [7, 44, 81]

    for month in range(1, 13):
        block_col = start_cols[(month - 1) % 4]
        block_row = start_rows[(month - 1) // 4]

        ws.merge_cells(start_row=block_row, start_column=block_col, end_row=block_row, end_column=block_col + 3)
        c = ws.cell(block_row, block_col, mesi[month - 1])
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill_month
        c.alignment = Alignment(horizontal="center")

        headers = ["Giorno", "Sett."] + persone
        for j, h in enumerate(headers):
            cell = ws.cell(block_row + 1, block_col + j, h)
            cell.font = Font(bold=True)
            cell.fill = fill_header
            cell.border = border
            cell.alignment = Alignment(horizontal="center")

        days = calendar.monthrange(anno, month)[1]
        for day in range(1, days + 1):
            row = block_row + 1 + day
            data = pd.Timestamp(year=anno, month=month, day=day).date()
            weekday = data.weekday()
            values = [day, giorni_it[weekday]]
            for persona in persone:
                values.append(lookup.get((data, persona), ""))

            for j, value in enumerate(values):
                cell = ws.cell(row, block_col + j, value)
                cell.border = border
                cell.alignment = Alignment(horizontal="center")
                if weekday >= 5:
                    cell.fill = fill_weekend
                elif value == "ASS":
                    cell.fill = fill_ass
                elif value == "PRE":
                    cell.fill = fill_pre
                elif value == "LAW":
                    cell.fill = fill_law

    # Validazione celle editabili nel calendario esportato
    dv = DataValidation(type="list", formula1='"ASS,PRE,LAW"', allow_blank=True)
    ws.add_data_validation(dv)
    for row in start_rows:
        for col in start_cols:
            for p_idx in range(len(persone)):
                dv.add(f"{get_column_letter(col + 2 + p_idx)}{row + 2}:{get_column_letter(col + 2 + p_idx)}{row + 32}")

    for col in range(1, 21):
        ws.column_dimensions[get_column_letter(col)].width = 13

    # Riepilogo leggibile
    ws_riep["A1"] = "RIEPILOGO PRESENZE"
    ws_riep["A1"].font = Font(bold=True, size=14)

    headers = ["Persona", "Periodo", "PRE", "LAW", "ASS", "Totale", "% PRE", "% LAW", "% ASS"]
    for j, h in enumerate(headers, start=1):
        cell = ws_riep.cell(3, j, h)
        cell.font = Font(bold=True)
        cell.fill = fill_header
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    df_r = df.copy()
    if not df_r.empty:
        df_r["data"] = pd.to_datetime(df_r["data"])
        df_r["periodo"] = df_r["data"].dt.to_period("Q").astype(str)
        df_r["codice"] = df_r["stato"].map(codici).fillna(df_r["stato"])
        pivot = (
            df_r.groupby(["persona", "periodo", "codice"])
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )
        for col in ["PRE", "LAW", "ASS"]:
            if col not in pivot.columns:
                pivot[col] = 0
        pivot["Totale"] = pivot[["PRE", "LAW", "ASS"]].sum(axis=1)
        for col in ["PRE", "LAW", "ASS"]:
            pivot[f"% {col}"] = (pivot[col] / pivot["Totale"]).fillna(0)

        pivot = pivot[["persona", "periodo", "PRE", "LAW", "ASS", "Totale", "% PRE", "% LAW", "% ASS"]]
        for i, (_, r) in enumerate(pivot.iterrows(), start=4):
            for j, value in enumerate(r.tolist(), start=1):
                cell = ws_riep.cell(i, j, value)
                cell.border = border
                cell.alignment = Alignment(horizontal="center")
                if j >= 7:
                    cell.number_format = "0.00%"

    for col in range(1, 10):
        ws_riep.column_dimensions[get_column_letter(col)].width = 15

    wb.save(buffer)
    return buffer.getvalue()



@st.cache_data(ttl=30)
def leggi_presenze_cache():
    dati = ws_presenze.get_all_records()

    if not dati:
        return pd.DataFrame(columns=["data", "persona", "stato"])

    df = pd.DataFrame(dati)
    df["data"] = pd.to_datetime(df["data"], errors="coerce")
    df = df.dropna(subset=["data"])

    return df


@st.cache_data(ttl=30)
def leggi_feste_cache():
    dati = ws_feste.get_all_records()

    if not dati:
        return pd.DataFrame(columns=["data", "descrizione"])

    df = pd.DataFrame(dati)
    df["data"] = pd.to_datetime(df["data"], errors="coerce")
    df = df.dropna(subset=["data"])

    return df


def leggi_presenze():
    return leggi_presenze_cache()


def leggi_feste_manuali():
    return leggi_feste_cache()


def svuota_cache_dati():
    leggi_presenze_cache.clear()
    leggi_feste_cache.clear()


def salva_presenze(df):
    df = df.copy()
    df["data"] = pd.to_datetime(df["data"]).dt.strftime("%Y-%m-%d")

    ws_presenze.clear()
    ws_presenze.update("A1", [["data", "persona", "stato"]])

    if not df.empty:
        ws_presenze.append_rows(df.values.tolist())

    svuota_cache_dati()


def salva_feste_manuali(df):
    df = df.copy()
    df["data"] = pd.to_datetime(df["data"]).dt.strftime("%Y-%m-%d")

    ws_feste.clear()
    ws_feste.update("A1", [["data", "descrizione"]])

    if not df.empty:
        ws_feste.append_rows(df.values.tolist())

    svuota_cache_dati()


def salva_riepilogo(riepilogo):
    intestazioni = [
        "persona",
        "trimestre",
        "Ufficio",
        "Smart",
        "Giorni lavorati",
        "% Ufficio",
        "% Smart",
        "Esito"
    ]

    ws_riepilogo.clear()
    ws_riepilogo.update("A1", [intestazioni])

    if not riepilogo.empty:
        ws_riepilogo.append_rows(riepilogo[intestazioni].values.tolist())


def calcola_riepilogo(df, df_feste):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df = df.copy()
    df["data"] = pd.to_datetime(df["data"], errors="coerce")
    df = df.dropna(subset=["data"])

    if df_feste.empty:
        date_feste_manuali = set()
    else:
        df_feste = df_feste.copy()
        df_feste["data"] = pd.to_datetime(df_feste["data"], errors="coerce")
        df_feste = df_feste.dropna(subset=["data"])
        date_feste_manuali = set(df_feste["data"].dt.date)

    df["weekend"] = df["data"].dt.weekday >= 5

    df["festivo_italia"] = df["data"].dt.date.apply(
        lambda x: x in festivita_italiane
    )

    df["festivo_manuale"] = df["data"].dt.date.apply(
        lambda x: x in date_feste_manuali
    )

    df_valido = df[
        (~df["weekend"]) &
        (~df["festivo_italia"]) &
        (~df["festivo_manuale"])
    ].copy()

    lavorati = df_valido[
        df_valido["stato"].isin(["Ufficio", "Smart"])
    ].copy()

    if lavorati.empty:
        return pd.DataFrame(), df_valido

    lavorati["trimestre"] = lavorati["data"].dt.to_period("Q").astype(str)

    riepilogo = (
        lavorati
        .groupby(["persona", "trimestre", "stato"])
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )

    for col in ["Ufficio", "Smart"]:
        if col not in riepilogo.columns:
            riepilogo[col] = 0

    riepilogo["Giorni lavorati"] = (
        riepilogo["Ufficio"] + riepilogo["Smart"]
    )

    riepilogo["% Ufficio"] = (
        riepilogo["Ufficio"] / riepilogo["Giorni lavorati"] * 100
    ).round(2)

    riepilogo["% Smart"] = (
        riepilogo["Smart"] / riepilogo["Giorni lavorati"] * 100
    ).round(2)

    riepilogo["Esito"] = riepilogo.apply(
        lambda r: "OK" if 40 <= r["% Ufficio"] <= 60 else "SFORO",
        axis=1
    )

    ordine_colonne = [
        "persona",
        "trimestre",
        "Ufficio",
        "Smart",
        "Giorni lavorati",
        "% Ufficio",
        "% Smart",
        "Esito"
    ]

    return riepilogo[ordine_colonne], df_valido


df = leggi_presenze()
df_feste = leggi_feste_manuali()
riepilogo, df_valido = calcola_riepilogo(df, df_feste)


st.title("📅 Smart Calendar")

tab1, tab2, tab3 = st.tabs([
    "Inserisci",
    "Festività",
    "Riepilogo"
])


with tab1:
    st.subheader("Nuova giornata")

    giorno = st.date_input("Data", value=date.today())
    persona = st.selectbox("Persona", PERSONE)
    stato = st.selectbox("Stato", STATI)

    if st.button("Salva presenza", use_container_width=True):
        nuova_riga = pd.DataFrame([{
            "data": pd.to_datetime(giorno),
            "persona": persona,
            "stato": stato
        }])

        df = df[
            ~(
                (df["data"] == pd.to_datetime(giorno)) &
                (df["persona"] == persona)
            )
        ]

        df = pd.concat([df, nuova_riga], ignore_index=True)
        df = df.sort_values(["data", "persona"])

        salva_presenze(df)

        riepilogo, df_valido = calcola_riepilogo(df, df_feste)
        salva_riepilogo(riepilogo)

        st.success("Presenza salvata")
        st.rerun()

    st.divider()

    st.subheader("Ultime presenze")

    if df.empty:
        st.info("Nessuna presenza inserita")
    else:
        ultime = df.sort_values("data", ascending=False).head(20)
        st.dataframe(ultime, use_container_width=True, hide_index=True)


with tab2:
    st.subheader("Aggiungi festività manuale")

    giorno_festa = st.date_input(
        "Data festività",
        value=date.today(),
        key="giorno_festa"
    )

    descrizione_festa = st.text_input(
        "Descrizione",
        placeholder="Es. Santo Patrono"
    )

    if st.button("Aggiungi festività", use_container_width=True):
        if descrizione_festa.strip() == "":
            st.warning("Inserisci una descrizione")
        else:
            nuova_festa = pd.DataFrame([{
                "data": pd.to_datetime(giorno_festa),
                "descrizione": descrizione_festa.strip()
            }])

            df_feste = df_feste[
                df_feste["data"] != pd.to_datetime(giorno_festa)
            ]

            df_feste = pd.concat([df_feste, nuova_festa], ignore_index=True)
            df_feste = df_feste.sort_values("data")

            salva_feste_manuali(df_feste)

            riepilogo, df_valido = calcola_riepilogo(df, df_feste)
            salva_riepilogo(riepilogo)

            st.success("Festività aggiunta")
            st.rerun()

    st.divider()

    st.subheader("Festività manuali")

    if df_feste.empty:
        st.info("Nessuna festività manuale inserita")
    else:
        st.dataframe(df_feste, use_container_width=True, hide_index=True)

        opzioni_feste = (
            df_feste["data"].dt.strftime("%Y-%m-%d")
            + " - "
            + df_feste["descrizione"]
        ).tolist()

        feste_da_eliminare = st.multiselect(
            "Festività da eliminare",
            options=opzioni_feste
        )

        if st.button("Elimina selezionate", use_container_width=True):
            date_da_eliminare = [
                x.split(" - ")[0] for x in feste_da_eliminare
            ]

            df_feste = df_feste[
                ~df_feste["data"].dt.strftime("%Y-%m-%d").isin(date_da_eliminare)
            ]

            salva_feste_manuali(df_feste)

            riepilogo, df_valido = calcola_riepilogo(df, df_feste)
            salva_riepilogo(riepilogo)

            st.success("Festività eliminate")
            st.rerun()


with tab3:
    st.subheader("Riepilogo trimestrale")

    if riepilogo.empty:
        st.info("Nessun giorno lavorato valido")
    else:
        st.dataframe(riepilogo, use_container_width=True, hide_index=True)

    st.divider()

    # Unico export disponibile: Excel formattato
    excel_formattato = genera_excel_formattato(df, df_feste, riepilogo, anno=2026)

    st.download_button(
        "Scarica Excel",
        data=excel_formattato,
        file_name="smart_calendar_2026_formattato.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )