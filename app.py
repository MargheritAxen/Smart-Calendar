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

    csv = df.to_csv(index=False).encode("utf-8")

    st.download_button(
        "Scarica CSV",
        data=csv,
        file_name="presenze_smart.csv",
        mime="text/csv",
        use_container_width=True
    )

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Presenze", index=False)
        df_feste.to_excel(writer, sheet_name="Festivita manuali", index=False)
        df_valido.to_excel(writer, sheet_name="Giorni validi", index=False)
        riepilogo.to_excel(writer, sheet_name="Riepilogo", index=False)

    st.download_button(
        "Scarica Excel",
        data=buffer.getvalue(),
        file_name="smart_calendar.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )