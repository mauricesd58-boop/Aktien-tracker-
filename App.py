import io
import pandas as pd
import streamlit as st
import yfinance as yf
from datetime import datetime
from streamlit_calendar import calendar
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

st.set_page_config(page_title="Aktien & Dividenden Tracker", layout="wide")

# --- 1. PASSWORT-SCHUTZ ---
def check_password():
    """Gibt True zurück, wenn der Nutzer das richtige Passwort eingegeben hat."""
    def password_entered():
        # Liest das Passwort aus den Secrets
        correct_password = st.secrets.get("passwords", {}).get("app_password", "1234")
        if st.session_state["password_input"] == correct_password:
            st.session_state["password_correct"] = True
            del st.session_state["password_input"]  # Passwort nicht im Session-State speichern
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # Erster Aufruf: Eingabefeld anzeigen
        st.title("🔒 Bitte anmelden")
        st.text_input(
            "Passwort eingeben:",
            type="password",
            on_change=password_entered,
            key="password_input"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Falsche Eingabe
        st.title("🔒 Bitte anmelden")
        st.text_input(
            "Passwort eingeben:",
            type="password",
            on_change=password_entered,
            key="password_input"
        )
        st.error("😕 Falsches Passwort. Bitte erneut versuchen.")
        return False
    else:
        # Passwort korrekt
        return True

# Überprüfung ausführen (bricht das Skript ab, wenn nicht angemeldet)
if not check_password():
    st.stop()

# --- AB HIER WIRD DIE EIGENTLICHE APP GELADEN ---
st.title("📈 Mein Aktien-Dashboard & Dividenden-Kalender")

FILE_NAME = "Aktieninvestments.xlsx"

# 2. GOOGLE DRIVE ANBINDUNG
@st.cache_resource
def get_drive_service():
    creds_dict = dict(st.secrets["google_credentials"])
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds)

def load_all_sheets_from_drive():
    service = get_drive_service()
    results = service.files().list(q=f"name='{FILE_NAME}' and trashed=false").execute()
    items = results.get('files', [])
    if not items:
        st.error(f"Datei '{FILE_NAME}' wurde nicht in Google Drive gefunden.")
        st.stop()
    
    file_id = items[0]['id']
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    
    fh.seek(0)
    all_sheets = pd.read_excel(fh, sheet_name=None)
    return all_sheets, file_id

def save_sheets_to_drive(all_sheets, file_id):
    service = get_drive_service()
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
    output.seek(0)
    media = MediaIoBaseUpload(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True
    )
    service.files().update(fileId=file_id, media_body=media).execute()

# Daten aus Drive laden
try:
    all_sheets, drive_file_id = load_all_sheets_from_drive()
except Exception as e:
    st.error("Fehler beim Laden der Datei aus Google Drive. Bitte erstelle den Service-Account in den Secrets.")
    st.stop()

# Ticker Mapping
TICKER_MAP = {
    "Capital southwest": "CSWC",
    "Capital southwest ": "CSWC",
    "Trinity Capital": "TRIN",
    "Trinity Capital ": "TRIN",
    "realty income": "O",
    "Basf": "BAS.DE",
    "eon": "EOAN.DE",
    "EOAN": "EOAN.DE",
    "Vodafone": "VOD",
    "Vodafone Group": "VOD",
    "ernst russ": "ERAG.DE",
    "Ernst Russ AG": "ERAG.DE",
    "ernst russ ag": "ERAG.DE",
    "lenovo": "LNVGY",
    "Lenovo": "LNVGY",
    "banco de Sabadell": "SAB.MC",
    "Banco de Sabadell SA": "SAB.MC",
    "meren energ": "MMENF",
    "Meren Energy": "MMENF"
}

# 3. NEUES INVESTMENT HINZUFÜGEN (SIDEBAR)
st.sidebar.header("➕ Neues Investment hinzufügen")
with st.sidebar.form("add_investment_form"):
    date_val = st.date_input("Kaufdatum", datetime.now())
    stock_name = st.text_input("Aktienname (z. B. Capital southwest)")
    shares_val = st.number_input("Stückzahl (Anteile)", min_value=0.0, step=0.01, format="%.4f")
    total_cost_val = st.number_input("Gesamtkosten (€)", min_value=0.0, step=0.01)
    
    submitted = st.form_submit_button("Speichern & auf Google Drive hochladen")
    
    if submitted and stock_name and shares_val > 0:
        year_str = str(date_val.year)
        cost_per_share = round(total_cost_val / shares_val, 2) if shares_val > 0 else 0.0
        
        new_row = {
            "Datum": date_val.strftime("%d.%m.%Y"),
            "Aktie": stock_name,
            "Anteile": shares_val,
            "Kosten/Anteil €": cost_per_share,
            "Gesamtkosten €": total_cost_val
        }
        
        if year_str in all_sheets:
            all_sheets[year_str] = pd.concat([all_sheets[year_str], pd.DataFrame([new_row])], ignore_index=True)
        else:
            all_sheets[year_str] = pd.DataFrame([new_row])
            
        save_sheets_to_drive(all_sheets, drive_file_id)
        st.sidebar.success(f"✅ {shares_val}x {stock_name} im Tab '{year_str}' gespeichert!")
        st.cache_data.clear()
        st.rerun()

# 4. AGGREGATION ALLER KÄUFE
portfolio_summary = {}

for sheet_name, df in all_sheets.items():
    for _, row in df.iterrows():
        raw_name = row.get("Aktie")
        shares = row.get("Anteile")
        cost = row.get("Gesamtkosten €") or row.get("Gesamtkosten")
        
        if pd.isna(raw_name) or pd.isna(shares) or float(shares) <= 0:
            continue
            
        name_clean = str(raw_name).strip()
        shares_num = float(shares)
        cost_num = float(cost) if pd.notna(cost) else 0.0
        
        if name_clean not in portfolio_summary:
            portfolio_summary[name_clean] = {"shares": 0.0, "cost": 0.0}
            
        portfolio_summary[name_clean]["shares"] += shares_num
        portfolio_summary[name_clean]["cost"] += cost_num

# 5. LIVE-KURSE & KALENDER
portfolio_rows = []
calendar_events = []

with st.spinner("Lade Live-Kurse und Dividendentermine..."):
    for name, data in portfolio_summary.items():
        shares = data["shares"]
        cost = data["cost"]
        ticker_symbol = TICKER_MAP.get(name)
        current_price = 0.0
        
        if ticker_symbol:
            try:
                t = yf.Ticker(ticker_symbol)
                current_price = t.fast_info.get("lastPrice", 0.0)
                cal = t.calendar
                
                if cal is not None and "Dividend Date" in cal:
                    div_date = str(cal["Dividend Date"])
                    div_rate = t.info.get("dividendRate", 0.0) or 0.0
                    est_payout = (div_rate / 4) * shares if div_rate else 0.0
                    
                    calendar_events.append({
                        "title": f"💰 {name}: ~{est_payout:.2f} €",
                        "start": div_date,
                        "end": div_date,
                        "color": "#2E7D32",
                        "allDay": True
                    })
            except Exception:
                pass
                
        current_val = shares * current_price
        pnl = current_val - cost
        pnl_pct = (pnl / cost * 100) if cost > 0 else 0.0
        
        portfolio_rows.append({
            "Aktie": name,
            "Anteile Gesamt": round(shares, 4),
            "Kaufwert Gesamt (€)": round(cost, 2),
            "Aktueller Kurs (€)": round(current_price, 2),
            "Gesamtwert (€)": round(current_val, 2),
            "Plus / Minus (€)": round(pnl, 2),
            "Performance": f"{pnl_pct:+.2f} %"
        })

df_display = pd.DataFrame(portfolio_rows)

# 6. DASHBOARD ANZEIGE
total_inv = df_display["Kaufwert Gesamt (€)"].sum()
total_val = df_display["Gesamtwert (€)"].sum()
total_pnl = total_val - total_inv
total_pnl_pct = (total_pnl / total_inv * 100) if total_inv > 0 else 0.0

col1, col2, col3 = st.columns(3)
col1.metric("Gesamt Investiert", f"{total_inv:,.2f} €")
col2.metric("Aktueller Depotwert", f"{total_val:,.2f} €")
col3.metric("Gesamtertrag (Plus/Minus)", f"{total_pnl:,.2f} €", delta=f"{total_pnl_pct:+.2f} %")

st.divider()

tab_cal, tab_table = st.tabs(["📅 Dividenden-Kalender", "📊 Depot Positionen"])

with tab_cal:
    st.subheader("Anstehende Auszahlungen im Kalender")
    calendar_options = {
        "headerToolbar": {
            "left": "prev,next today",
            "center": "title",
            "right": "dayGridMonth,listMonth",
        },
        "initialView": "dayGridMonth",
        "locale": "de",
    }
    calendar(events=calendar_events, options=calendar_options, key="div_cal")

with tab_table:
    st.subheader("Zusammengefasste Bestände aus allen Jahren")
    st.dataframe(df_display, use_container_width=True)
