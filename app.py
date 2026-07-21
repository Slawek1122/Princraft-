import streamlit as st
import pandas as pd
from datetime import datetime
import os
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

COLUMNS = [
    "Date",
    "Princraft warehouse location",
    "Job number",
    "Order number",
    "Design number",
    "quantity card",
    "quantity packs",
    "total boxes"
]

DATA_FILE = "data.xlsx"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587


def load_data():
    if os.path.exists(DATA_FILE):
        try:
            return pd.read_excel(DATA_FILE)
        except Exception:
            return pd.DataFrame(columns=COLUMNS)
    return pd.DataFrame(columns=COLUMNS)


def save_data(df):
    df.to_excel(DATA_FILE, index=False)


def get_used_locations(df):
    if df.empty or "Princraft warehouse location" not in df.columns:
        return set()
    return set(df["Princraft warehouse location"].dropna().astype(str).str.strip())


def get_next_free_location(available, used):
    for loc in available:
        if loc not in used:
            return loc
    return ""


def send_email(from_email: str, password: str, to_email: str, df: pd.DataFrame) -> bool:
    try:
        msg = MIMEMultipart()
        msg["From"] = from_email
        msg["To"] = to_email
        msg["Subject"] = f"Magazyn Princraft – dane z {datetime.now().strftime('%Y-%m-%d %H:%M')}"

        body = "W załączniku plik Excel z danymi magazynowymi."
        msg.attach(MIMEText(body, "plain"))

        buffer = BytesIO()
        df.to_excel(buffer, index=False)
        buffer.seek(0)

        part = MIMEBase("application", "octet-stream")
        part.set_payload(buffer.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment; filename=data.xlsx")
        msg.attach(part)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(from_email, password)
            server.send_message(msg)

        return True
    except Exception as e:
        st.error(f"Błąd wysyłania: {e}")
        return False


# ================== UI ==================
st.set_page_config(
    page_title="Magazyn Princraft",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="expanded"
)

# CSS mocno pod tablet / telefon
st.markdown("""
<style>
    /* ogólne */
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
        max-width: 900px;
    }
    
    h1 {
        font-size: 2rem !important;
        text-align: center;
        margin-bottom: 0.3rem !important;
    }
    
    /* przyciski */
    .stButton > button {
        width: 100%;
        height: 3.8rem;
        font-size: 1.35rem !important;
        font-weight: 700;
        border-radius: 12px;
    }
    
    /* pola tekstowe i liczby */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input {
        font-size: 1.25rem !important;
        padding: 0.85rem 0.9rem !important;
        border-radius: 10px !important;
    }
    
    /* text area */
    .stTextArea textarea {
        font-size: 1.15rem !important;
        line-height: 1.5 !important;
    }
    
    /* formularz */
    div[data-testid="stForm"] {
        border: 2px solid #d0d0d0;
        border-radius: 16px;
        padding: 1.5rem 1.2rem;
        background: #fafafa;
    }
    
    /* info i warning */
    .stAlert {
        font-size: 1.15rem !important;
        padding: 1rem 1.1rem !important;
        border-radius: 12px;
    }
    
    /* dataframe */
    .stDataFrame {
        font-size: 1.05rem;
    }
    
    /* sidebar */
    section[data-testid="stSidebar"] {
        min-width: 280px;
    }
    
    /* camera */
    div[data-testid="stCameraInput"] {
        border-radius: 12px;
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)

st.title("📦 Magazyn Princraft")
st.caption("Szybkie wpisywanie palet – bez dublowania roboty")

# ----- LOKALIZACJE (sidebar) -----
with st.sidebar:
    st.header("Lokalizacje")
    st.caption("Jedna lokalizacja = jeden wiersz. Aplikacja proponuje następną wolną.")

    default_locations = st.session_state.get("locations_text", "PB1G18H\nPB1G19H\nPB1G20H\nPB1G21H\nPB1G22H")
    locations_text = st.text_area(
        "Dostępne lokalizacje",
        value=default_locations,
        height=220,
        help="Jedna lokalizacja w wierszu"
    )

    if st.button("💾 Zapisz listę lokalizacji", use_container_width=True):
        st.session_state.locations_text = locations_text
        st.success("Zapisano")

    available_locations = [
        line.strip() for line in locations_text.strip().splitlines() if line.strip()
    ]
    st.session_state.available_locations = available_locations

    st.divider()
    st.caption(f"Lokalizacji na liście: {len(available_locations)}")

# ----- OBLICZENIE WOLNEJ LOKALIZACJI -----
df = load_data()
used = get_used_locations(df)
next_free = get_next_free_location(
    st.session_state.get("available_locations", available_locations),
    used
)

# ----- FORMULARZ -----
with st.form("form_paleta", clear_on_submit=False):
    st.subheader("Nowa paleta")

    if next_free:
        st.info(f"Proponowana wolna lokalizacja: **{next_free}**")
    else:
        st.warning("Brak wolnych lokalizacji. Wpisz ręcznie lub dodaj więcej w panelu bocznym.")

    location = st.text_input(
        "Lokalizacja magazynu",
        value=next_free,
        placeholder="np. PB1G18H"
    )
    job = st.text_input("Job number")
    order = st.text_input("Order number")
    design = st.text_input("Design number")

    col1, col2, col3 = st.columns(3)
    with col1:
        qty_card = st.number_input("quantity card", min_value=0, step=1, value=0)
    with col2:
        qty_packs = st.number_input("quantity packs", min_value=0, step=1, value=0)
    with col3:
        total_boxes = st.number_input("total boxes", min_value=0, step=1, value=0)

    submitted = st.form_submit_button("➕ DODAJ PALETĘ")

if submitted:
    if not location.strip() or not job.strip() or not order.strip() or not design.strip():
        st.error("Wypełnij wszystkie pola tekstowe.")
    else:
        df = load_data()
        new_row = {
            "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Princraft warehouse location": location.strip(),
            "Job number": job.strip(),
            "Order number": order.strip(),
            "Design number": design.strip(),
            "quantity card": int(qty_card),
            "quantity packs": int(qty_packs),
            "total boxes": int(total_boxes)
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_data(df)
        st.success(f"✅ Dodano paletę na lokalizacji **{location.strip()}**")
        st.balloons()
        st.rerun()

# ----- APARAT (zdjęcie ticketa) -----
st.divider()
st.subheader("📷 Zdjęcie ticketa (opcjonalnie)")
st.caption("Zrób zdjęcie paletowego ticketa. Na razie tylko podgląd – OCR można dorobić później.")

photo = st.camera_input("Zrób zdjęcie ticketa", label_visibility="collapsed")

if photo is not None:
    st.image(photo, caption="Zdjęcie ticketa", use_container_width=True)
    st.info("Zdjęcie zapisane w sesji. Możesz teraz ręcznie przepisać dane z ticketa do formularza powyżej.")

# ----- PODGLĄD -----
st.divider()
st.subheader("Ostatnie wpisy")

df = load_data()

if df.empty:
    st.info("Brak danych. Dodaj pierwszą paletę.")
else:
    st.dataframe(df.tail(10), use_container_width=True, hide_index=True)

    show_all = st.checkbox("Pokaż całą tabelę")
    if show_all:
        st.dataframe(df, use_container_width=True, hide_index=True)

    st.caption(f"Łącznie wpisów: {len(df)}")

    # ----- POBIERZ -----
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    st.download_button(
        label="📥 POBIERZ EXCEL",
        data=buffer.getvalue(),
        file_name=f"magazyn_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # ----- WYŚLIJ MAILEM -----
    st.divider()
    st.subheader("Wyślij na maila")

    with st.expander("Ustawienia wysyłki (Gmail)"):
        from_email = st.text_input("Twój email (nadawca)", placeholder="twoj@gmail.com")
        from_password = st.text_input("Hasło aplikacji Gmail", type="password", help="Nie zwykłe hasło – tylko hasło aplikacji z konta Google")

    email_to = st.text_input("Adres email odbiorcy", placeholder="np. magazin@firma.pl")

    if st.button("📧 WYŚLIJ EXCEL", use_container_width=True):
        if not from_email or not from_password:
            st.warning("Podaj swój email i hasło aplikacji w ustawieniach powyżej.")
        elif not email_to or "@" not in email_to:
            st.warning("Podaj poprawny adres email odbiorcy.")
        else:
            with st.spinner("Wysyłanie..."):
                if send_email(from_email.strip(), from_password, email_to.strip(), df):
                    st.success(f"Wysłano na {email_to}")

st.divider()
st.caption("Aplikacja magazynowa • dane zapisane lokalnie w data.xlsx")
