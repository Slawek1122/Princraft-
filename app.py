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
        msg["Subject"] = f"Princraft Warehouse – data from {datetime.now().strftime('%Y-%m-%d %H:%M')}"

        body = "Please find the warehouse data Excel file attached."
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
        st.error(f"Send error: {e}")
        return False


# ================== UI ==================
st.set_page_config(
    page_title="Princraft Warehouse",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Dark / professional theme CSS optimized for tablet
st.markdown("""
<style>
    /* background */
    .stApp {
        background-color: #0f172a;
        color: #e2e8f0;
    }
    
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
        max-width: 900px;
    }
    
    h1, h2, h3, .stMarkdown, .stCaption, label {
        color: #f1f5f9 !important;
    }
    
    h1 {
        font-size: 2rem !important;
        text-align: center;
        margin-bottom: 0.3rem !important;
    }
    
    /* buttons */
    .stButton > button {
        width: 100%;
        height: 3.8rem;
        font-size: 1.35rem !important;
        font-weight: 700;
        border-radius: 12px;
        background-color: #3b82f6 !important;
        color: white !important;
        border: none !important;
    }
    .stButton > button:hover {
        background-color: #2563eb !important;
    }
    
    /* inputs */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input {
        font-size: 1.25rem !important;
        padding: 0.85rem 0.9rem !important;
        border-radius: 10px !important;
        background-color: #1e293b !important;
        color: #f1f5f9 !important;
        border: 1px solid #334155 !important;
    }
    
    .stTextArea textarea {
        font-size: 1.15rem !important;
        line-height: 1.5 !important;
        background-color: #1e293b !important;
        color: #f1f5f9 !important;
        border: 1px solid #334155 !important;
    }
    
    /* form */
    div[data-testid="stForm"] {
        border: 2px solid #334155;
        border-radius: 16px;
        padding: 1.5rem 1.2rem;
        background: #1e293b;
    }
    
    /* alerts */
    .stAlert {
        font-size: 1.15rem !important;
        padding: 1rem 1.1rem !important;
        border-radius: 12px;
    }
    
    /* sidebar */
    section[data-testid="stSidebar"] {
        background-color: #0f172a;
        min-width: 280px;
    }
    
    section[data-testid="stSidebar"] * {
        color: #e2e8f0 !important;
    }
    
    /* dataframe */
    .stDataFrame {
        font-size: 1.05rem;
    }
    
    /* download button */
    .stDownloadButton > button {
        background-color: #10b981 !important;
    }
    .stDownloadButton > button:hover {
        background-color: #059669 !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📦 Princraft Warehouse")
st.caption("Fast pallet entry – no double work")

# ----- LOCATIONS (sidebar) -----
with st.sidebar:
    st.header("Locations")
    st.caption("One location per line. App suggests the next free one.")

    default_locations = st.session_state.get("locations_text", "PB1G18H\nPB1G19H\nPB1G20H\nPB1G21H\nPB1G22H")
    locations_text = st.text_area(
        "Available locations",
        value=default_locations,
        height=220,
        help="One location per line"
    )

    if st.button("💾 Save location list", use_container_width=True):
        st.session_state.locations_text = locations_text
        st.success("Saved")

    available_locations = [
        line.strip() for line in locations_text.strip().splitlines() if line.strip()
    ]
    st.session_state.available_locations = available_locations

    st.divider()
    st.caption(f"Locations on list: {len(available_locations)}")

# ----- NEXT FREE LOCATION -----
df = load_data()
used = get_used_locations(df)
next_free = get_next_free_location(
    st.session_state.get("available_locations", available_locations),
    used
)

# ----- FORM -----
with st.form("form_pallet", clear_on_submit=False):
    st.subheader("New pallet")

    if next_free:
        st.info(f"Suggested free location: **{next_free}**")
    else:
        st.warning("No free locations left. Enter manually or add more in the sidebar.")

    location = st.text_input(
        "Warehouse location",
        value=next_free,
        placeholder="e.g. PB1G18H"
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

    submitted = st.form_submit_button("➕ ADD PALLET")

if submitted:
    if not location.strip() or not job.strip() or not order.strip() or not design.strip():
        st.error("Please fill in all text fields.")
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
        st.success(f"✅ Pallet added at location **{location.strip()}**")
        st.balloons()
        st.rerun()

# ----- PREVIEW -----
st.divider()
st.subheader("Recent entries")

df = load_data()

if df.empty:
    st.info("No data yet. Add the first pallet.")
else:
    st.dataframe(df.tail(10), use_container_width=True, hide_index=True)

    show_all = st.checkbox("Show full table")
    if show_all:
        st.dataframe(df, use_container_width=True, hide_index=True)

    st.caption(f"Total entries: {len(df)}")

    # ----- DOWNLOAD -----
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    st.download_button(
        label="📥 DOWNLOAD EXCEL",
        data=buffer.getvalue(),
        file_name=f"warehouse_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # ----- EMAIL -----
    st.divider()
    st.subheader("Send by email")

    with st.expander("Email settings (Gmail)"):
        from_email = st.text_input("Your email (sender)", placeholder="you@gmail.com")
        from_password = st.text_input("Gmail App Password", type="password", help="Not your regular password – use an App Password from Google")

    email_to = st.text_input("Recipient email", placeholder="e.g. warehouse@company.com")

    if st.button("📧 SEND EXCEL", use_container_width=True):
        if not from_email or not from_password:
            st.warning("Enter your email and app password in the settings above.")
        elif not email_to or "@" not in email_to:
            st.warning("Enter a valid recipient email.")
        else:
            with st.spinner("Sending..."):
                if send_email(from_email.strip(), from_password, email_to.strip(), df):
                    st.success(f"Sent to {email_to}")

st.divider()
st.caption("Warehouse app • data saved locally in data.xlsx")
