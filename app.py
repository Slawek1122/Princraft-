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
            df = pd.read_excel(DATA_FILE)
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            return df[COLUMNS]
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


def parse_locations(text: str) -> list:
    return [line.strip() for line in text.strip().splitlines() if line.strip()]


def safe_int(value, default=0):
    try:
        if value is None or str(value).strip() == "":
            return default
        return int(value)
    except (ValueError, TypeError):
        return default


# ================== UI ==================
st.set_page_config(
    page_title="Princraft Warehouse",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    /* ===== base ===== */
    .stApp {
        background-color: #0f172a;
        color: #e2e8f0;
    }
    
    .block-container {
        padding-top: 0.8rem !important;
        padding-bottom: 1.5rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 1100px;
    }
    
    h1, h2, h3, .stMarkdown, .stCaption, label {
        color: #f1f5f9 !important;
    }
    
    h1 {
        font-size: 1.7rem !important;
        text-align: center;
        margin-bottom: 0.2rem !important;
    }
    
    /* ===== buttons – big touch targets ===== */
    .stButton > button {
        width: 100%;
        min-height: 3.6rem;
        font-size: 1.25rem !important;
        font-weight: 700;
        border-radius: 12px;
        background-color: #3b82f6 !important;
        color: white !important;
        border: none !important;
    }
    .stButton > button:hover {
        background-color: #2563eb !important;
    }
    
    /* ===== inputs ===== */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input {
        font-size: 1.2rem !important;
        padding: 0.8rem 0.85rem !important;
        border-radius: 10px !important;
        background-color: #1e293b !important;
        color: #f1f5f9 !important;
        border: 1px solid #334155 !important;
        min-height: 3rem;
    }
    
    .stTextArea textarea {
        font-size: 1.1rem !important;
        line-height: 1.45 !important;
        background-color: #1e293b !important;
        color: #f1f5f9 !important;
        border: 1px solid #334155 !important;
    }
    
    /* selectbox */
    div[data-baseweb="select"] > div {
        background-color: #1e293b !important;
        border-radius: 10px !important;
        min-height: 3rem !important;
        font-size: 1.15rem !important;
    }
    
    /* form */
    div[data-testid="stForm"] {
        border: 2px solid #334155;
        border-radius: 14px;
        padding: 1.2rem 1rem;
        background: #1e293b;
    }
    
    .stAlert {
        font-size: 1.1rem !important;
        padding: 0.9rem 1rem !important;
        border-radius: 10px;
    }
    
    /* sidebar */
    section[data-testid="stSidebar"] {
        background-color: #0f172a;
    }
    section[data-testid="stSidebar"] * {
        color: #e2e8f0 !important;
    }
    
    .stDataFrame {
        font-size: 1rem;
    }
    
    .stDownloadButton > button {
        background-color: #10b981 !important;
    }
    .stDownloadButton > button:hover {
        background-color: #059669 !important;
    }
    
    /* tabs – bigger on touch */
    button[data-baseweb="tab"] {
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        padding: 0.75rem 1rem !important;
    }
    
    /* ===== PHONE (portrait) ===== */
    @media (max-width: 600px) {
        h1 { font-size: 1.45rem !important; }
        .block-container {
            padding-left: 0.6rem !important;
            padding-right: 0.6rem !important;
        }
        .stButton > button {
            min-height: 3.4rem;
            font-size: 1.15rem !important;
        }
        button[data-baseweb="tab"] {
            font-size: 0.95rem !important;
            padding: 0.65rem 0.7rem !important;
        }
    }
    
    /* ===== TABLET portrait ===== */
    @media (min-width: 601px) and (max-width: 900px) {
        h1 { font-size: 1.8rem !important; }
        .stButton > button {
            min-height: 3.8rem;
            font-size: 1.3rem !important;
        }
    }
    
    /* ===== TABLET landscape / wide ===== */
    @media (min-width: 901px) {
        .block-container {
            max-width: 1100px;
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
        }
        h1 { font-size: 2rem !important; }
        .stButton > button {
            min-height: 3.6rem;
            font-size: 1.25rem !important;
        }
        /* form fields a bit tighter on wide screens */
        div[data-testid="stForm"] {
            padding: 1.4rem 1.4rem;
        }
    }
    
    /* hide streamlit branding a bit cleaner */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

st.title("📦 Princraft Warehouse")
st.caption("Fast pallet entry – no double work")

# ----- SIDEBAR -----
with st.sidebar:
    st.header("Default locations")

    st.subheader("Full size pallet")
    full_default = st.session_state.get("full_locations_text", "PB1G18H\nPB1G19H\nPB1G20H\nPB1G21H")
    full_text = st.text_area("Full size locations", value=full_default, height=140, key="full_ta")

    st.subheader("Small pallet")
    small_default = st.session_state.get("small_locations_text", "PB2A01H\nPB2A02H\nPB2A03H\nPB2A04H")
    small_text = st.text_area("Small locations", value=small_default, height=140, key="small_ta")

    if st.button("💾 Save location lists", use_container_width=True):
        st.session_state.full_locations_text = full_text
        st.session_state.small_locations_text = small_text
        st.success("Saved")

    full_locations = parse_locations(full_text)
    small_locations = parse_locations(small_text)

    st.session_state.full_locations = full_locations
    st.session_state.small_locations = small_locations

    st.divider()
    st.caption(f"Full size: {len(full_locations)} | Small: {len(small_locations)}")

df = load_data()
used = get_used_locations(df)

# ----- TABS -----
tab_full, tab_small, tab_aldi = st.tabs(["🟦 Full size pallet", "🟧 Small pallet", "🟨 Aldi"])


def render_form_with_list(pallet_label: str, available: list, key_prefix: str):
    next_free = get_next_free_location(available, used)

    with st.form(f"form_{key_prefix}", clear_on_submit=True):
        st.subheader(f"New {pallet_label}")

        if next_free:
            st.info(f"Suggested free location: **{next_free}**")
        else:
            st.warning("No free locations left. Choose manually or add more in the sidebar.")

        options = available if available else ["(no locations defined)"]
        default_idx = 0
        if next_free and next_free in options:
            default_idx = options.index(next_free)

        selected = st.selectbox(
            "Select location",
            options=options,
            index=default_idx,
            key=f"select_{key_prefix}"
        )

        location = st.text_input(
            "Or type location manually",
            value="",
            placeholder="e.g. PB1G18H",
            key=f"loc_{key_prefix}"
        )

        job = st.text_input("Job number", key=f"job_{key_prefix}")
        order = st.text_input("Order number", key=f"order_{key_prefix}")
        design = st.text_input("Design number", key=f"design_{key_prefix}")

        col1, col2, col3 = st.columns(3)
        with col1:
            qty_card = st.text_input("quantity card", value="", placeholder="0", key=f"card_{key_prefix}")
        with col2:
            qty_packs = st.text_input("quantity packs", value="", placeholder="0", key=f"packs_{key_prefix}")
        with col3:
            total_boxes = st.text_input("total boxes", value="", placeholder="0", key=f"boxes_{key_prefix}")

        submitted = st.form_submit_button(f"➕ ADD {pallet_label.upper()}")

    if submitted:
        # if manual field empty -> use selected from list
        final_location = location.strip() if location.strip() else (selected if selected != "(no locations defined)" else "")
        if not final_location or not job.strip() or not order.strip() or not design.strip():
            st.error("Please fill in all text fields.")
        else:
            current_df = load_data()
            used_locs = get_used_locations(current_df)
            force = st.session_state.get(f"force_dup_{key_prefix}", False)

            if final_location in used_locs and not force:
                st.session_state[f"pending_row_{key_prefix}"] = {
                    "Princraft warehouse location": final_location,
                    "Job number": job.strip(),
                    "Order number": order.strip(),
                    "Design number": design.strip(),
                    "quantity card": safe_int(qty_card),
                    "quantity packs": safe_int(qty_packs),
                    "total boxes": safe_int(total_boxes)
                }
                st.session_state[f"dup_warning_{key_prefix}"] = final_location
            else:
                new_row = {
                    "Princraft warehouse location": final_location,
                    "Job number": job.strip(),
                    "Order number": order.strip(),
                    "Design number": design.strip(),
                    "quantity card": safe_int(qty_card),
                    "quantity packs": safe_int(qty_packs),
                    "total boxes": safe_int(total_boxes)
                }
                current_df = pd.concat([current_df, pd.DataFrame([new_row])], ignore_index=True)
                save_data(current_df)
                st.session_state.pop(f"force_dup_{key_prefix}", None)
                st.session_state.pop(f"dup_warning_{key_prefix}", None)
                st.session_state.pop(f"pending_row_{key_prefix}", None)
                st.success(f"✅ {pallet_label} added at **{final_location}**")
                st.balloons()
                st.rerun()

    # duplicate confirmation UI
    if st.session_state.get(f"dup_warning_{key_prefix}"):
        loc = st.session_state[f"dup_warning_{key_prefix}"]
        st.warning(f"Location **{loc}** is already used. Add anyway?")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ Yes, add anyway", key=f"force_yes_{key_prefix}", use_container_width=True):
                pending = st.session_state.get(f"pending_row_{key_prefix}")
                if pending:
                    current_df = load_data()
                    current_df = pd.concat([current_df, pd.DataFrame([pending])], ignore_index=True)
                    save_data(current_df)
                st.session_state.pop(f"dup_warning_{key_prefix}", None)
                st.session_state.pop(f"pending_row_{key_prefix}", None)
                st.session_state.pop(f"force_dup_{key_prefix}", None)
                st.success(f"✅ Added at **{loc}** (duplicate allowed)")
                st.rerun()
        with c2:
            if st.button("❌ Cancel", key=f"force_no_{key_prefix}", use_container_width=True):
                st.session_state.pop(f"dup_warning_{key_prefix}", None)
                st.session_state.pop(f"pending_row_{key_prefix}", None)
                st.rerun()


def render_form_manual(pallet_label: str, key_prefix: str):
    with st.form(f"form_{key_prefix}", clear_on_submit=True):
        st.subheader(f"New {pallet_label}")

        location = st.text_input(
            "Warehouse location",
            placeholder="Type location manually",
            key=f"loc_{key_prefix}"
        )

        job = st.text_input("Job number", key=f"job_{key_prefix}")
        order = st.text_input("Order number", key=f"order_{key_prefix}")
        design = st.text_input("Design number", key=f"design_{key_prefix}")

        col1, col2, col3 = st.columns(3)
        with col1:
            qty_card = st.text_input("quantity card", value="", placeholder="0", key=f"card_{key_prefix}")
        with col2:
            qty_packs = st.text_input("quantity packs", value="", placeholder="0", key=f"packs_{key_prefix}")
        with col3:
            total_boxes = st.text_input("total boxes", value="", placeholder="0", key=f"boxes_{key_prefix}")

        submitted = st.form_submit_button(f"➕ ADD {pallet_label.upper()}")

    if submitted:
        final_location = location.strip()
        if not final_location or not job.strip() or not order.strip() or not design.strip():
            st.error("Please fill in all text fields.")
        else:
            current_df = load_data()
            used_locs = get_used_locations(current_df)

            if final_location in used_locs:
                st.session_state[f"pending_row_{key_prefix}"] = {
                    "Princraft warehouse location": final_location,
                    "Job number": job.strip(),
                    "Order number": order.strip(),
                    "Design number": design.strip(),
                    "quantity card": safe_int(qty_card),
                    "quantity packs": safe_int(qty_packs),
                    "total boxes": safe_int(total_boxes)
                }
                st.session_state[f"dup_warning_{key_prefix}"] = final_location
            else:
                new_row = {
                    "Princraft warehouse location": final_location,
                    "Job number": job.strip(),
                    "Order number": order.strip(),
                    "Design number": design.strip(),
                    "quantity card": safe_int(qty_card),
                    "quantity packs": safe_int(qty_packs),
                    "total boxes": safe_int(total_boxes)
                }
                current_df = pd.concat([current_df, pd.DataFrame([new_row])], ignore_index=True)
                save_data(current_df)
                st.session_state.pop(f"dup_warning_{key_prefix}", None)
                st.session_state.pop(f"pending_row_{key_prefix}", None)
                st.success(f"✅ {pallet_label} added at **{final_location}**")
                st.balloons()
                st.rerun()

    if st.session_state.get(f"dup_warning_{key_prefix}"):
        loc = st.session_state[f"dup_warning_{key_prefix}"]
        st.warning(f"Location **{loc}** is already used. Add anyway?")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ Yes, add anyway", key=f"force_yes_{key_prefix}", use_container_width=True):
                pending = st.session_state.get(f"pending_row_{key_prefix}")
                if pending:
                    current_df = load_data()
                    current_df = pd.concat([current_df, pd.DataFrame([pending])], ignore_index=True)
                    save_data(current_df)
                st.session_state.pop(f"dup_warning_{key_prefix}", None)
                st.session_state.pop(f"pending_row_{key_prefix}", None)
                st.success(f"✅ Added at **{loc}** (duplicate allowed)")
                st.rerun()
        with c2:
            if st.button("❌ Cancel", key=f"force_no_{key_prefix}", use_container_width=True):
                st.session_state.pop(f"dup_warning_{key_prefix}", None)
                st.session_state.pop(f"pending_row_{key_prefix}", None)
                st.rerun()


with tab_full:
    render_form_with_list("Full size pallet", st.session_state.get("full_locations", full_locations), "full")

with tab_small:
    render_form_with_list("Small pallet", st.session_state.get("small_locations", small_locations), "small")

with tab_aldi:
    render_form_manual("Aldi", "aldi")

# ----- PREVIEW -----
st.divider()
st.subheader("Recent entries")

df = load_data()

# Undo last + Start new list
col_a, col_b = st.columns(2)
with col_a:
    if st.button("↩️ Undo last entry", use_container_width=True, help="Remove the last added pallet"):
        current_df = load_data()
        if not current_df.empty:
            current_df = current_df.iloc[:-1]
            save_data(current_df)
            st.success("Last entry removed")
            st.rerun()
        else:
            st.info("Nothing to undo")
with col_b:
    if st.button("🗑️ Start new list", use_container_width=True, help="Clear all data and start fresh"):
        st.session_state["confirm_clear"] = True

if st.session_state.get("confirm_clear"):
    st.warning("This will delete ALL current entries. Are you sure?")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ Yes, clear everything", use_container_width=True):
            if os.path.exists(DATA_FILE):
                os.remove(DATA_FILE)
            st.session_state["confirm_clear"] = False
            st.success("List cleared. Starting fresh.")
            st.rerun()
    with c2:
        if st.button("❌ Cancel", use_container_width=True):
            st.session_state["confirm_clear"] = False
            st.rerun()

if df.empty:
    st.info("No data yet. Add the first pallet.")
else:
    st.dataframe(df.tail(10), use_container_width=True, hide_index=True)

    show_all = st.checkbox("Show full table")
    if show_all:
        st.dataframe(df, use_container_width=True, hide_index=True)

    st.caption(f"Total entries: {len(df)}")

    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    st.download_button(
        label="📥 DOWNLOAD EXCEL",
        data=buffer.getvalue(),
        file_name=f"warehouse_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

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
