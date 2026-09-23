from __future__ import annotations

import base64
import json
import os
import re
import uuid
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import streamlit as st


# =========================================================
# APP / DATA CONFIG
# =========================================================
APP_TITLE = "PB Hotel Pool Access"
TIMEZONE = ZoneInfo("Asia/Kuala_Lumpur")

PMS_SHEET_ID = "13v0hn3d0I6bDSBicLWq3zQBQcUR1Vs8pP-hnFCZV4HM"
PMS_SHEET_NAME = "POOL_PMS_LATEST"

# Morning fallback: the PMS snapshot already loaded to HK Room208 allocation.
# Pool Traffic only reads this sheet; it never changes maid allocation data.
HK_ALLOCATION_SHEET_ID = "1_HE-Lh4nFtQWEyNmMBJNFLPwK2oEoMdNH_17Yrg6LIY"
HK_ALLOCATION_SHEET_NAME = "Room208_Tasks"

TRAFFIC_SHEET_ID = "1YpPkbYZBFSXjpuAHBVx0aYPJDTAtmEyAGXhRrAyiHtM"
TRAFFIC_LOG_SHEET = "POOL_TRAFFIC_LOG"
TRAFFIC_ALERT_SHEET = "POOL_TRAFFIC_ALERTS"
TRAFFIC_CONFIG_SHEET = "POOL_TRAFFIC_CONFIG"

DEFAULT_CONFIG = {
    "MORNING_START": "07:00",
    "MORNING_END": "12:00",
    "EVENING_START": "14:00",
    "EVENING_END": "19:00",
    "ALERT_HOURS": "3",
    "MAX_TOWEL_PER_ENTRY": "2",
}

# Locked operating rule: maximum 2 pool towels per room for each session.
MAX_TOWELS_PER_ROOM_SESSION = 2

ACTIVE_PMS_PRIORITIES = {
    "INHOUSE",
    "STAY OVER",
    "DUE OUT",
    "DUE OUT + DUE IN",
    "DUE IN + DUE OUT",
}


# =========================================================
# BASIC HELPERS
# =========================================================
def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def upper(value: Any) -> str:
    return " ".join(clean_text(value).upper().replace("_", " ").split())


def now_local() -> datetime:
    return datetime.now(TIMEZONE)


def timestamp_text(value: datetime | None = None) -> str:
    value = value or now_local()
    return value.strftime("%Y-%m-%d %H:%M:%S")


def normalise_room(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
    except Exception:
        pass
    return text.lstrip("0") or "0"


def parse_hhmm(value: str, fallback: str) -> time:
    text = clean_text(value) or fallback
    try:
        return datetime.strptime(text, "%H:%M").time()
    except ValueError:
        return datetime.strptime(fallback, "%H:%M").time()


def new_id(prefix: str) -> str:
    return (
        f"{prefix}-{now_local():%Y%m%d-%H%M%S}-"
        f"{uuid.uuid4().hex[:6].upper()}"
    )


# =========================================================
# GOOGLE SHEETS CONNECTION
# =========================================================
def _normalise_private_key(info: dict[str, Any]) -> dict[str, Any]:
    data = dict(info)
    if isinstance(data.get("private_key"), str):
        data["private_key"] = data["private_key"].replace("\\n", "\n")
    return data


def _service_account_info() -> dict[str, Any] | None:
    app_dir = Path(__file__).resolve().parent

    # Optional explicit path for local/Synology deployment.
    explicit_path = os.getenv("AOC_GOOGLE_SERVICE_ACCOUNT_PATH", "").strip()
    candidates = []
    if explicit_path:
        candidates.append(Path(explicit_path))

    # Standalone public app: credential may sit beside app.py.
    # AOC commonly uses secrets.json; service_account.json is also supported.
    candidates.append(app_dir / "secrets.json")
    candidates.append(app_dir / "service_account.json")

    # Convenient local staging layout: public app folder inside/under AOC root.
    candidates.append(app_dir.parent / "secrets.json")
    candidates.append(app_dir.parent / "service_account.json")

    for local_file in candidates:
        if local_file.exists() and local_file.is_file():
            data = json.loads(local_file.read_text(encoding="utf-8"))
            return _normalise_private_key(data)

    # Streamlit Cloud / environment deployment using one-line Base64.
    # This avoids TOML multiline/private-key formatting issues.
    raw_b64 = os.getenv("GOOGLE_SERVICE_ACCOUNT_B64", "").strip()
    if not raw_b64:
        try:
            raw_b64 = clean_text(st.secrets.get("GOOGLE_SERVICE_ACCOUNT_B64", ""))
        except Exception:
            raw_b64 = ""
    if raw_b64:
        try:
            decoded = base64.b64decode(raw_b64).decode("utf-8")
            return _normalise_private_key(json.loads(decoded))
        except Exception as exc:
            raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_B64 tidak sah.") from exc

    # Optional environment deployment using raw JSON.
    raw_env = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw_env:
        return _normalise_private_key(json.loads(raw_env))

    # Streamlit Cloud / Synology secrets pattern used by AOC.
    try:
        if "gcp_service_account" in st.secrets:
            return _normalise_private_key(dict(st.secrets["gcp_service_account"]))
    except Exception:
        pass

    return None


@st.cache_resource(show_spinner=False)
def google_client():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as exc:
        raise RuntimeError(
            "Library Google belum lengkap. Install gspread dan google-auth."
        ) from exc

    info = _service_account_info()
    if not info:
        raise RuntimeError(
            "Google service account tidak dijumpai. Gunakan secrets.json / service_account.json "
            "untuk local test atau [gcp_service_account] dalam Streamlit secrets."
        )

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    credentials = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(credentials)


@st.cache_resource(show_spinner=False)
def worksheet(spreadsheet_id: str, sheet_name: str):
    return google_client().open_by_key(spreadsheet_id).worksheet(sheet_name)


def rows_as_dicts(ws) -> list[dict[str, str]]:
    values = ws.get_all_values()
    if not values:
        return []
    headers = [clean_text(v) for v in values[0]]
    result: list[dict[str, str]] = []
    for raw in values[1:]:
        row = {
            header: clean_text(raw[i]) if i < len(raw) else ""
            for i, header in enumerate(headers)
            if header
        }
        if any(row.values()):
            result.append(row)
    return result


def append_record(ws, record: dict[str, Any]) -> None:
    headers = [clean_text(v) for v in ws.row_values(1)]
    if not headers:
        raise RuntimeError("Header Google Sheet tidak dijumpai.")
    ws.append_row(
        [record.get(header, "") for header in headers],
        value_input_option="USER_ENTERED",
    )


def _backend_preflight() -> tuple[bool, str]:
    """Verify core Pool Traffic sheets before guest entry.

    PMS fallback is resolved only when the guest submits. This keeps the public
    form fast while still failing closed if a required PMS source is unreadable.
    """
    try:
        worksheet(PMS_SHEET_ID, PMS_SHEET_NAME)
        worksheet(TRAFFIC_SHEET_ID, TRAFFIC_LOG_SHEET)
        worksheet(TRAFFIC_SHEET_ID, TRAFFIC_ALERT_SHEET)
        return True, ""
    except Exception as exc:
        return False, clean_text(exc)


# =========================================================
# POOL CONFIG / SESSION
# =========================================================
@st.cache_data(ttl=60, show_spinner=False)
def load_config() -> dict[str, str]:
    config = dict(DEFAULT_CONFIG)
    try:
        ws = worksheet(TRAFFIC_SHEET_ID, TRAFFIC_CONFIG_SHEET)
        for row in rows_as_dicts(ws):
            key = upper(row.get("KEY"))
            value = clean_text(row.get("VALUE"))
            active = upper(row.get("ACTIVE"))
            if key and value and active not in {"FALSE", "0", "NO", "N"}:
                config[key] = value
    except Exception:
        # Operational fallback: session can still run with locked default values.
        pass
    return config


def current_session(now: datetime, config: dict[str, str]) -> tuple[str, str]:
    current = now.time().replace(second=0, microsecond=0)
    morning_start = parse_hhmm(config.get("MORNING_START", ""), "07:00")
    morning_end = parse_hhmm(config.get("MORNING_END", ""), "12:00")
    evening_start = parse_hhmm(config.get("EVENING_START", ""), "14:00")
    evening_end = parse_hhmm(config.get("EVENING_END", ""), "19:00")

    if morning_start <= current < morning_end:
        return "MORNING", "Morning Session · 7:00 AM – 12:00 PM"
    if evening_start <= current < evening_end:
        return "EVENING", "Evening Session · 2:00 PM – 7:00 PM"

    if current < morning_start:
        return "CLOSED", "Registration opens at 7:00 AM"
    if morning_end <= current < evening_start:
        return "CLOSED", "Morning session has ended. Evening session opens at 2:00 PM"
    return "CLOSED", "Pool registration is closed for today"


# =========================================================
# PMS VERIFICATION
# =========================================================
def _snapshot_id_from_timestamp(prefix: str, value: str) -> str:
    digits = re.sub(r"[^0-9]", "", clean_text(value))
    if len(digits) >= 14:
        return f"{prefix}-{digits[:14]}"
    if len(digits) >= 8:
        return f"{prefix}-{digits[:8]}"
    return f"{prefix}-{now_local():%Y%m%d}"


@st.cache_data(ttl=30, show_spinner=False)
def load_pool_latest_snapshot(work_date: str) -> dict[str, Any]:
    """Return POOL_LATEST only when that snapshot was uploaded today."""
    ws = worksheet(PMS_SHEET_ID, PMS_SHEET_NAME)
    rows = rows_as_dicts(ws)

    last_updated = ""
    for row in rows:
        uploaded_at = clean_text(row.get("UPLOADED_AT"))
        if uploaded_at and uploaded_at > last_updated:
            last_updated = uploaded_at

    # A previous-day Pool PMS must never override today's HK allocation snapshot.
    if not last_updated or last_updated[:10] != work_date:
        return {
            "available": False,
            "source": "POOL_LATEST",
            "snapshot_id": "",
            "last_updated": last_updated,
            "rooms": {},
        }

    room_map: dict[str, dict[str, str]] = {}
    for row in rows:
        if clean_text(row.get("UPLOADED_AT")) != last_updated:
            # The current implementation replaces the sheet. This guard also
            # protects us if upload history is added to the same tab later.
            continue
        room = normalise_room(row.get("ROOM_NO"))
        if room:
            room_map[room] = row

    return {
        "available": bool(room_map),
        "source": "POOL_LATEST",
        "snapshot_id": _snapshot_id_from_timestamp("POOL", last_updated),
        "last_updated": last_updated,
        "rooms": room_map,
    }


@st.cache_data(ttl=30, show_spinner=False)
def load_hk_allocation_snapshot(work_date: str) -> dict[str, Any]:
    """Read today's locked HK allocation as the morning PMS baseline.

    This is read-only. Pool Traffic does not change Room208_Tasks, maid assignment,
    status, or any HK workflow field.
    """
    ws = worksheet(HK_ALLOCATION_SHEET_ID, HK_ALLOCATION_SHEET_NAME)
    rows = [
        row for row in rows_as_dicts(ws)
        if clean_text(row.get("DATE")) == work_date
    ]

    if not rows:
        return {
            "available": False,
            "source": "HK_ALLOCATION",
            "snapshot_id": "",
            "last_updated": "",
            "rooms": {},
        }

    # One daily load is enforced by HK. If legacy/duplicate batches exist, use
    # the newest loaded batch as the baseline instead of mixing snapshots.
    latest_row = max(rows, key=lambda row: clean_text(row.get("LOADED_TIME")))
    batch_id = clean_text(latest_row.get("BATCH_ID"))
    loaded_time = clean_text(latest_row.get("LOADED_TIME"))
    selected = [
        row for row in rows
        if not batch_id or clean_text(row.get("BATCH_ID")) == batch_id
    ]

    room_map: dict[str, dict[str, str]] = {}
    for row in selected:
        room = normalise_room(row.get("ROOM_NO"))
        if room:
            room_map[room] = row

    return {
        "available": bool(room_map),
        "source": "HK_ALLOCATION",
        "snapshot_id": batch_id or _snapshot_id_from_timestamp("HK", loaded_time),
        "last_updated": loaded_time,
        "rooms": room_map,
    }


def select_pms_snapshot(work_date: str) -> dict[str, Any]:
    """Locked priority: today's POOL_LATEST -> today's HK_ALLOCATION -> NONE."""
    pool_latest = load_pool_latest_snapshot(work_date)
    if pool_latest.get("available"):
        return pool_latest

    hk_allocation = load_hk_allocation_snapshot(work_date)
    if hk_allocation.get("available"):
        return hk_allocation

    return {
        "available": False,
        "source": "NONE",
        "snapshot_id": "",
        "last_updated": "",
        "rooms": {},
    }


def pms_room_is_active(row: dict[str, str] | None) -> bool:
    if not row:
        return False

    priority = upper(row.get("PRIORITY"))
    house_status = upper(row.get("HOUSE_STATUS"))
    guest_name = clean_text(row.get("GUEST_NAME"))

    if priority in ACTIVE_PMS_PRIORITIES:
        return True

    # Defensive fallback for a current occupied-clean room where Priority may
    # be blank in an unusual export. Due In is deliberately NOT accepted.
    if house_status == "OC" and guest_name and priority != "DUE IN":
        return True

    return False


# =========================================================
# TRAFFIC LOG / DUPLICATE CONTROL
# =========================================================
def _safe_int(value: Any) -> int:
    try:
        return int(float(clean_text(value) or 0))
    except (TypeError, ValueError):
        return 0


def room_session_summary(room_no: str, work_date: str, session: str) -> dict[str, Any]:
    """Return cumulative users and towels for one room in one pool session."""
    ws = worksheet(TRAFFIC_SHEET_ID, TRAFFIC_LOG_SHEET)
    matched: list[dict[str, str]] = []

    for row in rows_as_dicts(ws):
        if (
            normalise_room(row.get("ROOM_NO")) == room_no
            and clean_text(row.get("WORK_DATE")) == work_date
            and upper(row.get("SESSION")) == session
        ):
            matched.append(row)

    if not matched:
        return {
            "exists": False,
            "adult": 0,
            "children": 0,
            "total_guest": 0,
            "towels": 0,
            "primary": {},
        }

    # The base registration keeps the original verification/PMS audit fields.
    primary = next(
        (
            row for row in matched
            if upper(row.get("VERIFY_STATUS")) in {"VERIFIED", "NON RECORDED"}
            and upper(row.get("SOURCE")) == "PUBLIC QR"
        ),
        matched[0],
    )

    return {
        "exists": True,
        "adult": sum(_safe_int(row.get("ADULT")) for row in matched),
        "children": sum(_safe_int(row.get("CHILDREN")) for row in matched),
        "total_guest": sum(_safe_int(row.get("TOTAL_GUEST")) for row in matched),
        "towels": sum(_safe_int(row.get("TOWEL_QTY")) for row in matched),
        "primary": primary,
    }


def room_already_registered(room_no: str, work_date: str, session: str) -> bool:
    return bool(room_session_summary(room_no, work_date, session).get("exists"))


def save_verified(
    *,
    room_no: str,
    adult: int,
    children: int,
    towel_qty: int,
    session: str,
    pms_row: dict[str, str],
    pms_snapshot: dict[str, Any],
) -> str:
    traffic_id = new_id("POOL")
    now = now_local()
    pms_source = clean_text(pms_snapshot.get("source")) or "NONE"
    record = {
        "TRAFFIC_ID": traffic_id,
        "TIMESTAMP": timestamp_text(now),
        "WORK_DATE": now.strftime("%Y-%m-%d"),
        "SESSION": session,
        "ROOM_NO": room_no,
        # PMS guest name is intentionally not exposed or copied to traffic log.
        "GUEST_NAME": "",
        "ADULT": adult,
        "CHILDREN": children,
        "TOTAL_GUEST": adult + children,
        "TOWEL_QTY": towel_qty,
        "VERIFY_STATUS": "VERIFIED",
        "PMS_LAST_UPDATED": clean_text(pms_snapshot.get("last_updated")),
        "PMS_MATCH_PRIORITY": clean_text(pms_row.get("PRIORITY")),
        "SOURCE": "PUBLIC_QR",
        "REMARKS": f"Room verified against {pms_source} snapshot.",
        "PMS_SOURCE": pms_source,
        "PMS_SNAPSHOT_ID": clean_text(pms_snapshot.get("snapshot_id")),
        "PMS_VERIFIED_AT": timestamp_text(now),
    }
    append_record(worksheet(TRAFFIC_SHEET_ID, TRAFFIC_LOG_SHEET), record)
    return traffic_id


def save_session_update(
    *,
    room_no: str,
    adult: int,
    children: int,
    towel_qty: int,
    session: str,
    base_registration: dict[str, str],
) -> str:
    """Append new users/towels to an existing room session without duplicating the base visit."""
    now = now_local()
    traffic_id = new_id("POOLUPD")
    record = {
        "TRAFFIC_ID": traffic_id,
        "TIMESTAMP": timestamp_text(now),
        "WORK_DATE": now.strftime("%Y-%m-%d"),
        "SESSION": session,
        "ROOM_NO": room_no,
        "GUEST_NAME": "",
        "ADULT": adult,
        "CHILDREN": children,
        "TOTAL_GUEST": adult + children,
        "TOWEL_QTY": towel_qty,
        # Keep verification KPIs based on the original registration only.
        "VERIFY_STATUS": "SESSION_UPDATE",
        "PMS_LAST_UPDATED": clean_text(base_registration.get("PMS_LAST_UPDATED")),
        "PMS_MATCH_PRIORITY": clean_text(base_registration.get("PMS_MATCH_PRIORITY")),
        "SOURCE": "PUBLIC_QR_SESSION_UPDATE",
        "REMARKS": "Additional users/towels for existing room session; base registration not duplicated.",
        "PMS_SOURCE": clean_text(base_registration.get("PMS_SOURCE")),
        "PMS_SNAPSHOT_ID": clean_text(base_registration.get("PMS_SNAPSHOT_ID")),
        "PMS_VERIFIED_AT": clean_text(base_registration.get("PMS_VERIFIED_AT")),
    }
    append_record(worksheet(TRAFFIC_SHEET_ID, TRAFFIC_LOG_SHEET), record)
    return traffic_id


def save_non_recorded(
    *,
    room_no: str,
    guest_name: str,
    adult: int,
    children: int,
    towel_qty: int,
    session: str,
    pms_snapshot: dict[str, Any],
    alert_hours: int,
) -> str:
    now = now_local()
    traffic_id = new_id("POOL")
    alert_id = new_id("PALERT")

    pms_source = clean_text(pms_snapshot.get("source")) or "NONE"

    traffic_record = {
        "TRAFFIC_ID": traffic_id,
        "TIMESTAMP": timestamp_text(now),
        "WORK_DATE": now.strftime("%Y-%m-%d"),
        "SESSION": session,
        "ROOM_NO": room_no,
        "GUEST_NAME": guest_name,
        "ADULT": adult,
        "CHILDREN": children,
        "TOTAL_GUEST": adult + children,
        "TOWEL_QTY": towel_qty,
        "VERIFY_STATUS": "NON_RECORDED",
        "PMS_LAST_UPDATED": clean_text(pms_snapshot.get("last_updated")),
        "PMS_MATCH_PRIORITY": "",
        "SOURCE": "PUBLIC_QR",
        "REMARKS": f"Room not verified against {pms_source} snapshot; guest name supplied.",
        "PMS_SOURCE": pms_source,
        "PMS_SNAPSHOT_ID": clean_text(pms_snapshot.get("snapshot_id")),
        "PMS_VERIFIED_AT": "",
    }
    append_record(worksheet(TRAFFIC_SHEET_ID, TRAFFIC_LOG_SHEET), traffic_record)

    alert_record = {
        "ALERT_ID": alert_id,
        "TRAFFIC_ID": traffic_id,
        "CREATED_AT": timestamp_text(now),
        "EXPIRES_AT": timestamp_text(now + timedelta(hours=alert_hours)),
        "ROOM_NO": room_no,
        "GUEST_NAME": guest_name,
        "ADULT": adult,
        "CHILDREN": children,
        "TOWEL_QTY": towel_qty,
        "STATUS": "ACTIVE",
        "DISMISSED_BY": "",
        "DISMISSED_AT": "",
        "SOURCE": "POOL_TRAFFIC",
        "REMARKS": "NON_RECORDED guest registration. PA/SV acknowledgement only; not a task.",
    }
    append_record(worksheet(TRAFFIC_SHEET_ID, TRAFFIC_ALERT_SHEET), alert_record)
    return traffic_id


# =========================================================
# OPTIONAL QR KEY GATE
# =========================================================
def qr_access_allowed() -> bool:
    try:
        required = clean_text(st.secrets.get("POOL_QR_KEY", ""))
    except Exception:
        required = ""

    if not required:
        return True

    try:
        provided = clean_text(st.query_params.get("key", ""))
    except Exception:
        provided = ""

    return bool(provided) and provided == required


# =========================================================
# UI
# =========================================================
st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🏊",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      #MainMenu, footer {visibility:hidden;}
      header {visibility:hidden; height:0px;}
      .block-container {max-width:560px; padding-top:18px; padding-bottom:36px;}
      .pool-hero {
        border-radius:24px;
        padding:24px 22px;
        background:linear-gradient(145deg,#0B3558,#087F8C);
        color:white;
        margin-bottom:16px;
        box-shadow:0 10px 30px rgba(0,0,0,.12);
      }
      .pool-brand {font-size:12px; letter-spacing:.12em; opacity:.78; font-weight:700;}
      .pool-title {font-size:30px; font-weight:800; line-height:1.08; margin-top:6px;}
      .pool-sub {font-size:15px; opacity:.9; margin-top:7px;}
      div.stButton > button, div[data-testid="stFormSubmitButton"] > button {
        min-height:52px;
        border-radius:14px;
        font-weight:750;
      }
      [data-testid="stNumberInput"] input, [data-testid="stTextInput"] input {
        min-height:48px;
        border-radius:12px;
      }
      .small-note {font-size:12px; opacity:.68; text-align:center; margin-top:14px;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="pool-hero">
      <div class="pool-brand">PAYA BUNGA HOTEL</div>
      <div class="pool-title">🏊 PB Hotel Pool Access</div>
      <div class="pool-sub">Guest Registration</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if not qr_access_allowed():
    st.error("This QR link is not valid. Please scan the Pool Access QR displayed at the swimming pool.")
    st.stop()

config = load_config()
now = now_local()
session, session_label = current_session(now, config)

if session == "CLOSED":
    st.info(f"🕒 {session_label}")
    st.caption("Pool operating sessions: 7:00 AM–12:00 PM and 2:00 PM–7:00 PM.")
    st.stop()

st.success(f"🟢 {session_label}")

# Fail closed if the Google backend is unavailable. A system/credential outage
# must never be misclassified as a NON_RECORDED guest.
backend_ok, backend_error = _backend_preflight()
if not backend_ok:
    st.error("Pool registration is temporarily unavailable. Please see the Pool Attendant.")
    st.stop()

# Show final success as a clean completion screen.
success_data = st.session_state.get("pool_success")
if isinstance(success_data, dict):
    st.markdown("### ✅ Welcome to PB Hotel Pool!")
    st.success("We're delighted to have you with us. Enjoy your swim.")
    st.write(
        f"**Room {success_data.get('room', '')} · "
        f"{success_data.get('session', '').title()} Session**"
    )
    st.write(
        f"**Guests:** {success_data.get('guests', 0)}  ·  "
        f"**Towels:** {success_data.get('towels', 0)}"
    )
    st.warning(
        "⚠️ **Child Safety Reminder**  \n"
        "Please supervise children continuously. Avoid phone distractions and "
        "keep young children within arm’s reach at all times in and around the pool."
    )
    if st.button("Register another room", width="stretch"):
        st.session_state.pop("pool_success", None)
        st.rerun()
    st.stop()

update_success = st.session_state.get("pool_update_success")
if isinstance(update_success, dict):
    st.markdown("### ✅ Pool Session Updated")
    st.success(
        f"Room {update_success.get('room', '')} · "
        f"{str(update_success.get('session', '')).title()} Session"
    )
    st.write(
        f"**Session users:** {update_success.get('total_guest', 0)}  ·  "
        f"**Towels:** {update_success.get('towels', 0)}/{MAX_TOWELS_PER_ROOM_SESSION}"
    )
    st.caption("Only newly joined users were added. The original room registration was not duplicated.")
    if st.button("Register another room", key="pool_update_next", width="stretch"):
        st.session_state.pop("pool_update_success", None)
        st.rerun()
    st.stop()


existing_session = st.session_state.get("pool_existing_session")
if isinstance(existing_session, dict):
    existing_room = normalise_room(existing_session.get("room"))
    existing_work_date = clean_text(existing_session.get("work_date"))
    existing_session_name = upper(existing_session.get("session"))

    try:
        current_config = load_config()
        live_session, _ = current_session(now_local(), current_config)
        fresh = room_session_summary(existing_room, existing_work_date, existing_session_name)
    except Exception:
        st.error("Current room session could not be loaded. Please see the Pool Attendant.")
        st.stop()

    if not fresh.get("exists"):
        st.session_state.pop("pool_existing_session", None)
        st.warning("The earlier room registration could not be found. Please register again.")
        st.rerun()

    current_adult = _safe_int(fresh.get("adult"))
    current_children = _safe_int(fresh.get("children"))
    current_guests = _safe_int(fresh.get("total_guest"))
    current_towels = _safe_int(fresh.get("towels"))
    remaining_towels = max(0, MAX_TOWELS_PER_ROOM_SESSION - current_towels)
    suggested_adult = max(0, min(_safe_int(existing_session.get("suggested_adult")), 6))
    suggested_children = max(0, min(_safe_int(existing_session.get("suggested_children")), 6))
    suggested_towel = max(0, _safe_int(existing_session.get("suggested_towel")))

    st.markdown(f"### Room {existing_room} · {existing_session_name.title()} Session")
    st.info("This room is already registered for the current session.")
    st.write(
        f"**Users recorded:** {current_guests} "
        f"(Adult {current_adult} · Children {current_children})"
    )
    st.write(
        f"**Pool towels:** {current_towels}/{MAX_TOWELS_PER_ROOM_SESSION}"
    )
    st.caption(
        "If NEW users from this room join later in the same session, add only those new users below. "
        "Do not add people who were already counted earlier."
    )

    if live_session != existing_session_name:
        st.warning("This session is no longer active. A new session must use a new registration.")
        if st.button("← Back", key="existing_session_closed_back", width="stretch"):
            st.session_state.pop("pool_existing_session", None)
            st.rerun()
        st.stop()

    with st.form("pool_existing_session_update", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            add_adult = st.selectbox(
                "New Adult",
                list(range(0, 7)),
                index=suggested_adult,
            )
        with c2:
            add_children = st.selectbox(
                "New Children",
                list(range(0, 7)),
                index=suggested_children,
            )

        if remaining_towels > 0:
            towel_options = list(range(0, remaining_towels + 1))
            towel_index = min(suggested_towel, remaining_towels)
            add_towel = st.selectbox(
                "Additional Pool Towel",
                towel_options,
                index=towel_index,
                help=(
                    f"Room {existing_room} has {remaining_towels} towel allowance remaining "
                    f"for this session. Maximum is {MAX_TOWELS_PER_ROOM_SESSION}."
                ),
            )
        else:
            add_towel = 0
            st.warning(
                f"Towel limit reached: {current_towels}/{MAX_TOWELS_PER_ROOM_SESSION}. "
                "No additional towel can be issued in this session."
            )

        confirm_new_users = st.checkbox(
            "I confirm any guest numbers added above are NEW users not previously counted.",
            value=False,
        )
        update_submit = st.form_submit_button(
            "UPDATE THIS SESSION",
            type="primary",
            width="stretch",
        )

    if update_submit:
        add_adult = int(add_adult)
        add_children = int(add_children)
        add_towel = int(add_towel)
        add_guests = add_adult + add_children

        if add_guests <= 0 and add_towel <= 0:
            st.info("No new users or towels were added. Existing counts remain unchanged.")
        elif add_guests > 0 and not confirm_new_users:
            st.error("Please confirm that the guest numbers are NEW users for this session.")
        else:
            try:
                # Re-read before append so two phones cannot push towels above 2.
                latest = room_session_summary(existing_room, existing_work_date, existing_session_name)
                latest_towels = _safe_int(latest.get("towels"))
                latest_remaining = max(0, MAX_TOWELS_PER_ROOM_SESSION - latest_towels)

                if add_towel > latest_remaining:
                    st.error(
                        f"Only {latest_remaining} towel(s) remain for this room in the current session."
                    )
                    st.stop()

                save_session_update(
                    room_no=existing_room,
                    adult=add_adult,
                    children=add_children,
                    towel_qty=add_towel,
                    session=existing_session_name,
                    base_registration=dict(latest.get("primary") or {}),
                )

                final_guest = _safe_int(latest.get("total_guest")) + add_guests
                final_towels = latest_towels + add_towel
                st.session_state.pop("pool_existing_session", None)
                st.session_state["pool_update_success"] = {
                    "room": existing_room,
                    "session": existing_session_name,
                    "total_guest": final_guest,
                    "towels": final_towels,
                }
                st.rerun()
            except Exception:
                st.error("Session update could not be saved. Please see the Pool Attendant.")

    if st.button("← Back", key="existing_session_back", width="stretch"):
        st.session_state.pop("pool_existing_session", None)
        st.rerun()
    st.stop()


pending = st.session_state.get("pool_pending_nonrecorded")
if isinstance(pending, dict):
    st.warning("Room verification is not available. Please enter the guest name to continue.")
    st.caption("A Pool Attendant may verify this registration separately.")

    with st.form("non_recorded_guest_name", clear_on_submit=False):
        st.text_input("Room No.", value=pending["room_no"], disabled=True)
        guest_name = st.text_input(
            "Guest Name *",
            placeholder="Enter guest name",
            max_chars=80,
        )
        submit_nonrecorded = st.form_submit_button(
            "Continue Registration",
            type="primary",
            width="stretch",
        )

    if submit_nonrecorded:
        guest_name = clean_text(guest_name)
        if not guest_name:
            st.error("Please enter the guest name.")
        else:
            try:
                work_date = now_local().strftime("%Y-%m-%d")
                current_config = load_config()
                live_session, _ = current_session(now_local(), current_config)
                if live_session == "CLOSED":
                    st.error("Pool registration session has ended. Please see the Pool Attendant.")
                    st.stop()

                existing = room_session_summary(pending["room_no"], work_date, live_session)
                if existing.get("exists"):
                    st.session_state.pop("pool_pending_nonrecorded", None)
                    st.session_state["pool_existing_session"] = {
                        "room": pending["room_no"],
                        "work_date": work_date,
                        "session": live_session,
                        "suggested_adult": int(pending["adult"]),
                        "suggested_children": int(pending["children"]),
                        "suggested_towel": int(pending["towel_qty"]),
                    }
                    st.rerun()

                alert_hours = int(float(current_config.get("ALERT_HOURS", "3") or 3))
                save_non_recorded(
                    room_no=pending["room_no"],
                    guest_name=guest_name,
                    adult=int(pending["adult"]),
                    children=int(pending["children"]),
                    towel_qty=int(pending["towel_qty"]),
                    session=live_session,
                    pms_snapshot=dict(pending.get("pms_snapshot") or {}),
                    alert_hours=max(1, alert_hours),
                )
                st.session_state.pop("pool_pending_nonrecorded", None)
                st.session_state["pool_success"] = {
                    "room": pending["room_no"],
                    "session": live_session,
                    "guests": int(pending["adult"]) + int(pending["children"]),
                    "towels": int(pending["towel_qty"]),
                }
                st.rerun()
            except Exception:
                st.error("Registration could not be saved. Please see the Pool Attendant.")

    if st.button("← Back", width="stretch"):
        st.session_state.pop("pool_pending_nonrecorded", None)
        st.rerun()
    st.stop()

st.markdown("### Register your visit")
st.caption(
    "One base registration per room per pool session. If NEW users join later, "
    "scan again and add only the new users. Maximum 2 pool towels per room per session."
)

max_towel = MAX_TOWELS_PER_ROOM_SESSION

with st.form("pool_guest_registration", clear_on_submit=False):
    room_number = st.text_input(
        "Room No. *",
        placeholder="801",
        max_chars=4,
        help="Enter numbers only, for example 801.",
    )

    c1, c2 = st.columns(2)
    with c1:
        adult = st.selectbox("Adult", list(range(0, 7)), index=1)
    with c2:
        children = st.selectbox("Children", list(range(0, 7)), index=0)

    towel_qty = st.selectbox(
        "Pool Towel",
        list(range(0, max_towel + 1)),
        index=0,
        help="Select the number of pool towels requested for this registration.",
    )

    submit = st.form_submit_button(
        "CHECK IN POOL",
        type="primary",
        width="stretch",
    )

if submit:
    raw_room = clean_text(room_number).replace(" ", "")
    room_no = normalise_room(raw_room)
    adult = int(adult)
    children = int(children)
    towel_qty = int(towel_qty)

    if not raw_room:
        st.error("Please enter your room number.")
    elif not re.fullmatch(r"\d{3,4}", raw_room):
        st.error("Enter a valid room number using numbers only, for example 801.")
    elif adult + children <= 0:
        st.error("Please select at least one guest.")
    else:
        try:
            current_now = now_local()
            live_session, _ = current_session(current_now, load_config())
            if live_session == "CLOSED":
                st.error("Pool registration session has ended. Please see the Pool Attendant.")
                st.stop()

            work_date = current_now.strftime("%Y-%m-%d")
            existing = room_session_summary(room_no, work_date, live_session)
            if existing.get("exists"):
                st.session_state["pool_existing_session"] = {
                    "room": room_no,
                    "work_date": work_date,
                    "session": live_session,
                    "suggested_adult": adult,
                    "suggested_children": children,
                    "suggested_towel": towel_qty,
                }
                st.rerun()

            # Guest sees no extra step. Backend chooses the best PMS reference:
            # today's Pool PMS first; otherwise today's locked HK allocation.
            pms_snapshot = select_pms_snapshot(work_date)
            pms_rooms = dict(pms_snapshot.get("rooms") or {})
            pms_row = pms_rooms.get(room_no)

            if pms_room_is_active(pms_row):
                save_verified(
                    room_no=room_no,
                    adult=adult,
                    children=children,
                    towel_qty=towel_qty,
                    session=live_session,
                    pms_row=pms_row or {},
                    pms_snapshot=pms_snapshot,
                )
                st.session_state["pool_success"] = {
                    "room": room_no,
                    "session": live_session,
                    "guests": adult + children,
                    "towels": towel_qty,
                }
                st.rerun()
            else:
                st.session_state["pool_pending_nonrecorded"] = {
                    "room_no": room_no,
                    "adult": adult,
                    "children": children,
                    "towel_qty": towel_qty,
                    # Freeze the exact PMS source used for this failed check so
                    # later uploads cannot silently change the audit trail.
                    "pms_snapshot": {
                        "source": clean_text(pms_snapshot.get("source")),
                        "snapshot_id": clean_text(pms_snapshot.get("snapshot_id")),
                        "last_updated": clean_text(pms_snapshot.get("last_updated")),
                    },
                }
                st.rerun()

        except Exception:
            # A backend outage is not the same as a room failing PMS verification.
            # Do not create a false NON_RECORDED registration.
            st.error("Pool registration is temporarily unavailable. Please see the Pool Attendant.")

st.markdown(
    '<div class="small-note">Pool registration is used for operational attendance and towel control.</div>',
    unsafe_allow_html=True,
)
