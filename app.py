"""Faculty Guide Selection - first-come, first-served (Streamlit + SQLite)."""

from __future__ import annotations

import html
import io
import os
import re
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterator

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------
DB_PATH = Path(__file__).with_name("guide_selection.db")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")
MAX_STUDENTS_PER_FACULTY = 5

INTERESTS = [
    "Marketing",
    "Finance",
    "Human Resources",
    "Operations Management",
    "Analytics",
    "Banking",
    "Systems",
    "Other",
]

# Used only for the "Recommended Match" highlight. The faculty data stored in
# the database is left exactly as supplied.
SPECIALIZATION_ALIASES = {
    "human resource": "Human Resources",
    "human resources": "Human Resources",
    "analysis": "Analytics",
    "analytics": "Analytics",
}

STUDENT_NAMES = [
    "Abimanyu C", "Abishek. G", "Akilesh P", "Arunachalam S", "Asvika M",
    "Aswinkumar K", "Deepikha K S", "Divyadevi R", "Elamathi S",
    "Elan Thendral", "Govarthanan K", "Harani V S", "Jaisudhan M.V",
    "Jeevitha M", "Joheesvara Kc", "Kaviya S", "Kaviya S D", "Kavya D",
    "Kiruthika K", "Lakshmi Shankar S", "Madhumithra P", "Mahalakshmi T",
    "Manoranjith N", "Mathan Nivas M", "Mathielavarasan S", "Mouli R",
    "Naresh V", "Nevathika R", "Poorani K", "Poorvisha Meenakshi K",
    "Priyadharshini P", "Priyanga C S", "Priyanka J", "Purnashri K R",
    "Reashma S", "Sakthi M", "Saranya K R", "Sparjan Steve Mathew",
    "Subashini R", "Sudeekshaa V", "Suganesan S", "Vaishna Surya S",
    "Vikas V", "Yoshidha S",
]

FACULTY_DATA = [
    ("Dr Adhinarayanan B", "Human Resource, Marketing"),
    ("Dr Satheesh Kumar T", "Operations Management, Analytics"),
    ("Prof Senthil Kumar N", "Finance, Marketing"),
    ("Prof Nandhini B", "Human Resources, Systems"),
    ("Prof Mageswaran J", "Marketing, Finance"),
    ("Prof Dhanabalu S N", "Finance, Banking"),
    ("Prof Saranya S", "Finance, Analytics"),
    ("Prof Aishwariya M R", "Finance, Human Resources"),
    ("Prof Suganesh S", "Marketing, Analysis"),
]

GMAIL_PATTERN = re.compile(r"^[A-Za-z0-9._%+-]+@bitsathy\.ac\.in$", re.IGNORECASE)
REGISTER_PATTERN = re.compile(r"^[A-Za-z0-9]{6,15}$")


class SelectionError(Exception):
    """A problem with a submission that can be shown directly to the user."""


# --------------------------------------------------------------------------
# Database layer
# --------------------------------------------------------------------------
@contextmanager
def get_connection() -> Iterator[sqlite3.Connection]:
    """Open a connection in manual-transaction mode (we call BEGIN ourselves)."""
    conn = sqlite3.connect(DB_PATH, timeout=30, isolation_level=None)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
    finally:
        conn.close()


def init_db() -> None:
    """Create tables and insert seed data only if the tables are empty."""
    with get_connection() as conn:
        conn.execute("PRAGMA journal_mode = WAL")
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS students (
                id   INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE
            );

            CREATE TABLE IF NOT EXISTS faculty (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                name           TEXT NOT NULL UNIQUE,
                specialization TEXT NOT NULL,
                capacity       INTEGER NOT NULL DEFAULT 5
            );

            CREATE TABLE IF NOT EXISTS selections (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id      INTEGER NOT NULL UNIQUE
                                REFERENCES students(id),
                faculty_id      INTEGER NOT NULL REFERENCES faculty(id),
                register_number TEXT NOT NULL UNIQUE,
                gmail           TEXT NOT NULL,
                project_interest TEXT NOT NULL,
                submitted_at    TEXT NOT NULL
            );

            -- Last line of defence: the database itself refuses a 6th student.
            CREATE TRIGGER IF NOT EXISTS enforce_faculty_capacity
            BEFORE INSERT ON selections
            WHEN (SELECT COUNT(*) FROM selections
                  WHERE faculty_id = NEW.faculty_id)
                 >= (SELECT capacity FROM faculty WHERE id = NEW.faculty_id)
            BEGIN
                SELECT RAISE(ABORT, 'FACULTY_FULL');
            END;
            """
        )
        if conn.execute("SELECT COUNT(*) FROM students").fetchone()[0] == 0:
            conn.executemany(
                "INSERT INTO students (name) VALUES (?)",
                [(name,) for name in STUDENT_NAMES],
            )
        if conn.execute("SELECT COUNT(*) FROM faculty").fetchone()[0] == 0:
            conn.executemany(
                "INSERT INTO faculty (name, specialization, capacity) "
                "VALUES (?, ?, ?)",
                [(n, s, MAX_STUDENTS_PER_FACULTY) for n, s in FACULTY_DATA],
            )


def fetch_faculty_overview() -> list[dict]:
    """Every faculty member with live seat counts."""
    query = """
        SELECT f.id, f.name, f.specialization, f.capacity,
               COUNT(sel.id) AS filled
        FROM faculty f
        LEFT JOIN selections sel ON sel.faculty_id = f.id
        GROUP BY f.id
        ORDER BY f.id
    """
    with get_connection() as conn:
        rows = conn.execute(query).fetchall()
    return [
        {
            "id": r["id"],
            "name": r["name"],
            "specialization": r["specialization"],
            "capacity": r["capacity"],
            "filled": r["filled"],
            "remaining": r["capacity"] - r["filled"],
            "is_full": r["filled"] >= r["capacity"],
        }
        for r in rows
    ]


def fetch_pending_students() -> list[dict]:
    """Students who have not submitted a selection yet."""
    query = """
        SELECT s.id, s.name
        FROM students s
        LEFT JOIN selections sel ON sel.student_id = s.id
        WHERE sel.id IS NULL
        ORDER BY s.id
    """
    with get_connection() as conn:
        return [dict(r) for r in conn.execute(query).fetchall()]


def fetch_selections_df() -> pd.DataFrame:
    query = """
        SELECT st.name             AS "Student Name",
               sel.register_number AS "Register Number",
               sel.gmail           AS "Gmail",
               sel.project_interest AS "Project Interest",
               f.name              AS "Selected Faculty",
               sel.submitted_at    AS "Submission Time"
        FROM selections sel
        JOIN students st ON st.id = sel.student_id
        JOIN faculty f   ON f.id = sel.faculty_id
        ORDER BY sel.submitted_at, sel.id
    """
    with get_connection() as conn:
        df = pd.read_sql_query(query, conn)
    df["Submission Time"] = df["Submission Time"].str[:19]
    return df


def fetch_student_count() -> int:
    with get_connection() as conn:
        return conn.execute("SELECT COUNT(*) FROM students").fetchone()[0]


def submit_selection(
    student_id: int,
    register_number: str,
    gmail: str,
    project_interest: str,
    faculty_id: int,
) -> dict:
    """Record a selection atomically (first-come, first-served).

    BEGIN IMMEDIATE takes SQLite's write lock before we count seats, so two
    near-simultaneous submissions are processed one after the other: the
    second one sees the first one's row and is rejected if the faculty is
    full. The trigger in init_db() is a second safety net.
    """
    submitted_at = datetime.now().isoformat(sep=" ", timespec="milliseconds")
    try:
        with get_connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            try:
                if conn.execute(
                    "SELECT 1 FROM selections WHERE student_id = ?",
                    (student_id,),
                ).fetchone():
                    raise SelectionError(
                        "You have already submitted a guide selection. "
                        "Each student can submit only once."
                    )
                student = conn.execute(
                    "SELECT name FROM students WHERE id = ?", (student_id,)
                ).fetchone()
                if student is None:
                    raise SelectionError("Student not found in the list.")
                faculty = conn.execute(
                    "SELECT name, capacity FROM faculty WHERE id = ?",
                    (faculty_id,),
                ).fetchone()
                if faculty is None:
                    raise SelectionError("Selected faculty does not exist.")
                filled = conn.execute(
                    "SELECT COUNT(*) FROM selections WHERE faculty_id = ?",
                    (faculty_id,),
                ).fetchone()[0]
                if filled >= faculty["capacity"]:
                    raise SelectionError(
                        f"Sorry, {faculty['name']} just became FULL. "
                        "Please choose another available faculty."
                    )
                conn.execute(
                    "INSERT INTO selections (student_id, faculty_id, "
                    "register_number, gmail, project_interest, submitted_at) "
                    "VALUES (?, ?, ?, ?, ?, ?)",
                    (student_id, faculty_id, register_number, gmail,
                     project_interest, submitted_at),
                )
                conn.execute("COMMIT")
            except BaseException:
                if conn.in_transaction:
                    conn.execute("ROLLBACK")
                raise
    except sqlite3.IntegrityError as exc:
        message = str(exc)
        if "FACULTY_FULL" in message:
            raise SelectionError(
                "That faculty just became FULL. Please choose another."
            ) from exc
        if "selections.student_id" in message:
            raise SelectionError(
                "You have already submitted a guide selection."
            ) from exc
        if "selections.register_number" in message:
            raise SelectionError(
                "This register number has already been used by another "
                "submission."
            ) from exc
        raise SelectionError("Could not save your selection.") from exc
    except sqlite3.Error as exc:
        raise SelectionError(
            "The database is busy or unavailable. Please try again in a "
            "few seconds."
        ) from exc

    return {
        "student": student["name"],
        "faculty": faculty["name"],
        "time": submitted_at[:19],
    }


def reset_selections() -> None:
    """Delete all selections (demo reset). Students and faculty are kept."""
    with get_connection() as conn:
        conn.execute("DELETE FROM selections")


# --------------------------------------------------------------------------
# Matching and validation helpers
# --------------------------------------------------------------------------
def normalise_term(term: str) -> str:
    cleaned = term.strip()
    return SPECIALIZATION_ALIASES.get(cleaned.lower(), cleaned.title())


def is_recommended(interest: str | None, specialization: str) -> bool:
    """True if the faculty's specialization contains the student's interest."""
    if not interest or interest == "Other":
        return False
    terms = {normalise_term(t) for t in specialization.split(",")}
    return interest in terms


def validate_inputs(
    student_name: str | None,
    register_number: str,
    gmail: str,
    interest: str | None,
    faculty_id: int | None,
    valid_names: set[str],
) -> list[str]:
    errors: list[str] = []
    if not student_name:
        errors.append("Please select your name.")
    elif student_name not in valid_names:
        errors.append("Your name is not in the student list.")
    if not register_number.strip():
        errors.append("Please enter your register number.")
    elif not REGISTER_PATTERN.match(register_number.strip()):
        errors.append(
            "Register number should be 6-15 letters/digits with no spaces "
            "or symbols."
        )
    if not gmail.strip():
        errors.append("Please enter your Gmail address.")
    elif not GMAIL_PATTERN.match(gmail.strip()):
        errors.append("Please enter a valid Gmail address (name@gmail.com).")
    if not interest:
        errors.append("Please select your project interest area.")
    if faculty_id is None:
        errors.append("Please choose one faculty member.")
    return errors


# --------------------------------------------------------------------------
# Excel / CSV export
# --------------------------------------------------------------------------
def faculty_summary_df(faculty: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Faculty": f["name"],
                "Specialization": f["specialization"],
                "Seats Filled": f["filled"],
                "Seats Remaining": f["remaining"],
                "Status": "FULL" if f["is_full"] else "Available",
            }
            for f in faculty
        ]
    )


def style_sheet(sheet) -> None:
    header_fill = PatternFill("solid", fgColor="1F3A5F")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for column in sheet.columns:
        longest = max(len(str(c.value)) if c.value is not None else 0
                      for c in column)
        sheet.column_dimensions[get_column_letter(column[0].column)].width = (
            min(longest + 3, 50)
        )
    sheet.freeze_panes = "A2"


def build_excel(
    selections: pd.DataFrame, faculty: pd.DataFrame, pending: pd.DataFrame
) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        sheets = {
            "Allocations": selections,
            "Faculty Summary": faculty,
            "Not Submitted": pending,
        }
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
            style_sheet(writer.sheets[sheet_name])
    return buffer.getvalue()


# --------------------------------------------------------------------------
# UI helpers
# --------------------------------------------------------------------------
CSS = """
<style>
.block-container {max-width: 1100px; padding-top: 2rem;}
.hero {padding: 1.2rem 1.5rem; border-radius: 14px; color: #fff;
       background: linear-gradient(135deg, #1F3A5F, #3A7BD5);
       margin-bottom: 1.2rem;}
.hero h1 {margin: 0; font-size: 1.8rem; color: #fff;}
.hero p {margin: .3rem 0 0; opacity: .9;}
.fac-card {border: 1.5px solid rgba(128,128,128,.35); border-radius: 12px;
           padding: 14px 16px; margin-bottom: 14px; min-height: 170px;
           background: rgba(128,128,128,.07);}
.fac-card.rec {border-color: #F5A623; background: rgba(245,166,35,.12);
               box-shadow: 0 0 0 2px rgba(245,166,35,.35);}
.fac-card.full {opacity: .55;}
.fac-name {font-weight: 700; font-size: 1.02rem; margin: 6px 0 2px;}
.fac-spec {font-size: .88rem; opacity: .8; margin-bottom: 8px;}
.fac-seats {font-weight: 600; font-size: .95rem;}
.fac-left {font-size: .85rem; opacity: .85;}
.badge {display: inline-block; padding: 2px 9px; border-radius: 999px;
        font-size: .72rem; font-weight: 700; margin-right: 6px;}
.badge.avail {background: #1E9E5A; color: #fff;}
.badge.full {background: #D64545; color: #fff;}
.badge.rec {background: #F5A623; color: #222;}
.bar {height: 7px; border-radius: 4px; background: rgba(128,128,128,.25);
      margin: 6px 0;}
.bar > div {height: 100%; border-radius: 4px; background: #3A7BD5;}
.bar.full > div {background: #D64545;}
</style>
"""


def seats_text(remaining: int) -> str:
    if remaining <= 0:
        return "No seats remaining"
    return f"{remaining} seat{'s' if remaining != 1 else ''} remaining"


def faculty_card_html(faculty: dict, interest: str | None) -> str:
    recommended = (not faculty["is_full"]) and is_recommended(
        interest, faculty["specialization"]
    )
    classes = "fac-card"
    if recommended:
        classes += " rec"
    if faculty["is_full"]:
        classes += " full"
    status = (
        '<span class="badge full">FULL</span>'
        if faculty["is_full"]
        else '<span class="badge avail">Available</span>'
    )
    rec_badge = '<span class="badge rec">★ Recommended Match</span>' if (
        recommended) else ""
    percent = int(100 * faculty["filled"] / faculty["capacity"])
    bar_class = "bar full" if faculty["is_full"] else "bar"
    return (
        f'<div class="{classes}">{status}{rec_badge}'
        f'<div class="fac-name">{html.escape(faculty["name"])}</div>'
        f'<div class="fac-spec">{html.escape(faculty["specialization"])}</div>'
        f'<div class="fac-seats">{faculty["filled"]} / '
        f'{faculty["capacity"]} students</div>'
        f'<div class="{bar_class}"><div style="width:{percent}%"></div></div>'
        f'<div class="fac-left">{seats_text(faculty["remaining"])}</div>'
        "</div>"
    )


def render_faculty_cards(faculty: list[dict], interest: str | None) -> None:
    for start in range(0, len(faculty), 3):
        columns = st.columns(3)
        for column, member in zip(columns, faculty[start:start + 3]):
            column.markdown(
                faculty_card_html(member, interest), unsafe_allow_html=True
            )


# --------------------------------------------------------------------------
# Student page
# --------------------------------------------------------------------------
def render_student_page() -> None:
    st.markdown(
        '<div class="hero"><h1>Choose Your Project Guide</h1>'
        "<p>First-come, first-served · One faculty per student · "
        "Max 5 students per faculty</p></div>",
        unsafe_allow_html=True,
    )

    flash_error = st.session_state.pop("flash_error", None)
    if flash_error:
        st.error(flash_error)

    faculty = fetch_faculty_overview()
    confirmation = st.session_state.get("confirmation")

    if confirmation:
        st.success(
            f"✅ Submission successful, {confirmation['student']}! "
            f"Your project guide is **{confirmation['faculty']}**. "
            f"(Recorded at {confirmation['time']})"
        )
        st.info(
            "You will decide your project title together with your guide "
            "later."
        )
        if st.button("Next student"):
            st.session_state.pop("confirmation")
            st.session_state["form_id"] = st.session_state.get("form_id", 0) + 1
            st.rerun()
        st.subheader("Current faculty availability")
        render_faculty_cards(faculty, None)
        return

    pending = fetch_pending_students()
    if not pending:
        st.success("All students have submitted their guide selection.")
        render_faculty_cards(faculty, None)
        return

    name_to_id = {s["name"]: s["id"] for s in pending}
    fid = st.session_state.get("form_id", 0)

    st.subheader("1. Your details")
    left, right = st.columns(2)
    student_name = left.selectbox(
        "Student name", list(name_to_id), index=None,
        placeholder="Select your name", key=f"name_{fid}",
    )
    register_number = right.text_input(
        "Register number", key=f"reg_{fid}"
    )
    gmail = left.text_input(
        "Gmail address", placeholder="yourname@gmail.com", key=f"mail_{fid}"
    )
    interest = right.selectbox(
        "Project interest / area", INTERESTS, index=None,
        placeholder="Select an area", key=f"interest_{fid}",
    )

    st.subheader("2. Available faculty")
    if interest:
        st.caption(
            "Faculty marked ★ Recommended Match work in your interest area. "
            "This is only a suggestion; you may choose any available "
            "faculty."
        )
    render_faculty_cards(faculty, interest)

    st.subheader("3. Select one faculty")
    available = [f for f in faculty if not f["is_full"]]
    if not available:
        st.error("All faculty members are full.")
        return
    labels = {
        f["id"]: f"{f['name']} — {f['specialization']} "
                 f"({seats_text(f['remaining'])})"
        for f in available
    }
    faculty_id = st.radio(
        "Available faculty", list(labels), index=None,
        format_func=labels.get, key=f"faculty_{fid}",
    )

    if st.button("Submit selection", type="primary"):
        errors = validate_inputs(
            student_name, register_number, gmail, interest, faculty_id,
            set(name_to_id),
        )
        if errors:
            for message in errors:
                st.error(message)
            return
        try:
            result = submit_selection(
                name_to_id[student_name],
                register_number.strip().upper(),
                gmail.strip().lower(),
                interest,
                faculty_id,
            )
        except SelectionError as exc:
            st.session_state["flash_error"] = str(exc)
            st.rerun()
        st.session_state["confirmation"] = result
        st.rerun()


# --------------------------------------------------------------------------
# Admin page
# --------------------------------------------------------------------------
def render_admin_login() -> None:
    st.markdown(
        '<div class="hero"><h1>Admin Login</h1>'
        "<p>Staff access only</p></div>",
        unsafe_allow_html=True,
    )
    password = st.text_input("Admin password", type="password")
    if st.button("Log in", type="primary"):
        if password == ADMIN_PASSWORD:
            st.session_state["admin_ok"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")


def render_admin_dashboard() -> None:
    st.markdown(
        '<div class="hero"><h1>Admin Dashboard</h1>'
        "<p>Live allocation status</p></div>",
        unsafe_allow_html=True,
    )
    faculty = fetch_faculty_overview()
    selections = fetch_selections_df()
    pending = pd.DataFrame(fetch_pending_students())
    pending = (
        pending.rename(columns={"name": "Student Name"})[["Student Name"]]
        if not pending.empty
        else pd.DataFrame(columns=["Student Name"])
    )

    total_students = fetch_student_count()
    submitted = len(selections)
    total_seats = sum(f["capacity"] for f in faculty)
    allocated = sum(f["filled"] for f in faculty)

    row1 = st.columns(4)
    row1[0].metric("Total students", total_students)
    row1[1].metric("Submitted", submitted)
    row1[2].metric("Remaining students", total_students - submitted)
    row1[3].metric("Total faculty", len(faculty))
    row2 = st.columns(4)
    row2[0].metric("Total seats", total_seats)
    row2[1].metric("Allocated seats", allocated)
    row2[2].metric("Available seats", total_seats - allocated)
    row2[3].metric("Full faculty", sum(f["is_full"] for f in faculty))

    if st.button("🔄 Refresh"):
        st.rerun()

    tab_faculty, tab_selections, tab_pending, tab_export = st.tabs(
        ["Faculty", "All selections", "Not submitted", "Export & tools"]
    )
    faculty_df = faculty_summary_df(faculty)
    with tab_faculty:
        st.dataframe(faculty_df, hide_index=True)
        full_names = [f["name"] for f in faculty if f["is_full"]]
        if full_names:
            st.warning("FULL: " + ", ".join(full_names))
        else:
            st.info("No faculty is full yet.")
    with tab_selections:
        if selections.empty:
            st.info("No selections have been submitted yet.")
        else:
            st.dataframe(selections, hide_index=True)
    with tab_pending:
        st.write(f"{len(pending)} student(s) have not submitted yet.")
        st.dataframe(pending, hide_index=True)
    with tab_export:
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button(
            "⬇ Download Excel (.xlsx)",
            data=build_excel(selections, faculty_df, pending),
            file_name=f"guide_allocations_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument."
                 "spreadsheetml.sheet",
        )
        st.download_button(
            "⬇ Download CSV (.csv)",
            data=selections.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"guide_allocations_{stamp}.csv",
            mime="text/csv",
        )
        with st.expander("Demo reset (deletes all selections)"):
            st.warning("This clears every submission. It cannot be undone.")
            confirm = st.checkbox("Yes, delete all selections")
            if st.button("Reset selections", disabled=not confirm):
                reset_selections()
                st.session_state.pop("confirmation", None)
                st.success("All selections cleared.")
                st.rerun()

    if st.sidebar.button("Log out"):
        st.session_state["admin_ok"] = False
        st.rerun()


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------
@st.cache_resource
def setup_database() -> bool:
    init_db()
    return True


def main() -> None:
    st.set_page_config(
        page_title="Guide Selection", page_icon="🎓", layout="wide"
    )
    st.markdown(CSS, unsafe_allow_html=True)
    try:
        setup_database()
        page = st.sidebar.radio("Go to", ["Student", "Admin"])
        if page == "Student":
            render_student_page()
        elif st.session_state.get("admin_ok"):
            render_admin_dashboard()
        else:
            render_admin_login()
    except sqlite3.Error as exc:
        st.error(f"Database error: {exc}. Please try again or contact staff.")


if __name__ == "__main__":
    main()
