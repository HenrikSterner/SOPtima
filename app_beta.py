"""SOPtima beta: delvis tidsplan og dynamisk manuel omplanlægning.

Kør med: streamlit run app_beta.py

Denne fil lader den eksisterende app være urørt. Den genbruger dens input,
fordeling og eksportfunktioner, men erstatter tidsplansberegningen med en
variant der beholder den del af planen, som kan nås på dagen.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime, time as dt_time, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any
import base64
import hashlib
import html
import json
import re
import zipfile
from urllib.parse import quote

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import app as core


def _original_core_function(name: str, beta_global_name: str) -> Any:
    """Find den oprindelige core-funktion, også ved hot-reload af Streamlit."""
    current = getattr(core, name)
    # Hvis en ældre app_beta allerede har patched core, ligger dens gemte
    # reference i beta-funktionens globals. Det gør opgradering uden manuel
    # procesgenstart sikker.
    candidate = getattr(current, "__globals__", {}).get(beta_global_name)
    if callable(candidate) and candidate is not current:
        return candidate
    return current


# Streamlit kan genkøre scriptet i samme Python-proces. Efter første kørsel er
# ``core.make_schedule`` derfor allerede beta-wrapperen. Gem de oprindelige
# funktioner på core-modulet, så en genkørsel aldrig bygger en beta-plan oven på
# en tidligere beta-plan (hvilket bl.a. gav gentagne Aktivitets-ID/Status-
# kolonner i Arrow-konverteringen).
if not hasattr(core, "_SOPTIMA_BETA_ORIGINALS"):
    core._SOPTIMA_BETA_ORIGINALS = {
        "make_schedule": _original_core_function("make_schedule", "_CORE_MAKE_SCHEDULE"),
        "make_schedule_excel": _original_core_function("make_schedule_excel", "_CORE_MAKE_SCHEDULE_EXCEL"),
        "make_schedule_html": _original_core_function("make_schedule_html", "_CORE_MAKE_SCHEDULE_HTML"),
        "make_student_schedule_html": _original_core_function("make_student_schedule_html", "_CORE_MAKE_STUDENT_HTML"),
        "parse_students": _original_core_function("parse_students", "_CORE_PARSE_STUDENTS"),
        "parse_teachers": _original_core_function("parse_teachers", "_CORE_PARSE_TEACHERS"),
        "parse_schedule_upload": _original_core_function("parse_schedule_upload", "_CORE_PARSE_SCHEDULE_UPLOAD"),
        "make_schedule_input_template": _original_core_function("make_schedule_input_template", "_CORE_SCHEDULE_TEMPLATE"),
    }
_CORE_ORIGINALS = core._SOPTIMA_BETA_ORIGINALS
_CORE_MAKE_SCHEDULE = _CORE_ORIGINALS["make_schedule"]
_CORE_MAKE_SCHEDULE_EXCEL = _CORE_ORIGINALS["make_schedule_excel"]
_CORE_MAKE_SCHEDULE_HTML = _CORE_ORIGINALS["make_schedule_html"]
_CORE_MAKE_STUDENT_HTML = _CORE_ORIGINALS["make_student_schedule_html"]
_CORE_PARSE_STUDENTS = _CORE_ORIGINALS["parse_students"]
_CORE_PARSE_TEACHERS = _CORE_ORIGINALS["parse_teachers"]
_CORE_PARSE_SCHEDULE_UPLOAD = _CORE_ORIGINALS["parse_schedule_upload"]
_CORE_SCHEDULE_TEMPLATE = _CORE_ORIGINALS["make_schedule_input_template"]

KNOWN_TEACHER_EMAILS = {
    "hst": "hst@nextkbh.dk",
    "matr": "matr@nextkbh.dk",
}

_COMPONENT_PATH = Path(__file__).with_name("beta_drag_component")
drag_timeline = components.declare_component("soptima_drag_timeline", path=str(_COMPONENT_PATH))


def _combine(value: dt_time) -> datetime:
    return datetime.combine(date.today(), value)


def _parse_clock(value: Any) -> datetime:
    return datetime.combine(date.today(), datetime.strptime(str(value), "%H:%M").time())


def _clock(value: datetime) -> str:
    return value.strftime("%H:%M")


def _unique_columns(columns: Any) -> list[Any]:
    """Returnér kolonnerne i samme rækkefølge uden dubletter.

    Pandas tillader dublerede labels, men Streamlit/pyarrow gør ikke. Det er
    især vigtigt ved Streamlit-genkørsler, hvor en allerede beta-beriget
    grundplan kan blive sendt ind igen.
    """
    result: list[Any] = []
    seen: set[tuple[type, str]] = set()
    for column in list(columns):
        marker = (type(column), repr(column))
        if marker in seen:
            continue
        seen.add(marker)
        result.append(column)
    return result


def _deduplicate_frame_columns(frame: pd.DataFrame) -> pd.DataFrame:
    """Fjern gentagne DataFrame-kolonner deterministisk (første vinder)."""
    if frame.empty and not frame.columns.has_duplicates:
        return frame
    return frame.loc[:, ~frame.columns.duplicated(keep="first")].copy()


_EMAIL_PATTERN = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")


def normalise_teams_email(value: Any) -> str:
    """Returnér en gyldig Microsoft 365-adresse eller en tom streng."""
    email = core.repair_text(value).strip().casefold()
    return email if _EMAIL_PATTERN.fullmatch(email) else ""


def _email_column(columns: list[Any], *prefixes: str) -> Any | None:
    candidates = []
    for column in columns:
        key = core.normal_key(column)
        if not any(token in key for token in ("mail", "email", "upn")):
            continue
        score = sum(prefix in key for prefix in prefixes)
        candidates.append((score, column))
    if not candidates:
        return None
    return max(candidates, key=lambda item: item[0])[1]


def parse_students_beta(table: pd.DataFrame) -> list[dict[str, Any]]:
    students = _CORE_PARSE_STUDENTS(table)
    columns = list(table.columns)
    name_col = core.column_for(columns, "elevnavn") or core.column_for(columns, "navn")
    class_col = core.column_for(columns, "klasse") or core.column_for(columns, "hold")
    email_col = _email_column(columns, "elev", "studer")
    rows_by_identity: dict[tuple[str, str], list[str]] = defaultdict(list)
    if name_col is not None and email_col is not None:
        for _, row in table.dropna(how="all").iterrows():
            name = core.value_at(row, name_col)
            if not name:
                continue
            identity = (core.normal_key(name), core.normal_key(core.value_at(row, class_col)))
            rows_by_identity[identity].append(normalise_teams_email(row.get(email_col, "")))
    for student in students:
        identity = (core.normal_key(student.get("name")), core.normal_key(student.get("className")))
        matches = rows_by_identity.get(identity, [])
        student["teams_email"] = matches.pop(0) if matches else normalise_teams_email(student.get("teams_email", ""))
    return students


def parse_teachers_beta(
    table: pd.DataFrame,
    technical: dict[str, str] | None = None,
) -> tuple[list[dict[str, Any]], dict[str, int]]:
    teachers, capacities = _CORE_PARSE_TEACHERS(table, technical)
    columns = list(table.columns)
    id_col = core.column_for(columns, "initial") or core.column_for(columns, "lærer")
    email_col = _email_column(columns, "lærer", "vejleder", "teacher")
    emails: dict[str, str] = {}
    if id_col is not None and email_col is not None:
        for _, row in table.dropna(how="all").iterrows():
            teacher_id = core.normal_key(row.get(id_col, ""))
            email = normalise_teams_email(row.get(email_col, ""))
            if teacher_id and email:
                emails.setdefault(teacher_id, email)
    for teacher in teachers:
        teacher["teams_email"] = emails.get(teacher["id"], KNOWN_TEACHER_EMAILS.get(teacher["id"], ""))
    return teachers, capacities


def parse_schedule_upload_beta(uploaded: Any) -> dict[str, Any]:
    """Udvid direkte tidsplan-upload med valgfri Teams-mailkolonner."""
    parsed = _CORE_PARSE_SCHEDULE_UPLOAD(uploaded)
    for teacher in parsed["teachers"]:
        teacher["teams_email"] = KNOWN_TEACHER_EMAILS.get(teacher["id"], "")
    for student in parsed["students"]:
        student["teams_email"] = ""
    try:
        table = core.read_uploaded_table(uploaded).dropna(how="all")
        columns = list(table.columns)
        name_col = core.column_for(columns, "elev") or core.column_for(columns, "navn")
        class_col = core.column_for(columns, "klasse") or core.column_for(columns, "hold")
        student_email_col = _email_column(columns, "elev", "studer")
        teacher_email_columns = [
            core.column_for(columns, "lærer", str(index), "mail")
            or core.column_for(columns, "vejleder", str(index), "mail")
            or core.column_for(columns, "teacher", str(index), "email")
            for index in (1, 2)
        ]
        teacher_columns = [
            core.column_for(columns, "lærer", str(index)) or core.column_for(columns, "vejleder", str(index))
            for index in (1, 2)
        ]
        student_lookup = {
            (core.normal_key(student["name"]), core.normal_key(student.get("className"))): student
            for student in parsed["students"]
        }
        teacher_lookup = {teacher["id"]: teacher for teacher in parsed["teachers"]}
        for _, row in table.iterrows():
            if name_col is not None and student_email_col is not None:
                identity = (core.normal_key(row.get(name_col, "")), core.normal_key(row.get(class_col, "")))
                if identity in student_lookup:
                    student_lookup[identity]["teams_email"] = normalise_teams_email(row.get(student_email_col, ""))
            for teacher_col, teacher_email_col in zip(teacher_columns, teacher_email_columns):
                if teacher_col is None or teacher_email_col is None:
                    continue
                teacher_id, _ = core.teacher_identity(row.get(teacher_col, ""))
                email = normalise_teams_email(row.get(teacher_email_col, ""))
                if teacher_id in teacher_lookup and email:
                    teacher_lookup[teacher_id]["teams_email"] = email
    except (ImportError, OSError, ValueError, TypeError):
        # Grundparserens resultat er stadig anvendeligt; manglende adresser
        # markeres senere som en deaktiveret Teams-handling.
        pass
    return parsed


def make_schedule_input_template_beta() -> bytes:
    columns = [
        "Elev", "Elev Teams-email", "Klasse", "Lærer 1", "Lærer 1 Teams-email", "Fag 1",
        "Lærer 2", "Lærer 2 Teams-email", "Fag 2", "", "Bemærkninger/ændringer fra lærerne",
    ]
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(columns=columns).to_excel(writer, index=False, sheet_name="Tidsplan-input")
        sheet = writer.book["Tidsplan-input"]
        sheet.freeze_panes = "A2"
        widths = (24, 30, 14, 24, 30, 22, 24, 30, 22, 4, 45)
        for index, width in enumerate(widths, 1):
            sheet.column_dimensions[sheet.cell(1, index).column_letter].width = width
    return output.getvalue()


def stable_activity_id(student: dict[str, Any], index: int, used: set[str]) -> str:
    raw_identity = core.repair_text(student.get("id"))
    if raw_identity:
        token = re.sub(r"[^a-zA-Z0-9_-]+", "-", raw_identity).strip("-")
    else:
        fingerprint = "|".join((
            core.repair_text(student.get("name")), core.repair_text(student.get("className")),
            core.repair_text(student.get("projectTitle")),
        ))
        token = hashlib.sha256(fingerprint.encode("utf-8")).hexdigest()[:10]
    candidate = f"A-{token or index + 1}"
    suffix = 2
    while candidate.casefold() in used:
        candidate = f"A-{token or index + 1}-{suffix}"
        suffix += 1
    used.add(candidate.casefold())
    return candidate


def _overlaps(start: datetime, end: datetime, other_start: datetime, other_end: datetime) -> bool:
    return start < other_end and end > other_start


def _source_key(student: dict[str, Any], assigned: list[str | None], teacher_map: dict[str, dict[str, Any]]) -> tuple[str, ...]:
    return (
        str(student.get("name", "")),
        str(student.get("className", "")),
        str(student.get("projectTitle", "")),
        core.teacher_label(assigned[0], teacher_map),
        core.teacher_label(assigned[1], teacher_map),
    )


def _row_key(row: pd.Series) -> tuple[str, ...]:
    return (
        str(row.get("Elev", "")),
        str(row.get("Klasse", "")),
        str(row.get("Projekttitel", "")),
        str(row.get("Vejleder 1", "")),
        str(row.get("Vejleder 2", "")),
    )


def _activity_records(
    base: dict[str, Any],
    students: list[dict[str, Any]],
    teachers: list[dict[str, Any]],
    solution: dict[str, Any],
) -> list[dict[str, Any]]:
    """Knyt den oprindelige elevrækkefølge til rækkerne i grundplanen."""
    teacher_map = {teacher["id"]: teacher for teacher in teachers}
    available: dict[tuple[str, ...], list[pd.Series]] = defaultdict(list)
    for _, row in _deduplicate_frame_columns(base["students"]).iterrows():
        available[_row_key(row)].append(row)

    activities: list[dict[str, Any]] = []
    used_ids: set[str] = set()
    for index, student in enumerate(students):
        assigned = list(solution["assignments"][index])
        key = _source_key(student, assigned, teacher_map)
        matches = available.get(key, [])
        if not matches:
            raise ValueError(f"Beta-planen kan ikke finde vejledningen for {student.get('name', 'eleven')}.")
        row = matches.pop(0).copy()
        activities.append({
            "id": stable_activity_id(student, index, used_ids),
            "student": student,
            "assigned": tuple(str(item) for item in assigned if item),
            "pair": tuple(sorted(set(str(item) for item in assigned if item))),
            "row": row,
            "base_start": _parse_clock(row["Start"]),
        })
    return activities


def _global_breaks(base: dict[str, Any], plan_start: datetime, plan_end: datetime) -> list[tuple[datetime, datetime]]:
    intervals: list[tuple[datetime, datetime]] = []
    for _, row in base["timeline"].iterrows():
        if str(row.get("Type", "")) not in {"Pause", "Frokostpause"}:
            continue
        start, end = _parse_clock(row["Start"]), _parse_clock(row["Slut"])
        if _overlaps(start, end, plan_start, plan_end):
            intervals.append((max(start, plan_start), min(end, plan_end)))
    return intervals


def _find_first_slot(
    activity: dict[str, Any],
    duration: int,
    plan_start: datetime,
    plan_end: datetime,
    placed: list[dict[str, Any]],
    breaks: list[tuple[datetime, datetime]],
    teacher_blocks: dict[str, list[tuple[dt_time, dt_time]]],
    transition_minutes: int = 0,
    preferred: datetime | None = None,
    allow_fallback: bool = True,
) -> datetime | None:
    """Find første lovlige femminuttersfelt for aktiviteten."""

    def is_legal(start: datetime) -> bool:
        end = start + timedelta(minutes=duration)
        if start < plan_start or end > plan_end:
            return False
        if any(_overlaps(start, end, pause_start, pause_end) for pause_start, pause_end in breaks):
            return False
        for teacher_id in activity["assigned"]:
            for block_start, block_end in teacher_blocks.get(teacher_id, []):
                if _overlaps(start, end, _combine(block_start), _combine(block_end)):
                    return False
        for other in placed:
            if set(activity["assigned"]) & set(other["activity"]["assigned"]):
                transition = timedelta(minutes=transition_minutes)
                if _overlaps(start, end + transition, other["start"], other["end"] + transition):
                    return False
        return True

    if preferred is not None:
        if is_legal(preferred):
            return preferred
        if not allow_fallback:
            return None
    candidate = plan_start
    while candidate + timedelta(minutes=duration) <= plan_end:
        if is_legal(candidate):
            return candidate
        candidate += timedelta(minutes=5)
    return None


def _build_beta_schedule(
    base: dict[str, Any],
    students: list[dict[str, Any]],
    teachers: list[dict[str, Any]],
    solution: dict[str, Any],
    start_time: dt_time,
    end_time: dt_time,
    student_minutes: int,
    transition_minutes: int,
    teacher_blocks_raw: dict[str, Any] | None,
    locked_activities: dict[str, str],
    version: int = 1,
) -> dict[str, Any]:
    plan_start, plan_end = _combine(start_time), _combine(end_time)
    teacher_map = {teacher["id"]: teacher for teacher in teachers}
    activities = _activity_records(base, students, teachers, solution)
    teacher_blocks = core.normalise_teacher_blocks(teacher_blocks_raw, teachers)
    breaks = _global_breaks(base, plan_start, plan_end)
    placed: list[dict[str, Any]] = []
    unplaced: list[dict[str, Any]] = []

    # Manuelle flytninger placeres først og bliver dermed låste præmisser for
    # den lokale genberegning. Resten forsøger først at beholde sin gamle tid.
    ordered = sorted(
        activities,
        key=lambda activity: (activity["id"] not in locked_activities, activity["base_start"], activity["id"]),
    )
    for activity in ordered:
        requested = None
        if activity["id"] in locked_activities:
            requested = _parse_clock(locked_activities[activity["id"]])
        # Kun manuelt låste flytninger har en fast ønsket tid. Alle øvrige
        # aktiviteter finder det tidligst mulige lovlige tidspunkt.
        preferred = requested
        slot = _find_first_slot(
            activity, student_minutes, plan_start, plan_end, placed, breaks, teacher_blocks,
            transition_minutes, preferred,
            allow_fallback=requested is None,
        )
        if slot is None:
            if requested is not None:
                reason = "Den manuelt valgte tid kan ikke rummes uden konflikt eller falder uden for dagen."
            elif activity["base_start"] >= plan_end:
                reason = "Falder efter den valgte sluttid."
            else:
                reason = "Kan ikke placeres uden konflikt med vejledere, pauser eller spærringer."
            unplaced.append({
                "activity": activity,
                "reason": reason,
                "status": "konflikt" if requested is not None else "uplaceret",
                "locked": requested is not None,
            })
            continue
        placed.append({
            "activity": activity,
            "start": slot,
            "end": slot + timedelta(minutes=student_minutes),
            "locked": activity["id"] in locked_activities,
        })

    placed.sort(key=lambda item: (item["start"], item["activity"]["id"]))
    student_rows: list[dict[str, Any]] = []
    teacher_rows: list[dict[str, Any]] = []
    timeline_rows: list[dict[str, Any]] = []
    pair_counts: Counter[tuple[str, ...]] = Counter()
    pair_ranges: dict[tuple[str, ...], list[str]] = {}
    for item in placed:
        activity, start, end = item["activity"], item["start"], item["end"]
        student, assigned = activity["student"], activity["assigned"]
        row = dict(activity["row"])
        row.update({
            "Aktivitets-ID": activity["id"],
            "Status": "Låst" if item["locked"] else "Planlagt",
            "Start": _clock(start), "Slut": _clock(end), "Tid": f"{_clock(start)}–{_clock(end)}",
        })
        student_rows.append(row)
        pair_label = core._schedule_pair_label(activity["pair"], teacher_map)
        timeline_rows.append({
            "Type": "Vejledning", "Start": _clock(start), "Slut": _clock(end),
            "Varighed (min.)": student_minutes, "Elev": student["name"],
            "Klasse": student.get("className", ""), "Lærer(e)": pair_label,
            "Information": student.get("projectTitle", "") or "SOP-vejledning",
            "Bemærkning": student.get("scheduleNotes", ""),
            "Aktivitets-ID": activity["id"], "Status": "Låst" if item["locked"] else "Planlagt",
        })
        pair_counts[activity["pair"]] += 1
        time_label = row["Tid"]
        pair_ranges.setdefault(activity["pair"], [time_label, time_label])[1] = time_label
        for teacher_id in dict.fromkeys(assigned):
            subjects = [
                subject for subject_index, subject in enumerate(student.get("subjects", [])[:2])
                if subject_index < len(assigned) and assigned[subject_index] == teacher_id
            ]
            co_teachers = [other for other in dict.fromkeys(assigned) if other != teacher_id]
            teacher_rows.append({
                "Type": "Vejledning", "Start": _clock(start), "Slut": _clock(end), "Tid": time_label,
                "Lærer": core.teacher_label(teacher_id, teacher_map), "Initialer": teacher_id,
                "Elev": student["name"], "Klasse": student.get("className", ""),
                "Fag": " · ".join(dict.fromkeys(subjects)),
                "Medvejleder": " + ".join(core.teacher_label(other, teacher_map) for other in co_teachers) or "—",
                "Projekttitel": student.get("projectTitle", ""),
                "Bemærkning": student.get("scheduleNotes", ""),
                "Aktivitets-ID": activity["id"], "Status": "Låst" if item["locked"] else "Planlagt",
            })

    # Bevar pauser, frokost og spærringer fra grundplanen, men kun når de
    # berører den valgte dag. Den visuelle oversigt viser fortsat kun vejledning.
    def break_is_between_guidance(row_type: str, row_start: datetime, row_end: datetime) -> bool:
        if row_type not in {"Pause", "Frokostpause"}:
            return True
        return (
            any(item["end"] <= row_start for item in placed)
            and any(item["start"] >= row_end for item in placed)
        )

    for _, row in base["teachers"].iterrows():
        if str(row.get("Type", "")) == "Vejledning":
            continue
        try:
            row_start, row_end = _parse_clock(row["Start"]), _parse_clock(row["Slut"])
        except (TypeError, ValueError):
            continue
        if _overlaps(row_start, row_end, plan_start, plan_end) and break_is_between_guidance(str(row.get("Type", "")), row_start, row_end):
            teacher_rows.append(dict(row))
    for _, row in base["timeline"].iterrows():
        if str(row.get("Type", "")) == "Vejledning":
            continue
        try:
            row_start, row_end = _parse_clock(row["Start"]), _parse_clock(row["Slut"])
        except (TypeError, ValueError):
            continue
        if _overlaps(row_start, row_end, plan_start, plan_end) and break_is_between_guidance(str(row.get("Type", "")), row_start, row_end):
            timeline_rows.append(dict(row))

    unplaced_rows: list[dict[str, Any]] = []
    for item in unplaced:
        activity, student = item["activity"], item["activity"]["student"]
        assigned = activity["assigned"]
        unplaced_rows.append({
            "Aktivitets-ID": activity["id"],
            "Status": "Konflikt" if item["status"] == "konflikt" else "Uplaceret",
            "Årsag": item["reason"],
            "Elev": student["name"], "Klasse": student.get("className", ""),
            "Fag 1": student.get("subjects", ["", ""])[0],
            "Vejleder 1": core.teacher_label(assigned[0], teacher_map) if assigned else "",
            "Fag 2": student.get("subjects", ["", ""])[1] if len(student.get("subjects", [])) > 1 else "",
            "Vejleder 2": core.teacher_label(assigned[1], teacher_map) if len(assigned) > 1 else "",
            "Lærerpar": core._schedule_pair_label(activity["pair"], teacher_map),
            "Projekttitel": student.get("projectTitle", ""), "Varighed (min.)": student_minutes,
        })

    # Grundplanen kan komme fra en tidligere Streamlit-genkørsel. Deduplikér
    # både indgående og nye felter før DataFrame-oprettelse, ellers afviser
    # pyarrow planen som ugyldig ved ``st.dataframe``.
    base_students = _deduplicate_frame_columns(base["students"])
    base_teachers = _deduplicate_frame_columns(base["teachers"])
    base_timeline = _deduplicate_frame_columns(base["timeline"])
    student_columns = _unique_columns(list(base_students.columns) + ["Aktivitets-ID", "Status"])
    teacher_columns = _unique_columns(list(base_teachers.columns) + ["Aktivitets-ID", "Status"])
    timeline_columns = _unique_columns(list(base_timeline.columns) + ["Aktivitets-ID", "Status"])
    unplaced_columns = [
        "Aktivitets-ID", "Status", "Årsag", "Elev", "Klasse", "Fag 1", "Vejleder 1",
        "Fag 2", "Vejleder 2", "Lærerpar", "Projekttitel", "Varighed (min.)",
    ]
    student_frame = _deduplicate_frame_columns(pd.DataFrame(student_rows, columns=student_columns)).sort_values("Start", kind="stable")
    teacher_frame = _deduplicate_frame_columns(pd.DataFrame(teacher_rows, columns=teacher_columns)).sort_values(["Start", "Lærer"], kind="stable")
    timeline_frame = _deduplicate_frame_columns(pd.DataFrame(timeline_rows, columns=timeline_columns)).sort_values("Start", kind="stable")
    pair_rows = [
        {"Lærerpar": core._schedule_pair_label(pair, teacher_map), "Antal elever": count,
         "Tidsblok": " → ".join(pair_ranges[pair])}
        for pair, count in sorted(pair_counts.items(), key=lambda item: pair_ranges[item[0]][0])
    ]
    settings = dict(base["settings"])
    last_end = max((item["end"] for item in placed), default=plan_start)
    settings.update({
        "Start": _clock(plan_start),
        "Slut": _clock(plan_end),
        "Planlagte aktiviteter": len(placed),
        "Uplacerede aktiviteter": len(unplaced),
        "Manuelt låste aktiviteter": len(locked_activities),
        "Planlagt tidsforbrug (min.)": max(0, int((last_end - plan_start).total_seconds() // 60)),
    })
    activity_rows: list[dict[str, Any]] = []
    metadata: dict[str, dict[str, Any]] = {}
    placed_by_id = {item["activity"]["id"]: item for item in placed}
    unplaced_by_id = {item["activity"]["id"]: item for item in unplaced}
    for activity in activities:
        activity_id = activity["id"]
        placement = placed_by_id.get(activity_id)
        rejected = unplaced_by_id.get(activity_id)
        student = activity["student"]
        status = "låst" if placement and placement["locked"] else "planlagt" if placement else rejected["status"]
        start_label = _clock(placement["start"]) if placement else ""
        end_label = _clock(placement["end"]) if placement else ""
        reason = rejected["reason"] if rejected else ""
        activity_rows.append({
            "activity_id": activity_id,
            "elev_id": core.repair_text(student.get("id")),
            "lærer_ids": ",".join(activity["assigned"]),
            "varighed": student_minutes,
            "start": start_label,
            "slut": end_label,
            "status": status,
            "låst": bool(placement and placement["locked"]),
            "årsag": reason,
        })
        metadata[activity_id] = {
            "student": student,
            "student_email": normalise_teams_email(student.get("teams_email") or student.get("email")),
            "teacher_ids": list(activity["assigned"]),
            "teacher_emails": [
                normalise_teams_email(
                    teacher_map.get(teacher_id, {}).get("teams_email")
                    or teacher_map.get(teacher_id, {}).get("email")
                    or KNOWN_TEACHER_EMAILS.get(teacher_id, "")
                )
                for teacher_id in activity["assigned"]
            ],
            "subjects": list(student.get("subjects", []))[:2],
            "teacher_names": [core.teacher_label(teacher_id, teacher_map) for teacher_id in activity["assigned"]],
            "class_name": student.get("className", ""),
            "student_name": student.get("name", ""),
        }
    return {
        "students": student_frame.reset_index(drop=True),
        "teachers": teacher_frame.reset_index(drop=True),
        "timeline": timeline_frame.reset_index(drop=True),
        "pairs": _deduplicate_frame_columns(pd.DataFrame(pair_rows, columns=_unique_columns(base["pairs"].columns))),
        "settings": settings,
        "unplaced": pd.DataFrame(unplaced_rows, columns=unplaced_columns),
        "activities": pd.DataFrame(activity_rows),
        "activity_metadata": metadata,
        "beta": {
            "plan_start": _clock(plan_start), "plan_end": _clock(plan_end),
            "locked": dict(locked_activities), "version": version,
            "transition_minutes": transition_minutes,
            "breaks": [(_clock(start), _clock(end)) for start, end in breaks],
            "teacher_blocks": {
                teacher_id: [(start.strftime("%H:%M"), end.strftime("%H:%M")) for start, end in intervals]
                for teacher_id, intervals in teacher_blocks.items() if intervals
            },
        },
    }


def _fallback_base_schedule(
    students: list[dict[str, Any]], teachers: list[dict[str, Any]], solution: dict[str, Any],
    start_time: dt_time, end_time: dt_time, student_minutes: int, pause_count: int,
    pause_minutes: int, lunch_mode: str, lunch_start_time: dt_time, lunch_minutes: int,
    teacher_blocks: dict[str, Any] | None,
) -> dict[str, Any]:
    """Lav et grundlag, når selv en udvidet enkeltdagsplan ikke kan rumme alle."""
    teacher_map = {teacher["id"]: teacher for teacher in teachers}
    plan_start, plan_end = _combine(start_time), _combine(end_time)
    student_rows = []
    for index, student in enumerate(students):
        assigned = solution["assignments"][index]
        pair = tuple(sorted(set(assigned)))
        student_rows.append({
            "Start": _clock(plan_start), "Slut": _clock(plan_start + timedelta(minutes=student_minutes)),
            "Tid": f"{_clock(plan_start)}–{_clock(plan_start + timedelta(minutes=student_minutes))}",
            "Elev": student["name"], "Klasse": student.get("className", ""),
            "Fag 1": student.get("subjects", ["", ""])[0],
            "Vejleder 1": core.teacher_label(assigned[0], teacher_map),
            "Fag 2": student.get("subjects", ["", ""])[1] if len(student.get("subjects", [])) > 1 else "",
            "Vejleder 2": core.teacher_label(assigned[1], teacher_map),
            "Lærerpar": core._schedule_pair_label(pair, teacher_map),
            "Projekttitel": student.get("projectTitle", ""),
            "Bemærkninger/ændringer fra lærerne": student.get("scheduleNotes", ""),
        })
    teacher_columns = [
        "Type", "Start", "Slut", "Tid", "Lærer", "Initialer", "Elev", "Klasse", "Fag",
        "Medvejleder", "Projekttitel", "Bemærkning",
    ]
    timeline_columns = [
        "Type", "Start", "Slut", "Varighed (min.)", "Elev", "Klasse", "Lærer(e)",
        "Information", "Bemærkning",
    ]
    teacher_rows: list[dict[str, Any]] = []
    timeline_rows: list[dict[str, Any]] = []

    if lunch_mode == "Fast tidspunkt for alle lærere":
        lunch_start = _combine(lunch_start_time)
    else:
        midpoint = plan_start + (plan_end - plan_start) / 2
        rounded = int((midpoint - plan_start).total_seconds() // 300) * 5
        lunch_start = plan_start + timedelta(minutes=rounded)
    lunch_start = min(max(lunch_start, plan_start), plan_end - timedelta(minutes=lunch_minutes))
    lunch_end = lunch_start + timedelta(minutes=lunch_minutes)
    intervals: list[tuple[str, datetime, datetime]] = [("Frokostpause", lunch_start, lunch_end)]
    segments = [(plan_start, lunch_start), (lunch_end, plan_end)]
    for pause_index in range(max(0, pause_count)):
        segment_start, segment_end = segments[pause_index % len(segments)]
        candidate = segment_start + (segment_end - segment_start) / 2
        candidate = plan_start + timedelta(minutes=int((candidate - plan_start).total_seconds() // 300) * 5)
        pause_end = candidate + timedelta(minutes=pause_minutes)
        if pause_minutes and candidate > plan_start and pause_end < plan_end and not _overlaps(candidate, pause_end, lunch_start, lunch_end):
            intervals.append(("Pause", candidate, pause_end))
    for interval_type, interval_start, interval_end in intervals:
        timeline_rows.append({
            "Type": interval_type, "Start": _clock(interval_start), "Slut": _clock(interval_end),
            "Varighed (min.)": int((interval_end - interval_start).total_seconds() // 60),
            "Elev": "", "Klasse": "", "Lærer(e)": "Alle lærere" if interval_type == "Frokostpause" else "",
            "Information": interval_type, "Bemærkning": "",
        })
        for teacher in teachers:
            teacher_rows.append({
                "Type": interval_type, "Start": _clock(interval_start), "Slut": _clock(interval_end),
                "Tid": f"{_clock(interval_start)}–{_clock(interval_end)}",
                "Lærer": core.teacher_label(teacher["id"], teacher_map), "Initialer": teacher["id"],
                "Elev": "", "Klasse": "", "Fag": interval_type, "Medvejleder": "—",
                "Projekttitel": "", "Bemærkning": "Fælles planlagt pause",
            })
    for teacher_id, blocked in core.normalise_teacher_blocks(teacher_blocks, teachers).items():
        for block_start, block_end in blocked:
            teacher_rows.append({
                "Type": "Spærret", "Start": block_start.strftime("%H:%M"), "Slut": block_end.strftime("%H:%M"),
                "Tid": f"{block_start.strftime('%H:%M')}–{block_end.strftime('%H:%M')}",
                "Lærer": core.teacher_label(teacher_id, teacher_map), "Initialer": teacher_id,
                "Elev": "", "Klasse": "", "Fag": "Ikke tilgængelig", "Medvejleder": "—",
                "Projekttitel": "", "Bemærkning": "Lærerens angivne spærring",
            })
    student_columns = [
        "Start", "Slut", "Tid", "Elev", "Klasse", "Fag 1", "Vejleder 1", "Fag 2",
        "Vejleder 2", "Lærerpar", "Projekttitel", "Bemærkninger/ændringer fra lærerne",
    ]
    return {
        "students": pd.DataFrame(student_rows, columns=student_columns),
        "teachers": pd.DataFrame(teacher_rows, columns=teacher_columns),
        "timeline": pd.DataFrame(timeline_rows, columns=timeline_columns),
        "pairs": pd.DataFrame(columns=["Lærerpar", "Antal elever", "Tidsblok"]),
        "settings": {
            "Start": start_time.strftime("%H:%M"), "Slut": end_time.strftime("%H:%M"),
            "Minutter pr. elev": student_minutes, "Antal pauser": pause_count,
            "Minutter pr. pause": pause_minutes, "Frokostpause (min.)": lunch_minutes,
            "Frokostpause": lunch_mode, "Planlagt tidsforbrug (min.)": 0,
            "Forsøg at undgå lærerhuller": "Nej", "Lærerhuller (runder)": 0,
        },
    }


def make_schedule_beta(
    students: list[dict[str, Any]], teachers: list[dict[str, Any]], solution: dict[str, Any],
    start_time: dt_time, end_time: dt_time, student_minutes: int, pause_count: int,
    pause_minutes: int, transition_minutes: int, group_pairs: bool,
    lunch_mode: str = "Fast tidspunkt for alle lærere", lunch_start_time: dt_time = dt_time(12, 0),
    lunch_minutes: int = 30, search_attempts: int = 80, progress_callback: Any = None,
    avoid_teacher_gaps: bool = False, teacher_blocks: dict[str, Any] | None = None,
    floating_pauses: bool = False, locked_activities: dict[str, str] | None = None,
    expected_version: int | None = None,
) -> dict[str, Any]:
    """Generér en delvis plan; resten returneres som ``unplaced``."""
    if start_time >= end_time:
        raise ValueError("Sluttidspunktet skal ligge efter starttidspunktet.")
    if lunch_mode == "Fast tidspunkt for alle lærere":
        if not (_combine(start_time) <= _combine(lunch_start_time) and _combine(lunch_start_time) + timedelta(minutes=lunch_minutes) <= _combine(end_time)):
            raise ValueError("Den faste frokostpause skal ligge inden for tidsrummet og kunne rumme hele pausen.")
    # Grundplanen beregnes til sent på dagen, så vi får alle aktiviteter og
    # derefter kan beholde den del, som reelt kan ligge inden for valgt sluttid.
    extended_end = dt_time(23, 55)
    if end_time >= extended_end:
        extended_end = end_time
    try:
        base = _CORE_MAKE_SCHEDULE(
            students, teachers, solution, start_time, extended_end, student_minutes, pause_count,
            pause_minutes, transition_minutes, group_pairs, lunch_mode=lunch_mode,
            lunch_start_time=lunch_start_time, lunch_minutes=lunch_minutes,
            search_attempts=search_attempts, progress_callback=progress_callback,
            avoid_teacher_gaps=avoid_teacher_gaps, teacher_blocks=teacher_blocks,
            floating_pauses=floating_pauses,
        )
    except ValueError as error:
        if "Tidsplanen kan ikke afsluttes" not in str(error):
            raise
        base = _fallback_base_schedule(
            students, teachers, solution, start_time, end_time, student_minutes, pause_count,
            pause_minutes, lunch_mode, lunch_start_time, lunch_minutes, teacher_blocks,
        )
        if progress_callback:
            progress_callback(1.0)
    context_marker = core.stable_signature([
        [(student.get("id", ""), student.get("name", ""), student.get("className", "")) for student in students],
        solution.get("assignments", []), start_time, end_time, student_minutes, pause_count,
        pause_minutes, transition_minutes, lunch_mode, lunch_start_time, lunch_minutes,
        teacher_blocks, floating_pauses,
    ])
    context_changed = st.session_state.get("beta_context_marker") != context_marker
    current_schedule = st.session_state.get("v2_schedule")
    current_version = 0
    if not context_changed and isinstance(current_schedule, dict):
        current_version = int(current_schedule.get("beta", {}).get("version", 0))
    if expected_version is not None and current_version not in {0, expected_version}:
        raise ValueError(
            "Tidsplanen er ændret siden din seneste visning. Genindlæs planen, før flytningen gentages."
        )
    if context_changed:
        st.session_state["beta_context_marker"] = context_marker
        st.session_state["beta_locked_activities"] = {}
        st.session_state["beta_undo_stack"] = []
        st.session_state["beta_redo_stack"] = []
        st.session_state["beta_change_log"] = []
    locks = dict(locked_activities if locked_activities is not None else st.session_state.get("beta_locked_activities", {}))
    result = _build_beta_schedule(
        base, students, teachers, solution, start_time, end_time, student_minutes,
        transition_minutes, teacher_blocks, locks, version=current_version + 1
    )
    st.session_state["beta_schedule_context"] = {
        "students": students, "teachers": teachers, "solution": solution,
        "start_time": start_time, "end_time": end_time, "student_minutes": student_minutes,
        "pause_count": pause_count, "pause_minutes": pause_minutes,
        "transition_minutes": transition_minutes, "group_pairs": group_pairs,
        "lunch_mode": lunch_mode, "lunch_start_time": lunch_start_time,
        "lunch_minutes": lunch_minutes, "search_attempts": search_attempts,
        "avoid_teacher_gaps": avoid_teacher_gaps, "teacher_blocks": teacher_blocks,
        "floating_pauses": floating_pauses,
    }
    return result


def assess_move(schedule: dict[str, Any], activity_id: str, target_start: str) -> dict[str, str]:
    """Klassificér en manuel flytning som grøn, gul eller rød."""
    activities = schedule.get("activities", pd.DataFrame())
    metadata = schedule.get("activity_metadata", {})
    matches = activities[activities["activity_id"] == activity_id] if not activities.empty else pd.DataFrame()
    if matches.empty or activity_id not in metadata:
        return {"status": "red", "reason": "Aktiviteten findes ikke i den aktuelle plan."}
    try:
        target = _parse_clock(target_start)
    except (TypeError, ValueError):
        return {"status": "red", "reason": "Starttidspunktet er ugyldigt."}
    if target.minute % 5:
        return {"status": "red", "reason": "Starttidspunktet skal ligge på et femminuttersinterval."}
    duration = int(matches.iloc[0]["varighed"])
    target_end = target + timedelta(minutes=duration)
    plan_start = _parse_clock(schedule["beta"]["plan_start"])
    plan_end = _parse_clock(schedule["beta"]["plan_end"])
    if target < plan_start or target_end > plan_end:
        return {"status": "red", "reason": "Aktiviteten falder uden for dagens tidsrum."}
    for break_start, break_end in schedule["beta"].get("breaks", []):
        if _overlaps(target, target_end, _parse_clock(break_start), _parse_clock(break_end)):
            return {"status": "red", "reason": "Aktiviteten overlapper en pause eller frokost."}
    teacher_ids = set(metadata[activity_id]["teacher_ids"])
    for teacher_id in teacher_ids:
        for block_start, block_end in schedule["beta"].get("teacher_blocks", {}).get(teacher_id, []):
            if _overlaps(target, target_end, _parse_clock(block_start), _parse_clock(block_end)):
                return {"status": "red", "reason": f"Tidspunktet rammer en spærring for {teacher_id}."}
    conflicts = []
    locked_conflicts = []
    transition = timedelta(minutes=int(schedule["beta"].get("transition_minutes", 0)))
    for _, other in activities.iterrows():
        other_id = str(other["activity_id"])
        if other_id == activity_id or not core.repair_text(other.get("start")):
            continue
        other_teachers = set(metadata.get(other_id, {}).get("teacher_ids", []))
        if not teacher_ids & other_teachers:
            continue
        other_start, other_end = _parse_clock(other["start"]), _parse_clock(other["slut"])
        if _overlaps(target, target_end + transition, other_start, other_end + transition):
            conflicts.append(other_id)
            if bool(other.get("låst")):
                locked_conflicts.append(other_id)
    if locked_conflicts:
        return {
            "status": "red",
            "reason": "Tidspunktet kolliderer med en anden manuelt låst aktivitet: " + ", ".join(locked_conflicts),
        }
    if conflicts:
        return {
            "status": "yellow",
            "reason": "Flytningen er mulig, men disse aktiviteter skal genplaceres: " + ", ".join(conflicts),
        }
    return {"status": "green", "reason": "Placeringen er konfliktfri."}


def move_validity_map(schedule: dict[str, Any]) -> dict[str, dict[str, str]]:
    """Lav farvekortet til træk-og-slip-komponenten."""
    start, end = _parse_clock(schedule["beta"]["plan_start"]), _parse_clock(schedule["beta"]["plan_end"])
    result: dict[str, dict[str, str]] = {}
    activity_ids = schedule.get("activities", pd.DataFrame()).get("activity_id", pd.Series(dtype=str)).tolist()
    candidate = start
    choices = []
    while candidate < end:
        choices.append(_clock(candidate))
        candidate += timedelta(minutes=5)
    for activity_id in activity_ids:
        result[str(activity_id)] = {
            choice: assess_move(schedule, str(activity_id), choice)["status"] for choice in choices
        }
    return result


def apply_manual_move(
    schedule: dict[str, Any],
    context: dict[str, Any],
    locks: dict[str, str],
    activity_id: str,
    target_start: str,
    mode: str = "dynamic",
) -> tuple[dict[str, Any], dict[str, str], dict[str, str]]:
    """Validér en flytning og returnér ny plan, nye låse og vurderingen."""
    assessment = assess_move(schedule, activity_id, target_start)
    if assessment["status"] == "red":
        raise ValueError(assessment["reason"])
    if mode == "strict" and assessment["status"] == "yellow":
        raise ValueError("Streng tilstand flytter ikke andre aktiviteter. " + assessment["reason"])
    if mode not in {"strict", "dynamic"}:
        raise ValueError("Flyttetilstanden skal være 'strict' eller 'dynamic'.")
    updated_locks = dict(locks)
    updated_locks[activity_id] = target_start
    updated = make_schedule_beta(
        **context,
        locked_activities=updated_locks,
        expected_version=int(schedule["beta"]["version"]),
    )
    return updated, updated_locks, assessment


def push_lock_history(state: Any, current_locks: dict[str, str]) -> None:
    state.setdefault("beta_undo_stack", []).append(dict(current_locks))
    state["beta_redo_stack"] = []


def undo_lock_history(state: Any, current_locks: dict[str, str]) -> dict[str, str] | None:
    history = state.setdefault("beta_undo_stack", [])
    if not history:
        return None
    target = dict(history.pop())
    state.setdefault("beta_redo_stack", []).append(dict(current_locks))
    return target


def redo_lock_history(state: Any, current_locks: dict[str, str]) -> dict[str, str] | None:
    redo = state.setdefault("beta_redo_stack", [])
    if not redo:
        return None
    target = dict(redo.pop())
    state.setdefault("beta_undo_stack", []).append(dict(current_locks))
    return target


def teams_chat_link(schedule: dict[str, Any], activity_id: str) -> tuple[str, str]:
    """Returnér Teams-link og eventuel fejltekst for en planlagt aktivitet."""
    metadata = schedule.get("activity_metadata", {}).get(activity_id)
    activities = schedule.get("activities", pd.DataFrame())
    matches = activities[activities["activity_id"] == activity_id] if not activities.empty else pd.DataFrame()
    if not metadata or matches.empty or not core.repair_text(matches.iloc[0].get("start")):
        return "", "Aktiviteten har ikke et bekræftet tidspunkt."
    missing = []
    student_email = normalise_teams_email(metadata.get("student_email"))
    if not student_email:
        missing.append(metadata.get("student_name") or "eleven")
    teacher_emails = []
    for teacher_id, teacher_name, email in zip(
        metadata.get("teacher_ids", []), metadata.get("teacher_names", []), metadata.get("teacher_emails", [])
    ):
        valid = normalise_teams_email(email)
        if not valid:
            missing.append(teacher_name or teacher_id)
        else:
            teacher_emails.append(valid)
    if missing:
        return "", "Teams-link kan ikke oprettes: mangler Teams-adresse for " + ", ".join(missing) + "."
    users = list(dict.fromkeys([*teacher_emails, student_email]))
    encoded_users = ",".join(quote(email, safe="") for email in users)
    topic = quote(
        f"SOP - {metadata.get('student_name', '')} - {metadata.get('class_name', '')}",
        safe="",
    )
    return f"https://teams.microsoft.com/l/chat/0/0?users={encoded_users}&topicName={topic}", ""


def _safe_docx_name(value: Any) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]', "-", core.repair_text(value))
    return re.sub(r"\s+", " ", cleaned).strip(" .") or "Elev"


def _word_documents(schedule: dict[str, Any], progress_callback: Any = None) -> dict[str, dict[str, Any]]:
    """Generér én Word-fil for hver planlagt aktivitet."""
    result: dict[str, dict[str, Any]] = {}
    frame = schedule.get("students", pd.DataFrame())
    if frame.empty:
        return result
    try:
        template_bytes = core.DEFAULT_TEMPLATE.read_bytes() if core.DEFAULT_TEMPLATE.exists() else None
    except OSError:
        template_bytes = None
    name_counts = Counter(core.normal_key(value) for value in frame["Elev"].tolist())
    used_names: set[str] = set()
    total = len(frame)
    for position, (_, row) in enumerate(frame.iterrows(), 1):
        activity_id = str(row.get("Aktivitets-ID", ""))
        if not activity_id:
            if progress_callback:
                progress_callback(position, total)
            continue
        student_name = core.repair_text(row.get("Elev"))
        base = f"SOP-{_safe_docx_name(student_name)}"
        if name_counts[core.normal_key(student_name)] > 1:
            base += f" - {_safe_docx_name(row.get('Klasse') or activity_id)}"
        filename = f"{base}.docx"
        suffix = 2
        while filename.casefold() in used_names:
            filename = f"{base} - {suffix}.docx"
            suffix += 1
        used_names.add(filename.casefold())
        values = {
            "Elevnavn": student_name,
            "Klasse": core.repair_text(row.get("Klasse")),
            "Fag1 og niveau": core.repair_text(row.get("Fag 1")),
            "Vejleder fag 1": core.repair_text(row.get("Vejleder 1")),
            "Fag2 og niveau": core.repair_text(row.get("Fag 2")),
            "Vejleder fag 2": core.repair_text(row.get("Vejleder 2")),
        }
        try:
            if template_bytes is None:
                raise OSError("Word-skabelonen findes ikke")
            content = core.replace_docx_placeholders(template_bytes, values)
        except (ImportError, OSError, TypeError, ValueError):
            from docx import Document

            document = Document()
            document.add_heading(f"SOP – {student_name}", 0)
            document.add_paragraph(f"Klasse: {values['Klasse']}")
            document.add_paragraph(f"{values['Fag1 og niveau']}: {values['Vejleder fag 1']}")
            document.add_paragraph(f"{values['Fag2 og niveau']}: {values['Vejleder fag 2']}")
            output = BytesIO()
            document.save(output)
            content = output.getvalue()
        # Skabeloner kan mangle pladsholdere for elev og klasse. Indsæt derfor
        # altid et entydigt resumé forrest, så alle krævede data er til stede.
        from docx import Document

        completed = Document(BytesIO(content))
        summary_lines = [
            f"Elev: {student_name}",
            f"Klasse: {values['Klasse']}",
            f"{values['Fag1 og niveau']}: {values['Vejleder fag 1']}",
            f"{values['Fag2 og niveau']}: {values['Vejleder fag 2']}",
        ]
        if completed.paragraphs:
            first_paragraph = completed.paragraphs[0]
            for line in reversed(summary_lines):
                first_paragraph.insert_paragraph_before(line)
        else:
            for line in summary_lines:
                completed.add_paragraph(line)
        completed_output = BytesIO()
        completed.save(completed_output)
        content = completed_output.getvalue()
        result[activity_id] = {"filename": filename, "content": content}
        if progress_callback:
            progress_callback(position, total)
    return result


def _action_html(
    schedule: dict[str, Any],
    activity_id: str,
    documents: dict[str, dict[str, Any]],
    *,
    embedded_word: bool,
    dynamic_word: bool = False,
) -> str:
    document = documents.get(activity_id)
    if dynamic_word:
        word_action = (
            f'<a class="action action-word" href="#" '
            f'data-dynamic-word-id="{html.escape(activity_id, quote=True)}">Download Word</a>'
        )
    elif document:
        if embedded_word:
            word_href = "#"
            word_data = f' data-word-id="{html.escape(activity_id, quote=True)}"'
        else:
            word_href = "Word/" + quote(document["filename"])
            word_data = ""
        word_action = (
            f'<a class="action action-word" download="{html.escape(document["filename"], quote=True)}" '
            f'href="{word_href}"{word_data}>Download Word</a>'
        )
    else:
        word_action = '<span class="action action-disabled" title="Dokumentet kan først hentes, når aktiviteten er planlagt.">Download Word</span>'
    teams_href, teams_error = teams_chat_link(schedule, activity_id)
    if teams_href:
        teams_action = (
            f'<a class="action action-teams" href="{html.escape(teams_href, quote=True)}" '
            'target="_blank" rel="noopener">Åbn Teams-chat</a>'
        )
    else:
        teams_action = (
            f'<span class="action action-disabled" title="{html.escape(teams_error, quote=True)}">Åbn Teams-chat</span>'
        )
    return f'<div class="actions">{word_action}{teams_action}</div>'


def _dynamic_word_script(schedule: dict[str, Any]) -> str:
    """Indlejr Word-skabelonen og opret en elevudgave af den ved klik."""
    frame = schedule.get("students", pd.DataFrame())
    name_counts = Counter(core.normal_key(value) for value in frame.get("Elev", pd.Series(dtype=str)).tolist())
    name_class_counts = Counter(
        (core.normal_key(row.get("Elev")), core.normal_key(row.get("Klasse")))
        for _, row in frame.fillna("").iterrows()
    )
    values: dict[str, dict[str, str]] = {}
    for _, row in frame.fillna("").iterrows():
        activity_id = str(row.get("Aktivitets-ID", ""))
        if not activity_id:
            continue
        name = core.repair_text(row.get("Elev"))
        filename = f"SOP-{_safe_docx_name(name)}"
        if name_counts[core.normal_key(name)] > 1:
            filename += f"-{_safe_docx_name(row.get('Klasse') or activity_id)}"
        if name_class_counts[(core.normal_key(name), core.normal_key(row.get("Klasse")))] > 1:
            filename += f"-{_safe_docx_name(activity_id)}"
        values[activity_id] = {
            "filename": f"{filename}.docx",
            "elev": name,
            "klasse": core.repair_text(row.get("Klasse")),
            "fag1": core.repair_text(row.get("Fag 1")),
            "vejleder1": core.repair_text(row.get("Vejleder 1")),
            "fag2": core.repair_text(row.get("Fag 2")),
            "vejleder2": core.repair_text(row.get("Vejleder 2")),
        }
    # JSON-dataen ligger i et script-element. Kod HTML-tegnene eksplicit, så et
    # elevnavn aldrig kan afslutte eller oprette HTML, selv hvis filen åbnes
    # direkte i en browser.
    payload = json.dumps(values, ensure_ascii=True, separators=(",", ":"))
    payload = payload.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    try:
        template_bytes = core.DEFAULT_TEMPLATE.read_bytes()
    except OSError:
        template_bytes = b""
    template_payload = base64.b64encode(template_bytes).decode("ascii")
    script = r"""<script id="dynamic-word-values" type="application/json">__WORD_VALUES__</script>
<script id="dynamic-word-template" type="text/plain">__WORD_TEMPLATE__</script>
<script>
const dynamicWordValues=JSON.parse(document.getElementById('dynamic-word-values').textContent);
const wordEncoder=new TextEncoder();
const wordDecoder=new TextDecoder();
const wordCrcTable=(()=>{const table=new Uint32Array(256);for(let i=0;i<256;i++){let c=i;for(let j=0;j<8;j++)c=(c&1)?0xedb88320^(c>>>1):c>>>1;table[i]=c>>>0;}return table;})();
function wordCrc32(bytes){let c=0xffffffff;for(const value of bytes)c=wordCrcTable[(c^value)&255]^(c>>>8);return(c^0xffffffff)>>>0;}
function wordXml(value){return String(value||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');}
function wordBase64(value){const text=atob(value);const bytes=new Uint8Array(text.length);for(let i=0;i<text.length;i++)bytes[i]=text.charCodeAt(i);return bytes;}
function wordZip(files){
  const local=[],central=[];let offset=0,centralSize=0;
  for(const file of files){
    const nameBytes=wordEncoder.encode(file.name),data=file.data,crc=file.crc===undefined?wordCrc32(data):file.crc,method=file.method||0,flags=file.flags||0,size=file.size===undefined?data.length:file.size;
    const header=new Uint8Array(30);const view=new DataView(header.buffer);
    view.setUint32(0,0x04034b50,true);view.setUint16(4,20,true);view.setUint16(6,flags,true);view.setUint16(8,method,true);view.setUint16(10,file.time||0,true);view.setUint16(12,file.date||0,true);view.setUint32(14,crc,true);view.setUint32(18,data.length,true);view.setUint32(22,size,true);view.setUint16(26,nameBytes.length,true);
    local.push(header,nameBytes,data);
    const directory=new Uint8Array(46);const dirView=new DataView(directory.buffer);
    dirView.setUint32(0,0x02014b50,true);dirView.setUint16(4,20,true);dirView.setUint16(6,20,true);dirView.setUint16(8,flags,true);dirView.setUint16(10,method,true);dirView.setUint16(12,file.time||0,true);dirView.setUint16(14,file.date||0,true);dirView.setUint32(16,crc,true);dirView.setUint32(20,data.length,true);dirView.setUint32(24,size,true);dirView.setUint16(28,nameBytes.length,true);dirView.setUint32(42,offset,true);
    central.push(directory,nameBytes);offset+=header.length+nameBytes.length+data.length;centralSize+=directory.length+nameBytes.length;
  }
  const end=new Uint8Array(22);const endView=new DataView(end.buffer);endView.setUint32(0,0x06054b50,true);endView.setUint16(8,Object.keys(files).length,true);endView.setUint16(10,Object.keys(files).length,true);endView.setUint32(12,centralSize,true);endView.setUint32(16,offset,true);
  const total=offset+centralSize+end.length,out=new Uint8Array(total);let cursor=0;
  for(const chunk of [...local,...central,end]){out.set(chunk,cursor);cursor+=chunk.length;}return out;
}
function wordTemplateFiles(bytes){
  const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);let eocd=-1;
  for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--){if(view.getUint32(i,true)===0x06054b50){eocd=i;break;}}
  if(eocd<0)throw new Error('Word-skabelonen er ikke en gyldig DOCX-fil.');
  const count=view.getUint16(eocd+10,true),start=view.getUint32(eocd+16,true),files=[];let cursor=start;
  for(let index=0;index<count;index++){
    if(view.getUint32(cursor,true)!==0x02014b50)throw new Error('Word-skabelonen kan ikke læses.');
    const flags=view.getUint16(cursor+8,true),method=view.getUint16(cursor+10,true),time=view.getUint16(cursor+12,true),date=view.getUint16(cursor+14,true),crc=view.getUint32(cursor+16,true),compressedSize=view.getUint32(cursor+20,true),size=view.getUint32(cursor+24,true),nameLength=view.getUint16(cursor+28,true),extraLength=view.getUint16(cursor+30,true),commentLength=view.getUint16(cursor+32,true),localOffset=view.getUint32(cursor+42,true);
    const name=wordDecoder.decode(bytes.slice(cursor+46,cursor+46+nameLength));const localNameLength=view.getUint16(localOffset+26,true),localExtraLength=view.getUint16(localOffset+28,true),dataStart=localOffset+30+localNameLength+localExtraLength;
    files.push({name,flags,method,time,date,crc,size,data:bytes.slice(dataStart,dataStart+compressedSize)});cursor+=46+nameLength+extraLength+commentLength;
  }return files;
}
async function wordInflate(file){
  if(file.method===0)return file.data;
  if(file.method!==8||typeof DecompressionStream==='undefined')throw new Error('Browseren kan ikke åbne Word-skabelonen. Brug en opdateret Chrome eller Edge.');
  return new Uint8Array(await new Response(new Blob([file.data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
}
async function createWordDocument(item){
  const template=wordBase64(document.getElementById('dynamic-word-template').textContent.trim());if(!template.length)throw new Error('SOP_skabelon.docx mangler i eksporten. Eksportér planen igen.');
  const files=wordTemplateFiles(template),documentPart=files.find(file=>file.name==='word/document.xml');if(!documentPart)throw new Error('SOP_skabelon.docx mangler word/document.xml.');
  let source=wordDecoder.decode(await wordInflate(documentPart));
  const values={'Elevnavn':item.elev,'Klasse':item.klasse,'Fag1 og niveau':item.fag1,'Vejleder fag 1':item.vejleder1,'Fag2 og niveau':item.fag2,'Vejleder fag 2':item.vejleder2};
  for(const [placeholder,value] of Object.entries(values))source=source.split(placeholder).join(wordXml(value));
  documentPart.data=wordEncoder.encode(source);documentPart.method=0;documentPart.flags=0;documentPart.crc=undefined;documentPart.size=documentPart.data.length;
  return wordZip(files);
}
document.querySelectorAll('[data-dynamic-word-id]').forEach(link=>link.addEventListener('click',async event=>{
  event.preventDefault();const item=dynamicWordValues[link.dataset.dynamicWordId];if(!item)return;const original=link.textContent;link.textContent='Opretter Word …';link.setAttribute('aria-busy','true');
  try{const url=URL.createObjectURL(new Blob([await createWordDocument(item)],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}));const download=document.createElement('a');download.href=url;download.download=item.filename;download.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}catch(error){alert(error.message||'Word-dokumentet kunne ikke oprettes.');}
  finally{link.textContent=original;link.removeAttribute('aria-busy');}
}));
</script>"""
    return script.replace("__WORD_VALUES__", payload).replace("__WORD_TEMPLATE__", template_payload)


def _unplaced_section(schedule: dict[str, Any]) -> str:
    frame = schedule.get("unplaced", pd.DataFrame())
    if frame.empty:
        return ""
    table = frame.to_html(index=False, escape=True, classes="schedule-table")
    return (
        '<section class="unplaced"><h2>Kan ikke placeres inden for det valgte tidsrum</h2>'
        '<p>Kræver anden tid eller dag. Rækkerne er ikke kalenderaftaler.</p>'
        f"{table}</section>"
    )


def _append_unplaced_html(factory: Any, schedule: dict[str, Any]) -> str:
    document = factory(schedule)
    section = _unplaced_section(schedule)
    if not section:
        return document
    style = "<style>.unplaced{margin:2rem 0;padding:1rem;border:2px solid #a85b37;background:#fff8f2}.unplaced h2{margin-top:0}.schedule-table{width:100%;border-collapse:collapse}.schedule-table th,.schedule-table td{padding:.45rem;border:1px solid #ddd;text-align:left}</style>"
    return document.replace("</body>", f"{style}{section}</body>")


def make_schedule_excel_beta(schedule: dict[str, Any]) -> bytes:
    result = _CORE_MAKE_SCHEDULE_EXCEL(schedule)
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows

    workbook = load_workbook(BytesIO(result))
    # Indstillingerne bruges internt af programmet, men hører ikke hjemme i
    # den plan, der deles med lærerne.
    if "Indstillinger" in workbook.sheetnames:
        del workbook["Indstillinger"]

    frame = schedule.get("unplaced", pd.DataFrame())
    if frame.empty:
        output = BytesIO()
        workbook.save(output)
        return output.getvalue()

    name = core.excel_sheet_name("Kræver anden dag", set(workbook.sheetnames))
    worksheet = workbook.create_sheet(name)
    for row in dataframe_to_rows(frame, index=False, header=True):
        worksheet.append(list(row))
    worksheet.freeze_panes = "A2"
    for column in worksheet.columns:
        letter = column[0].column_letter
        worksheet.column_dimensions[letter].width = min(45, max(12, max(len(str(cell.value or "")) for cell in column) + 2))
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def _teacher_html_with_actions(
    schedule: dict[str, Any], *, embedded_word: bool,
    documents: dict[str, dict[str, Any]] | None = None, progress_callback: Any = None,
    dynamic_word: bool = False,
) -> str:
    document = _CORE_MAKE_SCHEDULE_HTML(schedule)
    # Core-eksporten viser ellers alle beregnings- og inputindstillinger i
    # introduktionen. De skal ikke følge med den delbare lærerplan.
    document = re.sub(r"<ul>.*?</ul>", "", document, count=1, flags=re.S)
    documents = documents if documents is not None else (
        {} if dynamic_word else _word_documents(schedule, progress_callback=progress_callback)
    )
    document = document.replace(
        "<th>Bemærkning</th>",
        "<th>Bemærkning</th><th>Handlinger</th>",
        1,
    )
    teacher_columns = ["Tid", "Elev", "Klasse", "Fag", "Medvejleder", "Projekttitel", "Bemærkning"]
    teacher_frame = schedule["teachers"].sort_values(["Lærer", "Start", "Slut"], kind="stable")
    teacher_frame = teacher_frame[teacher_frame["Type"] == "Vejledning"]
    for row in teacher_frame.fillna("").to_dict("records"):
        search_terms = core.normal_key(f"{row.get('Lærer', '')} {row.get('Initialer', '')}")
        cells = "".join(f"<td>{html.escape(str(row.get(column, '')))}</td>" for column in teacher_columns)
        old_row = f'<tr data-teacher="{html.escape(search_terms, quote=True)}">{cells}</tr>'
        action = _action_html(
            schedule, str(row.get("Aktivitets-ID", "")), documents,
            embedded_word=embedded_word, dynamic_word=dynamic_word,
        )
        new_row = f'<tr data-teacher="{html.escape(search_terms, quote=True)}">{cells}<td>{action}</td></tr>'
        document = document.replace(old_row, new_row, 1)
    action_style = """<style>
.actions{display:flex;flex-wrap:wrap;gap:6px;min-width:225px}.action{display:inline-block;padding:6px 8px;border-radius:6px;font-size:12px;font-weight:700;text-decoration:none;white-space:nowrap}.action-word{background:#e7f4f1;color:#075f5b}.action-teams{background:#6264a7;color:#fff}.action-disabled{background:#eceff0;color:#899397;cursor:not-allowed}
</style>"""
    document = document.replace("</head>", f"{action_style}</head>")
    if dynamic_word:
        document = document.replace("</body>", f"{_dynamic_word_script(schedule)}</body>")
    elif embedded_word and documents:
        payload = {
            activity_id: {
                "filename": item["filename"],
                "data": base64.b64encode(item["content"]).decode("ascii"),
            }
            for activity_id, item in documents.items()
        }
        payload_json = json.dumps(payload, ensure_ascii=True, separators=(",", ":")).replace("</", "<\\/")
        word_script = f"""<script id="word-documents" type="application/json">{payload_json}</script>
<script>
const wordDocuments=JSON.parse(document.getElementById('word-documents').textContent);
document.querySelectorAll('[data-word-id]').forEach(link=>link.addEventListener('click',event=>{{
  event.preventDefault();const item=wordDocuments[link.dataset.wordId];if(!item)return;
  const binary=atob(item.data),bytes=new Uint8Array(binary.length);for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);
  const url=URL.createObjectURL(new Blob([bytes],{{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}}));
  const download=document.createElement('a');download.href=url;download.download=item.filename;download.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
}}));
</script>"""
        document = document.replace("</body>", f"{word_script}</body>")
    return _append_unplaced_html(lambda _: document, schedule)


def make_schedule_html_beta(schedule: dict[str, Any], progress_callback: Any = None) -> str:
    """Selvstændig lærer-HTML med Word-generering ved klik og Teams-links."""
    return _teacher_html_with_actions(schedule, embedded_word=False, dynamic_word=True)


def make_teacher_export_package(schedule: dict[str, Any], progress_callback: Any = None) -> bytes:
    """Pak én lærer-HTML, som opretter det enkelte Word-dokument ved klik."""
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "Vejledningsplan.html",
            _teacher_html_with_actions(schedule, embedded_word=False, dynamic_word=True).encode("utf-8"),
        )
    return output.getvalue()


def make_student_schedule_html_beta(schedule: dict[str, Any]) -> str:
    return _append_unplaced_html(_CORE_MAKE_STUDENT_HTML, schedule)


def _render_beta_planner() -> None:
    schedule = st.session_state.get("v2_schedule")
    context = st.session_state.get("beta_schedule_context")
    if not schedule or not context or "beta" not in schedule:
        return
    st.divider()
    st.subheader("Beta · dynamisk planjustering")
    planned_count = len(schedule["students"])
    unplaced = schedule.get("unplaced", pd.DataFrame())
    metric_a, metric_b, metric_c = st.columns(3)
    metric_a.metric("Planlagt", planned_count)
    metric_b.metric("Kræver anden dag", len(unplaced))
    metric_c.metric("Låste aktiviteter", len(st.session_state.get("beta_locked_activities", {})))
    if not unplaced.empty:
        st.warning("Disse aktiviteter kan ikke nås inden for dagens tidsrum. De indgår ikke i elev- eller lærerplanen.")
        st.dataframe(unplaced, hide_index=True, use_container_width=True)

    activities = schedule.get("activities", pd.DataFrame())
    metadata = schedule.get("activity_metadata", {})
    if activities.empty:
        return
    labels = {
        str(row["activity_id"]): (
            f"{row['activity_id']} · {metadata.get(str(row['activity_id']), {}).get('student_name', '')} "
            f"({metadata.get(str(row['activity_id']), {}).get('class_name', '')}) · "
            f"{row['start'] or row['status']}"
        )
        for _, row in activities.iterrows()
    }
    mode_label = st.radio(
        "Når aktiviteten flyttes",
        ["Juster planen dynamisk", "Lås kun den flyttede aktivitet"],
        horizontal=True,
        key="beta_move_mode",
        help="Dynamisk tilstand genplacerer konfliktramte aktiviteter. Streng tilstand accepterer kun en allerede ledig tid.",
    )
    move_mode = "dynamic" if mode_label == "Juster planen dynamisk" else "strict"
    left, middle, right = st.columns([2, 2, 1])
    with left:
        selected_id = st.selectbox("Aktivitet, der skal flyttes", list(labels), format_func=labels.get, key="beta_move_activity")
    start_dt, end_dt = _parse_clock(schedule["beta"]["plan_start"]), _parse_clock(schedule["beta"]["plan_end"])
    choices: list[str] = []
    candidate = start_dt
    while candidate < end_dt:
        choices.append(_clock(candidate))
        candidate += timedelta(minutes=5)
    current_start = str(activities.loc[activities["activity_id"] == selected_id, "start"].iloc[0])
    with middle:
        selected_time = st.selectbox("Nyt starttidspunkt", choices, index=choices.index(current_start) if current_start in choices else 0, key="beta_move_time")
    with right:
        st.write("")
        st.write("")
        apply_move = st.button("Flyt og justér", type="primary", key="beta_apply_move", use_container_width=True)
    assessment = assess_move(schedule, selected_id, selected_time)
    if assessment["status"] == "green":
        st.success("Grøn placering: " + assessment["reason"])
    elif assessment["status"] == "yellow":
        st.warning("Gul placering: " + assessment["reason"])
    else:
        st.error("Rød placering: " + assessment["reason"])

    history_columns = st.columns(4)
    undo = history_columns[0].button(
        "Fortryd", key="beta_undo", disabled=not st.session_state.get("beta_undo_stack"), use_container_width=True
    )
    redo = history_columns[1].button(
        "Gentag", key="beta_redo", disabled=not st.session_state.get("beta_redo_stack"), use_container_width=True
    )
    optimise = history_columns[2].button("Optimér hele planen", key="beta_optimise_all", use_container_width=True)
    reset = history_columns[3].button("Fjern alle låse", key="beta_reset_moves", use_container_width=True)

    def save_change(
        updated_schedule: dict[str, Any], updated_locks: dict[str, str],
        *, action: str, old_time: str = "", new_time: str = "", activity_id: str = "",
        push_history: bool = True,
    ) -> None:
        previous_locks = dict(st.session_state.get("beta_locked_activities", {}))
        if push_history:
            push_lock_history(st.session_state, previous_locks)
        st.session_state["beta_locked_activities"] = dict(updated_locks)
        st.session_state["v2_schedule"] = updated_schedule
        try:
            actor = core.repair_text(getattr(st.user, "email", "") or getattr(st.user, "name", ""))
        except Exception:
            actor = ""
        st.session_state.setdefault("beta_change_log", []).append({
            "Tidspunkt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Bruger": actor or "Lokal bruger",
            "Handling": action,
            "Aktivitets-ID": activity_id,
            "Fra": old_time,
            "Til": new_time,
            "Planversion": updated_schedule["beta"]["version"],
        })
        st.rerun()

    if apply_move:
        try:
            updated, updated_locks, _ = apply_manual_move(
                schedule, context, dict(st.session_state.get("beta_locked_activities", {})),
                selected_id, selected_time, move_mode,
            )
            save_change(
                updated, updated_locks, action="Flyt aktivitet", activity_id=selected_id,
                old_time=current_start, new_time=selected_time,
            )
        except ValueError as error:
            st.error(str(error))
    if undo:
        target_locks = undo_lock_history(
            st.session_state, dict(st.session_state.get("beta_locked_activities", {}))
        )
        if target_locks is not None:
            updated = make_schedule_beta(
                **context, locked_activities=target_locks, expected_version=int(schedule["beta"]["version"])
            )
            save_change(updated, target_locks, action="Fortryd", push_history=False)
    if redo:
        target_locks = redo_lock_history(
            st.session_state, dict(st.session_state.get("beta_locked_activities", {}))
        )
        if target_locks is not None:
            updated = make_schedule_beta(
                **context, locked_activities=target_locks, expected_version=int(schedule["beta"]["version"])
            )
            save_change(updated, target_locks, action="Gentag", push_history=False)
    if optimise:
        locks = dict(st.session_state.get("beta_locked_activities", {}))
        updated = make_schedule_beta(
            **context, locked_activities=locks, expected_version=int(schedule["beta"]["version"])
        )
        save_change(updated, locks, action="Optimér hele planen", push_history=False)
    if reset:
        updated = make_schedule_beta(
            **context, locked_activities={}, expected_version=int(schedule["beta"]["version"])
        )
        save_change(updated, {}, action="Fjern alle låse")
    st.caption("Flytningen låser aktiviteten. SOPtima bevarer først de øvrige tider og flytter kun aktiviteter, der nu er i konflikt. En ugyldig eller umulig flytning ender i blokken for aktiviteter, der kræver anden tid eller dag.")

    st.subheader("Træk og slip i tidsplanen")
    st.info(
        "Træk den blå blok med elevens navn vandret til et nyt klokkeslæt, og slip den på lærerens række. "
        "Farverne viser, om placeringen er ledig (grøn), kræver omplanlægning (gul) eller er ugyldig (rød). "
        "Når du slipper, genberegnes planen automatisk. Hvis træk-og-slip ikke virker i din browser, kan du "
        "bruge felterne ovenfor: vælg aktivitet, vælg nyt starttidspunkt, og tryk ‘Flyt og justér’."
    )
    drag_items = [
        {"id": row["Aktivitets-ID"], "label": f"{row['Elev']} · {row['Klasse']}", "start": row["Start"], "end": row["Slut"]}
        for _, row in schedule["students"].iterrows()
    ]
    dropped = drag_timeline(
        items=drag_items, start=schedule["beta"]["plan_start"], end=schedule["beta"]["plan_end"], step=5,
        validity=move_validity_map(schedule), version=schedule["beta"]["version"],
        key="beta_drag_timeline",
        default=None,
    )
    if isinstance(dropped, dict) and dropped.get("id") in labels and dropped.get("start") in choices:
        drag_key = f"{dropped.get('version')}:{dropped['id']}:{dropped['start']}"
        if st.session_state.get("beta_last_drag") != drag_key:
            st.session_state["beta_last_drag"] = drag_key
            if int(dropped.get("version", -1)) != int(schedule["beta"]["version"]):
                st.error("Tidsplanen er blevet ændret. Træk aktiviteten igen i den aktuelle version.")
            else:
                try:
                    dragged_id, dragged_time = str(dropped["id"]), str(dropped["start"])
                    old_time = str(activities.loc[activities["activity_id"] == dragged_id, "start"].iloc[0])
                    updated, updated_locks, _ = apply_manual_move(
                        schedule, context, dict(st.session_state.get("beta_locked_activities", {})),
                        dragged_id, dragged_time, move_mode,
                    )
                    save_change(
                        updated, updated_locks, action="Træk og slip", activity_id=dragged_id,
                        old_time=old_time, new_time=dragged_time,
                    )
                except ValueError as error:
                    st.error(str(error))

    package_key = f"{st.session_state.get('beta_context_marker', '')}:{schedule['beta']['version']}"
    if st.session_state.get("beta_teacher_package_key") != package_key:
        st.session_state["beta_teacher_package_key"] = package_key
        st.session_state["beta_teacher_package"] = None
    package = st.session_state.get("beta_teacher_package")
    if package is None and st.button(
        "Klargør lærer-HTML med Word-links", type="primary", key="beta_prepare_teacher_package"
    ):
        progress = st.progress(0.2, text="Pakker lærer-HTML …")
        package = make_teacher_export_package(schedule)
        progress.progress(1.0, text="Lærer-HTML med Word-links er klar til download")
        st.session_state["beta_teacher_package"] = package
    if package is not None:
        st.download_button(
            "Download lærer-HTML med Word-links",
            data=st.session_state["beta_teacher_package"],
            file_name="Vejledningsplan-med-Word-links.zip",
            mime="application/zip",
            key="beta_download_teacher_package",
        )
    changes = st.session_state.get("beta_change_log", [])
    if changes:
        with st.expander("Ændringslog", expanded=False):
            st.dataframe(pd.DataFrame(changes), hide_index=True, use_container_width=True)


def activate_beta() -> None:
    """Aktivér beta-funktionerne uden at ændre app.py på disken."""
    core.make_schedule = make_schedule_beta
    core.make_schedule_excel = make_schedule_excel_beta
    core.make_schedule_html = make_schedule_html_beta
    core.make_student_schedule_html = make_student_schedule_html_beta
    core.parse_students = parse_students_beta
    core.parse_teachers = parse_teachers_beta
    core.parse_schedule_upload = parse_schedule_upload_beta
    core.make_schedule_input_template = make_schedule_input_template_beta


if __name__ == "__main__":
    activate_beta()
    core.main_v2()
    _render_beta_planner()
