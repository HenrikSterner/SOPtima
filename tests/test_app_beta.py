from __future__ import annotations

import io
import zipfile
from datetime import datetime, time as clock

import pandas as pd
import pytest
from docx import Document
from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest

import app
import app_beta as beta


def student(index: int, *, name: str | None = None, email: str = "") -> dict:
    return {
        "id": f"E{index:03d}",
        "name": name or f"Elev {index}",
        "className": "S 2024q",
        "subjects": ["Dansk A", "Historie B"],
        "subjectsWithLevel": ["Dansk A", "Historie B"],
        "wishes": [],
        "projectTitle": f"Projekt {index}",
        "projectDescription": "",
        "scheduleNotes": "",
        "teams_email": email,
    }


def teacher(teacher_id: str, name: str, email: str = "") -> dict:
    return {
        "id": teacher_id,
        "name": name,
        "subjects": ["Dansk", "Historie"],
        "holds": [],
        "teams_email": email,
    }


def make_beta_schedule(
    pairs: list[tuple[str, str]],
    *,
    start: clock = clock(8, 0),
    end: clock = clock(12, 0),
    minutes: int = 20,
    transition: int = 0,
    pause_count: int = 0,
    pause_minutes: int = 0,
    teacher_blocks=None,
    locks=None,
    students_override=None,
    teachers_override=None,
):
    ids = sorted({teacher_id for pair in pairs for teacher_id in pair})
    teachers = teachers_override or [teacher(teacher_id, teacher_id.upper(), f"{teacher_id}@nextkbh.dk") for teacher_id in ids]
    students = students_override or [student(index + 1, email=f"next{27000 + index}@edu.nextkbh.dk") for index in range(len(pairs))]
    solution = {"assignments": [list(pair) for pair in pairs]}
    schedule = beta.make_schedule_beta(
        students,
        teachers,
        solution,
        start,
        end,
        minutes,
        pause_count,
        pause_minutes,
        transition,
        True,
        lunch_mode="Fast tidspunkt for alle lærere",
        lunch_start_time=clock(10, 0) if end > clock(10, 5) else clock(8, 40),
        lunch_minutes=5,
        search_attempts=5,
        teacher_blocks=teacher_blocks,
        floating_pauses=True,
        locked_activities=locks or {},
    )
    context = {
        "students": students,
        "teachers": teachers,
        "solution": solution,
        "start_time": start,
        "end_time": end,
        "student_minutes": minutes,
        "pause_count": pause_count,
        "pause_minutes": pause_minutes,
        "transition_minutes": transition,
        "group_pairs": True,
        "lunch_mode": "Fast tidspunkt for alle lærere",
        "lunch_start_time": clock(10, 0) if end > clock(10, 5) else clock(8, 40),
        "lunch_minutes": 5,
        "search_attempts": 5,
        "avoid_teacher_gaps": True,
        "teacher_blocks": teacher_blocks,
        "floating_pauses": True,
    }
    return schedule, context


def minutes(value: str) -> int:
    parsed = datetime.strptime(value, "%H:%M")
    return parsed.hour * 60 + parsed.minute


def assert_no_teacher_conflicts(schedule: dict) -> None:
    frame = schedule["teachers"]
    frame = frame[frame["Type"] == "Vejledning"]
    for _, group in frame.groupby("Initialer"):
        rows = group.sort_values("Start").to_dict("records")
        for left, right in zip(rows, rows[1:]):
            assert minutes(left["Slut"]) <= minutes(right["Start"])


def test_importing_beta_does_not_patch_production_module():
    assert app.make_schedule is beta._CORE_MAKE_SCHEDULE
    assert app.make_schedule_html is beta._CORE_MAKE_SCHEDULE_HTML


def test_partial_schedule_returns_planned_and_unplaced_instead_of_error():
    schedule, _ = make_beta_schedule([("a", "b")] * 4, end=clock(9, 0))
    assert 0 < len(schedule["students"]) < 4
    assert len(schedule["students"]) + len(schedule["unplaced"]) == 4
    assert set(schedule["unplaced"]["Status"]) == {"Uplaceret"}
    assert schedule["settings"]["Slut"] == "09:00"
    assert_no_teacher_conflicts(schedule)


def test_beta_deduplicates_columns_from_a_previous_streamlit_rerun(monkeypatch):
    students = [student(1)]
    teachers = [teacher("a", "A"), teacher("b", "B")]
    solution = {"assignments": [["a", "b"]]}
    original_make_schedule = beta._CORE_MAKE_SCHEDULE
    def compatible_base(*args, **kwargs):
        kwargs["lunch_mode"] = "Fast tidspunkt for alle l" + chr(230) + "rere"
        return original_make_schedule(*args, **kwargs)
    monkeypatch.setattr(beta, "_CORE_MAKE_SCHEDULE", compatible_base)
    original_beta_schedule = beta.make_schedule_beta
    def compatible_beta(*args, **kwargs):
        kwargs["lunch_mode"] = "Fast tidspunkt for alle l" + chr(230) + "rere"
        return original_beta_schedule(*args, **kwargs)
    monkeypatch.setattr(beta, "make_schedule_beta", compatible_beta)
    base = beta._CORE_MAKE_SCHEDULE(
        students, teachers, solution, clock(8), clock(12), 20, 0, 0, 0, True,
        lunch_mode="Fast tidspunkt for alle lÃ¦rere", lunch_start_time=clock(10),
        lunch_minutes=5, search_attempts=2,
    )
    # Efter en script-genkørsel kunne en gammel beta-plan tidligere blive brugt
    # som grundplan og få de samme kolonner tilføjet igen.
    base["students"] = pd.concat([base["students"], base["students"]], axis=1)
    base["teachers"] = pd.concat([base["teachers"], base["teachers"]], axis=1)
    base["timeline"] = pd.concat([base["timeline"], base["timeline"]], axis=1)
    monkeypatch.setattr(beta, "_CORE_MAKE_SCHEDULE", lambda *args, **kwargs: base)
    schedule = beta.make_schedule_beta(
        students, teachers, solution, clock(8), clock(12), 20, 0, 0, 0, True,
        lunch_mode="Fast tidspunkt for alle lÃ¦rere", lunch_start_time=clock(10),
        lunch_minutes=5, search_attempts=2,
    )
    assert schedule["students"].columns.is_unique
    assert schedule["teachers"].columns.is_unique
    assert schedule["timeline"].columns.is_unique


def test_activity_schema_and_ids_are_stable_when_input_order_changes():
    students = [student(1), student(2), student(3)]
    teachers = [teacher("a", "A"), teacher("b", "B")]
    first, _ = make_beta_schedule([("a", "b")] * 3, students_override=students, teachers_override=teachers)
    second_students = [students[2], students[0], students[1]]
    second, _ = make_beta_schedule([("a", "b")] * 3, students_override=second_students, teachers_override=teachers)
    required = {"activity_id", "elev_id", "lærer_ids", "varighed", "start", "slut", "status", "låst", "årsag"}
    assert required <= set(first["activities"].columns)
    mapping_first = dict(zip(first["activities"]["elev_id"], first["activities"]["activity_id"]))
    mapping_second = dict(zip(second["activities"]["elev_id"], second["activities"]["activity_id"]))
    assert mapping_first == mapping_second


def test_parallel_rounds_transition_and_teacher_blocks_are_hard_constraints():
    schedule, _ = make_beta_schedule(
        [("a", "b"), ("c", "d"), ("a", "b")],
        transition=5,
        teacher_blocks={"a": [(clock(8, 0), clock(8, 30))]},
    )
    a_rows = schedule["teachers"]
    a_rows = a_rows[(a_rows["Type"] == "Vejledning") & (a_rows["Initialer"] == "a")].sort_values("Start")
    assert all(minutes(value) >= 8 * 60 + 30 for value in a_rows["Start"])
    rows = a_rows.to_dict("records")
    for left, right in zip(rows, rows[1:]):
        assert minutes(right["Start"]) - minutes(left["Slut"]) >= 5
    assert_no_teacher_conflicts(schedule)


def test_independent_teacher_pairs_can_run_in_parallel():
    schedule, _ = make_beta_schedule([("a", "b"), ("c", "d")])
    assert schedule["students"]["Start"].nunique() == 1


def test_move_assessment_and_dynamic_repair():
    schedule, context = make_beta_schedule([("a", "b")] * 3)
    ids = schedule["activities"]["activity_id"].tolist()
    first_start = schedule["activities"].set_index("activity_id").loc[ids[0], "start"]
    assessment = beta.assess_move(schedule, ids[1], first_start)
    assert assessment["status"] == "yellow"
    with pytest.raises(ValueError, match="Streng tilstand"):
        beta.apply_manual_move(schedule, context, {}, ids[1], first_start, "strict")
    updated, locks, result = beta.apply_manual_move(schedule, context, {}, ids[1], first_start, "dynamic")
    assert result["status"] == "yellow"
    assert locks[ids[1]] == first_start
    moved = updated["activities"].set_index("activity_id")
    assert moved.loc[ids[1], "start"] == first_start
    assert moved.loc[ids[1], "status"] == "låst"
    assert_no_teacher_conflicts(updated)


def test_red_move_is_rejected_for_break_and_five_minute_grid():
    schedule, context = make_beta_schedule([("a", "b")] * 3)
    activity_id = schedule["activities"].iloc[0]["activity_id"]
    lunch_start = schedule["beta"]["breaks"][0][0]
    assert beta.assess_move(schedule, activity_id, lunch_start)["status"] == "red"
    assert beta.assess_move(schedule, activity_id, "08:03")["status"] == "red"
    with pytest.raises(ValueError):
        beta.apply_manual_move(schedule, context, {}, activity_id, lunch_start, "dynamic")


def test_move_validity_map_contains_all_colours_needed_by_component():
    schedule, _ = make_beta_schedule([("a", "b")] * 3)
    activity_id = schedule["activities"].iloc[1]["activity_id"]
    validity = beta.move_validity_map(schedule)[activity_id]
    assert "yellow" in set(validity.values())
    assert "red" in set(validity.values())
    assert "green" in set(validity.values())


def test_pauses_are_only_exported_between_guidance_sessions():
    schedule, _ = make_beta_schedule(
        [("a", "b")] * 6,
        end=clock(12, 30),
        pause_count=2,
        pause_minutes=10,
    )
    guidance = schedule["students"]
    pause_rows = schedule["timeline"][schedule["timeline"]["Type"] == "Pause"]
    for _, pause in pause_rows.iterrows():
        assert any(guidance["Slut"].map(minutes) <= minutes(pause["Start"]))
        assert any(guidance["Start"].map(minutes) >= minutes(pause["Slut"]))


def test_extreme_overflow_uses_fallback_and_never_discards_all_results():
    pairs = [("a", "b")] * 55
    schedule, _ = make_beta_schedule(pairs, end=clock(9, 0))
    assert len(schedule["students"]) + len(schedule["unplaced"]) == 55
    assert len(schedule["students"]) > 0
    assert len(schedule["unplaced"]) > 0


def test_teams_link_contains_student_and_both_teacher_addresses():
    students = [student(1, email="next27888@edu.nextkbh.dk")]
    teachers = [
        teacher("hst", "Henrik Sterner", "hst@nextkbh.dk"),
        teacher("matr", "Mathias Ramberg", "matr@nextkbh.dk"),
    ]
    schedule, _ = make_beta_schedule(
        [("hst", "matr")], students_override=students, teachers_override=teachers
    )
    activity_id = schedule["activities"].iloc[0]["activity_id"]
    link, error = beta.teams_chat_link(schedule, activity_id)
    assert not error
    assert "hst%40nextkbh.dk" in link
    assert "matr%40nextkbh.dk" in link
    assert "next27888%40edu.nextkbh.dk" in link
    assert "topicName=SOP%20-%20Elev%201%20-%20S%202024q" in link


def test_missing_teams_email_disables_action_with_explanation():
    students = [student(1, email="")]
    teachers = [teacher("a", "A", "a@nextkbh.dk"), teacher("b", "B", "")]
    schedule, _ = make_beta_schedule([("a", "b")], students_override=students, teachers_override=teachers)
    activity_id = schedule["activities"].iloc[0]["activity_id"]
    link, error = beta.teams_chat_link(schedule, activity_id)
    assert link == ""
    assert "mangler Teams-adresse" in error
    teacher_html = beta.make_schedule_html_beta(schedule)
    assert "action-disabled" in teacher_html
    assert "mangler Teams-adresse" in teacher_html


def test_teacher_html_has_word_and_teams_actions_but_unplaced_has_none():
    schedule, _ = make_beta_schedule([("a", "b")] * 4, end=clock(9, 0))
    teacher_html = beta.make_schedule_html_beta(schedule)
    assert "Download Word" in teacher_html
    assert "Åbn Teams-chat" in teacher_html
    assert 'id="word-documents"' in teacher_html
    assert "data-word-id=" in teacher_html
    assert "Kan ikke placeres inden for det valgte tidsrum" in teacher_html
    assert teacher_html.count("Download Word") == len(schedule["teachers"][schedule["teachers"]["Type"] == "Vejledning"])


def test_teacher_export_package_contains_html_and_valid_unique_word_files():
    students = [
        student(1, name="Samme/Navn", email="one@edu.nextkbh.dk"),
        student(2, name="Samme/Navn", email="two@edu.nextkbh.dk"),
    ]
    teachers = [teacher("a", "Anna", "a@nextkbh.dk"), teacher("b", "Bo", "b@nextkbh.dk")]
    schedule, _ = make_beta_schedule(
        [("a", "b"), ("a", "b")], students_override=students, teachers_override=teachers
    )
    package = beta.make_teacher_export_package(schedule)
    with zipfile.ZipFile(io.BytesIO(package)) as archive:
        names = archive.namelist()
        assert "Vejledningsplan.html" in names
        word_names = [name for name in names if name.startswith("Word/")]
        assert len(word_names) == 2
        assert len(set(word_names)) == 2
        assert all("/" not in name.removeprefix("Word/") for name in word_names)
        document = Document(io.BytesIO(archive.read(word_names[0])))
        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        assert "Samme/Navn" in text
        package_html = archive.read("Vejledningsplan.html").decode("utf-8")
        assert "Word/SOP%20-%20Samme-Navn" in package_html


def test_excel_and_both_html_exports_include_unplaced_block():
    schedule, _ = make_beta_schedule([("a", "b")] * 4, end=clock(9, 0))
    workbook = load_workbook(io.BytesIO(beta.make_schedule_excel_beta(schedule)), read_only=True)
    assert "Kræver anden dag" in workbook.sheetnames
    assert "Kan ikke placeres inden for det valgte tidsrum" in beta.make_schedule_html_beta(schedule)
    assert "Kan ikke placeres inden for det valgte tidsrum" in beta.make_student_schedule_html_beta(schedule)


def test_email_columns_are_parsed_and_known_teacher_addresses_are_filled():
    student_frame = pd.DataFrame({
        "Elevnavn": ["Ane"], "Klasse": ["3q"], "Fag 1": ["Dansk A"],
        "Fag 2": ["Historie B"], "Elev Teams-email": ["NEXT1@EDU.NEXTKBH.DK"],
    })
    parsed_students = beta.parse_students_beta(student_frame)
    assert parsed_students[0]["teams_email"] == "next1@edu.nextkbh.dk"
    teacher_frame = pd.DataFrame({
        "Initialer": ["hst", "matr"], "Lærernavn": ["Henrik", "Mathias"],
        "Fag": ["Dansk", "Historie"], "Lærer Teams-email": ["", "matr@nextkbh.dk"],
    })
    parsed_teachers, _ = beta.parse_teachers_beta(teacher_frame)
    emails = {item["id"]: item["teams_email"] for item in parsed_teachers}
    assert emails == {"hst": "hst@nextkbh.dk", "matr": "matr@nextkbh.dk"}


def test_schedule_template_contains_teams_email_columns():
    workbook = load_workbook(io.BytesIO(beta.make_schedule_input_template_beta()), read_only=True)
    headers = [cell.value for cell in next(workbook["Tidsplan-input"].iter_rows())]
    assert "Elev Teams-email" in headers
    assert "Lærer 1 Teams-email" in headers
    assert "Lærer 2 Teams-email" in headers


def test_single_student_word_filename_is_exact_requirement():
    schedule, _ = make_beta_schedule([("a", "b")])
    documents = beta._word_documents(schedule)
    assert [item["filename"] for item in documents.values()] == ["SOP - Elev 1.docx"]


def test_unplaced_activity_has_neither_teams_link_nor_word_document():
    schedule, _ = make_beta_schedule([("a", "b")] * 4, end=clock(9, 0))
    unplaced_id = schedule["unplaced"].iloc[0]["Aktivitets-ID"]
    link, error = beta.teams_chat_link(schedule, unplaced_id)
    assert link == ""
    assert "bekræftet tidspunkt" in error
    assert unplaced_id not in beta._word_documents(schedule)


def test_teacher_html_escapes_student_content_and_action_attributes():
    dangerous = student(1, name='</td><script>alert("x")</script>', email="safe@edu.nextkbh.dk")
    teachers = [teacher("a", "A", "a@nextkbh.dk"), teacher("b", "B", "b@nextkbh.dk")]
    schedule, _ = make_beta_schedule([("a", "b")], students_override=[dangerous], teachers_override=teachers)
    teacher_html = beta.make_schedule_html_beta(schedule)
    assert "<script>alert" not in teacher_html
    assert "&lt;/td&gt;&lt;script&gt;" in teacher_html


def test_lock_history_supports_undo_redo_and_clears_redo_on_new_move():
    state = {}
    beta.push_lock_history(state, {})
    current = {"A-E001": "08:20"}
    assert beta.undo_lock_history(state, current) == {}
    assert beta.redo_lock_history(state, {}) == current
    state["beta_redo_stack"] = [{"old": "value"}]
    beta.push_lock_history(state, current)
    assert state["beta_redo_stack"] == []


def test_stale_plan_version_is_rejected(monkeypatch):
    fake_state = {}
    monkeypatch.setattr(beta.st, "session_state", fake_state)
    schedule, context = make_beta_schedule([("a", "b")] * 2)
    fake_state["v2_schedule"] = schedule
    fake_state["beta_context_marker"] = beta.core.stable_signature([
        [(item.get("id", ""), item.get("name", ""), item.get("className", "")) for item in context["students"]],
        context["solution"].get("assignments", []), context["start_time"], context["end_time"],
        context["student_minutes"], context["pause_count"], context["pause_minutes"],
        context["transition_minutes"], context["lunch_mode"], context["lunch_start_time"],
        context["lunch_minutes"], context["teacher_blocks"], context["floating_pauses"],
    ])
    with pytest.raises(ValueError, match="ændret siden"):
        beta.make_schedule_beta(**context, locked_activities={}, expected_version=999)


def test_invalid_settings_still_raise_clear_errors():
    students = [student(1)]
    teachers = [teacher("a", "A"), teacher("b", "B")]
    with pytest.raises(ValueError, match="Sluttidspunktet"):
        beta.make_schedule_beta(
            students, teachers, {"assignments": [["a", "b"]]}, clock(10), clock(9),
            20, 0, 0, 0, False,
        )


def test_beta_streamlit_app_starts_without_exceptions():
    rendered = AppTest.from_file("app_beta.py").run(timeout=45)
    assert not rendered.exception
