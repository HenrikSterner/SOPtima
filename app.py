from __future__ import annotations

import io
import html
import json
import re
import unicodedata
import zipfile
from collections import Counter, defaultdict
from datetime import date, datetime, time as dt_time, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st


ROOT = Path(__file__).resolve().parent
HTML_SOURCE = ROOT / "vejlederfordelingBETA.html"
DEFAULT_TEACHERS_FILE = ROOT / "lærere.csv"
DEFAULT_TEMPLATE = ROOT / "SOP - SKABELON.docx"
TEST_STUDENTS_FILE = ROOT / "testdata" / "sop_test_elever.csv"
TEST_TEACHERS_FILE = ROOT / "testdata" / "sop_test_lærere.csv"
NO_WISHES_LABEL = "Ingen ønsker"


def repair_text(value: Any) -> str:
    """Gør også ældre filer med dobbeltkodet UTF-8 læsbare."""
    text = "" if value is None else str(value).strip()
    if text.casefold() in {"nan", "nat", "none"}:
        return ""
    if any(marker in text for marker in ("Ã", "Â", "�")):
        try:
            fixed = text.encode("latin1").decode("utf-8")
            if fixed.count("�") < text.count("�"):
                return fixed
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return text


def normal_key(value: Any) -> str:
    return re.sub(r"\s+", " ", repair_text(value)).strip().casefold()


def canonical_subject(value: Any) -> str:
    subject = normal_key(value)
    subject = re.sub(r"\s+[abc]\*?$", "", subject)
    if subject in {"komm/info", "komm/it", "kommunikation og it", "kommunikation og it"}:
        return "kommunikation og it"
    repaired = {
        "idï¿½historie": "idehistorie",
        "idéhistorie": "idehistorie",
        "idrï¿½t": "idraet",
        "idræt": "idraet",
        "erhvervsï¿½konomi": "erhvervsoekonomi",
        "erhvervsøkonomi": "erhvervsoekonomi",
    }
    if subject in repaired:
        return repaired[subject]
    subject = "".join(char for char in unicodedata.normalize("NFD", subject) if unicodedata.category(char) != "Mn")
    return (
        subject.replace("æ", "ae")
        .replace("ø", "oe")
        .replace("å", "aa")
        .replace("Ã¦", "ae")
        .replace("Ã¸", "oe")
        .replace("Ã¥", "aa")
    )


def extract_embedded(name: str, terminator: str) -> Any:
    if not HTML_SOURCE.exists():
        return None
    source = HTML_SOURCE.read_text(encoding="utf-8")
    match = re.search(rf"const {name} = (.*?);\s*{terminator}", source, re.S)
    if not match:
        return None
    try:
        return json.loads(match.group(1))
    except json.JSONDecodeError:
        return None


@st.cache_data(show_spinner=False)
def embedded_data() -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, int]]:
    students = extract_embedded("STUDENTS", r"const TEACHERS") or []
    teachers = extract_embedded("TEACHERS", r"const DOCX_TEMPLATE_PARTS") or []
    capacities = extract_embedded("TEACHER_CAPACITIES", r"const teacherCapacities") or {}
    # Cloud-demoen behøver ikke den gamle HTML-kilde; brug de fiktive CSV-data som fallback.
    if not students and TEST_STUDENTS_FILE.exists():
        students = parse_students(read_path_table(TEST_STUDENTS_FILE))
    if not teachers and TEST_TEACHERS_FILE.exists():
        teachers, capacities = parse_teachers(read_path_table(TEST_TEACHERS_FILE))
    for student in students:
        student["name"] = repair_text(student.get("name"))
        student["className"] = repair_text(student.get("className"))
        student["subjects"] = [repair_text(item) for item in student.get("subjects", [])]
        student["wishes"] = [normal_key(item) for item in student.get("wishes", []) if item]
    for teacher in teachers:
        teacher["id"] = normal_key(teacher.get("id"))
        teacher["name"] = repair_text(teacher.get("name") or teacher["id"])
        teacher["subjects"] = [repair_text(item) for item in teacher.get("subjects", [])]
        teacher["holds"] = [repair_text(item) for item in teacher.get("holds", []) if repair_text(item)]
    return students, teachers, {normal_key(k): int(v) for k, v in capacities.items()}


def read_uploaded_table(uploaded: Any) -> pd.DataFrame:
    data = uploaded.getvalue()
    filename = uploaded.name.casefold()
    if filename.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(io.BytesIO(data), sheet_name=0, dtype=object)
    try:
        return pd.read_csv(io.BytesIO(data), sep=None, engine="python", dtype=object, encoding="utf-8-sig")
    except UnicodeDecodeError:
        return pd.read_csv(io.BytesIO(data), sep=None, engine="python", dtype=object, encoding="cp1252")


def read_path_table(path: Path) -> pd.DataFrame:
    if path.suffix.casefold() in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path, sheet_name=0, dtype=object)
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return pd.read_csv(path, sep=None, engine="python", dtype=object, encoding=encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Kunne ikke læse {path.name}.")


def column_for(columns: list[Any], *patterns: str) -> Any | None:
    normalised = {column: normal_key(column) for column in columns}
    for column, key in normalised.items():
        if all(pattern in key for pattern in patterns):
            return column
    return None


def value_at(row: pd.Series, column: Any | None) -> str:
    return repair_text(row.get(column, "")) if column is not None else ""


def parse_students(table: pd.DataFrame) -> list[dict[str, Any]]:
    table = table.dropna(how="all").copy()
    columns = list(table.columns)
    name_col = column_for(columns, "elevnavn") or column_for(columns, "navn")
    class_col = column_for(columns, "klasse") or column_for(columns, "hold")
    subject_cols = []
    wish_cols = []
    for index in (1, 2):
        subject_cols.append(
            column_for(columns, "fag", str(index))
            or column_for(columns, "subject", str(index))
            or (columns[index + 1] if len(columns) > index + 1 else None)
        )
        wish_cols.append(
            column_for(columns, "ønske", str(index))
            or column_for(columns, "ønsk", str(index))
            or column_for(columns, "onske", str(index))
            or column_for(columns, "wish", str(index))
        )
    if name_col is None or any(item is None for item in subject_cols):
        raise ValueError("Elevarket skal indeholde kolonnerne Navn, Klasse, Fag 1 og Fag 2.")

    title_col = column_for(columns, "projekttitel") or column_for(columns, "projekt", "titel") or column_for(columns, "projecttitle") or column_for(columns, "projektnavn")
    description_col = column_for(columns, "projektbeskrivelse") or column_for(columns, "projekt", "beskrivelse") or column_for(columns, "projectdescription") or column_for(columns, "beskrivelse") or column_for(columns, "description")
    students = []
    for row_number, (_, row) in enumerate(table.iterrows(), 1):
        name = value_at(row, name_col)
        subjects = [value_at(row, column) for column in subject_cols]
        if not name or not any(subjects):
            continue
        wishes = [normal_key(value_at(row, column)) for column in wish_cols]
        students.append(
            {
                "id": f"E{len(students) + 1:03d}",
                "name": name,
                "className": value_at(row, class_col) or "Ukendt klasse/hold",
                "subjects": subjects,
                "wishes": [wish for wish in wishes if wish],
                "subjectsWithLevel": subjects[:],
                "projectTitle": value_at(row, title_col),
                "projectDescription": value_at(row, description_col),
            }
        )
    if not students:
        raise ValueError("Der blev ikke fundet nogen elever i elevarket.")
    return students


def subject_from_course(value: Any, technical_subjects: dict[str, str]) -> str | None:
    subject = repair_text(value)
    if not subject or normal_key(subject).startswith("bro"):
        return None
    if normal_key(subject) in technical_subjects:
        return technical_subjects[normal_key(subject)]
    if "teknikfag" in normal_key(subject):
        # Bevar den konkrete retning, så fx "Digitalt design ... teknikfag"
        # kan matches direkte mod elevens fag uden en separat oversigt.
        subject = re.sub(r"\s+[ABC]\*?$", "", subject, flags=re.I)
        return subject.strip() or "Teknikfag"
    subject = re.sub(r"^\d+(?:vf)?\s*\.\s*", "", subject, flags=re.I)
    subject = re.sub(r"^\d+(?:vf)?\s+", "", subject, flags=re.I)
    subject = re.sub(r"\s+[ABC]\*?$", "", subject, flags=re.I)
    return subject.strip() or None


def parse_teachers(table: pd.DataFrame, technical: dict[str, str] | None = None) -> tuple[list[dict[str, Any]], dict[str, int]]:
    technical = technical or {}
    columns = list(table.columns)
    id_col = column_for(columns, "initial") or column_for(columns, "lærer")
    name_col = column_for(columns, "lærernavn") or column_for(columns, "navn")
    course_col = column_for(columns, "fag")
    capacity_col = column_for(columns, "max") or column_for(columns, "kapac")
    if id_col is None or course_col is None:
        raise ValueError("Lærerarket skal indeholde kolonnerne Initialer og Fag.")
    class_col = column_for(columns, "klasse")
    hold_columns = [column for column in columns if normal_key(column).startswith("hold")][:4]
    if not hold_columns and class_col is not None:
        hold_columns = [class_col]
    teachers_by_id: dict[str, dict[str, Any]] = {}
    capacities: dict[str, int] = {}
    subject_columns = [column for column in columns if normal_key(column).startswith("fag")]
    for _, row in table.iterrows():
        teacher_id = normal_key(row.get(id_col, ""))
        if not teacher_id:
            continue
        if capacity_col is not None:
            found = re.search(r"\d+", value_at(row, capacity_col))
            if found:
                capacities[teacher_id] = int(found.group())
        subjects = []
        course_values = [row.get(course_col)] + [row.get(column) for column in subject_columns if column != course_col]
        for course in course_values:
            subject = subject_from_course(course, technical)
            if class_col is not None and "teknikfag" in normal_key(course):
                detailed = subject_from_course(row.get(class_col), technical)
                if detailed and detailed != "Teknikfag":
                    subject = detailed
            if subject and canonical_subject(subject) not in {canonical_subject(item) for item in subjects}:
                subjects.append(subject)
        if not subjects:
            continue
        teachers_by_id.setdefault(teacher_id, {"id": teacher_id, "name": value_at(row, name_col) or teacher_id, "subjects": [], "holds": []})
        for subject in subjects:
            if canonical_subject(subject) not in {canonical_subject(item) for item in teachers_by_id[teacher_id]["subjects"]}:
                teachers_by_id[teacher_id]["subjects"].append(subject)
        for column in hold_columns:
            hold = value_at(row, column)
            if hold and normal_key(hold) not in {normal_key(item) for item in teachers_by_id[teacher_id]["holds"]}:
                teachers_by_id[teacher_id]["holds"].append(hold)
    teachers = sorted(teachers_by_id.values(), key=lambda item: item["id"])
    return teachers, capacities


def parse_teacher_upload(uploaded: Any, technical: dict[str, str] | None = None) -> tuple[list[dict[str, Any]], dict[str, int]]:
    """Læs både lærerlisten på ark 1 og max-tal på ark 2 i timefagfordeling.xlsx."""
    technical = technical or {}
    filename = uploaded.name.casefold()
    if not filename.endswith((".xlsx", ".xlsm", ".xls")):
        return parse_teachers(read_uploaded_table(uploaded), technical)
    sheets = pd.read_excel(io.BytesIO(uploaded.getvalue()), sheet_name=None, dtype=object)
    first = next(iter(sheets.values()))
    teachers, capacities = parse_teachers(first, technical)
    if len(sheets) > 1:
        second = list(sheets.values())[1]
        for _, row in second.iterrows():
            if len(row) < 2:
                continue
            teacher_id = normal_key(row.iloc[0])
            found = re.search(r"\d+", repair_text(row.iloc[1]))
            if teacher_id and found:
                value = int(found.group())
                capacities[teacher_id] = min(capacities.get(teacher_id, value), value)
    return teachers, capacities


def teacher_label(teacher_id: str | None, teachers_by_id: dict[str, dict[str, Any]]) -> str:
    if not teacher_id:
        return "Ikke tildelt"
    teacher = teachers_by_id.get(normal_key(teacher_id))
    if not teacher:
        return teacher_id
    return f"{teacher['name']} ({teacher['id']})" if normal_key(teacher["name"]) != teacher["id"] else teacher["id"]


def build_candidates(teachers: list[dict[str, Any]]) -> dict[str, list[str]]:
    candidates: dict[str, list[str]] = defaultdict(list)
    for teacher in teachers:
        for subject in teacher["subjects"]:
            if teacher["id"] not in candidates[canonical_subject(subject)]:
                candidates[canonical_subject(subject)].append(teacher["id"])
    return dict(candidates)


def wished_count(student: dict[str, Any], assigned: list[str | None]) -> int:
    return sum(bool(teacher_id and teacher_id in student["wishes"]) for teacher_id in assigned)


def optimize(
    students: list[dict[str, Any]],
    teachers: list[dict[str, Any]],
    capacities: dict[str, int],
    K: int,
    double_limit: int,
    use_global_k: bool,
    allow_over_capacity: bool,
    lock_teacher_max: bool,
    prioritize_pairs: bool,
    prioritize_classes: bool,
    attempts: int = 120,
    progress_callback: Any = None,
) -> dict[str, Any]:
    candidates = build_candidates(teachers)
    teacher_ids = [teacher["id"] for teacher in teachers]

    def candidate_for(subject: Any) -> list[str]:
        return candidates.get(canonical_subject(subject), [])

    def capacity_of(teacher_id: str) -> int:
        value = int(capacities.get(teacher_id, 0))
        return min(K, value) if use_global_k else value

    permit_over = use_global_k and allow_over_capacity and not lock_teacher_max
    best = None
    seed_base = K * 100003 + len(students) * 97 + double_limit * 7919
    demand = Counter(canonical_subject(subject) for student in students for subject in student["subjects"])

    def make_attempt(seed: int) -> dict[str, Any]:
        rng = __import__("random").Random(seed)
        assignments: list[list[str | None]] = [[None, None] for _ in students]
        loads = {teacher_id: 0 for teacher_id in teacher_ids}
        double_loads = {teacher_id: 0 for teacher_id in teacher_ids}

        def capacity_available(teacher_id: str, index: int) -> bool:
            return (
                teacher_id in assignments[index]
                or loads[teacher_id] < capacity_of(teacher_id)
                or (permit_over and loads[teacher_id] < K)
            )

        def pressure(subject: str) -> float:
            ids = candidate_for(subject)
            available = sum(K if permit_over else capacity_of(teacher_id) for teacher_id in ids)
            return demand[canonical_subject(subject)] / max(1, available)

        def can_share(index: int, slot: int, teacher_id: str) -> bool:
            other = assignments[index][1 - slot]
            if teacher_id != other:
                return True
            return assignments[index][0] == teacher_id and assignments[index][1] == teacher_id or double_loads[teacher_id] < double_limit

        def assign(index: int, slot: int, teacher_id: str | None) -> None:
            before = assignments[index][:]
            before_set = {item for item in before if item}
            before_double = before[0] if before[0] and before[0] == before[1] else None
            assignments[index][slot] = teacher_id
            after = assignments[index]
            after_set = {item for item in after if item}
            after_double = after[0] if after[0] and after[0] == after[1] else None
            for teacher_id_ in before_set | after_set:
                if teacher_id_ in before_set and teacher_id_ not in after_set:
                    loads[teacher_id_] -= 1
                if teacher_id_ not in before_set and teacher_id_ in after_set:
                    loads[teacher_id_] += 1
            if before_double and before_double != after_double:
                double_loads[before_double] -= 1
            if after_double and before_double != after_double:
                double_loads[after_double] += 1

        coverage = []
        for index, student in enumerate(students):
            options = []
            for slot in (0, 1):
                allowed = set(candidate_for(student["subjects"][slot]))
                options.extend({"slot": slot, "teacher": wish} for wish in dict.fromkeys(student["wishes"]) if wish in allowed)
            coverage.append({"index": index, "options": options, "pressure": max(pressure(item) for item in student["subjects"]), "scarcity": min(len(candidate_for(item)) for item in student["subjects"]), "tie": rng.random()})
        coverage.sort(key=lambda item: (-item["pressure"], len(item["options"]), item["scarcity"], item["tie"]))
        for item in coverage:
            index = item["index"]
            student = students[index]
            choices = [option for option in item["options"] if can_share(index, option["slot"], option["teacher"]) and capacity_available(option["teacher"], index)]
            choices.sort(key=lambda option: (
                int(not candidate_for(student["subjects"][1 - option["slot"]]) or len([item for item in candidate_for(student["subjects"][1 - option["slot"]]) if item != option["teacher"]]) == 0),
                loads[option["teacher"]],
                -len([item for item in candidate_for(student["subjects"][1 - option["slot"]]) if item != option["teacher"]]),
                rng.random(),
            ))
            if choices:
                assign(index, choices[0]["slot"], choices[0]["teacher"])

        remaining = []
        for index, student in enumerate(students):
            for slot in (0, 1):
                if not assignments[index][slot]:
                    subject = student["subjects"][slot]
                    remaining.append({"index": index, "slot": slot, "count": len(candidate_for(subject)), "pressure": pressure(subject), "tie": rng.random()})
        remaining.sort(key=lambda item: (-item["pressure"], item["count"], item["tie"]))
        for item in remaining:
            index, slot = item["index"], item["slot"]
            if assignments[index][slot]:
                continue
            student = students[index]
            choices = [teacher_id for teacher_id in candidate_for(student["subjects"][slot]) if can_share(index, slot, teacher_id) and capacity_available(teacher_id, index)]
            other_slot = 1 - slot
            def choice_key(teacher_id: str) -> tuple[Any, ...]:
                other = candidate_for(student["subjects"][other_slot])
                future_bad = 0 if any(item != teacher_id and capacity_available(item, index) for item in other) else 1
                return (max(0, loads[teacher_id] + 1 - capacity_of(teacher_id)), future_bad, int(teacher_id not in student["wishes"]), loads[teacher_id], rng.random())
            choices.sort(key=choice_key)
            if choices:
                assign(index, slot, choices[0])

        for _ in range(8):
            moved = False
            overloaded = sorted((item for item in loads if loads[item] > capacity_of(item)), key=lambda item: -loads[item])
            for teacher_id in overloaded:
                held = []
                for index, student in enumerate(students):
                    for slot in (0, 1):
                        if assignments[index][slot] == teacher_id:
                            before = wished_count(student, assignments[index])
                            held.append((2 if teacher_id in student["wishes"] and before == 1 else 1 if teacher_id in student["wishes"] else 0, rng.random(), index, slot))
                held.sort()
                for _, _, index, slot in held:
                    if loads[teacher_id] <= capacity_of(teacher_id):
                        break
                    student = students[index]
                    alternatives = [item for item in candidate_for(student["subjects"][slot]) if item != teacher_id and can_share(index, slot, item) and capacity_available(item, index)]
                    alternatives.sort(key=lambda item: (int(item not in student["wishes"]), loads[item], rng.random()))
                    if alternatives:
                        assign(index, slot, alternatives[0])
                        moved = True
            if not moved:
                break

        for _ in range(4):
            improved = False
            order = sorted(range(len(students)), key=lambda index: (wished_count(students[index], assignments[index]), rng.random()))
            for index in order:
                student = students[index]
                for slot in (0, 1):
                    current = assignments[index][slot]
                    if not current or current in student["wishes"]:
                        continue
                    preferred = [item for item in dict.fromkeys(student["wishes"]) if item in candidate_for(student["subjects"][slot]) and can_share(index, slot, item) and capacity_available(item, index)]
                    preferred.sort(key=lambda item: loads[item])
                    if preferred:
                        assign(index, slot, preferred[0])
                        improved = True
            if not improved:
                break

        for _ in range(3):
            improved = False
            for index, student in enumerate(students):
                wishes = list(dict.fromkeys(student["wishes"]))
                if len(wishes) < 2:
                    continue
                current_count = wished_count(student, assignments[index])
                pairs = [(wishes[0], wishes[1]), (wishes[1], wishes[0])]
                for pair in pairs:
                    if pair[0] not in candidate_for(student["subjects"][0]) or pair[1] not in candidate_for(student["subjects"][1]) or wished_count(student, list(pair)) <= current_count:
                        continue
                    changes = Counter(item for item in pair)
                    changes.subtract(item for item in assignments[index] if item)
                    if all(loads[item] + delta <= capacity_of(item) or (permit_over and loads[item] + delta <= K) for item, delta in changes.items()):
                        assign(index, 0, None)
                        assign(index, 1, None)
                        assign(index, 0, pair[0])
                        assign(index, 1, pair[1])
                        improved = True
                        break
            if not improved:
                break

        stats = score_solution(students, teachers, capacities, assignments, loads, K, use_global_k, prioritize_pairs, prioritize_classes)
        return {"assignments": assignments, "loads": loads, "double_loads": double_loads, "stats": stats, "K": K, "double_limit": double_limit, "use_global_k": use_global_k, "allow_over_capacity": permit_over, "lock_teacher_max": lock_teacher_max}

    attempts = max(1, int(attempts))
    if progress_callback:
        progress_callback(0.0)
    for attempt in range(attempts):
        candidate = make_attempt(seed_base + attempt * 7919)
        if best is None or candidate["stats"]["score"] < best["stats"]["score"]:
            best = candidate
        if progress_callback:
            progress_callback((attempt + 1) / attempts)
    return best


def update_algorithm_progress(progress_bar: Any, label: str, value: float) -> None:
    """Opdater en Streamlit-progressbar med en ensartet procenttekst."""
    bounded = max(0.0, min(1.0, float(value)))
    progress_bar.progress(bounded, text=f"{label}: {round(bounded * 100)} %")


def score_solution(students: list[dict[str, Any]], teachers: list[dict[str, Any]], capacities: dict[str, int], assignments: list[list[str | None]], loads: dict[str, int], K: int, use_global_k: bool, prioritize_pairs: bool, prioritize_classes: bool) -> dict[str, Any]:
    candidates = build_candidates(teachers)
    def cap(teacher_id: str) -> int:
        value = capacities.get(teacher_id, 0)
        return min(K, value) if use_global_k else value
    unassigned = sum(len([item for item in assigned if item is None]) > 0 for assigned in assignments)
    none = one = both = desired_total = capacity_blocked_slots = 0
    blocked_students = set()
    pair_counts = Counter()
    class_counts = Counter()
    for index, assigned in enumerate(assignments):
        student = students[index]
        count = wished_count(student, assigned)
        desired_total += count
        if count == 0:
            none += 1
        elif count == 1:
            one += 1
        else:
            both += 1
        for slot, teacher_id in enumerate(assigned):
            if not teacher_id:
                possible = candidates.get(canonical_subject(student["subjects"][slot]), [])
                if possible and all(loads[item] >= cap(item) for item in possible):
                    capacity_blocked_slots += 1
                    blocked_students.add(student["id"])
        for teacher_id in set(item for item in assigned if item):
            class_counts[f"{student.get('className', 'Ukendt')}|{teacher_id}"] += 1
        if assigned[0] and assigned[1] and assigned[0] != assigned[1]:
            pair_counts[f"{'|'.join(sorted(canonical_subject(item) for item in student['subjects']))}::{ '|'.join(sorted(assigned)) }"] += 1
    overload = sum(max(0, load - cap(teacher_id)) for teacher_id, load in loads.items())
    over_teachers = sum(load > cap(teacher_id) for teacher_id, load in loads.items())
    k_overload = sum(max(0, load - K) for load in loads.values()) if use_global_k else 0
    k_over_teachers = sum(load > K for load in loads.values()) if use_global_k else 0
    max_load = max(loads.values(), default=0)
    load_squares = sum(load * load for load in loads.values())
    pair_score = sum(value * (value - 1) // 2 for value in pair_counts.values())
    class_score = sum(value * (value - 1) // 2 for value in class_counts.values())
    score = k_overload * 1e15 + unassigned * 1e12 + overload * 1e9 + none * 1e6 + (len(students) * 2 - desired_total) * 1e4 - (pair_score * .05 if prioritize_pairs else 0) - (class_score * .05 if prioritize_classes else 0) + max_load * 100 + load_squares
    return {"score": score, "unassigned": int(unassigned), "none": none, "one": one, "both": both, "desired_total": desired_total, "overload": overload, "over_teachers": over_teachers, "k_overload": k_overload, "k_over_teachers": k_over_teachers, "max_load": max_load, "same_teacher_pair_score": pair_score, "same_class_teacher_score": class_score, "capacity_blocked_slots": capacity_blocked_slots, "capacity_blocked_students": len(blocked_students)}


def subject_stats(students: list[dict[str, Any]], teachers: list[dict[str, Any]], capacities: dict[str, int], K: int, use_global_k: bool) -> pd.DataFrame:
    candidates = build_candidates(teachers)
    grouped: dict[str, dict[str, Any]] = {}
    for student in students:
        for subject in student["subjects"]:
            key = canonical_subject(subject)
            grouped.setdefault(key, {"Fag": subject, "Elever": set()})["Elever"].add(student["id"])
    rows = []
    for key, item in grouped.items():
        ids = candidates.get(key, [])
        capacity = sum(min(K, capacities.get(teacher_id, 0)) if use_global_k else capacities.get(teacher_id, 0) for teacher_id in ids)
        rows.append({"Fag": item["Fag"], "Elever": len(item["Elever"]), "Lærere": len(ids), "Kapacitet": capacity, "Elever pr. lærer": round(len(item["Elever"]) / len(ids), 1) if ids else None, "Lærerinitialer": ", ".join(ids), "Mangler kapacitet": max(0, len(item["Elever"]) - capacity)})
    return pd.DataFrame(rows).sort_values(["Mangler kapacitet", "Elever pr. lærer", "Elever"], ascending=[False, False, False]) if rows else pd.DataFrame()


def make_export(
    students: list[dict[str, Any]],
    teachers: list[dict[str, Any]],
    solution: dict[str, Any],
    capacities: dict[str, int],
    use_global_k: bool,
    selected_sheets: list[str] | None = None,
) -> bytes:
    teacher_map = {teacher["id"]: teacher for teacher in teachers}
    rows = []
    for index, student in enumerate(students):
        assigned = solution["assignments"][index]
        rows.append({"Elev-ID": student["id"], "Elev": student["name"], "Klasse": student.get("className", ""), "Fag 1": student["subjects"][0], "Vejleder fag 1": teacher_label(assigned[0], teacher_map), "Vejleder-ID 1": assigned[0] or "", "Fag 2": student["subjects"][1], "Vejleder fag 2": teacher_label(assigned[1], teacher_map), "Vejleder-ID 2": assigned[1] or "", "Ønsker opfyldt": wished_count(student, assigned), "Projekttitel": student.get("projectTitle", ""), "Projektbeskrivelse": student.get("projectDescription", "")})
    distribution = pd.DataFrame(rows)
    load_rows = []
    for teacher in sorted(teachers, key=lambda item: teacher_label(item["id"], teacher_map)):
        load = solution["loads"].get(teacher["id"], 0)
        limit = min(solution["K"], capacities.get(teacher["id"], 0)) if use_global_k else capacities.get(teacher["id"], 0)
        load_rows.append({"Lærer": teacher_label(teacher["id"], teacher_map), "Initialer": teacher["id"], "Fag": " · ".join(teacher["subjects"]), "Antal elever": load, "Max": limit, "Over max": max(0, load - limit), "Dobbeltvejledninger": solution["double_loads"].get(teacher["id"], 0)})
    loads = pd.DataFrame(load_rows)
    stats = subject_stats(students, teachers, capacities, solution["K"], use_global_k)
    unassigned = distribution[distribution[["Vejleder-ID 1", "Vejleder-ID 2"]].eq("").any(axis=1)]
    if selected_sheets is not None:
        sheet_frames = {
            "Fordeling": distribution,
            "Lærerbelastning": loads,
            "Fagstatistik": stats,
            "Ikke tildelte": unassigned,
            "Elevdata": students_to_frame(students, teachers),
            "Lærerdata": teachers_to_frame(teachers, capacities),
        }
        sheet_order = ["Fordeling", "Lærerbelastning", "Fagstatistik", "Ikke tildelte", "Elevdata", "Lærerdata"]
        selected = [name for name in sheet_order if name in selected_sheets]
        if not selected:
            raise ValueError("Vælg mindst ét ark til Excel-filen.")
        selected_output = io.BytesIO()
        with pd.ExcelWriter(selected_output, engine="openpyxl") as writer:
            for sheet_name in selected:
                sheet_frames[sheet_name].to_excel(writer, index=False, sheet_name=sheet_name)
            for sheet in writer.book.worksheets:
                sheet.freeze_panes = "A2"
                sheet.auto_filter.ref = sheet.dimensions
                for column in sheet.columns:
                    width = min(45, max(12, max(len(str(cell.value or "")) for cell in column) + 2))
                    sheet.column_dimensions[column[0].column_letter].width = width
        return selected_output.getvalue()
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        distribution.to_excel(writer, index=False, sheet_name="Fordeling")
        loads.to_excel(writer, index=False, sheet_name="Lærerbelastning")
        stats.to_excel(writer, index=False, sheet_name="Fagstatistik")
        unassigned.to_excel(writer, index=False, sheet_name="Ikke tildelte")
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for column in sheet.columns:
                width = min(45, max(12, max(len(str(cell.value or "")) for cell in column) + 2))
                sheet.column_dimensions[column[0].column_letter].width = width
    return output.getvalue()


def _clock_label(value: datetime) -> str:
    return value.strftime("%H:%M")


def _schedule_pair_label(teacher_ids: tuple[str, ...], teacher_map: dict[str, dict[str, Any]]) -> str:
    return " + ".join(teacher_label(teacher_id, teacher_map) for teacher_id in teacher_ids) or "Mangler vejleder"


def make_schedule(
    students: list[dict[str, Any]],
    teachers: list[dict[str, Any]],
    solution: dict[str, Any],
    start_time: dt_time,
    end_time: dt_time,
    student_minutes: int,
    pause_count: int,
    pause_minutes: int,
    transition_minutes: int,
    group_pairs: bool,
    search_attempts: int = 80,
    progress_callback: Any = None,
) -> dict[str, Any]:
    """Lav en parallel vejledningsplan ud fra den aktuelle fordeling."""
    if start_time >= end_time:
        raise ValueError("Sluttidspunktet skal ligge efter starttidspunktet.")
    if student_minutes < 1:
        raise ValueError("Vejledningstiden pr. elev skal være mindst ét minut.")
    if pause_count < 0 or pause_minutes < 0 or transition_minutes < 0:
        raise ValueError("Pauser og skiftetid kan ikke være negative.")
    search_attempts = max(1, int(search_attempts))
    if progress_callback:
        progress_callback(0.0)

    teacher_map = {teacher["id"]: teacher for teacher in teachers}
    sessions = []
    for index, student in enumerate(students):
        assigned = solution["assignments"][index]
        if any(not teacher_id for teacher_id in assigned):
            raise ValueError("Fordelingen indeholder elever uden to vejledere. Ret fordelingen, før tidsplanen laves.")
        pair = tuple(sorted(set(assigned)))
        sessions.append({"student": student, "assigned": assigned, "pair": pair, "index": index})
    if not sessions:
        raise ValueError("Der er ingen elever i fordelingen.")

    if group_pairs:
        grouped: dict[tuple[str, ...], list[dict[str, Any]]] = {}
        first_index = {}
        for session in sessions:
            grouped.setdefault(session["pair"], []).append(session)
            first_index.setdefault(session["pair"], session["index"])
        pairs = list(grouped)
        rng = __import__("random").Random(20260913)
        candidate_orders = [pairs]
        candidate_orders.extend(rng.sample(pairs, len(pairs)) for _ in range(search_attempts - 1))
        best_order = pairs
        best_order_score = None
        for order_index, candidate_order in enumerate(candidate_orders):
            next_round: dict[str, int] = {}
            starts = []
            for pair in candidate_order:
                block_start = max((next_round.get(teacher_id, 0) for teacher_id in pair), default=0)
                starts.append(block_start)
                for teacher_id in pair:
                    next_round[teacher_id] = block_start + len(grouped[pair])
            score = (max((starts[i] + len(grouped[pair]) for i, pair in enumerate(candidate_order)), default=0), sum(starts), tuple(first_index[pair] for pair in candidate_order))
            if best_order_score is None or score < best_order_score:
                best_order = candidate_order
                best_order_score = score
            if progress_callback:
                progress_callback(0.65 * (order_index + 1) / len(candidate_orders))
        pair_order = best_order
        ordered_sessions = [session for pair in pair_order for session in grouped[pair]]
    else:
        ordered_sessions = sessions
        if progress_callback:
            progress_callback(0.65)

    # Hver lærer er en ressource. Elever med forskellige lærerpar kan derfor
    # ligge samtidig, mens et lærerpar aldrig får overlappende elever.
    teacher_next_round: dict[str, int] = {}
    for session in ordered_sessions:
        session["round"] = max((teacher_next_round.get(teacher_id, 0) for teacher_id in session["pair"]), default=0)
        for teacher_id in session["pair"]:
            teacher_next_round[teacher_id] = session["round"] + 1
    last_round = max(session["round"] for session in ordered_sessions)
    actual_pause_count = min(pause_count, last_round)
    pause_after = {
        max(1, min(last_round, round(index * (last_round + 1) / (actual_pause_count + 1))))
        for index in range(1, actual_pause_count + 1)
    }
    while len(pause_after) < actual_pause_count:
        pause_after.add(next(position for position in range(1, last_round + 1) if position not in pause_after))

    plan_start = datetime.combine(date.today(), start_time)
    day_end = datetime.combine(date.today(), end_time)
    round_starts = []
    timeline_rows = []
    cursor = plan_start
    for round_number in range(last_round + 1):
        round_starts.append(cursor)
        round_end = cursor + timedelta(minutes=student_minutes)
        if round_number < last_round:
            next_start = round_end + timedelta(minutes=transition_minutes)
            if round_number + 1 in pause_after:
                pause_end = next_start + timedelta(minutes=pause_minutes)
                timeline_rows.append({
                    "_sort": (round_number, 1), "Type": "Pause", "Start": _clock_label(next_start),
                    "Slut": _clock_label(pause_end), "Varighed (min.)": pause_minutes, "Elev": "", "Klasse": "",
                    "Lærer(e)": "", "Information": "Pause",
                })
                next_start = pause_end
            cursor = next_start
        else:
            cursor = round_end
    if cursor > day_end:
        required = int((cursor - plan_start).total_seconds() // 60)
        available = int((day_end - plan_start).total_seconds() // 60)
        raise ValueError(f"Tidsplanen kræver mindst {required} minutter, men tidsrummet rummer kun {available} minutter.")

    session_rows = []
    teacher_rows = []
    pair_counts = Counter()
    pair_first_last: dict[tuple[str, ...], list[str]] = {}
    for session in ordered_sessions:
        session_start = round_starts[session["round"]]
        session_end = session_start + timedelta(minutes=student_minutes)
        student = session["student"]
        assigned = session["assigned"]
        pair = session["pair"]
        pair_label = _schedule_pair_label(pair, teacher_map)
        time_label = f"{_clock_label(session_start)}–{_clock_label(session_end)}"
        session_rows.append({
            "Start": _clock_label(session_start), "Slut": _clock_label(session_end), "Tid": time_label,
            "Elev": student["name"], "Klasse": student.get("className", ""), "Fag 1": student["subjects"][0],
            "Vejleder 1": teacher_label(assigned[0], teacher_map), "Fag 2": student["subjects"][1],
            "Vejleder 2": teacher_label(assigned[1], teacher_map), "Lærerpar": pair_label,
            "Projekttitel": student.get("projectTitle", ""),
        })
        timeline_rows.append({
            "_sort": (session["round"], 0), "Type": "Vejledning", "Start": _clock_label(session_start),
            "Slut": _clock_label(session_end), "Varighed (min.)": student_minutes, "Elev": student["name"],
            "Klasse": student.get("className", ""), "Lærer(e)": pair_label,
            "Information": student.get("projectTitle", "") or "SOP-vejledning",
        })
        pair_counts[pair] += 1
        pair_first_last.setdefault(pair, [time_label, time_label])
        pair_first_last[pair][1] = time_label
        for teacher_id in dict.fromkeys(assigned):
            subject_slots = [subject for slot, subject in enumerate(student["subjects"][:2]) if assigned[slot] == teacher_id]
            co_teachers = list(dict.fromkeys(other_id for other_id in assigned if other_id != teacher_id))
            teacher_rows.append({
                "Start": _clock_label(session_start), "Slut": _clock_label(session_end), "Tid": time_label,
                "Lærer": teacher_label(teacher_id, teacher_map), "Initialer": teacher_id, "Elev": student["name"],
                "Klasse": student.get("className", ""), "Fag": " · ".join(dict.fromkeys(subject_slots)),
                "Medvejleder": " + ".join(teacher_label(other_id, teacher_map) for other_id in co_teachers) or "—",
                "Projekttitel": student.get("projectTitle", ""),
            })
    timeline_rows.sort(key=lambda row: row["_sort"])
    for row in timeline_rows:
        row.pop("_sort", None)
    session_rows.sort(key=lambda row: row["Start"])
    teacher_rows.sort(key=lambda row: (row["Start"], row["Lærer"]))

    student_columns = ["Start", "Slut", "Tid", "Elev", "Klasse", "Fag 1", "Vejleder 1", "Fag 2", "Vejleder 2", "Lærerpar", "Projekttitel"]
    teacher_columns = ["Start", "Slut", "Tid", "Lærer", "Initialer", "Elev", "Klasse", "Fag", "Medvejleder", "Projekttitel"]
    timeline_columns = ["Type", "Start", "Slut", "Varighed (min.)", "Elev", "Klasse", "Lærer(e)", "Information"]
    pair_rows = [
        {"Lærerpar": _schedule_pair_label(pair, teacher_map), "Antal elever": count, "Tidsblok": " → ".join(pair_first_last[pair])}
        for pair, count in sorted(pair_counts.items(), key=lambda item: pair_first_last[item[0]][0])
    ]
    total_minutes = int((cursor - plan_start).total_seconds() // 60)
    if progress_callback:
        progress_callback(1.0)
    return {
        "students": pd.DataFrame(session_rows, columns=student_columns),
        "teachers": pd.DataFrame(teacher_rows, columns=teacher_columns),
        "timeline": pd.DataFrame(timeline_rows, columns=timeline_columns),
        "pairs": pd.DataFrame(pair_rows, columns=["Lærerpar", "Antal elever", "Tidsblok"]),
        "settings": {
            "Start": _clock_label(datetime.combine(date.today(), start_time)),
            "Slut": _clock_label(datetime.combine(date.today(), end_time)),
            "Minutter pr. elev": student_minutes,
            "Antal pauser": actual_pause_count,
            "Minutter pr. pause": pause_minutes,
            "Minutter mellem elever": transition_minutes,
            "Lærerpar samlet": "Ja" if group_pairs else "Nej",
            "Optimeringsdybde": search_attempts,
            "Planlagt tidsforbrug (min.)": total_minutes,
        },
    }


def make_teacher_gantt_html(schedule: dict[str, Any], selected_teacher: str | None = None) -> str:
    """Lav en kompakt, visuel Gantt-plan for én eller alle lærere."""
    frame = schedule["teachers"]
    if frame.empty:
        return '<div class="gantt-empty">Der er ingen vejledninger at vise.</div>'

    def clock_minutes(value: Any) -> int:
        parsed = datetime.strptime(str(value), "%H:%M")
        return parsed.hour * 60 + parsed.minute

    settings = schedule["settings"]
    plan_start = clock_minutes(settings["Start"])
    plan_end = clock_minutes(settings["Slut"])
    visible_end = max(plan_end, max(clock_minutes(value) for value in frame["Slut"]))
    total_minutes = max(1, visible_end - plan_start)
    names = sorted(frame["Lærer"].dropna().unique().tolist(), key=str.casefold)
    if selected_teacher and selected_teacher != "Alle lærere":
        names = [selected_teacher] if selected_teacher in names else []

    palette = ["#0f817a", "#df8b2d", "#b44438", "#5367a8", "#8d5aa7", "#4f8b69", "#bd5f83"]

    def position(value: int) -> float:
        return max(0.0, min(100.0, (value - plan_start) / total_minutes * 100))

    ticks = list(range(plan_start, visible_end + 1, 30))
    if ticks[-1] != visible_end:
        ticks.append(visible_end)
    tick_markup = "".join(
        f'<span class="gantt-tick" style="left:{position(tick):.3f}%">{tick // 60:02d}:{tick % 60:02d}</span>'
        for tick in ticks
    )
    rows = []
    for name in names:
        teacher_rows = frame[frame["Lærer"] == name].sort_values(["Start", "Slut"])
        blocks = []
        for _, row in teacher_rows.iterrows():
            start = clock_minutes(row["Start"])
            end = clock_minutes(row["Slut"])
            subject = str(row.get("Fag", "Vejledning"))
            student = str(row.get("Elev", "Elev"))
            color = palette[sum(ord(char) for char in subject) % len(palette)]
            tooltip = html.escape(f"{row['Start']}–{row['Slut']} · {student} · {subject}", quote=True)
            blocks.append(
                f'<div class="gantt-block" title="{tooltip}" style="left:{position(start):.3f}%;width:{max(0.8, position(end) - position(start)):.3f}%;background:{color}">'
                f'<span>{html.escape(row["Start"])}–{html.escape(row["Slut"])}</span>'
                f'<strong>{html.escape(student)}</strong><small>{html.escape(subject)}</small></div>'
            )
        rows.append(
            f'<div class="gantt-row"><div class="gantt-name">{html.escape(name)}<small>{len(teacher_rows)} elev(er)</small></div>'
            f'<div class="gantt-track">{"".join(blocks)}</div></div>'
        )

    if not rows:
        return '<div class="gantt-empty">Vælg en lærer med planlagte vejledninger.</div>'
    return f'''<div class="teacher-gantt">
<style>
.teacher-gantt{{border:1px solid #dce4e2;border-radius:14px;background:#fff;padding:16px;overflow:auto;box-shadow:0 8px 24px #1933300b}}
.gantt-intro{{display:flex;justify-content:space-between;gap:16px;align-items:flex-end;margin-bottom:14px;color:#66777f;font-size:13px}}
.gantt-intro strong{{display:block;color:#172126;font:700 21px Georgia,serif;margin-bottom:3px}}
.gantt-scale{{margin-left:178px;height:28px;min-width:760px;position:relative;border-bottom:1px solid #dce4e2}}
.gantt-tick{{position:absolute;bottom:5px;transform:translateX(-50%);font-size:11px;color:#66777f;white-space:nowrap}}
.gantt-row{{display:flex;min-width:938px;min-height:72px;border-bottom:1px solid #edf1f0}}
.gantt-row:last-child{{border-bottom:0}}
.gantt-name{{width:162px;flex:0 0 162px;padding:15px 12px 8px 0;font-weight:700;color:#172126;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.gantt-name small{{display:block;color:#66777f;font-size:11px;font-weight:400;margin-top:3px}}
.gantt-track{{position:relative;flex:1;margin:9px 0;background:repeating-linear-gradient(to right,#f4f8f7 0,#f4f8f7 calc(8.333% - 1px),#dfe9e6 calc(8.333% - 1px),#dfe9e6 8.333%);border-radius:8px;min-height:54px}}
.gantt-block{{position:absolute;top:5px;height:44px;border-radius:7px;padding:4px 7px;box-sizing:border-box;color:white;overflow:hidden;box-shadow:0 2px 5px #1721262b;line-height:1.15;font-size:10px;min-width:25px}}
.gantt-block span,.gantt-block strong,.gantt-block small{{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.gantt-block strong{{font-size:11px;margin-top:2px}}
.gantt-block small{{opacity:.88;margin-top:2px}}
</style>
<div class="gantt-intro"><div><strong>Visuel lærerplan</strong><span>Blokkene viser elevens vejledningstid. Hold musen over en blok for detaljer.</span></div><span>{html.escape(str(settings["Start"]))}–{html.escape(str(settings["Slut"]))}</span></div>
<div class="gantt-scale">{tick_markup}</div>
{"".join(rows)}
</div>'''


def make_schedule_excel(schedule: dict[str, Any]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        schedule["students"].to_excel(writer, index=False, sheet_name="Elevplan")
        schedule["teachers"].to_excel(writer, index=False, sheet_name="Lærerplan")
        schedule["timeline"].to_excel(writer, index=False, sheet_name="Tidslinje")
        schedule["pairs"].to_excel(writer, index=False, sheet_name="Lærerpar")
        pd.DataFrame([schedule["settings"]]).to_excel(writer, index=False, sheet_name="Indstillinger")
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for column in sheet.columns:
                sheet.column_dimensions[column[0].column_letter].width = min(55, max(12, max(len(str(cell.value or "")) for cell in column) + 2))
    return output.getvalue()


def make_schedule_html(schedule: dict[str, Any]) -> str:
    settings = schedule["settings"]
    setting_lines = "".join(f"<li><strong>{html.escape(str(key))}:</strong> {html.escape(str(value))}</li>" for key, value in settings.items())
    gantt_html = make_teacher_gantt_html(schedule, "Alle lærere")
    return f"""<!doctype html>
<html lang="da">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Vejledningsplan</title>
<style>
body{{margin:0;background:#f7fbf9;color:#172126;font:15px/1.45 Arial,sans-serif}}
main{{max-width:1400px;margin:auto;padding:28px 20px 60px}}
h1,h2{{font-family:Georgia,serif}} h1{{font-size:38px;margin:0 0 8px}} h2{{margin-top:30px}}
.intro,.card{{background:white;border:1px solid #dce4e2;border-radius:14px;padding:18px;margin:14px 0;box-shadow:0 8px 24px #1933300b}}
.intro p{{margin:4px 0;color:#66777f}} ul{{display:flex;flex-wrap:wrap;gap:8px 28px;padding-left:20px}}
.table-wrap{{overflow:auto;background:white;border:1px solid #dce4e2;border-radius:12px}}
table{{border-collapse:collapse;width:100%;min-width:760px}} th,td{{padding:9px 11px;border-bottom:1px solid #e2e9e7;text-align:left;vertical-align:top}} th{{background:#e7f4f1;color:#075f5b;white-space:nowrap}} tr:last-child td{{border-bottom:0}}
@media print{{body{{background:white}} main{{padding:0}} .card,.table-wrap{{box-shadow:none}} h2{{break-before:page}}}}
</style></head>
<body><main>
<section class="intro"><h1>Vejledningsplan</h1><p>Detaljeret tidsplan for elever og lærere baseret på den beregnede SOP-fordeling.</p><ul>{setting_lines}</ul></section>
<section class="card"><h2>Visuel lærerplan</h2>{gantt_html}</section>
<section class="card"><h2>Elevplan</h2><div class="table-wrap">{schedule["students"].to_html(index=False, escape=True, border=0)}</div></section>
<section class="card"><h2>Lærerplan</h2><div class="table-wrap">{schedule["teachers"].to_html(index=False, escape=True, border=0)}</div></section>
<section class="card"><h2>Samlet tidslinje</h2><div class="table-wrap">{schedule["timeline"].to_html(index=False, escape=True, border=0)}</div></section>
<section class="card"><h2>Lærerpar</h2><div class="table-wrap">{schedule["pairs"].to_html(index=False, escape=True, border=0)}</div></section>
</main></body></html>"""


def replace_docx_placeholders(template: bytes, values: dict[str, str]) -> bytes:
    from docx import Document
    source = io.BytesIO(template)
    document = Document(source)
    paragraphs = list(document.paragraphs)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
    for paragraph in paragraphs:
        full = "".join(run.text for run in paragraph.runs)
        for placeholder, replacement in values.items():
            if placeholder in full:
                full = full.replace(placeholder, replacement)
        if paragraph.runs:
            paragraph.runs[0].text = full
            for run in paragraph.runs[1:]:
                run.text = ""
    result = io.BytesIO()
    document.save(result)
    return result.getvalue()


def make_docx_zip(students: list[dict[str, Any]], teachers: list[dict[str, Any]], solution: dict[str, Any]) -> bytes:
    teacher_map = {teacher["id"]: teacher for teacher in teachers}
    files = io.BytesIO()
    with zipfile.ZipFile(files, "w", zipfile.ZIP_DEFLATED) as archive:
        for index, student in enumerate(students):
            assigned = solution["assignments"][index]
            values = {"Elevnavn": student["name"], "Klasse": student.get("className", ""), "Fag1 og niveau": student["subjects"][0], "Vejleder fag 1": teacher_label(assigned[0], teacher_map), "Fag2 og niveau": student["subjects"][1], "Vejleder fag 2": teacher_label(assigned[1], teacher_map)}
            if DEFAULT_TEMPLATE.exists():
                template = DEFAULT_TEMPLATE.read_bytes()
                content = replace_docx_placeholders(template, values)
            else:
                from docx import Document
                document = Document()
                document.add_heading(f"SOP – {student['name']}", 0)
                document.add_paragraph(f"Klasse: {student.get('className', '')}")
                document.add_paragraph(f"{student['subjects'][0]}: {teacher_label(assigned[0], teacher_map)}")
                document.add_paragraph(f"{student['subjects'][1]}: {teacher_label(assigned[1], teacher_map)}")
                output = io.BytesIO(); document.save(output); content = output.getvalue()
            filename = re.sub(r"[<>:\"/\\|?*]", "-", f"SOP - {student['name']}.docx")
            archive.writestr(filename, content)
    return files.getvalue()


def students_to_frame(students: list[dict[str, Any]], teachers: list[dict[str, Any]] | None = None) -> pd.DataFrame:
    """Lav en redigerbar tabel; vis lærerønsker som læsbare navne, når lærerlisten findes."""
    teacher_map = {teacher["id"]: teacher for teacher in (teachers or [])}
    rows = []
    for student in students:
        subjects = list(student.get("subjects", [])) + ["", ""]
        wishes = [teacher_label(normal_key(wish), teacher_map) if normal_key(wish) in teacher_map else repair_text(wish) for wish in student.get("wishes", [])]
        if not wishes:
            wishes = [NO_WISHES_LABEL]
        wishes = wishes + [""]
        rows.append(
            {
                "Elev-ID": student.get("id", ""),
                "Elevnavn": student.get("name", ""),
                "Klasse": student.get("className", ""),
                "Fag 1": subjects[0],
                "Fag 2": subjects[1],
                "Ønskevejleder 1": wishes[0],
                "Ønskevejleder 2": wishes[1],
                "Projekttitel": student.get("projectTitle", ""),
                "Projektbeskrivelse": student.get("projectDescription", ""),
            }
        )
    return pd.DataFrame(rows, columns=["Elev-ID", "Elevnavn", "Klasse", "Fag 1", "Fag 2", "Ønskevejleder 1", "Ønskevejleder 2", "Projekttitel", "Projektbeskrivelse"])


def wish_column_config(teachers: list[dict[str, Any]]) -> dict[str, Any]:
    """Giv elevarket en tydelig dropdown med alle lærere og ingen ønsker."""
    teacher_map = {teacher["id"]: teacher for teacher in teachers}
    options = [NO_WISHES_LABEL] + sorted((teacher_label(teacher_id, teacher_map) for teacher_id in teacher_map), key=str.casefold)
    return {
        "Ønskevejleder 1": st.column_config.SelectboxColumn("Ønskevejleder 1", options=options, required=False),
        "Ønskevejleder 2": st.column_config.SelectboxColumn("Ønskevejleder 2", options=options, required=False),
    }


def frame_to_students(frame: pd.DataFrame) -> list[dict[str, Any]]:
    if frame.empty:
        return []
    students = parse_students(frame)
    ids = [normal_key(value) for value in frame.get("Elev-ID", pd.Series(dtype=object)).tolist()]
    for index, student in enumerate(students):
        if index < len(ids) and ids[index]:
            student["id"] = ids[index]
    return students


def student_data_signature(students: list[dict[str, Any]]) -> list[tuple[Any, ...]]:
    """Sammenlign kun de felter, der kan redigeres i elevarket.

    Niveauoplysninger og forskel på store/små bogstaver i Elev-ID må ikke få
    editoren til at markere data som ændret ved hver genkørsel.
    """
    return [
        (
            normal_key(student.get("id", "")),
            repair_text(student.get("name", "")),
            repair_text(student.get("className", "")),
            tuple(repair_text(subject) for subject in list(student.get("subjects", []))[:2]),
            tuple(normal_key(wish) for wish in student.get("wishes", []) if normal_key(wish)),
            repair_text(student.get("projectTitle", "")),
            repair_text(student.get("projectDescription", "")),
        )
        for student in students
    ]


def teachers_to_frame(
    teachers: list[dict[str, Any]],
    capacities: dict[str, int],
    subject_count: int | None = None,
    hold_count: int | None = None,
) -> pd.DataFrame:
    # Tre fagkolonner er standard; brugeren kan udvide ved behov.
    existing_subject_count = max((len(teacher.get("subjects", [])) for teacher in teachers), default=0)
    subject_count = max(3, subject_count or 3, existing_subject_count)
    existing_hold_count = max((len(teacher.get("holds", [])) for teacher in teachers), default=0)
    # Op til fire hold kan vises. Eksisterende holdkolonner bevares, hvis brugeren
    # midlertidigt vælger et lavere antal i kontrollen.
    hold_count = min(4, max(0, hold_count if hold_count is not None else 4, existing_hold_count))
    columns = ["Initialer", "Lærernavn"] + [f"Fag {index}" for index in range(1, subject_count + 1)] + [f"Hold {index}" for index in range(1, hold_count + 1)] + ["Max"]
    rows = []
    for teacher in teachers:
        subjects = list(teacher.get("subjects", []))
        row = {"Initialer": teacher.get("id", ""), "Lærernavn": teacher.get("name", "")}
        row.update({f"Fag {index}": subjects[index - 1] if index <= len(subjects) else "" for index in range(1, subject_count + 1)})
        holds = list(teacher.get("holds", []))
        row.update({f"Hold {index}": holds[index - 1] if index <= len(holds) else "" for index in range(1, hold_count + 1)})
        row["Max"] = int(capacities.get(teacher.get("id", ""), 0))
        rows.append(row)
    return pd.DataFrame(rows, columns=columns)


def frame_to_teachers(frame: pd.DataFrame, technical: dict[str, str] | None = None) -> tuple[list[dict[str, Any]], dict[str, int]]:
    if frame.empty:
        return [], {}
    technical = technical or {}
    subject_columns = [column for column in frame.columns if normal_key(column).startswith("fag")]
    hold_columns = [column for column in frame.columns if normal_key(column).startswith("hold")]
    teachers_by_id: dict[str, dict[str, Any]] = {}
    capacities: dict[str, int] = {}
    for _, row in frame.iterrows():
        teacher_id = normal_key(row.get("Initialer", ""))
        if not teacher_id:
            continue
        name = value_at(row, "Lærernavn") or teacher_id
        teacher = teachers_by_id.setdefault(teacher_id, {"id": teacher_id, "name": name, "subjects": [], "holds": []})
        if teacher["name"] == teacher_id and name != teacher_id:
            teacher["name"] = name
        for column in subject_columns:
            subject = subject_from_course(row.get(column), technical)
            if subject and canonical_subject(subject) not in {canonical_subject(item) for item in teacher["subjects"]}:
                teacher["subjects"].append(subject)
        for column in hold_columns:
            hold = value_at(row, column)
            if hold and normal_key(hold) not in {normal_key(item) for item in teacher["holds"]}:
                teacher["holds"].append(hold)
        raw_max = value_at(row, "Max")
        found = re.search(r"\d+", raw_max)
        capacities[teacher_id] = int(found.group()) if found else 0
    return sorted(teachers_by_id.values(), key=lambda item: item["id"]), capacities


def resolve_wishes(students: list[dict[str, Any]], teachers: list[dict[str, Any]]) -> None:
    """Tillad både initialer og lærernavne i ønskekolonnerne."""
    lookup = {}
    teacher_map = {teacher["id"]: teacher for teacher in teachers}
    for teacher in teachers:
        lookup[normal_key(teacher["id"])] = teacher["id"]
        lookup[normal_key(teacher.get("name", ""))] = teacher["id"]
        # Elevarket viser etiketten "Navn (initialer)"; den skal også kunne læses tilbage.
        lookup[normal_key(teacher_label(teacher["id"], teacher_map))] = teacher["id"]
    no_wishes_key = normal_key(NO_WISHES_LABEL)
    for student in students:
        resolved = []
        for wish in student.get("wishes", []):
            key = normal_key(wish)
            if not key or key == no_wishes_key:
                continue
            resolved.append(lookup.get(key, key))
        student["wishes"] = resolved


def input_validation(students: list[dict[str, Any]], teachers: list[dict[str, Any]]) -> dict[str, list[dict[str, str]]]:
    teacher_ids = {teacher["id"] for teacher in teachers}
    candidates = build_candidates(teachers)
    teacher_map = {teacher["id"]: teacher for teacher in teachers}
    unknown_wishes = []
    missing_subjects = []
    incompatible_wishes = []
    partial_wishes = []
    for student in students:
        for wish in student.get("wishes", []):
            if wish not in teacher_ids:
                unknown_wishes.append({"Elev": student["name"], "Ønskevejleder": wish})
        for subject in student.get("subjects", []):
            if subject and not candidates.get(canonical_subject(subject), []):
                missing_subjects.append({"Elev": student["name"], "Fag": subject})
        for wish in student.get("wishes", []):
            if wish in teacher_ids:
                covered = [subject for subject in student.get("subjects", []) if wish in candidates.get(canonical_subject(subject), [])]
                uncovered = [subject for subject in student.get("subjects", []) if subject not in covered]
                if not covered:
                    incompatible_wishes.append({"Elev": student["name"], "Ønskevejleder": teacher_label(wish, teacher_map), "Elevens fag": " · ".join(student.get("subjects", [])), "Status": "Underviser ikke i elevens fag"})
                elif uncovered:
                    partial_wishes.append({"Elev": student["name"], "Ønskevejleder": teacher_label(wish, teacher_map), "Fag uden match": " · ".join(uncovered), "Status": "Underviser kun i ét af elevens fag"})
    def unique_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
        seen = set()
        result = []
        for row in rows:
            marker = tuple(sorted(row.items()))
            if marker not in seen:
                seen.add(marker)
                result.append(row)
        return result

    # Saml flere ukendte ønsker fra samme elev i én række, så eleven ikke
    # optræder dobbelt i advarslen.
    unknown_by_student: dict[str, list[str]] = defaultdict(list)
    for row in unknown_wishes:
        if row["Ønskevejleder"] not in unknown_by_student[row["Elev"]]:
            unknown_by_student[row["Elev"]].append(row["Ønskevejleder"])
    unknown_wishes = [
        {"Elev": student_name, "Ønskevejleder": " · ".join(wishes)}
        for student_name, wishes in unknown_by_student.items()
    ]

    return {
        "unknown_wishes": unknown_wishes,
        "missing_subjects": unique_rows(missing_subjects),
        "incompatible_wishes": unique_rows(incompatible_wishes),
        "partial_wishes": unique_rows(partial_wishes),
    }


def incompatible_wish_suggestions(students: list[dict[str, Any]], teachers: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Lav redigerbare forslag til ønsker, der ikke matcher nogen af elevens fag."""
    teacher_ids = {teacher["id"] for teacher in teachers}
    teacher_map = {teacher["id"]: teacher for teacher in teachers}
    candidates = build_candidates(teachers)
    rows = []
    for student_index, student in enumerate(students):
        for wish_index, wish in enumerate(student.get("wishes", [])):
            if wish in teacher_ids and any(wish in candidates.get(canonical_subject(subject), []) for subject in student.get("subjects", [])):
                continue
            possible_ids = []
            for subject in student.get("subjects", []):
                for teacher_id in candidates.get(canonical_subject(subject), []):
                    if teacher_id not in possible_ids:
                        possible_ids.append(teacher_id)
            rows.append({
                "_student_index": student_index,
                "_wish_index": wish_index,
                "_possible_ids": possible_ids,
                "Elev": student["name"],
                "Nuværende ønske": teacher_label(wish, teacher_map),
                "Elevens fag": " · ".join(student.get("subjects", [])),
                "Mulige lærere": " · ".join(teacher_label(teacher_id, teacher_map) for teacher_id in possible_ids) or "Ingen",
                "Vælg ny lærer": "",
            })
    return rows


def subject_coverage(students: list[dict[str, Any]], teachers: list[dict[str, Any]]) -> pd.DataFrame:
    """Viser tydeligt hvilke elevfag der har mindst én mulig lærer."""
    demand = Counter(canonical_subject(subject) for student in students for subject in student.get("subjects", []) if subject)
    candidates = build_candidates(teachers)
    teacher_map = {teacher["id"]: teacher for teacher in teachers}
    labels = {teacher_id: teacher_label(teacher_id, teacher_map) for teacher_id in teacher_map}
    subjects = {}
    for student in students:
        for subject in student.get("subjects", []):
            if subject:
                subjects.setdefault(canonical_subject(subject), subject)
    rows = []
    for key, subject in sorted(subjects.items(), key=lambda item: item[1].casefold()):
        ids = candidates.get(key, [])
        rows.append({
            "Fag": subject,
            "Antal elevpladser": demand.get(key, 0),
            "Mulige lærere": " · ".join(labels.get(teacher_id, teacher_id) for teacher_id in ids) or "Ingen",
            "Status": "Dækket" if ids else "MANGLER LÆRER",
        })
    return pd.DataFrame(rows, columns=["Fag", "Antal elevpladser", "Mulige lærere", "Status"])


def data_readiness(students: list[dict[str, Any]], teachers: list[dict[str, Any]]) -> list[str]:
    """Returnér fejl, som skal være rettet, før en fordeling må beregnes."""
    errors = []
    if not students:
        errors.append("Der er ikke indlæst nogen elever.")
    if not teachers:
        errors.append("Der er ikke indlæst nogen lærere.")
    for index, student in enumerate(students, 1):
        if not repair_text(student.get("name")):
            errors.append(f"Elev række {index} mangler elevnavn.")
        subjects = list(student.get("subjects", []))
        if len(subjects) < 2 or any(not repair_text(subject) for subject in subjects[:2]):
            errors.append(f"Elev {student.get('name') or index} skal have både Fag 1 og Fag 2.")
    for index, teacher in enumerate(teachers, 1):
        if not normal_key(teacher.get("id")):
            errors.append(f"Lærer række {index} mangler initialer.")
        if not any(repair_text(subject) for subject in teacher.get("subjects", [])):
            errors.append(f"Lærer {teacher.get('name') or index} mangler mindst ét fag.")

    validation = input_validation(students, teachers)
    if validation["unknown_wishes"]:
        errors.append(f"{len(validation['unknown_wishes'])} ønskevejleder(e) findes ikke i lærerlisten.")
    if validation["missing_subjects"]:
        errors.append(f"{len(validation['missing_subjects'])} elevfag har ingen lærer.")
    if validation["incompatible_wishes"]:
        errors.append(f"{len(validation['incompatible_wishes'])} ønsker peger på lærere uden elevens fag.")
    return errors


def make_input_export(students: list[dict[str, Any]], teachers: list[dict[str, Any]], capacities: dict[str, int]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        students_to_frame(students, teachers).to_excel(writer, index=False, sheet_name="Elevdata")
        teachers_to_frame(teachers, capacities).to_excel(writer, index=False, sheet_name="Lærerdata")
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for column in sheet.columns:
                width = min(55, max(12, max(len(str(cell.value or "")) for cell in column) + 2))
                sheet.column_dimensions[column[0].column_letter].width = width
    return output.getvalue()


def make_teacher_export(
    teachers: list[dict[str, Any]],
    capacities: dict[str, int],
    subject_count: int | None = None,
    hold_count: int | None = None,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        teachers_to_frame(teachers, capacities, subject_count, hold_count).to_excel(writer, index=False, sheet_name="Lærerdata")
        sheet = writer.book["Lærerdata"]
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            width = min(55, max(12, max(len(str(cell.value or "")) for cell in column) + 2))
            sheet.column_dimensions[column[0].column_letter].width = width
    return output.getvalue()


def try_excel_export(factory: Any) -> bytes | None:
    try:
        return factory()
    except ImportError:
        return None


def load_default_state(load_students: bool = True, load_teachers: bool = True) -> None:
    students, teachers, capacities = embedded_data()
    st.session_state.setdefault("students", students if load_students else [])
    st.session_state.setdefault("teachers", teachers if load_teachers else [])
    st.session_state.setdefault("capacities", capacities if load_teachers else {})
    st.session_state.setdefault("solution", None)


def legacy_main() -> None:
    st.set_page_config(page_title="SOPtima · demo med fiktive data", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")
    st.markdown("""
    <style>
      :root { --ink:#172126; --muted:#66777f; --brand:#075f5b; --soft:#e7f4f1; }
      .stApp { background: radial-gradient(circle at 8% 0%, rgba(15,129,122,.12), transparent 31rem), linear-gradient(180deg,#f7fbf9 0,#fffdf8 24rem); }
      h1,h2,h3 { font-family: Georgia, 'Times New Roman', serif; letter-spacing:-.025em; }
      .hero-kicker { color:var(--brand); font-size:.76rem; font-weight:800; letter-spacing:.14em; text-transform:uppercase; margin-bottom:.35rem; }
      .info-card { border:1px solid #dce4e2; border-radius:14px; padding:1rem 1.15rem; background:rgba(255,255,255,.78); color:#52625f; }
      .metric-card { border:1px solid #dce4e2; border-radius:14px; padding:1rem 1.1rem; background:rgba(255,255,255,.78); min-height:92px; }
      .metric-value { display:block; font:700 2rem/1 Georgia,serif; color:#172126; }
      .metric-label { color:#66777f; font-size:.72rem; font-weight:800; letter-spacing:.07em; text-transform:uppercase; }
      div[data-testid='stDataFrame'] { border:1px solid #dce4e2; border-radius:12px; overflow:hidden; }
      div[data-testid='stDataEditor'] { border:1px solid #dce4e2; border-radius:12px; overflow:hidden; }
    </style>
    """, unsafe_allow_html=True)
    load_default_state()
    st.session_state.setdefault("capacity_revision", 0)
    students = st.session_state["students"]
    teachers = st.session_state["teachers"]
    capacities = st.session_state["capacities"]
    teacher_map = {teacher["id"]: teacher for teacher in teachers}

    st.markdown('<div class="hero-kicker">SOPtima · planlægningsværktøj</div>', unsafe_allow_html=True)
    st.caption(f"{len(students)} elever · {len(teachers)} lærere · data lokalt på denne computer")
    st.markdown('<div class="info-card"><strong>Sådan fungerer fordelingen:</strong> Først reserveres gyldige elevønsker til elever med få muligheder. Derefter fordeles de mest begrænsede fag først under hensyn til K, individuelle max-tal og grænsen for dobbeltvejledning I. Der beregnes flere variationer, og den bedste løsning vises.</div>', unsafe_allow_html=True)

    st.markdown("### 1 · Indlæs data")
    with st.sidebar:
        st.header("Proces")
        st.markdown("**1 - Data**\n\nIndlæs elever og lærere")
        st.markdown("**2 - Regler**\n\nVælg K, I og prioriteringer")
        st.markdown("**3 - Lærermax**\n\nKontrollér kapaciteter")
        st.markdown("**4 - Resultat**\n\nGennemgå og eksportér")
        st.divider()
        st.caption("Alt behandles lokalt på denne computer.")

    with st.expander("1 - Indlæs egne data", expanded=False):
        st.write("Elevarket skal indeholde Navn, Klasse, Fag 1, Fag 2 og gerne to kolonner med vejlederønsker.")
        student_upload = st.file_uploader("Elevdata (.xlsx, .xlsm eller .csv)", type=["xlsx", "xlsm", "csv"], key="students_upload")
        teacher_upload = st.file_uploader("Lærerdata (.xlsx, .xlsm eller .csv) – valgfri", type=["xlsx", "xlsm", "csv"], key="teachers_upload")
        if st.button("Indlæs og klargør data", type="primary"):
            try:
                if student_upload is not None:
                    st.session_state["students"] = parse_students(read_uploaded_table(student_upload))
                if teacher_upload is not None:
                    parsed_teachers, parsed_caps = parse_teacher_upload(teacher_upload)
                    if parsed_teachers:
                        st.session_state["teachers"] = parsed_teachers
                        st.session_state["capacities"] = {teacher["id"]: parsed_caps.get(teacher["id"], 0) for teacher in parsed_teachers}
                        st.session_state["capacity_revision"] += 1
                st.session_state["solution"] = None
                st.success(f"Data klargjort: {len(st.session_state['students'])} elever og {len(st.session_state['teachers'])} lærere.")
                st.rerun()
            except Exception as error:
                st.error(str(error))

    st.markdown("### 2 · Vælg fordelingsregler")
    with st.expander("Regler for K, I og prioriteringer", expanded=True):
        rule_cols = st.columns(2)
        with rule_cols[0]:
            use_global = st.checkbox("Brug global K-grænse", value=True, key="main_use_global")
            K = st.slider("Maksimalt antal elever pr. lærer (K)", 1, 50, 18, key="main_k")
            double_limit = st.slider("Maksimale dobbeltvejledninger pr. lærer (I)", 0, 50, 0, key="main_double_limit")
            allow_over = st.checkbox("Tillad overskridelse af max (op til K)", value=True, disabled=not use_global, key="main_allow_over")
        with rule_cols[1]:
            lock_max = st.checkbox("Lås lærernes max-tal fast", value=False, key="main_lock_max")
            prioritize_pairs = st.checkbox("Prioritér samme vejlederpar", value=True, key="main_prioritize_pairs")
            prioritize_classes = st.checkbox("Saml elever fra samme klasse/hold", value=True, key="main_prioritize_classes")
            attempts = st.slider("Algoritmedybde for fordeling", 20, 180, 120, 10, key="main_attempts", help="Algoritmen forsøger at placere eleverne hos lærere, der dækker fagene, samtidig med at kapacitet og elevønsker respekteres. Tallet angiver, hvor mange forslag der afprøves; højere værdi kan give et bedre resultat, men tager længere tid.")
        st.caption("K er den globale grænse. I bestemmer, hvor mange elever der må få samme lærer til begge fag.")

    st.markdown("### 3 · Kontrollér lærermax og beregn")
    st.caption("Redigér kun kolonnen Max. Brug søgning for at finde en bestemt lærer eller et fag.")
    capacity_query = st.text_input("Søg i lærere", placeholder="Navn, initialer eller fag", key="main_capacity_query").casefold()
    sorted_teachers = sorted(teachers, key=lambda item: teacher_label(item["id"], teacher_map))
    capacity_rows = [{"Lærer": teacher_label(teacher["id"], teacher_map), "Initialer": teacher["id"], "Fag": " · ".join(teacher["subjects"]), "Max": int(capacities.get(teacher["id"], 0))} for teacher in sorted_teachers]
    if capacity_query:
        capacity_rows = [row for row in capacity_rows if capacity_query in " ".join(str(value) for value in row.values()).casefold()]
    edited_capacities = st.data_editor(
        pd.DataFrame(capacity_rows),
        width="stretch",
        hide_index=True,
        height=min(430, 38 + max(1, len(capacity_rows)) * 35),
        disabled=["Lærer", "Initialer", "Fag"],
        column_config={"Max": st.column_config.NumberColumn("Max", min_value=0, max_value=50, step=1, required=True)},
        key=f"main_capacity_editor_{st.session_state['capacity_revision']}",
    )
    for _, row in edited_capacities.iterrows():
        teacher_id = normal_key(row["Initialer"])
        if teacher_id in capacities:
            capacities[teacher_id] = max(0, min(50, int(row["Max"] or 0)))

    cap_action_cols = st.columns([1, 1, 1, 3])
    with cap_action_cols[0]:
        if st.button("Nulstil max"):
            _, _, default_caps = embedded_data()
            st.session_state["capacities"] = {teacher["id"]: default_caps.get(teacher["id"], 0) for teacher in teachers}
            st.session_state["capacity_revision"] += 1
            st.rerun()
    with cap_action_cols[1]:
        if st.button("Foreslå kapaciteter"):
            st.info("Forslag beregnes ved at optimere med ensartet loft. Tryk derefter på Beregn fordeling for at kontrollere resultatet.")
            progress_bar = st.progress(0.0, text="Beregner kapacitetsforslag: 0 %")
            suggestion = optimize(
                students, teachers, {teacher["id"]: 50 for teacher in teachers}, K, double_limit, use_global, allow_over,
                lock_max, prioritize_pairs, prioritize_classes, max(30, attempts // 2),
                progress_callback=lambda value: update_algorithm_progress(progress_bar, "Beregner kapacitetsforslag", value),
            )
            update_algorithm_progress(progress_bar, "Kapacitetsforslag færdigt", 1.0)
            st.session_state["capacities"] = dict(suggestion["loads"])
            st.session_state["capacity_revision"] += 1
            st.rerun()
    with cap_action_cols[2]:
        if st.button("Beregn fordeling", type="primary"):
            progress_bar = st.progress(0.0, text="Beregner fordeling: 0 %")
            with st.spinner("Beregner flere mulige fordelinger …"):
                st.session_state["solution"] = optimize(
                    students, teachers, capacities, K, double_limit, use_global, allow_over, lock_max,
                    prioritize_pairs, prioritize_classes, attempts,
                    progress_callback=lambda value: update_algorithm_progress(progress_bar, "Beregner fordeling", value),
                )
            update_algorithm_progress(progress_bar, "Fordeling færdig", 1.0)
            st.rerun()

    solution = st.session_state.get("solution")
    if solution is None:
        st.info("Vælg kapaciteter og tryk **Beregn fordeling** for at se resultatet.")
        return
    st.markdown("### 4 - Resultat og eksport")
    stats = solution["stats"]
    if stats["unassigned"] or stats["over_teachers"] or stats["capacity_blocked_slots"]:
        messages = []
        if stats["unassigned"]:
            messages.append(f"{stats['unassigned']} elever mangler mindst én vejleder.")
        if stats["capacity_blocked_slots"]:
            messages.append(f"{stats['capacity_blocked_slots']} fagpladser blev blokeret af lærermax.")
        if stats["over_teachers"]:
            messages.append(f"{stats['over_teachers']} lærere ligger over deres individuelle max.")
        st.warning(" ".join(messages))

    metric_data = [(len(students) - stats["unassigned"], "Elever fordelt"), (stats["both"], "Begge ønsker"), (stats["one"] + stats["both"], "Mindst ét ønske"), (stats["none"], "Ingen ønsker"), (stats["capacity_blocked_students"], "Kapacitetsblokerede")]
    cols = st.columns(5)
    for column, (value, label) in zip(cols, metric_data):
        column.markdown(f'<div class="metric-card"><span class="metric-value">{value}</span><span class="metric-label">{label}</span></div>', unsafe_allow_html=True)
    st.caption(f"Resultat ved {'K=' + str(solution['K']) if solution['use_global_k'] else 'lærernes egne max-tal'} og I={solution['double_limit']}. {stats['both']} elever får begge ønsker opfyldt.")

    tab_students, tab_teachers, tab_subjects, tab_unassigned, tab_analysis = st.tabs(["Elever", "Lærere", "Fagstatistik", "Ikke tildelte", "Analyse"])
    with tab_students:
        query = st.text_input("Søg i elever", key="student_query").casefold()
        student_rows = []
        for index, student in enumerate(students):
            assigned = solution["assignments"][index]
            search = " ".join([student["name"], student.get("className", ""), *student["subjects"], *(assigned_item or "" for assigned_item in assigned)]).casefold()
            if query and query not in search:
                continue
            student_rows.append({"Elev": student["name"], "Klasse": student.get("className", ""), "Fag 1": student["subjects"][0], "Vejleder 1": teacher_label(assigned[0], teacher_map), "Fag 2": student["subjects"][1], "Vejleder 2": teacher_label(assigned[1], teacher_map), "Ønsker": f"{wished_count(student, assigned)}/2"})
        st.dataframe(pd.DataFrame(student_rows), width="stretch", hide_index=True, height=560)
        st.caption("Manuel justering foretages via knapperne nedenfor. Værdierne skal være lærerinitialer.")
        with st.expander("Manuel justering af vejledere"):
            edit_rows = []
            for index, student in enumerate(students):
                assigned = solution["assignments"][index]
                edit_rows.append({"Elev": student["name"], "Fag 1": student["subjects"][0], "Vejleder 1": assigned[0] or "", "Fag 2": student["subjects"][1], "Vejleder 2": assigned[1] or ""})
            edited = st.data_editor(pd.DataFrame(edit_rows), width="stretch", hide_index=True, disabled=["Elev", "Fag 1", "Fag 2"], key="manual_editor")
            if st.button("Gem manuelle ændringer"):
                valid_ids = set(teacher_map)
                for index, row in edited.iterrows():
                    for slot in (1, 2):
                        value = normal_key(row[f"Vejleder {slot}"])
                        if value and value not in valid_ids:
                            st.error(f"Ukendt lærerinitial på række {index + 1}: {value}")
                            break
                        solution["assignments"][index][slot - 1] = value or None
                solution["loads"] = {teacher_id: 0 for teacher_id in teacher_map}
                solution["double_loads"] = {teacher_id: 0 for teacher_id in teacher_map}
                for assigned in solution["assignments"]:
                    for teacher_id in set(item for item in assigned if item):
                        solution["loads"][teacher_id] += 1
                    if assigned[0] and assigned[0] == assigned[1]:
                        solution["double_loads"][assigned[0]] += 1
                solution["stats"] = score_solution(students, teachers, capacities, solution["assignments"], solution["loads"], solution["K"], solution["use_global_k"], prioritize_pairs, prioritize_classes)
                st.session_state["solution"] = solution
                st.success("Fordelingen er opdateret.")
                st.rerun()
    with tab_teachers:
        teacher_rows = []
        for teacher in sorted(teachers, key=lambda item: (-solution["loads"].get(item["id"], 0), teacher_label(item["id"], teacher_map))):
            teacher_id = teacher["id"]
            limit = min(solution["K"], capacities.get(teacher_id, 0)) if solution["use_global_k"] else capacities.get(teacher_id, 0)
            pupils = [students[index]["name"] for index, assigned in enumerate(solution["assignments"]) if teacher_id in assigned]
            teacher_rows.append({"Lærer": teacher_label(teacher_id, teacher_map), "Fag": " · ".join(teacher["subjects"]), "Elever": len(pupils), "Max": limit, "Status": "Over max" if len(pupils) > limit else "OK", "Dobbelt": solution["double_loads"].get(teacher_id, 0)})
        st.dataframe(pd.DataFrame(teacher_rows), width="stretch", hide_index=True, height=600)
    with tab_subjects:
        subject_frame = subject_stats(students, teachers, capacities, solution["K"], solution["use_global_k"])
        st.dataframe(subject_frame, width="stretch", hide_index=True, height=560)
    with tab_unassigned:
        missing = []
        for index, student in enumerate(students):
            for slot, subject in enumerate(student["subjects"]):
                if not solution["assignments"][index][slot]:
                    missing.append({"Elev": student["name"], "Klasse": student.get("className", ""), "Fag": subject, "Status": "Mangler vejleder"})
        st.dataframe(pd.DataFrame(missing), width="stretch", hide_index=True, height=560)
        if not missing:
            st.success("Alle fag er tildelt.")
    with tab_analysis:
        st.write("K/I-afvejningen viser, hvordan flere kapacitetsvariationer påvirker antallet af opfyldte ønsker.")
        if st.button("Beregn K-kurve", key="calculate_curve"):
            curve = []
            progress = st.progress(0)
            for index, curve_k in enumerate(range(1, 51)):
                curve_solution = optimize(students, teachers, capacities, curve_k, double_limit, True, allow_over, lock_max, prioritize_pairs, prioritize_classes, max(3, attempts // 30))
                curve.append({"K": curve_k, "Begge ønsker": curve_solution["stats"]["both"], "Mindst ét ønske": curve_solution["stats"]["one"] + curve_solution["stats"]["both"], "Ingen ønsker": curve_solution["stats"]["none"]})
                progress.progress((index + 1) / 50)
            st.session_state["curve"] = pd.DataFrame(curve).set_index("K")
        if st.session_state.get("curve") is not None:
            st.line_chart(st.session_state["curve"], color=["#075f5b", "#e69f35", "#b44438"])
        else:
            st.info("Tryk på Beregn K-kurve for at lave analysen.")

    st.divider()
    st.subheader("Eksport")
    export_col, docx_col = st.columns(2)
    with export_col:
        st.download_button("Download fordeling som Excel", data=make_export(students, teachers, solution, capacities, solution["use_global_k"]), file_name="vejlederfordeling.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    with docx_col:
        if st.button("Forbered Word-filer"):
            with st.spinner("Genererer Word-filer …"):
                try:
                    st.session_state["docx_zip"] = make_docx_zip(students, teachers, solution)
                except ImportError:
                    st.error("Word-eksport kræver python-docx. Kør start_app.cmd igen, så installeres pakken.")
        if st.session_state.get("docx_zip"):
            st.download_button("Download SOP-filer (.zip)", data=st.session_state["docx_zip"], file_name="SOP-filer.zip", mime="application/zip")


if False and __name__ == "__main__":
    legacy_main()


def main() -> None:
    st.set_page_config(page_title="SOPtima · demo med fiktive data", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")
    st.markdown("""
    <style>
      :root { --ink:#172126; --muted:#66777f; --brand:#075f5b; }
      .stApp { background: radial-gradient(circle at 8% 0%, rgba(15,129,122,.12), transparent 31rem), linear-gradient(180deg,#f7fbf9 0,#fffdf8 24rem); }
      h1,h2,h3 { font-family: Georgia, 'Times New Roman', serif; letter-spacing:-.025em; }
      .hero-kicker { color:var(--brand); font-size:.76rem; font-weight:800; letter-spacing:.14em; text-transform:uppercase; margin-bottom:.35rem; }
      .info-card, .metric-card { border:1px solid #dce4e2; border-radius:14px; padding:1rem 1.15rem; background:rgba(255,255,255,.78); }
      .metric-card { min-height:92px; }
      .metric-value { display:block; font:700 2rem/1 Georgia,serif; color:#172126; }
      .metric-label { color:#66777f; font-size:.72rem; font-weight:800; letter-spacing:.07em; text-transform:uppercase; }
      div[data-testid='stDataFrame'], div[data-testid='stDataEditor'] { border:1px solid #dce4e2; border-radius:12px; overflow:hidden; }
    </style>
    """, unsafe_allow_html=True)
    load_default_state()
    for key, default in (("capacity_revision", 0), ("student_revision", 0), ("teacher_revision", 0), ("input_loaded", False)):
        st.session_state.setdefault(key, default)

    students = st.session_state["students"]
    teachers = st.session_state["teachers"]
    capacities = st.session_state["capacities"]
    resolve_wishes(students, teachers)
    teacher_map = {teacher["id"]: teacher for teacher in teachers}

    st.markdown('<div class="hero-kicker">SOPtima · planlægningsværktøj</div>', unsafe_allow_html=True)
    st.caption(f"{len(students)} elever · {len(teachers)} lærere · data behandles lokalt")
    st.markdown('<div class="info-card"><strong>Arbejdsgang:</strong> Indlæs og ret elev- og lærerdata i første fane. Vælg derefter regler og max-tal. Til sidst beregnes fordelingen og vises i en overskuelig tabel, som også kan downloades som Excel.</div>', unsafe_allow_html=True)

    with st.sidebar:
        st.header("Proces")
        st.markdown("**1 · Data**\n\nIndlæs og ret elever/lærere")
        st.markdown("**2 · Regler**\n\nVælg K, I og max-tal")
        st.markdown("**3 · Fordeling**\n\nBeregn og gennemgå resultatet")
        st.divider()
        st.caption("Alt behandles lokalt på denne computer.")

    data_tab, rules_tab, result_tab = st.tabs(["1 · Elev- og lærerdata", "2 · Fordelingsregler og max", "3 · Beregn fordeling"])

    with data_tab:
        st.subheader("Indlæs data")
        st.write("Indlæs elevdata og lærerdata hver for sig. Du kan derefter rette alle celler direkte i tabellerne nedenfor.")
        upload_cols = st.columns(2)
        with upload_cols[0]:
            student_upload = st.file_uploader("Elevdata (.xlsx, .xlsm eller .csv)", type=["xlsx", "xlsm", "csv"], key="students_upload")
        with upload_cols[1]:
            teacher_upload = st.file_uploader("Lærerdata (.xlsx, .xlsm eller .csv)", type=["xlsx", "xlsm", "csv"], key="teachers_upload")
        if st.button("Indlæs valgte filer", type="primary", key="load_input_files"):
            try:
                if student_upload is not None:
                    st.session_state["students"] = parse_students(read_uploaded_table(student_upload))
                    st.session_state["student_revision"] += 1
                if teacher_upload is not None:
                    parsed_teachers, parsed_caps = parse_teacher_upload(teacher_upload)
                    st.session_state["teachers"] = parsed_teachers
                    st.session_state["capacities"] = {teacher["id"]: parsed_caps.get(teacher["id"], 0) for teacher in parsed_teachers}
                    st.session_state["teacher_revision"] += 1
                    st.session_state["capacity_revision"] += 1
                resolve_wishes(st.session_state["students"], st.session_state["teachers"])
                readiness_errors = data_readiness(st.session_state["students"], st.session_state["teachers"])
                st.session_state["input_loaded"] = not readiness_errors
                st.session_state["solution"] = None
                st.success(f"Data indlæst: {len(st.session_state['students'])} elever og {len(st.session_state['teachers'])} lærere.")
                st.rerun()
            except Exception as error:
                st.error(str(error))

        input_export = try_excel_export(lambda: make_input_export(students, teachers, capacities))
        if input_export is not None:
            st.download_button("Download aktuelle inputark som Excel", data=input_export, file_name="vejlederfordeling_input.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_input")
        else:
            st.info("Excel-download kræver, at projektets afhængigheder er installeret (openpyxl mangler i dette miljø).")
        st.divider()
        st.subheader("Elevdata – redigerbar")
        st.caption("Påkrævede felter er Elevnavn, Klasse, Fag 1 og Fag 2. Ønskevejlederne skal være initialer eller lærernavne.")
        edited_students = st.data_editor(
            students_to_frame(students, teachers), width="stretch", hide_index=True, num_rows="dynamic",
            height=min(600, 80 + max(1, len(students)) * 38), disabled=["Elev-ID"],
            column_config={**wish_column_config(teachers), "Projektbeskrivelse": st.column_config.TextColumn("Projektbeskrivelse", width="large")},
            key=f"students_editor_{st.session_state['student_revision']}",
        )
        try:
            new_students = frame_to_students(edited_students)
            resolve_wishes(new_students, teachers)
            if student_data_signature(new_students) != student_data_signature(students):
                st.session_state["students"] = new_students
                st.session_state["input_loaded"] = False
                st.session_state["solution"] = None
                students = new_students
        except ValueError as error:
            st.error(str(error))

        st.subheader("Lærerdata – redigerbar")
        st.caption("Én række pr. lærer. Ret initialer, navn, fag og Max direkte. Brug pilen nederst i tabellen for at tilføje lærere.")
        edited_teachers = st.data_editor(
            teachers_to_frame(teachers, capacities), width="stretch", hide_index=True, num_rows="dynamic",
            height=min(500, 80 + max(1, len(teachers)) * 38),
            column_config={"Max": st.column_config.NumberColumn("Max", min_value=0, max_value=50, step=1, required=True)},
            key=f"teachers_editor_{st.session_state['teacher_revision']}",
        )
        new_teachers, new_capacities = frame_to_teachers(edited_teachers)
        if json.dumps(new_teachers, sort_keys=True, ensure_ascii=False) != json.dumps(teachers, sort_keys=True, ensure_ascii=False) or new_capacities != capacities:
            st.session_state["teachers"] = new_teachers
            st.session_state["capacities"] = new_capacities
            st.session_state["input_loaded"] = False
            st.session_state["solution"] = None
            teachers, capacities = new_teachers, new_capacities
        resolve_wishes(students, teachers)
        teacher_map = {teacher["id"]: teacher for teacher in teachers}

        validation = input_validation(students, teachers)
        if validation["unknown_wishes"]:
            st.warning(f"{len(validation['unknown_wishes'])} ønske(r) peger på en lærer, der ikke findes i lærerarket. Ret initialerne/navnet i elevtabellen.")
            st.dataframe(pd.DataFrame(validation["unknown_wishes"]), width="stretch", hide_index=True)
        if validation["missing_subjects"]:
            st.warning(f"{len(validation['missing_subjects'])} elevfag har ingen lærer i lærerarket. Ret fagene i elevtabellen eller tilføj faget på en lærer.")
            st.dataframe(pd.DataFrame(validation["missing_subjects"]).drop_duplicates(), width="stretch", hide_index=True)
        if validation["incompatible_wishes"]:
            st.info("Nogle ønsker findes, men den ønskede lærer har ikke et af elevens fag. Kontrollér fag og ønsker i tabellerne.")
            st.dataframe(pd.DataFrame(validation["incompatible_wishes"]).drop_duplicates(), width="stretch", hide_index=True)
        if not any(validation.values()):
            st.success("Elevønsker, elevfag og lærernes fag ser konsistente ud.")

        readiness_errors = data_readiness(students, teachers)
        if st.session_state["input_loaded"] and not readiness_errors:
            st.success("Data er godkendt og klar til beregning.")
        else:
            if readiness_errors:
                st.error("Fordeling er låst, indtil data er rettet: " + " ".join(readiness_errors[:8]))
            if st.button("Godkend data til beregning", type="primary", key="approve_input_data"):
                if readiness_errors:
                    st.error("Data kan ikke godkendes endnu. Ret fejlene ovenfor først.")
                else:
                    st.session_state["input_loaded"] = True
                    st.session_state["solution"] = None
                    st.success("Data er godkendt. Gå til fanen Beregn fordeling.")
                    st.rerun()

    with rules_tab:
        st.subheader("Fordelingsregler")
        rule_cols = st.columns(2)
        with rule_cols[0]:
            use_global = st.checkbox("Brug global K-grænse", value=True, key="main_use_global")
            K = st.slider("Maksimalt antal elever pr. lærer (K)", 1, 50, 18, key="main_k")
            double_limit = st.slider("Maksimale dobbeltvejledninger pr. lærer (I)", 0, 50, 0, key="main_double_limit")
            allow_over = st.checkbox("Tillad overskridelse af max (op til K)", value=True, disabled=not use_global, key="main_allow_over")
        with rule_cols[1]:
            lock_max = st.checkbox("Lås lærernes max-tal fast", value=False, key="main_lock_max")
            prioritize_pairs = st.checkbox("Prioritér samme vejlederpar", value=True, key="main_prioritize_pairs")
            prioritize_classes = st.checkbox("Saml elever fra samme klasse/hold", value=True, key="main_prioritize_classes")
            attempts = st.slider("Algoritmedybde for fordeling", 20, 180, 120, 10, key="main_attempts", help="Algoritmen forsøger at placere eleverne hos lærere, der dækker fagene, samtidig med at kapacitet og elevønsker respekteres. Tallet angiver, hvor mange forslag der afprøves; højere værdi kan give et bedre resultat, men tager længere tid.")
        st.caption("K er den globale grænse. I bestemmer, hvor mange elever der må få samme lærer til begge fag.")

        st.subheader("Lærernes max-tal")
        st.caption("Redigér Max direkte. Søgning filtrerer kun visningen.")
        capacity_query = st.text_input("Søg i lærere", placeholder="Navn, initialer eller fag", key="main_capacity_query").casefold()
        sorted_teachers = sorted(teachers, key=lambda item: teacher_label(item["id"], teacher_map))
        capacity_rows = [{"Lærer": teacher_label(teacher["id"], teacher_map), "Initialer": teacher["id"], "Fag": " · ".join(teacher["subjects"]), "Max": int(capacities.get(teacher["id"], 0))} for teacher in sorted_teachers]
        if capacity_query:
            capacity_rows = [row for row in capacity_rows if capacity_query in " ".join(str(value) for value in row.values()).casefold()]
        edited_capacities = st.data_editor(
            pd.DataFrame(capacity_rows), width="stretch", hide_index=True,
            height=min(430, 80 + max(1, len(capacity_rows)) * 35), disabled=["Lærer", "Initialer", "Fag"],
            column_config={"Max": st.column_config.NumberColumn("Max", min_value=0, max_value=50, step=1, required=True)},
            key=f"main_capacity_editor_{st.session_state['capacity_revision']}",
        )
        for _, row in edited_capacities.iterrows():
            teacher_id = normal_key(row["Initialer"])
            if teacher_id in capacities:
                capacities[teacher_id] = max(0, min(50, int(row["Max"] or 0)))
        action_cols = st.columns([1, 1, 2])
        with action_cols[0]:
            if st.button("Nulstil max", key="reset_max"):
                _, _, default_caps = embedded_data()
                st.session_state["capacities"] = {teacher["id"]: default_caps.get(teacher["id"], 0) for teacher in teachers}
                st.session_state["capacity_revision"] += 1
                st.rerun()
        with action_cols[1]:
            if st.button("Foreslå kapaciteter", key="suggest_max") and students and teachers:
                progress_bar = st.progress(0.0, text="Beregner kapacitetsforslag: 0 %")
                with st.spinner("Beregner forslag til kapaciteter …"):
                    suggestion = optimize(
                        students, teachers, {teacher["id"]: 50 for teacher in teachers}, K, double_limit, use_global, allow_over,
                        lock_max, prioritize_pairs, prioritize_classes, max(30, attempts // 2),
                        progress_callback=lambda value: update_algorithm_progress(progress_bar, "Beregner kapacitetsforslag", value),
                    )
                update_algorithm_progress(progress_bar, "Kapacitetsforslag færdigt", 1.0)
                st.session_state["capacities"] = dict(suggestion["loads"])
                st.session_state["capacity_revision"] += 1
                st.rerun()
        with action_cols[2]:
            st.caption("Ændringer i data eller regler påvirker resultatet, når du beregner igen.")

    with result_tab:
        st.subheader("Beregn fordeling")
        st.write("Gennemgå eventuelle advarsler i datafanen, og beregn derefter fordelingen. Resultatet kan downloades som et Excel-ark med flere faner.")
        validation = input_validation(students, teachers)
        readiness_errors = data_readiness(students, teachers)
        can_calculate = st.session_state["input_loaded"] and not readiness_errors
        if not st.session_state["input_loaded"]:
            st.warning("Data er ikke godkendt endnu. Gå til datafanen, indlæs eller ret data, og tryk **Godkend data til beregning**.")
        elif readiness_errors:
            st.error("Fordeling er låst, fordi data ikke er på korrekt form: " + " ".join(readiness_errors[:8]))
        if st.button("Beregn fordeling", type="primary", key="calculate_distribution", disabled=not can_calculate):
            if not students:
                st.error("Der er ingen elever at fordele.")
            elif not teachers:
                st.error("Der er ingen lærere at fordele på.")
            else:
                progress_bar = st.progress(0.0, text="Beregner fordeling: 0 %")
                with st.spinner("Beregner flere mulige fordelinger …"):
                    st.session_state["solution"] = optimize(
                        students, teachers, capacities, K, double_limit, use_global, allow_over, lock_max,
                        prioritize_pairs, prioritize_classes, attempts,
                        progress_callback=lambda value: update_algorithm_progress(progress_bar, "Beregner fordeling", value),
                    )
                update_algorithm_progress(progress_bar, "Fordeling færdig", 1.0)
                st.rerun()

        solution = st.session_state.get("solution")
        if solution is None:
            st.info("Tryk på **Beregn fordeling** for at se resultatet.")
            return

        stats = solution["stats"]
        if stats["unassigned"] or stats["over_teachers"] or stats["capacity_blocked_slots"]:
            messages = []
            if stats["unassigned"]:
                messages.append(f"{stats['unassigned']} elever mangler mindst én vejleder.")
            if stats["capacity_blocked_slots"]:
                messages.append(f"{stats['capacity_blocked_slots']} fagpladser blev blokeret af lærermax.")
            if stats["over_teachers"]:
                messages.append(f"{stats['over_teachers']} lærere ligger over deres individuelle max.")
            st.warning(" ".join(messages))

        metric_data = [(len(students) - stats["unassigned"], "Elever fordelt"), (stats["both"], "Begge ønsker"), (stats["one"] + stats["both"], "Mindst ét ønske"), (stats["none"], "Ingen ønsker"), (stats["capacity_blocked_students"], "Kapacitetsblokerede")]
        cols = st.columns(5)
        for column, (value, label) in zip(cols, metric_data):
            column.markdown(f'<div class="metric-card"><span class="metric-value">{value}</span><span class="metric-label">{label}</span></div>', unsafe_allow_html=True)
        st.caption(f"Resultat ved {'K=' + str(solution['K']) if solution['use_global_k'] else 'lærernes egne max-tal'} og I={solution['double_limit']}. {stats['both']} elever får begge ønsker opfyldt.")
        export_sheet_options = ["Fordeling", "Lærerbelastning", "Fagstatistik", "Ikke tildelte", "Elevdata", "Lærerdata"]
        selected_export_sheets = st.multiselect(
            "Vælg ark til Excel-filen",
            options=export_sheet_options,
            default=export_sheet_options[:4],
            help="Vælg ét eller flere ark. Rækkefølgen i Excel følger listen ovenfor.",
            key="selected_export_sheets",
        )
        if not selected_export_sheets:
            st.warning("Vælg mindst ét ark, før Excel-filen kan downloades.")
            result_export = None
        else:
            result_export = try_excel_export(lambda: make_export(students, teachers, solution, capacities, solution["use_global_k"], selected_export_sheets))
        if result_export is not None:
            st.download_button("Download fordeling som Excel", data=result_export, file_name="vejlederfordeling.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", key="download_result")
        elif selected_export_sheets:
            st.error("Excel-download kræver openpyxl. Installer projektets requirements.txt for at aktivere eksporten.")

        result_students, result_teachers, result_subjects, result_missing = st.tabs(["Fordeling", "Lærere", "Fagstatistik", "Ikke tildelte"])
        with result_students:
            query = st.text_input("Søg i elever", key="student_query").casefold()
            rows = []
            for index, student in enumerate(students):
                assigned = solution["assignments"][index]
                search = " ".join([student["name"], student.get("className", ""), *student["subjects"], *(assigned_item or "" for assigned_item in assigned)]).casefold()
                if query and query not in search:
                    continue
                rows.append({"Elev": student["name"], "Klasse": student.get("className", ""), "Fag 1": student["subjects"][0], "Vejleder 1": teacher_label(assigned[0], teacher_map), "Fag 2": student["subjects"][1], "Vejleder 2": teacher_label(assigned[1], teacher_map), "Ønsker": f"{wished_count(student, assigned)}/2", "Projekttitel": student.get("projectTitle", "")})
            st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True, height=560)
            st.caption("Fordelingsarket i Excel indeholder også projektbeskrivelse og vejlederinitialer.")
            with st.expander("Manuel justering af vejledere"):
                edit_rows = []
                for index, student in enumerate(students):
                    assigned = solution["assignments"][index]
                    edit_rows.append({"Elev": student["name"], "Fag 1": student["subjects"][0], "Vejleder 1": assigned[0] or "", "Fag 2": student["subjects"][1], "Vejleder 2": assigned[1] or ""})
                edited = st.data_editor(pd.DataFrame(edit_rows), width="stretch", hide_index=True, disabled=["Elev", "Fag 1", "Fag 2"], key="manual_editor")
                if st.button("Gem manuelle ændringer", key="save_manual"):
                    errors = []
                    valid_ids = set(teacher_map)
                    for index, row in edited.iterrows():
                        values = []
                        for slot in (1, 2):
                            value = normal_key(row[f"Vejleder {slot}"])
                            if value and value not in valid_ids:
                                errors.append(f"Række {index + 1}: ukendt lærer {value}")
                            values.append(value or None)
                        if not errors:
                            solution["assignments"][index] = values
                    if errors:
                        st.error("; ".join(errors[:5]))
                    else:
                        solution["loads"] = {teacher_id: 0 for teacher_id in teacher_map}
                        solution["double_loads"] = {teacher_id: 0 for teacher_id in teacher_map}
                        for assigned in solution["assignments"]:
                            for teacher_id in set(item for item in assigned if item):
                                solution["loads"][teacher_id] += 1
                            if assigned[0] and assigned[0] == assigned[1]:
                                solution["double_loads"][assigned[0]] += 1
                        solution["stats"] = score_solution(students, teachers, capacities, solution["assignments"], solution["loads"], solution["K"], solution["use_global_k"], prioritize_pairs, prioritize_classes)
                        st.session_state["solution"] = solution
                        st.success("Fordelingen er opdateret.")
                        st.rerun()
        with result_teachers:
            teacher_rows = []
            for teacher in sorted(teachers, key=lambda item: (-solution["loads"].get(item["id"], 0), teacher_label(item["id"], teacher_map))):
                teacher_id = teacher["id"]
                limit = min(solution["K"], capacities.get(teacher_id, 0)) if solution["use_global_k"] else capacities.get(teacher_id, 0)
                pupils = [students[index]["name"] for index, assigned in enumerate(solution["assignments"]) if teacher_id in assigned]
                teacher_rows.append({"Lærer": teacher_label(teacher_id, teacher_map), "Initialer": teacher_id, "Fag": " · ".join(teacher["subjects"]), "Elever": len(pupils), "Max": limit, "Status": "Over max" if len(pupils) > limit else "OK", "Dobbelt": solution["double_loads"].get(teacher_id, 0), "Elevliste": ", ".join(pupils)})
            st.dataframe(pd.DataFrame(teacher_rows), width="stretch", hide_index=True, height=600)
        with result_subjects:
            st.dataframe(subject_stats(students, teachers, capacities, solution["K"], solution["use_global_k"]), width="stretch", hide_index=True, height=560)
        with result_missing:
            missing = []
            overview = {}
            candidates = build_candidates(teachers)
            for index, student in enumerate(students):
                for slot, subject in enumerate(student["subjects"]):
                    if not solution["assignments"][index][slot]:
                        missing.append({
                            "Nr.": len(missing) + 1,
                            "student_index": index,
                            "slot": slot,
                            "Elev": student["name"],
                            "Klasse": student.get("className", ""),
                            "Fag": subject,
                            "Ønskevejledere": " · ".join(student.get("wishes", [])) or "Ingen registrerede ønsker",
                            "Mulige lærere": " · ".join(candidates.get(canonical_subject(subject), [])) or "Ingen",
                            "Tildel vejleder": "",
                        })
                        item = overview.setdefault(index, {"Elev": student["name"], "Klasse": student.get("className", ""), "Fag der mangler": [], "Antal": 0})
                        item["Fag der mangler"].append(subject)
                        item["Antal"] += 1

            if not missing:
                st.success("Alle elever og fag er tildelt.")
            else:
                st.warning(f"{len(overview)} elever mangler vejleder til {len(missing)} fagplads(er). Vælg en lærer direkte nedenfor og tryk Gem tildelinger.")
                st.markdown("**Overblik pr. elev**")
                overview_rows = list(overview.values())
                for item in overview_rows:
                    item["Fag der mangler"] = " · ".join(item["Fag der mangler"])
                st.dataframe(pd.DataFrame(overview_rows), width="stretch", hide_index=True, height=min(280, 70 + len(overview_rows) * 38))

                st.markdown("**Tildel manglende vejledere direkte**")
                st.caption("Vælg lærerinitialer i kolonnen Tildel vejleder. Kun lærere, der har det valgte fag, kan gemmes.")
                missing_frame = pd.DataFrame(missing)
                edited_missing = st.data_editor(
                    missing_frame[["Nr.", "Elev", "Klasse", "Fag", "Ønskevejledere", "Mulige lærere", "Tildel vejleder"]],
                    width="stretch", hide_index=True, height=min(560, 90 + len(missing) * 38),
                    disabled=["Nr.", "Elev", "Klasse", "Fag", "Ønskevejledere", "Mulige lærere"],
                    column_config={
                        "Tildel vejleder": st.column_config.SelectboxColumn(
                            "Tildel vejleder", options=[""] + sorted(teacher_map), required=False,
                            help="Vælg en lærer, der har dette fag.",
                        )
                    },
                    key="missing_assignments_editor",
                )
                if st.button("Gem tildelinger", type="primary", key="save_missing_assignments"):
                    errors = []
                    changes = []
                    for _, row in edited_missing.iterrows():
                        row_number = int(row["Nr."]) - 1
                        if row_number < 0 or row_number >= len(missing):
                            continue
                        item = missing[row_number]
                        teacher_id = normal_key(row["Tildel vejleder"])
                        if not teacher_id:
                            continue
                        if teacher_id not in teacher_map:
                            errors.append(f"{item['Elev']}: ukendt lærer {teacher_id}.")
                        elif teacher_id not in candidates.get(canonical_subject(item["Fag"]), []):
                            errors.append(f"{item['Elev']}: {teacher_id} har ikke faget {item['Fag']}.")
                        else:
                            changes.append((item["student_index"], item["slot"], teacher_id))
                    if errors:
                        st.error(" ".join(errors[:8]))
                    else:
                        for student_index, slot, teacher_id in changes:
                            solution["assignments"][student_index][slot] = teacher_id
                        solution["loads"] = {teacher_id: 0 for teacher_id in teacher_map}
                        solution["double_loads"] = {teacher_id: 0 for teacher_id in teacher_map}
                        for assigned in solution["assignments"]:
                            for teacher_id in set(item for item in assigned if item):
                                solution["loads"][teacher_id] += 1
                            if assigned[0] and assigned[0] == assigned[1]:
                                solution["double_loads"][assigned[0]] += 1
                        solution["stats"] = score_solution(students, teachers, capacities, solution["assignments"], solution["loads"], solution["K"], solution["use_global_k"], prioritize_pairs, prioritize_classes)
                        st.session_state["solution"] = solution
                        st.success(f"{len(changes)} tildeling(er) er gemt i den aktuelle fordeling.")
                        st.rerun()

        st.divider()
        st.subheader("Eksport")
        if st.button("Forbered Word-filer", key="prepare_docx"):
            with st.spinner("Genererer Word-filer …"):
                try:
                    st.session_state["docx_zip"] = make_docx_zip(students, teachers, solution)
                except ImportError:
                    st.error("Word-eksport kræver python-docx.")
        if st.session_state.get("docx_zip"):
            st.download_button("Download SOP-filer (.zip)", data=st.session_state["docx_zip"], file_name="SOP-filer.zip", mime="application/zip", key="download_docx")


def main_v2() -> None:
    """Trinopdelt arbejdsgang med navigation i venstremenuen."""
    st.set_page_config(page_title="SOPtima · demo med fiktive data", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")
    st.markdown("""
    <style>
      :root { --ink:#172126; --muted:#66777f; --brand:#075f5b; }
      .stApp { background: radial-gradient(circle at 8% 0%, rgba(15,129,122,.12), transparent 31rem), linear-gradient(180deg,#f7fbf9 0,#fffdf8 24rem); }
      h1,h2,h3 { font-family: Georgia, 'Times New Roman', serif; letter-spacing:-.025em; }
      .hero-kicker { color:var(--brand); font-size:.76rem; font-weight:800; letter-spacing:.14em; text-transform:uppercase; margin-bottom:.35rem; }
      .info-card, .metric-card { border:1px solid #dce4e2; border-radius:14px; padding:1rem 1.15rem; background:rgba(255,255,255,.78); }
      .metric-card { min-height:92px; }
      .metric-value { display:block; font:700 2rem/1 Georgia,serif; color:#172126; }
      .metric-label { color:#66777f; font-size:.72rem; font-weight:800; letter-spacing:.07em; text-transform:uppercase; }
      div[data-testid='stDataFrame'], div[data-testid='stDataEditor'] { border:1px solid #dce4e2; border-radius:12px; overflow:hidden; }
    </style>
    """, unsafe_allow_html=True)
    # Start uden indlæste data. Elevdata indlæses først, så elevhold kan bruges som reference.
    # Ryd kun det gamle indlejrede lærergrundlag ved overgang til den nye starttilstand.
    if "v2_empty_teacher_start" not in st.session_state:
        if st.session_state.get("teacher_revision", 0) == 0:
            st.session_state["teachers"] = []
            st.session_state["capacities"] = {}
        st.session_state["v2_empty_teacher_start"] = True
    load_default_state(load_students=False, load_teachers=False)
    for key, default in (
        ("capacity_revision", 0),
        ("student_revision", 0),
        ("teacher_revision", 0),
        ("input_loaded", False),
        ("solution", None),
        ("v2_schedule", None),
        ("v2_test_students_active", False),
        ("v2_test_teachers_active", False),
        ("v2_test_students_checkbox", True),
        ("v2_test_teachers_checkbox", True),
    ):
        st.session_state.setdefault(key, default)

    students = st.session_state["students"]
    teachers = st.session_state["teachers"]
    capacities = st.session_state["capacities"]
    resolve_wishes(students, teachers)
    teacher_map = {teacher["id"]: teacher for teacher in teachers}

    readiness_errors = data_readiness(students, teachers)
    if not students:
        current_step = 1
        next_step = "Start med at indlæse elevdata."
    elif not teachers:
        current_step = 2
        next_step = "Næste skridt er at indlæse lærerdata."
    elif not st.session_state["input_loaded"] or readiness_errors:
        current_step = 2
        next_step = "Godkend data øverst, og gå derefter til trin 4 for at beregne."
    elif st.session_state.get("solution") is None:
        current_step = 4
        next_step = "Data er godkendt — gå videre til Beregn."
    else:
        current_step = 5
        next_step = "Fordelingen er beregnet — gennemgå resultatet og eksportér."

    process_steps = [
        "1 · Elevdata",
        "2 · Lærerdata",
        "3 · Regler og max",
        "4 · Beregn",
        "5 · Resultat og eksport",
        "6 · Tidsplan",
    ]
    K = st.session_state.get("v2_k", 18)
    double_limit = st.session_state.get("v2_double_limit", 0)
    use_global = st.session_state.get("v2_use_global", True)
    allow_over = st.session_state.get("v2_allow_over", True)
    lock_max = st.session_state.get("v2_lock_max", False)
    prioritize_pairs = st.session_state.get("v2_prioritize_pairs", True)
    prioritize_classes = st.session_state.get("v2_prioritize_classes", True)
    attempts = st.session_state.get("v2_attempts", 120)

    with st.sidebar:
        st.title("SOPtima")
        st.caption("Demoversion · fiktive data")
        st.progress(current_step / 6, text=f"Trin {current_step} af 6")
        st.caption(next_step)
        active_step = st.radio("Gå til trin", process_steps, index=current_step - 1, key="v2_active_step", label_visibility="collapsed")
        st.subheader("Status")
        if students:
            student_status = "✅" if not readiness_errors and st.session_state["input_loaded"] else "⚠️"
            student_detail = f"{len(students)} elever"
            if readiness_errors:
                student_detail += f" · {len(readiness_errors)} fejl"
            elif st.session_state["input_loaded"]:
                student_detail += " · godkendt"
            else:
                student_detail += " · ikke godkendt"
        else:
            student_status = "⬜"
            student_detail = "Ikke indlæst"
        st.caption(f"{student_status} Elevdata · {student_detail}")
        teacher_detail = f"{len(teachers)} lærere" if teachers else "Ikke indlæst"
        st.caption(f"{'✅' if teachers else '⬜'} Lærerdata · {teacher_detail}")
        if st.session_state.get("solution") is not None:
            st.caption("✅ Fordeling beregnet")
        if students and teachers and not readiness_errors and not st.session_state["input_loaded"]:
            st.success("Data er klar til godkendelse.")
            if st.button("Godkend data til fordeling", type="primary", key="v2_sidebar_approve_data"):
                st.session_state["input_loaded"] = True
                st.session_state["solution"] = None
                st.session_state["v2_schedule"] = None
                st.rerun()
        elif students and teachers and st.session_state["input_loaded"]:
            st.success("Data er godkendt og klar til fordeling.")
        st.divider()
        st.subheader("Demo-data")
        st.caption("Alle personer, elever, lærere og opgaver i demoen er fiktive.")
        test_students_checked = st.checkbox("Indlæs 350 fiktive demoelever", key="v2_test_students_checkbox")
        test_teachers_checked = st.checkbox("Indlæs 85 fiktive demolærere", key="v2_test_teachers_checkbox")
        testdata_loaded = False
        if not test_students_checked:
            st.session_state["v2_test_students_active"] = False
        elif not st.session_state["v2_test_students_active"]:
            try:
                st.session_state["students"] = parse_students(read_path_table(TEST_STUDENTS_FILE))
                st.session_state["student_revision"] += 1
                st.session_state["input_loaded"] = False
                st.session_state["solution"] = None
                st.session_state["v2_schedule"] = None
                st.session_state["v2_test_students_active"] = True
                testdata_loaded = True
            except Exception as error:
                st.error(f"Testelever kunne ikke indlæses: {error}")
        if not test_teachers_checked:
            st.session_state["v2_test_teachers_active"] = False
        elif not st.session_state["v2_test_teachers_active"]:
            try:
                parsed_test_teachers, parsed_test_capacities = parse_teachers(read_path_table(TEST_TEACHERS_FILE))
                st.session_state["teachers"] = parsed_test_teachers
                st.session_state["capacities"] = {teacher["id"]: parsed_test_capacities.get(teacher["id"], 0) for teacher in parsed_test_teachers}
                st.session_state["teacher_revision"] += 1
                st.session_state["capacity_revision"] += 1
                st.session_state["input_loaded"] = False
                st.session_state["solution"] = None
                st.session_state["v2_schedule"] = None
                st.session_state["v2_test_teachers_active"] = True
                testdata_loaded = True
            except Exception as error:
                st.error(f"Testlærere kunne ikke indlæses: {error}")
        if testdata_loaded:
            st.rerun()
        st.caption("Alt behandles lokalt på denne computer.")

    if active_step == "2 · Lærerdata":
        st.title("2 · Lærerdata")
        st.subheader("Indlæs lærerdata")
        if students and teachers and not readiness_errors:
            if st.session_state["input_loaded"]:
                st.success("Data er godkendt og klar til beregning. Gå til trin 4 for at beregne fordelingen.")
            else:
                st.info("Data ser korrekt ud. Godkend data med knappen øverst i sidepanelet, og gå derefter til trin 4.")
        st.caption("Brug eventuelt elevdataens holdnavne som reference i lærerarket.")
        with st.expander("Se eksempel på lærerarket (kun illustration)", expanded=False):
            st.dataframe(pd.DataFrame([
                {"Initialer": "AB", "Lærernavn": "Anna Borg", "Fag 1": "Matematik", "Fag 2": "Fysik", "Fag 3": "Informatik", "Hold 1": "3a", "Hold 2": "3b", "Hold 3": "", "Hold 4": "", "Max": 18},
                {"Initialer": "CD", "Lærernavn": "Carsten Dahl", "Fag 1": "Dansk", "Fag 2": "Historie", "Fag 3": "Samfundsfag", "Hold 1": "2x", "Hold 2": "", "Hold 3": "", "Hold 4": "", "Max": 16},
            ]), width="stretch", hide_index=True)
            st.caption("Format: Initialer + mindst ét fag. Fag 1–3 vises som standard; op til 20 fag og 4 hold kan tilføjes.")
        st.caption("Næste skridt: vælg en lærerfil.")
        teacher_upload = st.file_uploader("Upload lærerdata (.xlsx, .xlsm eller .csv)", type=["xlsx", "xlsm", "csv"], key="v2_teachers_upload")
        if st.button("Indlæs lærerdata", type="primary" if teacher_upload is not None else "secondary", disabled=teacher_upload is None, key="v2_load_teachers"):
            try:
                parsed_teachers, parsed_caps = parse_teacher_upload(teacher_upload)
                st.session_state["teachers"] = parsed_teachers
                st.session_state["capacities"] = {teacher["id"]: parsed_caps.get(teacher["id"], 0) for teacher in parsed_teachers}
                st.session_state["teacher_revision"] += 1
                st.session_state["capacity_revision"] += 1
                st.session_state["input_loaded"] = False
                st.session_state["solution"] = None
                st.session_state["v2_schedule"] = None
                st.success(f"Lærerdata indlæst: {len(parsed_teachers)} lærere.")
                st.rerun()
            except Exception as error:
                st.error(str(error))

        st.subheader("Lærerdata – redigerbart Excel-ark")
        st.caption("Redigér direkte i arket. Ét fag pr. kolonne; brug pilen nederst for flere lærere.")
        current_subject_count = max(3, max((len(teacher.get("subjects", [])) for teacher in teachers), default=0))
        current_hold_count = max((len(teacher.get("holds", [])) for teacher in teachers), default=0)
        setting_columns = st.columns(2)
        with setting_columns[0]:
            subject_column_count = st.number_input(
                "Antal fagkolonner i lærerarket",
                min_value=3,
                max_value=20,
                value=min(20, current_subject_count),
                step=1,
                help="Start med 3. Vælg et højere antal, hvis en lærer skal have flere fag.",
                key=f"v2_subject_column_count_{st.session_state['teacher_revision']}",
            )
        with setting_columns[1]:
            hold_column_count = st.number_input(
                "Antal holdkolonner i lærerarket",
                min_value=0,
                max_value=4,
                value=min(4, current_hold_count),
                step=1,
                help="Vælg 0–4 holdkolonner. Eksisterende hold bevares, hvis du vælger færre kolonner.",
                key=f"v2_hold_column_count_{st.session_state['teacher_revision']}",
            )
        st.caption("Fag: mindst 3 kolonner og op til 20. Hold: valgfrit, fra 0 til 4 kolonner. Ryd en holdcelle i arket, hvis et eksisterende hold skal fjernes.")
        edited_teachers = st.data_editor(
            teachers_to_frame(teachers, capacities, int(subject_column_count), int(hold_column_count)), width="stretch", hide_index=True, num_rows="dynamic",
            height=min(500, 80 + max(1, len(teachers)) * 38),
            column_config={"Max": st.column_config.NumberColumn("Max", min_value=0, max_value=50, step=1, required=True)},
            key=f"v2_teachers_editor_{st.session_state['teacher_revision']}",
        )
        new_teachers, new_capacities = frame_to_teachers(edited_teachers)
        if json.dumps(new_teachers, sort_keys=True, ensure_ascii=False) != json.dumps(teachers, sort_keys=True, ensure_ascii=False) or new_capacities != capacities:
            st.session_state["teachers"] = new_teachers
            st.session_state["capacities"] = new_capacities
            st.session_state["input_loaded"] = False
            st.session_state["solution"] = None
            st.session_state["v2_schedule"] = None
            teachers, capacities = new_teachers, new_capacities
            teacher_map = {teacher["id"]: teacher for teacher in teachers}
        no_subject_teachers = [teacher["id"] for teacher in teachers if not teacher.get("subjects")]
        if no_subject_teachers:
            st.warning(f"{len(no_subject_teachers)} lærer(e) mangler fag og kan ikke få elever: {', '.join(no_subject_teachers)}")
        teacher_data_ready = bool(teachers) and not no_subject_teachers and all(normal_key(teacher.get("id")) for teacher in teachers)
        if teacher_data_ready:
            teacher_export = try_excel_export(
                lambda: make_teacher_export(teachers, capacities, int(subject_column_count), int(hold_column_count))
            )
            if teacher_export is not None:
                st.download_button(
                    "Download lærerdata som Excel",
                    data=teacher_export,
                    file_name="lærerdata.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="v2_download_teachers",
                )
        if students and teachers:
            st.subheader("Fagdækning blandt elevernes fag")
            st.dataframe(subject_coverage(students, teachers), width="stretch", hide_index=True)

        if students:
            readiness_errors = data_readiness(students, teachers)
            if st.session_state["input_loaded"] and not readiness_errors:
                st.success("Data er godkendt og klar til fordeling.")
            else:
                if readiness_errors:
                    st.error("Ret følgende før godkendelse: " + " ".join(readiness_errors[:8]))
                if st.button("Godkend data til fordeling", type="primary" if not readiness_errors else "secondary", disabled=bool(readiness_errors), key="v2_approve_data"):
                    st.session_state["input_loaded"] = True
                    st.session_state["solution"] = None
                    st.session_state["v2_schedule"] = None
                    st.success("Data er godkendt. Gå videre til fanen Regler og max.")
                    st.rerun()

    if active_step == "1 · Elevdata":
        st.title("1 · Elevdata")
        st.subheader("Indlæs elevdata først")
        with st.expander("Se eksempel på elevarket (kun illustration)", expanded=False):
            st.dataframe(pd.DataFrame([
                {"Elevnavn": "Emma Jensen", "Klasse": "3a", "Fag 1": "Matematik", "Fag 2": "Fysik", "Ønskevejleder 1": "AB", "Ønskevejleder 2": "CD"},
                {"Elevnavn": "Noah Hansen", "Klasse": "3a", "Fag 1": "Dansk", "Fag 2": "Historie", "Ønskevejleder 1": "CD", "Ønskevejleder 2": ""},
            ]), width="stretch", hide_index=True)
            st.caption("Format: Elevnavn, Klasse og Fag 1–2. Ønskevejledere er valgfri.")
        st.caption("Første skridt: vælg en elevfil.")
        student_upload = st.file_uploader("Upload elevdata (.xlsx, .xlsm eller .csv)", type=["xlsx", "xlsm", "csv"], key="v2_students_upload")
        if st.button("Indlæs elevdata", type="primary" if student_upload is not None else "secondary", disabled=student_upload is None, key="v2_load_students"):
            try:
                st.session_state["students"] = parse_students(read_uploaded_table(student_upload))
                st.session_state["student_revision"] += 1
                st.session_state["input_loaded"] = False
                st.session_state["solution"] = None
                st.session_state["v2_schedule"] = None
                st.success(f"Elevdata indlæst: {len(st.session_state['students'])} elever.")
                st.rerun()
            except Exception as error:
                st.error(str(error))

        st.subheader("Elevdata – redigerbart Excel-ark")
        st.caption("Redigér direkte i arket. Vælg en lærer eller **Ingen ønsker**.")
        edited_students = st.data_editor(
            students_to_frame(students, teachers), width="stretch", hide_index=True, num_rows="dynamic",
            height=min(600, 80 + max(1, len(students)) * 38), disabled=["Elev-ID"],
            column_config={**wish_column_config(teachers), "Projektbeskrivelse": st.column_config.TextColumn("Projektbeskrivelse", width="large")},
            key=f"v2_students_editor_{st.session_state['student_revision']}",
        )
        try:
            new_students = frame_to_students(edited_students)
            resolve_wishes(new_students, teachers)
            if student_data_signature(new_students) != student_data_signature(students):
                st.session_state["students"] = new_students
                st.session_state["input_loaded"] = False
                st.session_state["solution"] = None
                students = new_students
        except ValueError as error:
            st.error(str(error))

        resolve_wishes(students, teachers)
        teacher_map = {teacher["id"]: teacher for teacher in teachers}
        validation = input_validation(students, teachers)
        if validation["unknown_wishes"]:
            unknown_frame = pd.DataFrame(validation["unknown_wishes"]).drop_duplicates()
            st.warning(f"{len(unknown_frame)} ønsker kan ikke genkendes. Vælg en lærer eller **Ingen ønsker** i elevarket:")
            st.dataframe(unknown_frame, width="stretch", hide_index=True)
        if validation["missing_subjects"]:
            st.warning("Disse elevfag har ingen lærer endnu:")
            st.dataframe(pd.DataFrame(validation["missing_subjects"]).drop_duplicates(), width="stretch", hide_index=True)
        if validation["incompatible_wishes"] or validation["unknown_wishes"]:
            if validation["incompatible_wishes"]:
                st.warning("Disse ønskede lærere underviser ikke i elevens fag:")
                st.dataframe(pd.DataFrame(validation["incompatible_wishes"]).drop_duplicates(), width="stretch", hide_index=True)
            suggestion_rows = incompatible_wish_suggestions(students, teachers)
            if suggestion_rows:
                st.markdown("**Ret ønskerne med forslag fra lærerlisten**")
                st.caption("Vælg en foreslået lærer, eller vælg **Ingen ønsker** for at fjerne ønsket. Forslagene dækker mindst ét af elevens fag.")
                teacher_labels = {teacher["id"]: teacher_label(teacher["id"], teacher_map) for teacher in teachers}
                selected_suggestions = {}
                suggestion_table = pd.DataFrame(suggestion_rows)[["Elev", "Nuværende ønske", "Elevens fag", "Mulige lærere"]]
                st.dataframe(suggestion_table, width="stretch", hide_index=True)
                st.markdown("**Vælg ny lærer**")
                for row_index, item in enumerate(suggestion_rows):
                    possible_options = [NO_WISHES_LABEL] + [teacher_labels[teacher_id] for teacher_id in item["_possible_ids"]]
                    default_index = 1 if len(possible_options) > 1 else 0
                    selected_suggestions[row_index] = st.selectbox(
                        f"{item['Elev']} · {item['Nuværende ønske']}", possible_options, index=default_index,
                        key=f"v2_wish_suggestion_{st.session_state['teacher_revision']}_{st.session_state['student_revision']}_{row_index}",
                    )
                if st.button("Gem foreslåede lærere", type="primary", key="v2_save_wish_suggestions"):
                    label_to_id = {normal_key(label): teacher_id for teacher_id, label in teacher_labels.items()}
                    candidate_map = build_candidates(teachers)
                    changed_wishes = {index: list(student.get("wishes", [])) for index, student in enumerate(students)}
                    errors = []
                    changes = 0
                    for row_index, selected in selected_suggestions.items():
                        selected = repair_text(selected)
                        if not selected:
                            continue
                        item = suggestion_rows[row_index]
                        student_index = item["_student_index"]
                        wish_index = item["_wish_index"]
                        if normal_key(selected) == normal_key(NO_WISHES_LABEL):
                            changed_wishes[student_index][wish_index] = None
                            changes += 1
                            continue
                        teacher_id = label_to_id.get(normal_key(selected))
                        if not teacher_id:
                            errors.append(f"{item['Elev']}: kunne ikke genkende {selected}.")
                            continue
                        valid_for_subject = any(teacher_id in candidate_map.get(canonical_subject(subject), []) for subject in students[student_index].get("subjects", []))
                        if not valid_for_subject:
                            errors.append(f"{item['Elev']}: {selected} underviser ikke i elevens fag.")
                            continue
                        changed_wishes[student_index][wish_index] = teacher_id
                        changes += 1
                    if errors:
                        st.error(" ".join(errors[:8]))
                    elif changes:
                        for student_index, wishes in changed_wishes.items():
                            students[student_index]["wishes"] = [wish for wish in wishes if wish]
                        st.session_state["input_loaded"] = False
                        st.session_state["solution"] = None
                        st.session_state["v2_schedule"] = None
                        st.success(f"{changes} ønske(r) er opdateret med de valgte lærere.")
                        st.rerun()
        if validation["partial_wishes"]:
            st.caption(f"{len(validation['partial_wishes'])} ønsker matcher kun ét af elevens fag.")
            with st.expander("Se ønsker med delvist fagmatch"):
                st.dataframe(pd.DataFrame(validation["partial_wishes"]).drop_duplicates(), width="stretch", hide_index=True)
        if students and not any(validation[key] for key in ("unknown_wishes", "missing_subjects", "incompatible_wishes")):
            st.success("Elevønsker, elevfag og lærernes fag ser konsistente ud.")
        if students and teachers:
            st.subheader("Kontrol: har hvert fag en mulig lærer?")
            st.dataframe(subject_coverage(students, teachers), width="stretch", hide_index=True)

        input_export = try_excel_export(lambda: make_input_export(students, teachers, capacities))
        if input_export is not None:
            st.download_button("Download aktuelle inputark som Excel", data=input_export, file_name="vejlederfordeling_input.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="v2_download_input")

    if active_step == "3 · Regler og max":
        st.title("3 · Regler og max")
        st.subheader("Fordelingsregler")
        rule_cols = st.columns(2)
        with rule_cols[0]:
            use_global = st.checkbox("Brug global K-grænse", value=True, key="v2_use_global")
            K = st.slider("Maksimalt antal elever pr. lærer (K)", 1, 50, 18, key="v2_k")
            double_limit = st.slider("Maksimale dobbeltvejledninger pr. lærer (I)", 0, 50, 0, key="v2_double_limit")
            allow_over = st.checkbox("Tillad overskridelse af max (op til K)", value=True, disabled=not use_global, key="v2_allow_over")
        with rule_cols[1]:
            lock_max = st.checkbox("Lås lærernes max-tal fast", value=False, key="v2_lock_max")
            prioritize_pairs = st.checkbox("Prioritér samme vejlederpar", value=True, key="v2_prioritize_pairs")
            prioritize_classes = st.checkbox("Saml elever fra samme klasse/hold", value=True, key="v2_prioritize_classes")
            attempts = st.slider("Algoritmedybde for fordeling", 20, 180, 120, 10, key="v2_attempts", help="Algoritmen forsøger at placere eleverne hos lærere, der dækker fagene, samtidig med at kapacitet og elevønsker respekteres. Tallet angiver, hvor mange forslag der afprøves; højere værdi kan give et bedre resultat, men tager længere tid.")
        st.caption("K er den globale grænse. I bestemmer, hvor mange elever der må få samme lærer til begge fag.")
        st.subheader("Lærernes max-tal")
        max_action_cols = st.columns([1, 3])
        with max_action_cols[0]:
            if st.button(
                "Foreslå lærernes max-tal",
                type="primary",
                key="v2_suggest_max_top",
                disabled=not bool(students and teachers),
                help="Beregner et forslag til hvert maksimum ud fra elevfordelingen. Du kan efterfølgende redigere tallene manuelt.",
            ):
                progress_bar = st.progress(0.0, text="Beregner max-forslag: 0 %")
                with st.spinner("Beregner forslag til lærernes max-tal …"):
                    suggestion = optimize(
                        students, teachers, {teacher["id"]: 50 for teacher in teachers}, K, double_limit, use_global, allow_over,
                        lock_max, prioritize_pairs, prioritize_classes, max(30, attempts // 2),
                        progress_callback=lambda value: update_algorithm_progress(progress_bar, "Beregner max-forslag", value),
                    )
                update_algorithm_progress(progress_bar, "Max-forslag færdigt", 1.0)
                st.session_state["capacities"] = dict(suggestion["loads"])
                st.session_state["capacity_revision"] += 1
                st.session_state["solution"] = None
                st.session_state["v2_schedule"] = None
                st.success("Forslag til lærernes max-tal er indsat. Kontrollér og tilpas dem efter behov.")
                st.rerun()
        with max_action_cols[1]:
            st.caption("Forslaget fordeler belastningen så jævnt som muligt. Kontrollér altid tallene, før du beregner den endelige fordeling.")
        capacity_query = st.text_input("Søg i lærere", placeholder="Navn, initialer eller fag", key="v2_capacity_query").casefold()
        capacity_rows = [{"Lærer": teacher_label(teacher["id"], teacher_map), "Initialer": teacher["id"], "Fag": " · ".join(teacher["subjects"]), "Max": int(capacities.get(teacher["id"], 0))} for teacher in sorted(teachers, key=lambda item: teacher_label(item["id"], teacher_map))]
        if capacity_query:
            capacity_rows = [row for row in capacity_rows if capacity_query in " ".join(str(value) for value in row.values()).casefold()]
        edited_capacities = st.data_editor(
            pd.DataFrame(capacity_rows), width="stretch", hide_index=True, disabled=["Lærer", "Initialer", "Fag"],
            column_config={"Max": st.column_config.NumberColumn("Max", min_value=0, max_value=50, step=1, required=True)}, key="v2_capacity_editor",
        )
        for _, row in edited_capacities.iterrows():
            teacher_id = normal_key(row["Initialer"])
            if teacher_id in capacities:
                capacities[teacher_id] = max(0, min(50, int(row["Max"] or 0)))

    can_calculate = st.session_state["input_loaded"] and not data_readiness(students, teachers)
    if active_step == "4 · Beregn":
        st.title("4 · Beregn")
        st.subheader("Beregn fordeling")
        if not st.session_state["input_loaded"]:
            st.warning("Godkend først data i fanen Elevdata.")
        elif not can_calculate:
            st.error("Fordeling er låst, fordi data ikke er på korrekt form.")
        else:
            st.success("Data er godkendt. Du kan beregne fordelingen.")
        if st.button("Beregn fordeling", type="primary" if can_calculate else "secondary", key="v2_calculate", disabled=not can_calculate):
            progress_bar = st.progress(0.0, text="Beregner fordeling: 0 %")
            with st.spinner("Beregner flere mulige fordelinger …"):
                st.session_state["solution"] = optimize(
                    students, teachers, capacities, K, double_limit, use_global, allow_over, lock_max,
                    prioritize_pairs, prioritize_classes, attempts,
                    progress_callback=lambda value: update_algorithm_progress(progress_bar, "Beregner fordeling", value),
                )
                st.session_state["v2_schedule"] = None
            update_algorithm_progress(progress_bar, "Fordeling færdig", 1.0)
            st.success("Fordelingen er beregnet. Gå til fanen Resultat og eksport eller Tidsplan.")

    if active_step == "5 · Resultat og eksport":
        st.title("5 · Resultat og eksport")
        solution = st.session_state.get("solution")
        st.subheader("Resultat og eksport")
        if solution is None:
            st.caption("Beregn først en fordeling i menuen.")
        else:
            stats = solution["stats"]
            if stats["unassigned"] or stats["over_teachers"] or stats["capacity_blocked_slots"]:
                messages = []
                if stats["unassigned"]: messages.append(f"{stats['unassigned']} elever mangler mindst én vejleder.")
                if stats["capacity_blocked_slots"]: messages.append(f"{stats['capacity_blocked_slots']} fagpladser blev blokeret af lærermax.")
                if stats["over_teachers"]: messages.append(f"{stats['over_teachers']} lærere ligger over deres individuelle max.")
                st.warning(" ".join(messages))
            metric_data = [(len(students) - stats["unassigned"], "Elever fordelt"), (stats["both"], "Begge ønsker"), (stats["one"] + stats["both"], "Mindst ét ønske"), (stats["none"], "Ingen ønsker"), (stats["capacity_blocked_students"], "Kapacitetsblokerede")]
            cols = st.columns(5)
            for column, (value, label) in zip(cols, metric_data):
                column.markdown(f'<div class="metric-card"><span class="metric-value">{value}</span><span class="metric-label">{label}</span></div>', unsafe_allow_html=True)
            export_sheet_options = ["Fordeling", "Lærerbelastning", "Fagstatistik", "Ikke tildelte", "Elevdata", "Lærerdata"]
            selected_export_sheets = st.multiselect("Vælg ark til Excel-filen", export_sheet_options, default=export_sheet_options[:4], key="v2_export_sheets")
            result_export = try_excel_export(lambda: make_export(students, teachers, solution, capacities, solution["use_global_k"], selected_export_sheets)) if selected_export_sheets else None
            if result_export is not None:
                st.download_button("Download fordeling som Excel", data=result_export, file_name="vejlederfordeling.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", key="v2_download_result")

            result_students, result_teachers, result_subjects, result_missing = st.tabs(["Fordeling", "Lærere", "Fagstatistik", "Ikke tildelte"])
            with result_students:
                rows = []
                for index, student in enumerate(students):
                    assigned = solution["assignments"][index]
                    rows.append({"Elev": student["name"], "Klasse": student.get("className", ""), "Fag 1": student["subjects"][0], "Vejleder 1": teacher_label(assigned[0], teacher_map), "Fag 2": student["subjects"][1], "Vejleder 2": teacher_label(assigned[1], teacher_map), "Ønsker": f"{wished_count(student, assigned)}/2"})
                st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True, height=560)
            with result_teachers:
                rows = []
                for teacher in sorted(teachers, key=lambda item: (-solution["loads"].get(item["id"], 0), item["id"])):
                    teacher_id = teacher["id"]
                    limit = min(solution["K"], capacities.get(teacher_id, 0)) if solution["use_global_k"] else capacities.get(teacher_id, 0)
                    rows.append({"Lærer": teacher_label(teacher_id, teacher_map), "Initialer": teacher_id, "Fag": " · ".join(teacher["subjects"]), "Elever": solution["loads"].get(teacher_id, 0), "Max": limit, "Status": "Over max" if solution["loads"].get(teacher_id, 0) > limit else "OK"})
                st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True, height=560)
            with result_subjects:
                st.dataframe(subject_stats(students, teachers, capacities, solution["K"], solution["use_global_k"]), width="stretch", hide_index=True)
            with result_missing:
                missing = [{"Elev": student["name"], "Klasse": student.get("className", ""), "Fag": student["subjects"][slot], "Status": "Mangler vejleder"} for index, student in enumerate(students) for slot in (0, 1) if not solution["assignments"][index][slot]]
                st.dataframe(pd.DataFrame(missing), width="stretch", hide_index=True)
            with st.expander("Manuel justering af vejledere"):
                st.caption("Vælg en anden lærer i dropdownen. Kun lærere, der underviser i det valgte fag, accepteres. Vælg **Ingen ønsker** for at fjerne en tildeling.")
                assignment_options = [NO_WISHES_LABEL] + sorted((teacher_label(teacher["id"], teacher_map) for teacher in teachers), key=str.casefold)
                edit_rows = [{"Elev": student["name"], "Fag 1": student["subjects"][0], "Vejleder 1": teacher_label(solution["assignments"][index][0], teacher_map) if solution["assignments"][index][0] else NO_WISHES_LABEL, "Fag 2": student["subjects"][1], "Vejleder 2": teacher_label(solution["assignments"][index][1], teacher_map) if solution["assignments"][index][1] else NO_WISHES_LABEL} for index, student in enumerate(students)]
                edited = st.data_editor(
                    pd.DataFrame(edit_rows), width="stretch", hide_index=True, disabled=["Elev", "Fag 1", "Fag 2"],
                    column_config={
                        "Vejleder 1": st.column_config.SelectboxColumn("Vejleder 1", options=assignment_options),
                        "Vejleder 2": st.column_config.SelectboxColumn("Vejleder 2", options=assignment_options),
                    }, key="v2_manual_editor",
                )
                if st.button("Gem manuelle ændringer", key="v2_save_manual"):
                    candidates = build_candidates(teachers)
                    label_to_id = {normal_key(teacher_label(teacher["id"], teacher_map)): teacher["id"] for teacher in teachers}
                    errors = []
                    assignments = []
                    for index, row in edited.iterrows():
                        values = []
                        for slot in (1, 2):
                            selected = repair_text(row[f"Vejleder {slot}"])
                            teacher_id = None if not selected or normal_key(selected) == normal_key(NO_WISHES_LABEL) else label_to_id.get(normal_key(selected), normal_key(selected))
                            subject = students[index]["subjects"][slot - 1]
                            if teacher_id and teacher_id not in candidates.get(canonical_subject(subject), []):
                                errors.append(f"{students[index]['name']}: {selected} underviser ikke i {subject} eller findes ikke i lærerlisten.")
                                values.append(None)
                            else:
                                values.append(teacher_id)
                        assignments.append(values)
                    if errors:
                        st.error(" ".join(errors[:8]))
                    else:
                        solution["assignments"] = assignments
                        solution["loads"] = {teacher_id: 0 for teacher_id in teacher_map}
                        solution["double_loads"] = {teacher_id: 0 for teacher_id in teacher_map}
                        for assigned in assignments:
                            for teacher_id in set(item for item in assigned if item):
                                solution["loads"][teacher_id] += 1
                            if assigned[0] and assigned[0] == assigned[1]:
                                solution["double_loads"][assigned[0]] += 1
                        solution["stats"] = score_solution(students, teachers, capacities, assignments, solution["loads"], solution["K"], solution["use_global_k"], prioritize_pairs, prioritize_classes)
                        st.session_state["solution"] = solution
                        st.session_state["v2_schedule"] = None
                        st.success("Fordelingen er opdateret.")
                        st.rerun()
            st.divider()
            if st.button("Forbered Word-filer", key="v2_prepare_docx"):
                with st.spinner("Genererer Word-filer …"):
                    try:
                        st.session_state["v2_docx_zip"] = make_docx_zip(students, teachers, solution)
                    except ImportError:
                        st.error("Word-eksport kræver python-docx.")
            if st.session_state.get("v2_docx_zip"):
                st.download_button("Download SOP-filer (.zip)", data=st.session_state["v2_docx_zip"], file_name="SOP-filer.zip", mime="application/zip", key="v2_download_docx")

    if active_step == "6 · Tidsplan":
        st.title("6 · Tidsplan for vejledning")
        solution = st.session_state.get("solution")
        if solution is None:
            st.info("Beregn først en fordeling i fanen Beregn.")
        else:
            st.caption("Hver elev får én samlet vejledningstid med sine tildelte vejledere. Brug Elevplanen som elevens opslag og Lærerplanen som lærerens dagsorden.")
            time_columns = st.columns(4)
            with time_columns[0]:
                schedule_start = st.time_input("Starttidspunkt", value=dt_time(9, 0), key="v2_schedule_start")
            with time_columns[1]:
                schedule_end = st.time_input("Sluttidspunkt", value=dt_time(15, 0), key="v2_schedule_end")
            with time_columns[2]:
                schedule_student_minutes = st.number_input("Minutter pr. elev", min_value=1, max_value=180, value=20, step=5, key="v2_schedule_student_minutes")
            with time_columns[3]:
                schedule_transition_minutes = st.number_input("Minutter mellem elever", min_value=0, max_value=60, value=5, step=1, key="v2_schedule_transition_minutes")
            pause_columns = st.columns(3)
            with pause_columns[0]:
                schedule_pause_count = st.number_input("Antal pauser", min_value=0, max_value=20, value=2, step=1, key="v2_schedule_pause_count")
            with pause_columns[1]:
                schedule_pause_minutes = st.number_input("Minutter pr. pause", min_value=0, max_value=120, value=15, step=5, key="v2_schedule_pause_minutes")
            with pause_columns[2]:
                schedule_group_pairs = st.checkbox("Saml samme lærerpar mest muligt", value=True, key="v2_schedule_group_pairs", help="Elever med samme lærerpar lægges i sammenhængende blokke, så lærerne skifter færre gange.")
            schedule_depth = st.slider(
                "Algoritmedybde for tidsplan",
                10,
                300,
                80,
                10,
                key="v2_schedule_depth",
                help="Algoritmen afprøver så mange forskellige rækkefølger af lærerpar. En højere værdi kan samle lærerpar bedre, men tager længere tid. Indstillingen har størst effekt, når 'Saml samme lærerpar mest muligt' er slået til.",
            )
            if st.button("Lav eller opdater tidsplan", type="primary", key="v2_make_schedule"):
                try:
                    progress_bar = st.progress(0.0, text="Beregner tidsplan: 0 %")
                    st.session_state["v2_schedule"] = make_schedule(
                        students,
                        teachers,
                        solution,
                        schedule_start,
                        schedule_end,
                        int(schedule_student_minutes),
                        int(schedule_pause_count),
                        int(schedule_pause_minutes),
                        int(schedule_transition_minutes),
                        schedule_group_pairs,
                        search_attempts=int(schedule_depth),
                        progress_callback=lambda value: update_algorithm_progress(progress_bar, "Beregner tidsplan", value),
                    )
                    update_algorithm_progress(progress_bar, "Tidsplan færdig", 1.0)
                    st.success("Tidsplanen er opdateret.")
                except ValueError as error:
                    st.session_state["v2_schedule"] = None
                    st.error(str(error))
            schedule = st.session_state.get("v2_schedule")
            if schedule is not None:
                settings = schedule["settings"]
                metric_columns = st.columns(4)
                metric_columns[0].metric("Elever", len(schedule["students"]))
                metric_columns[1].metric("Lærerpar", len(schedule["pairs"]))
                metric_columns[2].metric("Planlagt tidsforbrug", f"{settings['Planlagt tidsforbrug (min.)']} min.")
                metric_columns[3].metric("Sidste sluttid", settings["Slut"])
                st.info("Elever finder deres tidspunkt i Elevplanen. Lærere kan vælge deres navn i Lærerplanen og se elev, fag og eventuel medvejleder.")
                schedule_students, schedule_teachers, schedule_timeline, schedule_pairs = st.tabs(["Elevplan", "Lærerplan", "Samlet tidslinje", "Lærerpar"])
                with schedule_students:
                    st.dataframe(schedule["students"], width="stretch", hide_index=True, height=560)
                with schedule_teachers:
                    teacher_options = ["Alle lærere"] + sorted(schedule["teachers"]["Lærer"].unique().tolist(), key=str.casefold)
                    selected_teacher = st.selectbox("Vis lærer", teacher_options, key="v2_schedule_teacher_filter")
                    teacher_frame = schedule["teachers"] if selected_teacher == "Alle lærere" else schedule["teachers"][schedule["teachers"]["Lærer"] == selected_teacher]
                    st.dataframe(teacher_frame, width="stretch", hide_index=True, height=560)
                    st.divider()
                    gantt_options = ["Alle lærere"] + sorted(schedule["teachers"]["Lærer"].unique().tolist(), key=str.casefold)
                    gantt_default = 1 if len(gantt_options) > 1 else 0
                    gantt_teacher = st.selectbox("Vis visuel dagsplan for", gantt_options, index=gantt_default, key="v2_schedule_gantt_teacher")
                    st.markdown(make_teacher_gantt_html(schedule, gantt_teacher), unsafe_allow_html=True)
                with schedule_timeline:
                    st.dataframe(schedule["timeline"], width="stretch", hide_index=True, height=560)
                with schedule_pairs:
                    st.dataframe(schedule["pairs"], width="stretch", hide_index=True)
                download_columns = st.columns(2)
                with download_columns[0]:
                    st.download_button(
                        "Download tidsplan som HTML",
                        data=make_schedule_html(schedule).encode("utf-8"),
                        file_name="vejledningsplan.html",
                        mime="text/html",
                        type="primary",
                        key="v2_download_schedule_html",
                    )
                with download_columns[1]:
                    schedule_export = try_excel_export(lambda: make_schedule_excel(schedule))
                    if schedule_export is not None:
                        st.download_button(
                            "Download tidsplan som Excel",
                            data=schedule_export,
                            file_name="vejledningsplan.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            key="v2_download_schedule_excel",
                        )
                    else:
                        st.error("Excel-download kræver openpyxl. Installer projektets requirements.txt for at aktivere eksporten.")


if __name__ == "__main__":
    main_v2()
