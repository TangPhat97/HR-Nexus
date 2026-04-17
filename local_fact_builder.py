from __future__ import annotations

import hashlib
import json
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Callable

import pandas as pd

ATTENDANCE_STATUSES = ("Có mặt", "Vắng có phép", "Vắng không phép", "Đi muộn")
APPLIED_OPTIONS = ("Có", "Một phần", "Không")
LOCAL_SYNC_USER = "LOCAL_POC"

Logger = Callable[[str], None]


@dataclass(frozen=True)
class LocalTrainingSyncResult:
    training_records: pd.DataFrame
    raw_participants: pd.DataFrame
    synced_rows: int
    passed_rows: int
    warning_rows: int
    failed_rows: int
    employees: pd.DataFrame


def _emit(logger: Logger | None, message: str) -> None:
    if logger is not None:
        logger(message)


def _text_scalar(value: Any) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()


def _normalize_token(value: Any) -> str:
    text = _text_scalar(value).replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii").lower()
    chars = [character if character.isalnum() else "_" for character in text]
    normalized = "".join(chars).strip("_")
    while "__" in normalized:
        normalized = normalized.replace("__", "_")
    return normalized


def _text_series(series: pd.Series | None) -> pd.Series:
    if series is None:
        return pd.Series(dtype="string")
    return series.astype("string").fillna("").str.strip()


def _normalize_series(series: pd.Series | None) -> pd.Series:
    return _text_series(series).map(_normalize_token)


def _number_series(series: pd.Series | None) -> pd.Series:
    if series is None:
        return pd.Series(dtype="float64")
    text = series.astype("string").fillna("").str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(text, errors="coerce")


def _iso_date_series(series: pd.Series | None) -> pd.Series:
    if series is None:
        return pd.Series(dtype="string")
    parsed = pd.to_datetime(series, errors="coerce", format="%Y-%m-%d")
    missing = parsed.isna()
    if missing.any():
        parsed.loc[missing] = pd.to_datetime(series[missing], errors="coerce", dayfirst=True)
    formatted = parsed.dt.strftime("%Y-%m-%d")
    return formatted.astype("string").fillna("")


def _derive_month_series(month_series: pd.Series | None, fallback_series: pd.Series | None) -> pd.Series:
    month_text = _text_series(month_series)
    parsed_month = pd.to_datetime(month_text, errors="coerce", format="%Y-%m-%d")
    missing_month = parsed_month.isna()
    if missing_month.any():
        parsed_month.loc[missing_month] = pd.to_datetime(month_text[missing_month], errors="coerce", dayfirst=True)
    fallback_text = _text_series(fallback_series)
    parsed_fallback = pd.to_datetime(fallback_text, errors="coerce", format="%Y-%m-%d")
    missing_fallback = parsed_fallback.isna()
    if missing_fallback.any():
        parsed_fallback.loc[missing_fallback] = pd.to_datetime(fallback_text[missing_fallback], errors="coerce", dayfirst=True)
    direct_month = parsed_month.dt.strftime("%Y-%m")
    fallback_month = parsed_fallback.dt.strftime("%Y-%m")
    month_pattern = month_text.str.match(r"^\d{4}-\d{2}$")
    result = month_text.where(month_pattern, direct_month)
    return result.astype("string").fillna("").mask(result.astype("string").fillna("").eq(""), fallback_month.astype("string").fillna(""))


def _hash_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:10].upper()


def _hash_series(series: pd.Series) -> pd.Series:
    return series.astype("string").fillna("").map(lambda value: _hash_text(str(value)))


def _stable_id_series(prefix: str, series: pd.Series) -> pd.Series:
    return prefix + "-" + _hash_series(series)


def _canonical_option_series(series: pd.Series, allowed_values: tuple[str, ...]) -> pd.Series:
    normalized_map = {_normalize_token(option): option for option in allowed_values}
    return _text_series(series).map(lambda value: normalized_map.get(_normalize_token(value), ""))


def _stringify_number_series(series: pd.Series) -> pd.Series:
    def _stringify(value: Any) -> str:
        if value is None or pd.isna(value):
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    return series.map(_stringify)


def _parse_json(value: Any) -> dict[str, Any]:
    raw = _text_scalar(value)
    if not raw:
        return {}
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def _build_lookup_table(frame: pd.DataFrame, key_columns: list[str]) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(columns=["_lookup_key"])
    base = frame.drop(columns=["__rowNumber"], errors="ignore")
    lookup_frames = []
    for column in key_columns:
        if column not in base.columns:
            continue
        candidate = base.copy()
        candidate["_lookup_key"] = _normalize_series(candidate[column])
        candidate = candidate.loc[candidate["_lookup_key"].ne("")]
        lookup_frames.append(candidate)
    if not lookup_frames:
        return pd.DataFrame(columns=["_lookup_key"])
    lookup = pd.concat(lookup_frames, ignore_index=True, sort=False)
    lookup = lookup.drop_duplicates("_lookup_key")
    lookup["_lookup_found"] = pd.array([True] * len(lookup), dtype="boolean")
    return lookup


def _build_session_lookup(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(columns=["_lookup_key"])
    base = frame.drop(columns=["__rowNumber"], errors="ignore")
    lookup_frames = []
    key_series_list = [
        _normalize_series(base["session_id"]),
        _normalize_series(base["session_code"]),
        _normalize_series(base["course_id"].mask(base["course_id"].eq(""), base["course_name"]) + "::" + base["training_date"]),
        _normalize_series(base["course_name"] + "::" + base["training_date"]),
    ]
    for key_series in key_series_list:
        candidate = base.copy()
        candidate["_lookup_key"] = key_series
        candidate = candidate.loc[candidate["_lookup_key"].ne("")]
        lookup_frames.append(candidate)
    lookup = pd.concat(lookup_frames, ignore_index=True, sort=False)
    lookup = lookup.drop_duplicates("_lookup_key")
    lookup["_lookup_found"] = pd.array([True] * len(lookup), dtype="boolean")
    return lookup


def _match_lookup_by_priority(
    source: pd.DataFrame,
    key_candidates: list[tuple[str, pd.Series]],
    lookup: pd.DataFrame,
    prefix: str,
) -> pd.DataFrame:
    base = source[["__rowNumber"]].drop_duplicates().copy()
    if lookup.empty:
        return base
    matched_frames = []
    for priority, (_, key_series) in enumerate(key_candidates):
        candidate = pd.DataFrame(
            {
                "__rowNumber": source["__rowNumber"],
                "_lookup_key": key_series.fillna(""),
                "_priority": priority,
            }
        )
        candidate = candidate.loc[candidate["_lookup_key"].ne("")]
        if candidate.empty:
            continue
        merged = candidate.merge(lookup, on="_lookup_key", how="left")
        merged = merged.loc[merged["_lookup_found"].astype("boolean").fillna(False)]
        if merged.empty:
            continue
        matched_frames.append(merged)
    if not matched_frames:
        return base
    resolved = pd.concat(matched_frames, ignore_index=True, sort=False)
    resolved = resolved.sort_values(["__rowNumber", "_priority"]).drop_duplicates("__rowNumber")
    value_columns = [column for column in resolved.columns if column not in {"__rowNumber", "_lookup_key", "_priority"}]
    renamed = resolved[["__rowNumber"] + value_columns].rename(columns={column: f"{prefix}{column}" for column in value_columns})
    return base.merge(renamed, on="__rowNumber", how="left")


def _normalize_employees(frame: pd.DataFrame) -> pd.DataFrame:
    df = frame.copy()
    return df.assign(
        row_status_key=_normalize_series(df.get("row_status", pd.Series(dtype="object"))),
        emp_id=_text_series(df.get("emp_id")),
        email=_text_series(df.get("email")).mask(_text_series(df.get("email")).eq("0"), ""),
        full_name=_text_series(df.get("full_name")),
        employment_status=_text_series(df.get("employment_status")),
        company=_text_series(df.get("company")),
        division=_text_series(df.get("division")),
        department=_text_series(df.get("department")),
        job_title=_text_series(df.get("job_title")),
        grade=_text_series(df.get("grade")),
        level=_text_series(df.get("level")),
        region=_text_series(df.get("region")),
    )


def _normalize_courses(frame: pd.DataFrame) -> pd.DataFrame:
    df = frame.copy()
    return df.assign(
        row_status_key=_normalize_series(df.get("row_status", pd.Series(dtype="object"))),
        course_id=_text_series(df.get("course_id")),
        course_name=_text_series(df.get("course_name")),
        course_category=_text_series(df.get("course_category")),
        platform=_text_series(df.get("platform")),
        delivery_type=_text_series(df.get("delivery_type")),
        training_format_default=_text_series(df.get("training_format_default")),
        training_unit=_text_series(df.get("training_unit")),
        target_audience=_text_series(df.get("target_audience")),
        company_scope=_text_series(df.get("company_scope")),
        duration_hours=_number_series(df.get("duration_hours")),
        cost_per_pax=_number_series(df.get("cost_per_pax")),
    )


def _normalize_sessions(frame: pd.DataFrame, courses: pd.DataFrame) -> pd.DataFrame:
    df = frame.copy().assign(
        row_status_key=_normalize_series(frame.get("row_status", pd.Series(dtype="object"))),
        session_id=_text_series(frame.get("session_id")),
        session_code=_text_series(frame.get("session_code")),
        course_id=_text_series(frame.get("course_id")),
        course_name=_text_series(frame.get("course_name")),
        delivery_type=_text_series(frame.get("delivery_type")),
        training_format=_text_series(frame.get("training_format")),
        location=_text_series(frame.get("location")),
        training_unit=_text_series(frame.get("training_unit")),
        target_audience=_text_series(frame.get("target_audience")),
        company_scope=_text_series(frame.get("company_scope")),
        training_date=_iso_date_series(frame.get("training_date")),
        training_month=_derive_month_series(frame.get("training_month"), frame.get("training_date")),
        class_count=_number_series(frame.get("class_count")),
        registered_count=_number_series(frame.get("registered_count")),
        actual_count=_number_series(frame.get("actual_count")),
        duration_hours=_number_series(frame.get("duration_hours")),
        estimated_cost=_number_series(frame.get("estimated_cost")),
        instructor_cost=_number_series(frame.get("instructor_cost")),
        organization_cost=_number_series(frame.get("organization_cost")),
        avg_cost_per_pax=_number_series(frame.get("avg_cost_per_pax")),
        iso_request_reference=_text_series(frame.get("iso_request_reference")),
        iso_budget_reference=_text_series(frame.get("iso_budget_reference")),
        iso_attendance_evidence=_text_series(frame.get("iso_attendance_evidence")),
        iso_other_reference=_text_series(frame.get("iso_other_reference")),
    )
    matched_courses = _match_lookup_by_priority(
        df.assign(course_id_key=_normalize_series(df["course_id"]), course_name_key=_normalize_series(df["course_name"])),
        [("course_id_key", _normalize_series(df["course_id"])), ("course_name_key", _normalize_series(df["course_name"]))],
        _build_lookup_table(courses, ["course_id", "course_name"]),
        "course_",
    )
    merged = df.merge(matched_courses, on="__rowNumber", how="left")
    session_key = merged["session_code"] + "::" + merged["course_id"] + "::" + merged["course_name"] + "::" + merged["training_date"]
    return merged.assign(
        session_id=merged["session_id"].mask(merged["session_id"].eq(""), _stable_id_series("SES", session_key)),
        course_id=merged["course_id"].mask(merged["course_id"].eq(""), merged["course_course_id"].fillna("")),
        course_name=merged["course_name"].mask(merged["course_name"].eq(""), merged["course_course_name"].fillna("")),
        delivery_type=merged["delivery_type"].mask(merged["delivery_type"].eq(""), merged["course_delivery_type"].fillna("")),
        training_format=merged["training_format"].mask(merged["training_format"].eq(""), merged["course_training_format_default"].fillna("")),
        training_unit=merged["training_unit"].mask(merged["training_unit"].eq(""), merged["course_training_unit"].fillna("")),
        target_audience=merged["target_audience"].mask(merged["target_audience"].eq(""), merged["course_target_audience"].fillna("")),
        company_scope=merged["company_scope"].mask(merged["company_scope"].eq(""), merged["course_company_scope"].fillna("")),
        duration_hours=merged["duration_hours"].fillna(merged["course_duration_hours"]),
        avg_cost_per_pax=merged["avg_cost_per_pax"].fillna(merged["course_cost_per_pax"]),
    )


def _normalize_raw(frame: pd.DataFrame) -> pd.DataFrame:
    df = frame.copy().assign(
        raw_id=_text_series(frame.get("raw_id")),
        session_id=_text_series(frame.get("session_id")),
        session_code=_text_series(frame.get("session_code")),
        course_id=_text_series(frame.get("course_id")),
        course_name=_text_series(frame.get("course_name")),
        training_date=_iso_date_series(frame.get("training_date")),
        emp_id=_text_series(frame.get("emp_id")),
        full_name=_text_series(frame.get("full_name")),
        email=_text_series(frame.get("email")).mask(_text_series(frame.get("email")).eq("0"), ""),
        attendance_status=_text_series(frame.get("attendance_status")),
        score=_number_series(frame.get("score")),
        satisfaction=_number_series(frame.get("satisfaction")),
        relevance=_number_series(frame.get("relevance")),
        nps=_number_series(frame.get("nps")),
        applied_on_job=_text_series(frame.get("applied_on_job")),
        manager_comment=_text_series(frame.get("manager_comment")),
        source_row_hash=_text_series(frame.get("source_row_hash")),
        row_status=_text_series(frame.get("row_status")),
    )
    hash_payload = pd.DataFrame(
        {
            "session_id": df["session_id"],
            "session_code": df["session_code"],
            "course_id": df["course_id"],
            "course_name": df["course_name"],
            "training_date": df["training_date"],
            "emp_id": df["emp_id"],
            "full_name": df["full_name"],
            "email": df["email"],
            "attendance_status": df["attendance_status"],
            "score": _stringify_number_series(df["score"]),
            "satisfaction": _stringify_number_series(df["satisfaction"]),
            "relevance": _stringify_number_series(df["relevance"]),
            "nps": _stringify_number_series(df["nps"]),
            "applied_on_job": df["applied_on_job"],
            "manager_comment": df["manager_comment"],
        }
    ).fillna("")
    raw_id_key = df["session_id"] + "::" + df["session_code"] + "::" + df["course_id"] + "::" + df["course_name"] + "::" + df["training_date"] + "::" + df["emp_id"] + "::" + df["email"] + "::" + df["full_name"]
    participant_identity = df["emp_id"].mask(df["emp_id"].eq(""), df["email"]).mask(df["emp_id"].eq("") & df["email"].eq(""), df["full_name"])
    return df.assign(
        source_row_hash=df["source_row_hash"].mask(df["source_row_hash"].eq(""), _hash_series(hash_payload.astype("string").fillna("").agg("::".join, axis=1))),
        raw_id=df["raw_id"].mask(df["raw_id"].eq(""), _stable_id_series("RAW", raw_id_key)),
        row_status_key=_normalize_series(df["row_status"]),
        emp_id_key=_normalize_series(df["emp_id"]),
        email_key=_normalize_series(df["email"]),
        full_name_key=_normalize_series(df["full_name"]),
        session_id_key=_normalize_series(df["session_id"]),
        session_code_key=_normalize_series(df["session_code"]),
        course_id_key=_normalize_series(df["course_id"]),
        course_name_key=_normalize_series(df["course_name"]),
        course_date_key=_normalize_series(df["course_id"].mask(df["course_id"].eq(""), df["course_name"]) + "::" + df["training_date"]),
        course_name_date_key=_normalize_series(df["course_name"] + "::" + df["training_date"]),
        participant_key=_normalize_series(
            df["session_id"].mask(df["session_id"].eq(""), df["session_code"]).mask(df["session_id"].eq("") & df["session_code"].eq(""), df["course_name"])
            + "::"
            + participant_identity
        ),
    ).loc[~_normalize_series(df["row_status"]).eq("inactive")].copy()


def _build_existing_raw_lookup(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(columns=["_lookup_key", "record_id", "created_at", "created_by"])
    working = frame.copy()
    source_series = working["source_type"] if "source_type" in working.columns else pd.Series("", index=working.index, dtype="object")
    working = working.loc[_normalize_series(source_series).eq("raw_participant")].copy()
    if working.empty:
        return pd.DataFrame(columns=["_lookup_key", "record_id", "created_at", "created_by"])
    metadata = working.get("metadata_json", pd.Series(dtype="object")).map(_parse_json)
    working["existing_raw_id"] = metadata.map(lambda item: _text_scalar(item.get("raw_id") if isinstance(item, dict) else ""))
    working["existing_participant_key"] = metadata.map(lambda item: _normalize_token(item.get("raw_participant_key") if isinstance(item, dict) else ""))
    fallback_key = _normalize_series(
        _text_series(working.get("session_id")).mask(_text_series(working.get("session_id")).eq(""), _text_series(working.get("session_code")))
        + "::"
        + _text_series(working.get("emp_id")).mask(_text_series(working.get("emp_id")).eq(""), _text_series(working.get("email")).mask(_text_series(working.get("email")).eq(""), _text_series(working.get("full_name"))))
    )
    lookup = pd.concat(
        [
            working.assign(_lookup_key=_normalize_series(working["existing_raw_id"]))[["_lookup_key", "record_id", "created_at", "created_by"]],
            working.assign(_lookup_key=working["existing_participant_key"].mask(working["existing_participant_key"].eq(""), fallback_key))[["_lookup_key", "record_id", "created_at", "created_by"]],
        ],
        ignore_index=True,
    )
    lookup = lookup.loc[lookup["_lookup_key"].ne("")].drop_duplicates("_lookup_key")
    lookup["_lookup_found"] = pd.array([True] * len(lookup), dtype="boolean")
    return lookup


def _merge_sheet_updates(original: pd.DataFrame, updates: pd.DataFrame) -> pd.DataFrame:
    merged = original.merge(updates, on="__rowNumber", how="left", suffixes=("", "__new"))
    for column in updates.columns:
        if column == "__rowNumber":
            continue
        if column not in merged.columns and f"{column}__new" not in merged.columns:
            continue
        new_column = f"{column}__new"
        if new_column not in merged.columns:
            merged[column] = merged[column]
            continue
        if column in merged.columns:
            merged[column] = merged[new_column].where(~merged[new_column].isna(), merged[column])
        else:
            merged[column] = merged[new_column]
        merged = merged.drop(columns=[new_column])
    return merged


def build_local_training_sync(
    employees: pd.DataFrame,
    courses: pd.DataFrame,
    training_sessions: pd.DataFrame,
    raw_participants: pd.DataFrame,
    existing_records: pd.DataFrame,
    logger: Logger | None = None,
) -> LocalTrainingSyncResult:
    raw_original = raw_participants.copy()
    active_raw = _normalize_raw(raw_original)
    if active_raw.empty:
        _emit(logger, "Khong co dong raw active de dong bo local.")
        return LocalTrainingSyncResult(existing_records, raw_original, 0, 0, 0, 0)

    employees_all = _normalize_employees(employees)
    employees_active = employees_all.loc[~employees_all["row_status_key"].eq("inactive")].copy()
    courses_all = _normalize_courses(courses)
    courses_active = courses_all.loc[~courses_all["row_status_key"].eq("inactive")].copy()
    sessions_active = _normalize_sessions(training_sessions, courses_active)
    sessions_active = sessions_active.loc[~sessions_active["row_status_key"].eq("inactive")].copy()

    merged = (
        active_raw
        .merge(_match_lookup_by_priority(active_raw, [("session_id", active_raw["session_id_key"]), ("session_code", active_raw["session_code_key"]), ("course_date", active_raw["course_date_key"]), ("course_name_date", active_raw["course_name_date_key"])], _build_session_lookup(sessions_active), "session_"), on="__rowNumber", how="left")
        .merge(_match_lookup_by_priority(active_raw, [("course_id", active_raw["course_id_key"]), ("course_name", active_raw["course_name_key"])], _build_lookup_table(courses_active, ["course_id", "course_name"]), "course_"), on="__rowNumber", how="left")
        .merge(_match_lookup_by_priority(active_raw, [("emp_id", active_raw["emp_id_key"]), ("email", active_raw["email_key"]), ("full_name", active_raw["full_name_key"])], _build_lookup_table(employees_active, ["emp_id", "email", "full_name"]), "employee_"), on="__rowNumber", how="left")
        .merge(_match_lookup_by_priority(active_raw, [("emp_id", active_raw["emp_id_key"]), ("email", active_raw["email_key"]), ("full_name", active_raw["full_name_key"])], _build_lookup_table(employees_all, ["emp_id", "email", "full_name"]), "employee_all_"), on="__rowNumber", how="left")
    )

    attendance = _canonical_option_series(merged["attendance_status"], ATTENDANCE_STATUSES)
    applied = _canonical_option_series(merged["applied_on_job"], APPLIED_OPTIONS)
    employee_present = merged[["emp_id", "email", "full_name"]].astype("string").fillna("").apply(lambda row: any(value.strip() for value in row), axis=1)
    emp_found = merged["employee__lookup_found"].astype("boolean").fillna(False)
    emp_all_found = merged["employee_all__lookup_found"].astype("boolean").fillna(False)
    ses_found = merged["session__lookup_found"].astype("boolean").fillna(False)

    # Logic for each row
    def _generate_audit_note(row):
        notes = []

        found_emp = row.get("employee__lookup_found")
        found_emp_all = row.get("employee_all__lookup_found")
        found_ses = row.get("session__lookup_found")
        
        # Safe boolean conversion
        is_emp_found = False if pd.isna(found_emp) else bool(found_emp)
        is_emp_all_found = False if pd.isna(found_emp_all) else bool(found_emp_all)
        is_ses_found = False if pd.isna(found_ses) else bool(found_ses)

        # Check Employee Link
        if not is_emp_found:
            if is_emp_all_found:
                notes.append("⚠️ NV đang 'Nghỉ việc' hoặc 'Khác'.")
            else:
                notes.append("❌ Không tìm thấy nhân viên trong Danh mục.")
        else:
            # Check Mismatches
            raw_email = _text_scalar(row["email"])
            master_email = _text_scalar(row["employee_email"])
            if raw_email and master_email and raw_email.lower() != master_email.lower():
                notes.append(f"⚠️ Lệch Email: Raw '{raw_email}' vs Danh mục '{master_email}'")
            
            raw_emp_id = _text_scalar(row["emp_id"])
            master_emp_id = _text_scalar(row["employee_emp_id"])
            # Strip trailing '.0' from float-cast IDs (Excel reads numbers as float)
            if raw_emp_id.endswith(".0"):
                raw_emp_id = raw_emp_id[:-2]
            if master_emp_id.endswith(".0"):
                master_emp_id = master_emp_id[:-2]
            if raw_emp_id and master_emp_id and raw_emp_id != master_emp_id:
                notes.append(f"⚠️ Lệch Mã NV: Raw '{raw_emp_id}' vs Danh mục '{master_emp_id}'")

        # Check Session/Course
        if not is_ses_found:
            notes.append("❌ Không tìm thấy Lớp/Khóa học trong Danh mục.")
        else:
            raw_course = _text_scalar(row["course_name"])
            master_course = _text_scalar(row["session_course_name"])
            if raw_course and master_course and _normalize_token(raw_course) != _normalize_token(master_course):
                notes.append(f"⚠️ Tên khóa học lệch: Raw '{raw_course}' vs Danh mục '{master_course}'")

        # Other QA checks
        if _text_series(pd.Series([row["attendance_status"]])).iloc[0] != "" and attendance[row.name] == "":
            notes.append("❌ Trạng thái điểm danh không hợp lệ.")
        if pd.notna(row["score"]) and row["score"] < 0:
            notes.append("❌ Điểm số không được âm.")

        if not notes:
            return "✅ San sang dong bo vao du lieu dao tao."
        return " | ".join(notes)

    final_notes = merged.apply(_generate_audit_note, axis=1)

    fail_mask = (
        ~employee_present
        | (~emp_found & ~emp_all_found & employee_present)
        | ~ses_found
        | (_text_series(merged["attendance_status"]).ne("") & attendance.eq(""))
        | (merged["score"].notna() & (merged["score"] < 0))
        | (merged["satisfaction"].notna() & ~merged["satisfaction"].between(1, 5))
        | (merged["relevance"].notna() & ~merged["relevance"].between(1, 5))
        | (merged["nps"].notna() & ~merged["nps"].between(0, 10))
        | (_text_series(merged["applied_on_job"]).ne("") & applied.eq(""))
        | (merged["participant_key"].ne("") & merged["participant_key"].duplicated(keep=False))
    )

    # Status logic: If there's any "❌" in notes, it's a FAIL. If only "⚠️", it's a WARNING (but still synced).
    def _derive_status(note, is_failed):
        if is_failed or "❌" in note:
            return "QA_FAILED"
        if "⚠️" in note:
            return "QA_WARNING"
        return "QA_PASSED"

    row_statuses = pd.Series([_derive_status(n, f) for n, f in zip(final_notes, fail_mask)], index=merged.index, dtype="string")

    merged = merged.assign(
        attendance_status=attendance.mask(attendance.eq(""), merged["attendance_status"]),
        applied_on_job=applied.mask(applied.eq(""), merged["applied_on_job"]),
        row_status_local=row_statuses,
        notes_local=final_notes,
        updated_at_local=datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        updated_by_local=LOCAL_SYNC_USER,
    )
    merged = merged.assign(
        session_id_resolved=merged["session_session_id"].mask(merged["session_session_id"].fillna("").eq(""), merged["session_id"]),
        session_code_resolved=merged["session_session_code"].mask(merged["session_session_code"].fillna("").eq(""), merged["session_code"]),
        course_id_resolved=merged["course_course_id"].mask(merged["course_course_id"].fillna("").eq(""), merged["session_course_id"].mask(merged["session_course_id"].fillna("").eq(""), merged["course_id"])),
        course_name_resolved=merged["course_course_name"].mask(merged["course_course_name"].fillna("").eq(""), merged["session_course_name"].mask(merged["session_course_name"].fillna("").eq(""), merged["course_name"])),
        training_date_resolved=merged["session_training_date"].mask(merged["session_training_date"].fillna("").eq(""), merged["training_date"]),
        training_month_resolved=merged["session_training_month"].mask(merged["session_training_month"].fillna("").eq(""), _derive_month_series(merged["training_date"], merged["training_date"])),
    )

    passed_rows = merged.loc[merged["row_status_local"].ne("QA_FAILED")].copy()
    next_records = _build_next_records(passed_rows, existing_records, _build_existing_raw_lookup(existing_records))
    source_series = existing_records["source_type"] if "source_type" in existing_records.columns else pd.Series("", index=existing_records.index, dtype="object")
    non_raw_existing = existing_records.loc[_normalize_series(source_series).ne("raw_participant")].copy()
    final_records = next_records.copy() if non_raw_existing.empty else pd.concat([non_raw_existing, next_records], ignore_index=True, sort=False)
    raw_updates = merged[["__rowNumber", "raw_id", "session_id_resolved", "session_code_resolved", "course_id_resolved", "course_name_resolved", "training_date_resolved", "source_row_hash", "row_status_local", "notes_local", "updated_at_local", "updated_by_local"]].rename(columns={"session_id_resolved": "session_id", "session_code_resolved": "session_code", "course_id_resolved": "course_id", "course_name_resolved": "course_name", "training_date_resolved": "training_date", "row_status_local": "row_status", "notes_local": "notes", "updated_at_local": "updated_at", "updated_by_local": "updated_by"})
    final_raw = _merge_sheet_updates(raw_original, raw_updates)
    failed_rows = int(merged["row_status_local"].eq("QA_FAILED").sum())
    
    employees_final = employees.copy()
    email_update_mask = (
        merged["employee__lookup_found"].astype("boolean").fillna(False)
        & _text_series(merged["email"]).ne("")
        & _text_series(merged["employee_email"]).eq("")
        & _text_series(merged["employee_emp_id"]).ne("")
    )
    if email_update_mask.any():
        updates = merged.loc[email_update_mask, ["employee_emp_id", "email"]].drop_duplicates("employee_emp_id")
        update_map = dict(zip(updates["employee_emp_id"], updates["email"]))
        emp_id_series = _text_series(employees_final.get("emp_id"))
        update_mask = emp_id_series.isin(update_map.keys())
        if update_mask.any():
            employees_final.loc[update_mask, "email"] = emp_id_series.loc[update_mask].map(update_map)

    _emit(logger, f"Local sync: {len(passed_rows)} dong hop le, {failed_rows} dong fail.")
    return LocalTrainingSyncResult(final_records, final_raw, len(passed_rows), len(passed_rows), 0, failed_rows, employees_final)


def _build_next_records(
    passed_rows: pd.DataFrame,
    existing_records: pd.DataFrame,
    existing_lookup: pd.DataFrame,
) -> pd.DataFrame:
    if passed_rows.empty:
        return existing_records.iloc[0:0].copy()
    matched_existing = _match_lookup_by_priority(
        passed_rows,
        [("raw_id", _normalize_series(passed_rows["raw_id"])), ("participant_key", passed_rows["participant_key"])],
        existing_lookup,
        "existing_",
    )
    merged = passed_rows.merge(matched_existing, on="__rowNumber", how="left")
    existing_record_id = merged["existing_record_id"] if "existing_record_id" in merged.columns else pd.Series("", index=merged.index, dtype="string")
    existing_created_at = merged["existing_created_at"] if "existing_created_at" in merged.columns else pd.Series("", index=merged.index, dtype="string")
    existing_created_by = merged["existing_created_by"] if "existing_created_by" in merged.columns else pd.Series("", index=merged.index, dtype="string")
    candidate_training_format = merged["training_format"] if "training_format" in merged.columns else pd.Series("", index=merged.index, dtype="string")
    now_text = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    training_date = merged["training_date_resolved"]
    training_month = merged["training_month_resolved"]
    record_key = merged["raw_id"] + "::" + merged["participant_key"]
    metadata_json = pd.DataFrame(
        {
            "raw_id": merged["raw_id"],
            "raw_participant_key": merged["participant_key"],
            "sync_fingerprint": _hash_series(record_key + "::" + merged["source_row_hash"]),
            "source_row_number": merged["__rowNumber"].astype("Int64").astype("string"),
            "notes": "Dong bo tu Data raw hoc vien (local)",
            "sync_mode": "local",
        }
    ).apply(lambda row: json.dumps(row.to_dict(), ensure_ascii=False), axis=1)
    return pd.DataFrame(
        {
            "record_id": existing_record_id.mask(existing_record_id.fillna("").eq(""), _stable_id_series("REC", record_key)),
            "source_type": "RAW_PARTICIPANT",
            "batch_id": "",
            "emp_id": merged["employee_emp_id"].mask(merged["employee_emp_id"].fillna("").eq(""), merged["emp_id"]),
            "full_name": merged["employee_full_name"].mask(merged["employee_full_name"].fillna("").eq(""), merged["full_name"]),
            "email": merged["employee_email"].mask(merged["employee_email"].fillna("").eq(""), merged["email"]),
            "employment_status": merged["employee_employment_status"].fillna(""),
            "company": merged["employee_company"].fillna(""),
            "division": merged["employee_division"].fillna(""),
            "department": merged["employee_department"].fillna(""),
            "job_title": merged["employee_job_title"].fillna(""),
            "grade": merged["employee_grade"].fillna(""),
            "level": merged["employee_level"].fillna(""),
            "course_id": merged["course_id_resolved"],
            "course_name": merged["course_name_resolved"],
            "course_category": merged["course_course_category"].fillna(""),
            "platform": merged["course_platform"].fillna(""),
            "duration_hours": merged["course_duration_hours"].fillna(merged["session_duration_hours"]),
            "cost_per_pax": merged["course_cost_per_pax"].fillna(merged["session_avg_cost_per_pax"]),
            "training_date": training_date,
            "training_format": merged["session_training_format"].mask(merged["session_training_format"].fillna("").eq(""), merged["course_training_format_default"].mask(merged["course_training_format_default"].fillna("").eq(""), candidate_training_format)),
            "attendance_status": merged["attendance_status"],
            "score": merged["score"],
            "satisfaction": merged["satisfaction"],
            "relevance": merged["relevance"],
            "nps": merged["nps"],
            "applied_on_job": merged["applied_on_job"],
            "manager_comment": merged["manager_comment"],
            "created_at": existing_created_at.mask(existing_created_at.fillna("").eq(""), now_text),
            "created_by": existing_created_by.mask(existing_created_by.fillna("").eq(""), LOCAL_SYNC_USER),
            "updated_at": now_text,
            "updated_by": LOCAL_SYNC_USER,
            "archive_year": pd.to_datetime(training_date, errors="coerce").dt.year,
            "row_status": "ACTIVE",
            "source_row_hash": merged["source_row_hash"],
            "qa_status": "QA_PASSED",
            "metadata_json": metadata_json,
            "session_id": merged["session_id_resolved"],
            "session_code": merged["session_code_resolved"],
            "training_month": training_month,
            "location": merged["session_location"].fillna(""),
            "delivery_type": merged["session_delivery_type"].mask(merged["session_delivery_type"].fillna("").eq(""), merged["course_delivery_type"]),
            "training_unit": merged["session_training_unit"].mask(merged["session_training_unit"].fillna("").eq(""), merged["course_training_unit"]),
            "target_audience": merged["session_target_audience"].mask(merged["session_target_audience"].fillna("").eq(""), merged["course_target_audience"]),
            "company_scope": merged["session_company_scope"].mask(merged["session_company_scope"].fillna("").eq(""), merged["course_company_scope"].mask(merged["course_company_scope"].fillna("").eq(""), merged["employee_company"])),
            "class_count": merged["session_class_count"],
            "registered_count": merged["session_registered_count"],
            "actual_count": merged["session_actual_count"],
            "estimated_cost": merged["session_estimated_cost"],
            "instructor_cost": merged["session_instructor_cost"],
            "organization_cost": merged["session_organization_cost"],
            "iso_request_reference": merged["session_iso_request_reference"].fillna(""),
            "iso_budget_reference": merged["session_iso_budget_reference"].fillna(""),
            "iso_attendance_evidence": merged["session_iso_attendance_evidence"].fillna(""),
            "iso_other_reference": merged["session_iso_other_reference"].fillna(""),
            "region": merged["employee_region"].fillna(""),
        }
    )
