from __future__ import annotations

import json
import math
import re
import unicodedata
import hashlib
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Iterable, Mapping

import numpy as np
import pandas as pd
import pytz

PASS_SCORE = 70
ACTIVE_STATUS = "active"
INACTIVE_STATUS = "inactive"
PUBLISHED_STATUS = "published"
QA_FAIL_KEY = "fail"
OPEN_KEY = "open"
PENDING_STATUS = "PENDING"
PROCESSING_STATUS = "PROCESSING"
MANAGED_JOB_TYPES = {"REFRESH_ANALYTICS", "SYNC_TRAINING_DATA"}

SHEET_NAMES = {
    "dashboard_exec": "Bảng điều khiển",
    "dashboard_operations": "Điều hành hệ thống",
    "analytics_course": "Phân tích khóa học",
    "analytics_department": "Phân tích phòng ban",
    "analytics_department_course": "Phòng ban theo khóa",
    "analytics_trend": "Phân tích xu hướng",
    "report_external_assignment": "Cử đi học bên ngoài",
    "report_session_reconciliation": "Đối chiếu lớp đào tạo",
    "analytics_grade_level": "Phân tích ngạch cấp bậc",
    "looker_studio_data": "Looker Studio Data",
}

DASHBOARD_EXEC_HEADERS = [
    "metric_key",
    "metric_label",
    "metric_value",
    "metric_unit",
    "target",
    "last_refreshed",
]

DASHBOARD_OPERATIONS_HEADERS = [
    "metric_key",
    "metric_label",
    "metric_value",
    "metric_unit",
    "status",
    "last_refreshed",
]

ANALYTICS_COURSE_HEADERS = [
    "course_id",
    "course_name",
    "course_category",
    "training_count",
    "unique_employees",
    "avg_score",
    "avg_satisfaction",
    "avg_nps",
    "applied_rate",
    "total_hours",
    "total_cost",
    "unique_sessions",
    "participant_count",
    "attended_count",
]


def _empty_frame() -> pd.DataFrame:
    return pd.DataFrame()


@dataclass(frozen=True)
class AnalyticsInputs:
    training_records: pd.DataFrame
    employees: pd.DataFrame = field(default_factory=_empty_frame)
    training_sessions: pd.DataFrame = field(default_factory=_empty_frame)
    raw_participants: pd.DataFrame = field(default_factory=_empty_frame)
    queue_jobs: pd.DataFrame = field(default_factory=_empty_frame)
    qa_results: pd.DataFrame = field(default_factory=_empty_frame)
    staging_hr: pd.DataFrame = field(default_factory=_empty_frame)
    staging_courses: pd.DataFrame = field(default_factory=_empty_frame)
    staging_training: pd.DataFrame = field(default_factory=_empty_frame)
    config_map: Mapping[str, Any] = field(default_factory=dict)
    raw_sync_summary: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class NormalizedAnalyticsInputs:
    fiscal_year: str
    last_refreshed: str
    training_records: pd.DataFrame
    active_training_records: pd.DataFrame
    employees: pd.DataFrame
    active_employees: pd.DataFrame
    training_sessions: pd.DataFrame
    active_training_sessions: pd.DataFrame
    raw_participants: pd.DataFrame
    active_raw_participants: pd.DataFrame
    queue_jobs: pd.DataFrame
    qa_results: pd.DataFrame
    staging_hr: pd.DataFrame
    staging_courses: pd.DataFrame
    staging_training: pd.DataFrame
    config_map: Mapping[str, Any]
    raw_sync_summary: Mapping[str, Any]
    department_headcount: pd.Series


def fetch_data(*_args: Any, **_kwargs: Any) -> AnalyticsInputs:
    raise NotImplementedError(
        "fetch_data() must live in an adapter layer. transform.py only handles pandas transformations."
    )


def write_data(*_args: Any, **_kwargs: Any) -> None:
    raise NotImplementedError(
        "write_data() must live in an adapter layer. transform.py only returns sheet-ready matrices."
    )


def transform_data(
    inputs: AnalyticsInputs,
    fiscal_year: str | None = None,
    last_refreshed: str | None = None,
) -> dict[str, list[list[Any]]]:
    normalized = normalize_inputs(inputs, fiscal_year=fiscal_year, last_refreshed=last_refreshed)
    return {
        SHEET_NAMES["dashboard_exec"]: build_dashboard_exec_matrix(normalized),
        SHEET_NAMES["dashboard_operations"]: build_dashboard_operations_matrix(normalized),
        SHEET_NAMES["analytics_course"]: build_course_matrix(normalized),
        SHEET_NAMES["analytics_department"]: build_department_canvas(normalized),
        SHEET_NAMES["analytics_department_course"]: build_department_course_canvas(normalized),
        SHEET_NAMES["analytics_trend"]: build_trend_canvas(normalized),
        SHEET_NAMES["report_external_assignment"]: build_external_assignment_canvas(normalized),
        SHEET_NAMES["report_session_reconciliation"]: build_session_reconciliation_canvas(normalized),
        SHEET_NAMES["analytics_grade_level"]: build_grade_level_canvas(normalized),
        SHEET_NAMES["looker_studio_data"]: build_looker_flat_matrix(normalized),
    }


def normalize_inputs(
    inputs: AnalyticsInputs,
    fiscal_year: str | None = None,
    last_refreshed: str | None = None,
) -> NormalizedAnalyticsInputs:
    config_map = dict(inputs.config_map or {})
    resolved_year = (
        _sanitize_scalar(fiscal_year)
        or _sanitize_scalar(config_map.get("CURRENT_FISCAL_YEAR"))
        or str(datetime.now().year)
    )
    if last_refreshed:
        resolved_refreshed = str(last_refreshed)
    else:
        vn_tz = pytz.timezone("Asia/Ho_Chi_Minh")
        resolved_refreshed = datetime.now(vn_tz).strftime("%Y-%m-%dT%H:%M:%S")

    training_records = _normalize_training_records(inputs.training_records)
    employees = _normalize_employees(inputs.employees)
    training_sessions = _normalize_training_sessions(inputs.training_sessions)
    raw_participants = _normalize_raw_participants(inputs.raw_participants)
    queue_jobs = _normalize_queue_jobs(inputs.queue_jobs)
    qa_results = _normalize_qa_results(inputs.qa_results)
    staging_hr = _normalize_staging(inputs.staging_hr)
    staging_courses = _normalize_staging(inputs.staging_courses)
    staging_training = _normalize_staging(inputs.staging_training)

    active_records = training_records.loc[training_records["row_status_key"].eq(ACTIVE_STATUS)].copy()
    active_employees = employees.loc[~employees["row_status_key"].eq(INACTIVE_STATUS)].copy()
    active_training_sessions = training_sessions.loc[
        ~training_sessions["row_status_key"].eq(INACTIVE_STATUS)
    ].copy()
    active_raw_participants = raw_participants.loc[
        ~raw_participants["row_status_key"].eq(INACTIVE_STATUS)
    ].copy()

    department_headcount = (
        active_employees.groupby("department_name", sort=True).size().rename("total_employees")
        if not active_employees.empty
        else pd.Series(dtype="int64", name="total_employees")
    )

    return NormalizedAnalyticsInputs(
        fiscal_year=resolved_year,
        last_refreshed=resolved_refreshed,
        training_records=training_records,
        active_training_records=active_records,
        employees=employees,
        active_employees=active_employees,
        training_sessions=training_sessions,
        active_training_sessions=active_training_sessions,
        raw_participants=raw_participants,
        active_raw_participants=active_raw_participants,
        queue_jobs=queue_jobs,
        qa_results=qa_results,
        staging_hr=staging_hr,
        staging_courses=staging_courses,
        staging_training=staging_training,
        config_map=config_map,
        raw_sync_summary=dict(inputs.raw_sync_summary or {}),
        department_headcount=department_headcount,
    )


def _prepare_frame(frame: pd.DataFrame | None, columns: Iterable[str]) -> pd.DataFrame:
    if frame is None:
        return pd.DataFrame({column: pd.Series(dtype="object") for column in columns})
    df = frame.copy()
    for column in columns:
        if column not in df.columns:
            df[column] = pd.NA
    return df.loc[:, list(columns)].copy()


def _sanitize_scalar(value: Any) -> str:
    if value is None or value is pd.NA:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
    return str(value).strip()


def _normalize_key(value: Any) -> str:
    sanitized = _sanitize_scalar(value)
    if not sanitized:
        return ""
    normalized = unicodedata.normalize("NFKD", sanitized)
    normalized = normalized.encode("ascii", "ignore").decode("ascii")
    normalized = normalized.lower()
    normalized = re.sub(r"[^a-z0-9]+", "_", normalized)
    normalized = re.sub(r"^_+|_+$", "", normalized)
    return normalized


def _parse_json_scalar(value: Any) -> dict[str, Any]:
    if isinstance(value, Mapping):
        return dict(value)
    raw = _sanitize_scalar(value)
    if not raw:
        return {}
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def _string_series_from(series: pd.Series) -> pd.Series:
    if series.empty:
        return pd.Series(dtype="string")
    return series.map(_sanitize_scalar).astype("string")


def _sanitize_series(series: pd.Series) -> pd.Series:
    return _string_series_from(series)


def _normalize_key_series(series: pd.Series) -> pd.Series:
    sanitized = _sanitize_series(series)
    if sanitized.empty:
        return sanitized
    normalized = sanitized.str.normalize("NFKD")
    normalized = normalized.str.encode("ascii", errors="ignore").str.decode("ascii")
    normalized = normalized.str.lower()
    normalized = normalized.str.replace(r"[^a-z0-9]+", "_", regex=True)
    normalized = normalized.str.replace(r"^_+|_+$", "", regex=True)
    return normalized


def _numeric_series(series: pd.Series) -> pd.Series:
    if series.empty:
        return pd.Series(dtype="float64")
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    normalized = _sanitize_series(series).str.replace(",", "", regex=False)
    return pd.to_numeric(normalized, errors="coerce")


def _iso_date_series(series: pd.Series) -> pd.Series:
    raw = _sanitize_series(series)
    if raw.empty:
        return raw.astype("object")
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    slash_mask = raw.str.match(r"^\d{1,2}/\d{1,2}/\d{4}$")
    if slash_mask.any():
        parsed.loc[slash_mask] = pd.to_datetime(raw.loc[slash_mask], format="%d/%m/%Y", errors="coerce")
    if (~slash_mask).any():
        parsed.loc[~slash_mask] = pd.to_datetime(series.loc[~slash_mask], errors="coerce")
    return parsed.dt.strftime("%Y-%m-%d").fillna("")


def _format_date_series(series: pd.Series) -> pd.Series:
    iso = _iso_date_series(series)
    return pd.Series(
        np.where(iso.eq(""), "", iso.str.slice(8, 10) + "/" + iso.str.slice(5, 7) + "/" + iso.str.slice(0, 4)),
        index=series.index,
    )


def _derive_training_month_series(series: pd.Series, fallback: pd.Series | None = None) -> pd.Series:
    raw = _sanitize_series(series)
    derived = pd.Series("", index=series.index, dtype="object")
    iso_from_value = _iso_date_series(series)
    derived.loc[iso_from_value.ne("")] = iso_from_value.loc[iso_from_value.ne("")].str.slice(0, 7)
    month_like = raw.str.match(r"^\d{4}-\d{2}$")
    derived.loc[derived.eq("") & month_like] = raw.loc[derived.eq("") & month_like]
    if fallback is not None:
        fallback_series = _string_series_from(fallback)
        derived.loc[derived.eq("") & fallback_series.ne("")] = fallback_series.loc[
            derived.eq("") & fallback_series.ne("")
        ].str.slice(0, 7)
    return derived


def _first_nonblank(series_list: Iterable[pd.Series], default: str = "") -> pd.Series:
    frame = pd.concat([_sanitize_series(series).replace("", pd.NA) for series in series_list], axis=1)
    if frame.empty:
        return pd.Series(dtype="object")
    return frame.bfill(axis=1).iloc[:, 0].fillna(default).astype("object")


def _first_nonblank_value(series: pd.Series) -> str:
    cleaned = _sanitize_series(series).replace("", pd.NA).dropna()
    return cleaned.iloc[0] if not cleaned.empty else ""


def _first_nonblank_numeric_or_blank(series: pd.Series) -> float | None:
    cleaned = pd.to_numeric(series, errors="coerce")
    cleaned = cleaned.loc[cleaned.notna()]
    return float(cleaned.iloc[0]) if not cleaned.empty else np.nan


def _unique_nonblank_count(series: pd.Series) -> int:
    cleaned = _sanitize_series(series).replace("", pd.NA)
    return int(cleaned.nunique(dropna=True))


def _mean_or_blank(series: pd.Series) -> float | None:
    cleaned = pd.to_numeric(series, errors="coerce").dropna()
    return round(float(cleaned.mean()), 2) if not cleaned.empty else np.nan


def _sum_or_zero(series: pd.Series) -> float:
    cleaned = pd.to_numeric(series, errors="coerce").fillna(0)
    return round(float(cleaned.sum()), 2)


def _sum_or_blank_if_all_missing(series: pd.Series) -> float | None:
    cleaned = pd.to_numeric(series, errors="coerce")
    if cleaned.notna().sum() == 0:
        return np.nan
    return round(float(cleaned.fillna(0).sum()), 2)


def _ratio_value(numerator: float | int, denominator: float | int) -> str:
    if not denominator:
        return ""
    pct = round(float(numerator) / float(denominator) * 100, 2)
    return f"{int(pct)}%" if pct.is_integer() else f"{pct}%"


def _ratio_series(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    numerator_values = pd.to_numeric(numerator, errors="coerce").fillna(0)
    denominator_values = pd.to_numeric(denominator, errors="coerce")
    result = pd.Series("", index=numerator.index, dtype="object")
    valid = denominator_values.notna() & denominator_values.ne(0)
    
    pct = (numerator_values.loc[valid] / denominator_values.loc[valid] * 100).round(2)
    result.loc[valid] = pct.apply(lambda x: f"{int(x)}%" if x.is_integer() else f"{x}%")
    return result


def _safe_divide_series(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    numerator_values = pd.to_numeric(numerator, errors="coerce")
    denominator_values = pd.to_numeric(denominator, errors="coerce")
    result = pd.Series(np.nan, index=numerator.index, dtype="float64")
    valid = numerator_values.notna() & denominator_values.notna() & denominator_values.ne(0)
    result.loc[valid] = (numerator_values.loc[valid] / denominator_values.loc[valid]).round(2)
    return result


def _assign_rank(df: pd.DataFrame, value_column: str, fallback_column: str, target_column: str) -> pd.DataFrame:
    ranked = df.copy()
    order = ranked.sort_values(
        [value_column, fallback_column],
        ascending=[False, True],
        na_position="last",
    ).index
    ranks = pd.Series(np.arange(1, len(ranked) + 1), index=order)
    ranked[target_column] = ranks.sort_index().to_numpy()
    return ranked


def _sort_desc(df: pd.DataFrame, value_column: str, fallback_column: str) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    return df.sort_values([value_column, fallback_column], ascending=[False, True], na_position="last").reset_index(
        drop=True
    )


def _format_timestamp_display(value: Any) -> str:
    raw = _sanitize_scalar(value)
    if not raw:
        return ""
    parsed = pd.to_datetime(raw, errors="coerce")
    if pd.isna(parsed):
        return raw
    return parsed.strftime("%d/%m/%Y %H:%M:%S")


def _frame_to_rows(frame: pd.DataFrame, columns: list[str]) -> list[list[Any]]:
    if frame.empty:
        return []
    subset = frame.loc[:, columns]
    values = subset.astype("object").where(pd.notna(subset), None).values.tolist()
    return _pythonize_rows(values)


def _blank_row(width: int) -> list[str]:
    return [""] * width


def _pad_rows(rows: list[list[Any]], width: int) -> list[list[Any]]:
    return [row + [""] * (width - len(row)) if len(row) < width else row[:width] for row in rows]


def _pythonize_rows(rows: list[list[Any]]) -> list[list[Any]]:
    return [[_to_python_cell(cell) for cell in row] for row in rows]


def _to_python_cell(value: Any) -> Any:
    if value is None or value is pd.NA:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (np.integer, int)):
        return int(value)
    if isinstance(value, (np.floating, float)):
        if math.isnan(value):
            return ""
        rounded = round(float(value), 2)
        return int(rounded) if rounded.is_integer() else rounded
    if pd.isna(value):
        return ""
    return value


def _normalize_training_records(frame: pd.DataFrame) -> pd.DataFrame:
    df = _prepare_frame(
        frame,
        [
            "row_status",
            "__rowNumber",
            "emp_id",
            "full_name",
            "email",
            "department",
            "division",
            "job_title",
            "grade",
            "level",
            "session_id",
            "session_code",
            "course_id",
            "course_name",
            "course_category",
            "training_date",
            "training_month",
            "delivery_type",
            "training_format",
            "training_unit",
            "location",
            "attendance_status",
            "plan_scope",
            "program_type",
            "score",
            "satisfaction",
            "nps",
            "applied_on_job",
            "duration_hours",
            "cost_per_pax",
            "estimated_cost",
            "avg_cost_per_pax",
            "registered_count",
            "actual_count",
        ],
    )
    training_date_iso = _iso_date_series(df["training_date"])
    training_month_key = _derive_training_month_series(df["training_month"], fallback=training_date_iso)
    course_id = _sanitize_series(df["course_id"])
    course_name = _sanitize_series(df["course_name"])
    course_group_key = _first_nonblank([course_id, course_name], "UNSPECIFIED")
    attendance_key = _normalize_key_series(df["attendance_status"])
    return df.assign(
        row_status_key=_normalize_key_series(df["row_status"]),
        row_number=_numeric_series(df["__rowNumber"]).fillna(0),
        emp_id=_sanitize_series(df["emp_id"]),
        full_name=_sanitize_series(df["full_name"]),
        email=_sanitize_series(df["email"]),
        department_name=_first_nonblank([_sanitize_series(df["department"])], "Không xác định"),
        division=_sanitize_series(df["division"]),
        job_title=_sanitize_series(df["job_title"]),
        grade_name=_first_nonblank([_sanitize_series(df["grade"])], "Không xác định"),
        level_name=_first_nonblank([_sanitize_series(df["level"])], "Không xác định"),
        session_id=_sanitize_series(df["session_id"]),
        session_code=_sanitize_series(df["session_code"]),
        course_id=course_id,
        course_name=course_name,
        course_category=_sanitize_series(df["course_category"]),
        training_date_iso=training_date_iso,
        training_month_key=training_month_key,
        year_key=training_month_key.str.slice(0, 4),
        delivery_type=_sanitize_series(df["delivery_type"]),
        delivery_type_key=_normalize_key_series(df["delivery_type"]),
        training_format=_sanitize_series(df["training_format"]),
        training_format_key=_normalize_key_series(df["training_format"]),
        training_unit=_sanitize_series(df["training_unit"]),
        location=_sanitize_series(df["location"]),
        attendance_status=_sanitize_series(df["attendance_status"]),
        attendance_key=attendance_key,
        attendance_is_present=attendance_key.eq(_normalize_key("Có mặt")).astype(int),
        attendance_is_excused_absent=attendance_key.eq(_normalize_key("Vắng có phép")).astype(int),
        attendance_is_unexcused_absent=attendance_key.eq(_normalize_key("Vắng không phép")).astype(int),
        attendance_is_late=attendance_key.eq(_normalize_key("Đi muộn")).astype(int),
        plan_scope=_sanitize_series(df["plan_scope"]),
        program_type=_sanitize_series(df["program_type"]),
        score=_numeric_series(df["score"]),
        satisfaction=_numeric_series(df["satisfaction"]),
        nps=_numeric_series(df["nps"]),
        applied_on_job=_sanitize_series(df["applied_on_job"]),
        applied_is_yes=_normalize_key_series(df["applied_on_job"]).eq(_normalize_key("Có")).astype(int),
        duration_hours=_numeric_series(df["duration_hours"]).fillna(0),
        cost_per_pax=_numeric_series(df["cost_per_pax"]).fillna(0),
        estimated_cost=_numeric_series(df["estimated_cost"]),
        avg_cost_per_pax=_numeric_series(df["avg_cost_per_pax"]),
        registered_count=_numeric_series(df["registered_count"]),
        actual_count=_numeric_series(df["actual_count"]),
        course_group_key=course_group_key,
        session_key=_first_nonblank(
            [
                _sanitize_series(df["session_id"]),
                _sanitize_series(df["session_code"]),
                course_group_key + "::" + training_date_iso,
            ],
            "",
        ),
        session_group_key=_first_nonblank(
            [
                _sanitize_series(df["session_id"]),
                _sanitize_series(df["session_code"]),
                course_id,
                course_name,
            ],
            "",
        ),
        participant_identity_key=_first_nonblank(
            [
                _normalize_key_series(df["emp_id"]),
                _normalize_key_series(df["email"]),
                _normalize_key_series(df["full_name"]),
            ],
            "",
        ),
    )


def _normalize_employees(frame: pd.DataFrame) -> pd.DataFrame:
    df = _prepare_frame(frame, ["row_status", "department", "emp_id", "email", "full_name", "grade", "level"])
    return df.assign(
        row_status_key=_normalize_key_series(df["row_status"]),
        department_name=_first_nonblank([_sanitize_series(df["department"])], "Không xác định"),
        emp_id=_sanitize_series(df["emp_id"]),
        email=_sanitize_series(df["email"]).mask(_sanitize_series(df["email"]).eq("0"), "").mask(_sanitize_series(df["email"]).eq(""), "no-email@gonsa.com.vn"),
        full_name=_sanitize_series(df["full_name"]),
        grade_name=_first_nonblank([_sanitize_series(df["grade"])], "Không xác định"),
        level_name=_first_nonblank([_sanitize_series(df["level"])], "Không xác định"),
    )


def _generate_gas_session_hash(session_code: str, course_id: str, course_name: str, training_date_iso: str) -> str:
    """Generate GAS-equivalent deterministic SES- ID"""
    raw_key = f"{session_code}::{course_id}::{course_name}::{training_date_iso}"
    json_str = json.dumps(raw_key, ensure_ascii=False)
    hash_hex = hashlib.sha256(json_str.encode('utf-8')).hexdigest()
    return f"SES-{hash_hex[:10].upper()}"

def _generate_gas_raw_hash(session_id: str, session_code: str, course_id: str, course_name: str, 
                           training_date: str, emp_id: str, email: str, full_name: str) -> str:
    """Generate GAS-equivalent deterministic RAW- ID"""
    raw_key = f"{session_id}::{session_code}::{course_id}::{course_name}::{training_date}::{emp_id}::{email}::{full_name}"
    json_str = json.dumps(raw_key, ensure_ascii=False)
    hash_hex = hashlib.sha256(json_str.encode('utf-8')).hexdigest()
    return f"RAW-{hash_hex[:10].upper()}"


def _normalize_training_sessions(frame: pd.DataFrame) -> pd.DataFrame:
    df = _prepare_frame(
        frame,
        [
            "row_status",
            "__rowNumber",
            "session_id",
            "session_code",
            "course_id",
            "course_name",
            "training_date",
            "training_month",
            "training_unit",
            "delivery_type",
            "training_format",
            "registered_count",
            "actual_count",
            "attendance_rate",
            "estimated_cost",
            "avg_cost_per_pax",
            "cost_per_pax",
            "total_hours",
        ],
    )
    training_date_iso = _iso_date_series(df["training_date"])
    training_month_key = _derive_training_month_series(df["training_month"], fallback=training_date_iso)
    course_id = _sanitize_series(df["course_id"])
    course_name = _sanitize_series(df["course_name"])
    session_code = _sanitize_series(df["session_code"])
    raw_session_id = _sanitize_series(df["session_id"])
    
    # Auto-generate session_id imitating GAS deterministic rules
    generated_session_id = pd.Series([
        _generate_gas_session_hash(sc, cid, cname, tdate) 
        for sc, cid, cname, tdate in zip(session_code, course_id, course_name, training_date_iso)
    ], index=df.index)
    
    final_session_id = _first_nonblank([raw_session_id, generated_session_id], "")

    return df.assign(
        row_status_key=_normalize_key_series(df["row_status"]),
        row_number=_numeric_series(df["__rowNumber"]).fillna(0),
        session_id=final_session_id,
        session_code=session_code,
        course_id=course_id,
        course_name=course_name,
        training_date_iso=training_date_iso,
        training_month_key=training_month_key,
        year_key=training_month_key.str.slice(0, 4),
        training_unit=_sanitize_series(df["training_unit"]),
        delivery_type=_sanitize_series(df["delivery_type"]),
        training_format=_sanitize_series(df["training_format"]),
        registered_count=_numeric_series(df["registered_count"]),
        actual_count=_numeric_series(df["actual_count"]),
        attendance_rate=_numeric_series(df["attendance_rate"]),
        estimated_cost=_numeric_series(df["estimated_cost"]),
        avg_cost_per_pax=_numeric_series(df["avg_cost_per_pax"]),
        cost_per_pax=_numeric_series(df["cost_per_pax"]),
        total_hours=_numeric_series(df["total_hours"]),
        session_group_key=_first_nonblank(
            [
                final_session_id,
                session_code,
                course_id,
                course_name,
            ],
            "",
        ),
    )


def _normalize_raw_participants(frame: pd.DataFrame) -> pd.DataFrame:
    df = _prepare_frame(
        frame,
        [
            "row_status",
            "session_id",
            "session_code",
            "course_id",
            "course_name",
            "training_date",
            "training_month",
            "emp_id",
            "full_name",
            "email",
            "attendance_status",
        ],
    )
    training_date_iso = _iso_date_series(df["training_date"])
    training_month_key = _derive_training_month_series(df["training_month"], fallback=training_date_iso)
    attendance_key = _normalize_key_series(df["attendance_status"])
    
    session_code = _sanitize_series(df["session_code"])
    raw_session_id = _sanitize_series(df["session_id"])
    course_id = _sanitize_series(df["course_id"])
    course_name = _sanitize_series(df["course_name"])
    
    generated_session_id = pd.Series([
        _generate_gas_session_hash(sc, cid, cname, tdate) 
        for sc, cid, cname, tdate in zip(session_code, course_id, course_name, training_date_iso)
    ], index=df.index)
    
    final_session_id = _first_nonblank([raw_session_id, generated_session_id], "")

    return df.assign(
        row_status_key=_normalize_key_series(df["row_status"]),
        session_id=final_session_id,
        session_code=session_code,
        course_id=course_id,
        course_name=course_name,
        training_date_iso=training_date_iso,
        training_month_key=training_month_key,
        year_key=training_month_key.str.slice(0, 4),
        emp_id=_sanitize_series(df["emp_id"]),
        full_name=_sanitize_series(df["full_name"]),
        email=_sanitize_series(df["email"]).mask(_sanitize_series(df["email"]).eq("0"), "").mask(_sanitize_series(df["email"]).eq(""), "no-email@gonsa.com.vn"),
        attendance_status=_sanitize_series(df["attendance_status"]),
        attendance_is_present=attendance_key.eq(_normalize_key("Có mặt")).astype(int),
        attendance_is_excused_absent=attendance_key.eq(_normalize_key("Vắng có phép")).astype(int),
        attendance_is_unexcused_absent=attendance_key.eq(_normalize_key("Vắng không phép")).astype(int),
        attendance_is_late=attendance_key.eq(_normalize_key("Đi muộn")).astype(int),
        participant_identity_key=_first_nonblank(
            [
                _normalize_key_series(df["emp_id"]),
                _normalize_key_series(df["email"]),
                _normalize_key_series(df["full_name"]),
            ],
            "",
        ),
    )


def _normalize_queue_jobs(frame: pd.DataFrame) -> pd.DataFrame:
    df = _prepare_frame(frame, ["job_type", "status", "updated_at", "payload_json", "checkpoint_json", "__rowNumber"])
    return df.assign(
        job_type=_sanitize_series(df["job_type"]),
        status=_sanitize_series(df["status"]),
        updated_at=_sanitize_series(df["updated_at"]),
        payload_json=df["payload_json"].fillna("").map(_parse_json_scalar),
        checkpoint_json=df["checkpoint_json"].fillna("").map(_parse_json_scalar),
        row_number=_numeric_series(df["__rowNumber"]).fillna(0),
    )


def _normalize_qa_results(frame: pd.DataFrame) -> pd.DataFrame:
    df = _prepare_frame(frame, ["severity", "status"])
    return df.assign(
        severity_key=_normalize_key_series(df["severity"]),
        status_key=_normalize_key_series(df["status"]),
    )


def _normalize_staging(frame: pd.DataFrame) -> pd.DataFrame:
    df = _prepare_frame(frame, ["row_status"])
    return df.assign(row_status_key=_normalize_key_series(df["row_status"]))


def build_dashboard_exec_matrix(inputs: NormalizedAnalyticsInputs) -> list[list[Any]]:
    records = inputs.active_training_records
    total_records = int(len(records))
    unique_employees = _unique_nonblank_count(records["emp_id"])
    unique_courses = _unique_nonblank_count(records["course_group_key"])
    unique_sessions = _unique_nonblank_count(records["session_key"])
    attended_count = int(records["attendance_is_present"].sum())
    scored_records = records.loc[records["score"].notna()]
    passed_count = int(scored_records["score"].ge(PASS_SCORE).sum())
    total_hours = _sum_or_zero(records["duration_hours"])
    total_cost = _sum_or_zero(records["cost_per_pax"])

    rows = [
        ["total_records", "Tổng bản ghi đào tạo", total_records, "bản ghi", "", inputs.last_refreshed],
        ["unique_employees", "Nhân viên đã đào tạo", unique_employees, "nhân viên", "", inputs.last_refreshed],
        ["unique_courses", "Khóa học đã sử dụng", unique_courses, "khóa học", "", inputs.last_refreshed],
        ["unique_sessions", "Lớp đào tạo đã tổ chức", unique_sessions, "lớp", "", inputs.last_refreshed],
        [
            "attendance_rate",
            "Tỷ lệ tham gia",
            _ratio_value(attended_count, total_records),
            "%",
            ">= 85%",
            inputs.last_refreshed,
        ],
        [
            "pass_rate",
            "Tỷ lệ đạt",
            _ratio_value(passed_count, len(scored_records)),
            "%",
            ">= 80%",
            inputs.last_refreshed,
        ],
        [
            "avg_satisfaction",
            "Hài lòng trung bình",
            _mean_or_blank(records["satisfaction"]),
            "/5",
            ">= 4.0",
            inputs.last_refreshed,
        ],
        [
            "avg_nps",
            "NPS trung bình",
            _mean_or_blank(records["nps"]),
            "/10",
            ">= 7.0",
            inputs.last_refreshed,
        ],
        ["total_hours", "Tổng giờ đào tạo", total_hours, "giờ", "", inputs.last_refreshed],
        ["total_cost", "Tổng chi phí đào tạo", total_cost, "VND", "", inputs.last_refreshed],
    ]
    return [DASHBOARD_EXEC_HEADERS] + _pythonize_rows(rows)


def build_dashboard_operations_matrix(inputs: NormalizedAnalyticsInputs) -> list[list[Any]]:
    records = inputs.active_training_records
    staging_rows = int(
        (~inputs.staging_hr["row_status_key"].eq(PUBLISHED_STATUS)).sum()
        + (~inputs.staging_courses["row_status_key"].eq(PUBLISHED_STATUS)).sum()
        + (~inputs.staging_training["row_status_key"].eq(PUBLISHED_STATUS)).sum()
    )
    qa_fail_rows = int(
        inputs.qa_results["severity_key"].eq(QA_FAIL_KEY).mul(inputs.qa_results["status_key"].eq(OPEN_KEY)).sum()
    )
    queue_backlog = int(inputs.queue_jobs["status"].isin([PENDING_STATUS, PROCESSING_STATUS]).sum())
    raw_sync_failed = int(float(inputs.raw_sync_summary.get("failed_rows", 0) or 0))
    applied_rate = _ratio_value(int(records["applied_is_yes"].sum()), len(records))
    avg_score = _mean_or_blank(records["score"])

    rows = [
        ["staging_rows", "Dòng đang chờ xử lý", staging_rows, "dòng", "", inputs.last_refreshed],
        ["qa_fail_rows", "Dòng có lỗi FAIL", qa_fail_rows, "dòng", "0", inputs.last_refreshed],
        ["queue_backlog", "Công việc trong queue", queue_backlog, "jobs", "0", inputs.last_refreshed],
        ["raw_sync_failed", "Data raw lỗi đồng bộ", raw_sync_failed, "dòng", "0", inputs.last_refreshed],
        ["applied_rate", "Tỷ lệ áp dụng công việc", applied_rate, "%", ">= 70%", inputs.last_refreshed],
        ["avg_score", "Điểm trung bình", avg_score, "/100", ">= 70", inputs.last_refreshed],
    ]
    rows.extend(_build_managed_sync_dashboard_rows(inputs))
    return [DASHBOARD_OPERATIONS_HEADERS] + _pythonize_rows(rows)


def build_course_matrix(inputs: NormalizedAnalyticsInputs) -> list[list[Any]]:
    course_rows = _build_course_dataframe(inputs.active_training_records)
    return [ANALYTICS_COURSE_HEADERS] + _frame_to_rows(course_rows, ANALYTICS_COURSE_HEADERS)


def _build_course_dataframe(records: pd.DataFrame) -> pd.DataFrame:
    grouped = records.groupby("course_group_key", sort=True, dropna=False)
    course_rows = grouped.agg(
        course_id=("course_id", _first_nonblank_value),
        course_name=("course_name", _first_nonblank_value),
        course_category=("course_category", _first_nonblank_value),
        training_count=("course_group_key", "size"),
        unique_employees=("emp_id", _unique_nonblank_count),
        avg_score=("score", _mean_or_blank),
        avg_satisfaction=("satisfaction", _mean_or_blank),
        avg_nps=("nps", _mean_or_blank),
        total_hours=("duration_hours", _sum_or_zero),
        total_cost=("cost_per_pax", _sum_or_zero),
        unique_sessions=("session_key", _unique_nonblank_count),
        participant_count=("course_group_key", "size"),
        attended_count=("attendance_is_present", "sum"),
        applied_count=("applied_is_yes", "sum"),
    ).reset_index(drop=True)
    if course_rows.empty:
        return pd.DataFrame(columns=ANALYTICS_COURSE_HEADERS)
    course_rows["applied_rate"] = _ratio_series(course_rows["applied_count"], course_rows["training_count"])
    return course_rows[ANALYTICS_COURSE_HEADERS].copy()


def _build_managed_sync_dashboard_rows(inputs: NormalizedAnalyticsInputs) -> list[list[Any]]:
    latest_job = _get_latest_managed_sync_job(inputs.queue_jobs)
    snapshot = _build_managed_sync_snapshot(latest_job)
    telemetry_rows = _build_managed_sync_telemetry_rows(latest_job, snapshot["checkpoint"], inputs.last_refreshed)
    rows = [
        [
            "sync_runtime_status",
            "Trạng thái đồng bộ",
            snapshot["message"],
            "",
            snapshot["state_status"],
            inputs.last_refreshed,
        ],
        [
            "sync_runtime_progress",
            "Tiến độ đồng bộ",
            snapshot["progress_value"],
            snapshot["progress_unit"],
            snapshot["progress_status"],
            inputs.last_refreshed,
        ],
    ]
    rows.extend(telemetry_rows)
    return rows


def _get_latest_managed_sync_job(queue_jobs: pd.DataFrame) -> Mapping[str, Any] | None:
    if queue_jobs.empty:
        return None
    candidates = queue_jobs.loc[
        queue_jobs["job_type"].isin(MANAGED_JOB_TYPES)
        & queue_jobs["status"].isin([PENDING_STATUS, PROCESSING_STATUS])
    ].copy()
    if candidates.empty:
        return None
    candidates["processing_rank"] = np.where(candidates["status"].eq(PROCESSING_STATUS), 1, 0)
    candidates["updated_at_ts"] = pd.to_datetime(candidates["updated_at"], errors="coerce").fillna(pd.Timestamp(0))
    candidates = candidates.sort_values(
        ["processing_rank", "updated_at_ts", "row_number"],
        ascending=[False, False, False],
    )
    return candidates.iloc[0].to_dict()


def _build_managed_sync_snapshot(job: Mapping[str, Any] | None) -> dict[str, Any]:
    if not job:
        return {
            "message": "Không có công việc đồng bộ nào đang chạy.",
            "progress_value": "",
            "progress_unit": "",
            "state_status": "IDLE",
            "progress_status": "IDLE",
            "checkpoint": _build_checkpoint({}),
        }

    payload = job.get("payload_json", {})
    checkpoint = _build_checkpoint(job.get("checkpoint_json") or {"mode": payload.get("mode", "quick")})
    current_status = _sanitize_scalar(job.get("status")) or PENDING_STATUS
    normalized_stage = _normalize_key(checkpoint.get("job_stage", "RAW_SYNC"))
    action_label = "đồng bộ dữ liệu đào tạo" if job.get("job_type") == "SYNC_TRAINING_DATA" else "làm mới data raw"
    total_rows = int(checkpoint.get("total_rows", 0) or 0)
    processed_rows = int(checkpoint.get("processed_rows", 0) or 0)
    processed_rows = min(processed_rows, total_rows or processed_rows)
    has_progress = total_rows > 0
    progress_value = round(processed_rows / total_rows * 100, 2) if has_progress else ""

    if normalized_stage == "analytics_rebuild":
        message = "Đang rebuild analytics..." if current_status == PROCESSING_STATUS else "Đang chờ rebuild analytics..."
        return {
            "message": message,
            "progress_value": 100 if has_progress else "",
            "progress_unit": "%" if has_progress else "",
            "state_status": current_status,
            "progress_status": current_status,
            "checkpoint": checkpoint,
        }

    message_prefix = "Đang " if current_status == PROCESSING_STATUS else "Đang chờ chạy tiếp "
    return {
        "message": f"{message_prefix}{action_label}: {processed_rows}/{total_rows} dòng",
        "progress_value": progress_value,
        "progress_unit": "%" if has_progress else "",
        "state_status": current_status,
        "progress_status": current_status,
        "checkpoint": checkpoint,
    }


def _build_managed_sync_telemetry_rows(
    job: Mapping[str, Any] | None,
    checkpoint: Mapping[str, Any],
    last_refreshed: str,
) -> list[list[Any]]:
    source_job = job or {}
    snapshot = _build_checkpoint(checkpoint)
    job_type = _sanitize_scalar(snapshot.get("sync_runtime_job_type") or source_job.get("job_type"))
    job_label = "đồng bộ dữ liệu đào tạo" if job_type == "SYNC_TRAINING_DATA" else (
        "làm mới data raw" if job_type == "REFRESH_ANALYTICS" else ""
    )
    return [
        ["sync_runtime_job_type", "Loai job dong bo", job_label, "", _sanitize_scalar(source_job.get("status")), last_refreshed],
        ["sync_runtime_started_at", "Bat dau luc", snapshot["sync_runtime_started_at"], "", "", last_refreshed],
        ["sync_runtime_last_block_at", "Block gan nhat luc", snapshot["sync_runtime_last_block_at"], "", "", last_refreshed],
        ["sync_runtime_block_duration_seconds", "Thoi gian block gan nhat", snapshot["sync_runtime_last_block_duration_seconds"], "giay", "", last_refreshed],
        ["sync_runtime_raw_total_seconds", "Raw sync tong", snapshot["sync_runtime_raw_total_seconds"], "giay", "", last_refreshed],
        ["sync_runtime_analytics_seconds", "Analytics rebuild", snapshot["sync_runtime_analytics_seconds"], "giay", "", last_refreshed],
        ["sync_runtime_total_seconds", "Tong thoi gian", snapshot["sync_runtime_total_seconds"], "giay", "", last_refreshed],
    ]


def _build_checkpoint(source: Mapping[str, Any]) -> dict[str, Any]:
    settings = dict(source or {})
    return {
        "mode": "full" if _normalize_key(settings.get("mode")) == "full" else "quick",
        "job_stage": _sanitize_scalar(settings.get("job_stage")) or "RAW_SYNC",
        "last_processed_row": int(float(settings.get("last_processed_row", 0) or 0)),
        "total_rows": int(float(settings.get("totalRows", settings.get("total_rows", 0)) or 0)),
        "processed_rows": int(float(settings.get("processed_rows", 0) or 0)),
        "passed_rows": int(float(settings.get("passed_rows", 0) or 0)),
        "warning_rows": int(float(settings.get("warning_rows", 0) or 0)),
        "failed_rows": int(float(settings.get("failed_rows", 0) or 0)),
        "updated_records": int(float(settings.get("updated_records", 0) or 0)),
        "inserted_records": int(float(settings.get("inserted_records", 0) or 0)),
        "removed_records": int(float(settings.get("removed_records", 0) or 0)),
        "chunk_size": int(float(settings.get("chunkSize", settings.get("chunk_size", 0)) or 0)),
        "last_run_at": _sanitize_scalar(settings.get("last_run_at")),
        "raw_sync_summary": settings.get("raw_sync_summary"),
        "sync_runtime_job_type": _sanitize_scalar(settings.get("sync_runtime_job_type")),
        "sync_runtime_current_stage": _sanitize_scalar(settings.get("sync_runtime_current_stage")),
        "sync_runtime_started_at": _sanitize_scalar(settings.get("sync_runtime_started_at")),
        "sync_runtime_completed_at": _sanitize_scalar(settings.get("sync_runtime_completed_at")),
        "sync_runtime_last_block_started_at": _sanitize_scalar(settings.get("sync_runtime_last_block_started_at")),
        "sync_runtime_last_block_at": _sanitize_scalar(settings.get("sync_runtime_last_block_at")),
        "sync_runtime_last_block_duration_seconds": float(settings.get("sync_runtime_last_block_duration_seconds", 0) or 0),
        "sync_runtime_raw_total_seconds": float(settings.get("sync_runtime_raw_total_seconds", 0) or 0),
        "sync_runtime_analytics_started_at": _sanitize_scalar(settings.get("sync_runtime_analytics_started_at")),
        "sync_runtime_analytics_completed_at": _sanitize_scalar(settings.get("sync_runtime_analytics_completed_at")),
        "sync_runtime_analytics_seconds": float(settings.get("sync_runtime_analytics_seconds", 0) or 0),
        "sync_runtime_total_seconds": float(settings.get("sync_runtime_total_seconds", 0) or 0),
    }


def _build_standard_report_canvas(
    *,
    title: str,
    meta_pairs: list[tuple[str, Any]],
    kpi_items: list[tuple[str, Any]],
    sections: list[dict[str, Any]],
    width: int,
) -> list[list[Any]]:
    canvas: list[list[Any]] = []
    canvas.extend(_pad_rows([[title]], width))
    canvas.extend(_pad_rows([[value for pair in meta_pairs for value in pair]], width))
    canvas.append(_blank_row(width))
    if kpi_items:
        canvas.extend(_pad_rows([["KPI"]], width))
        columns_per_row = 3
        for start in range(0, len(kpi_items), columns_per_row):
            chunk = kpi_items[start:start + columns_per_row]
            label_row: list[Any] = []
            value_row: list[Any] = []
            for label, value in chunk:
                label_row.extend([label, "", "", ""])
                value_row.extend([value, "", "", ""])
            canvas.extend(_pad_rows([label_row, value_row], width))
    canvas.append(_blank_row(width))
    for index, section in enumerate(sections):
        canvas.extend(_pad_rows([[section["title"]]], width))
        canvas.extend(_pad_rows([section["header"]], width))
        rows = section.get("rows") or [["Không có dữ liệu phù hợp"]]
        canvas.extend(_pad_rows(rows, width))
        if index != len(sections) - 1:
            canvas.append(_blank_row(width))
    return _pythonize_rows(canvas)


def build_department_canvas(inputs: NormalizedAnalyticsInputs) -> list[list[Any]]:
    report_data = _build_department_report_data(inputs)
    sections = [
        {
            "title": "TOP 5 PHÒNG BAN PHỦ ĐÀO TẠO CAO NHẤT",
            "header": ["Phòng ban", "Tỷ lệ phủ đào tạo", "Nhân sự đã đào tạo", "Tổng nhân sự"],
            "rows": _frame_to_rows(
                report_data["rankings"]["top_coverage"],
                ["department", "coverage_rate", "trained_employees", "total_employees"],
            ),
        },
        {
            "title": "TOP 5 PHÒNG BAN TỔNG GIỜ ĐÀO TẠO CAO NHẤT",
            "header": ["Phòng ban", "Tổng giờ đào tạo", "Số lượt đào tạo", "Số nhân viên"],
            "rows": _frame_to_rows(
                report_data["rankings"]["top_hours"],
                ["department", "total_hours", "training_count", "unique_employees"],
            ),
        },
        {
            "title": "TOP 5 PHÒNG BAN TỔNG CHI PHÍ CAO NHẤT",
            "header": ["Phòng ban", "Tổng chi phí", "Tỷ lệ phủ đào tạo", "Nhân sự đã đào tạo"],
            "rows": _frame_to_rows(
                report_data["rankings"]["top_cost"],
                ["department", "total_cost", "coverage_rate", "trained_employees"],
            ),
        },
        {
            "title": "TOP 5 PHÒNG BAN CẦN CHÚ Ý",
            "header": ["Phòng ban", "Tỷ lệ phủ đào tạo", "Nhân sự đã đào tạo", "Tổng nhân sự"],
            "rows": _frame_to_rows(
                report_data["rankings"]["watchlist"],
                ["department", "coverage_rate", "trained_employees", "total_employees"],
            ),
        },
        {
            "title": "CHI TIẾT PHÒNG BAN",
            "header": [
                "Phòng ban",
                "Số lượt đào tạo",
                "Số nhân viên",
                "Số khóa học",
                "Điểm trung bình",
                "Hài lòng trung bình",
                "Tỷ lệ áp dụng",
                "Tổng giờ đào tạo",
                "Tổng chi phí",
                "Tổng nhân sự",
                "Nhân sự đã đào tạo",
                "Tỷ lệ phủ đào tạo",
                "Số lớp",
                "Giờ / nhân viên",
                "Chi phí / nhân viên",
                "Xếp hạng phủ đào tạo",
                "Xếp hạng chi phí",
            ],
            "rows": _frame_to_rows(
                report_data["detail_rows"],
                [
                    "department",
                    "training_count",
                    "unique_employees",
                    "unique_courses",
                    "avg_score",
                    "avg_satisfaction",
                    "applied_rate",
                    "total_hours",
                    "total_cost",
                    "total_employees",
                    "trained_employees",
                    "coverage_rate",
                    "unique_sessions",
                    "hours_per_employee",
                    "cost_per_employee",
                    "coverage_rank",
                    "cost_rank",
                ],
            ),
        },
    ]
    return _build_standard_report_canvas(
        title="BÁO CÁO PHÂN TÍCH PHÒNG BAN",
        meta_pairs=[
            ("Năm báo cáo", report_data["fiscal_year"]),
            ("Nguồn dữ liệu", "Dữ liệu đào tạo"),
            ("Làm mới lúc", _format_timestamp_display(report_data["refreshed_at"])),
        ],
        kpi_items=[
            ("Số phòng ban có đào tạo", report_data["kpis"]["department_count"]),
            ("Tổng nhân sự", report_data["kpis"]["total_employees"]),
            ("Nhân sự đã đào tạo", report_data["kpis"]["trained_employees"]),
            ("Tỷ lệ phủ đào tạo toàn công ty", report_data["kpis"]["company_coverage_rate"]),
            ("Tổng giờ đào tạo", report_data["kpis"]["total_hours"]),
            ("Tổng chi phí", report_data["kpis"]["total_cost"]),
        ],
        sections=sections,
        width=17,
    )


def build_department_course_canvas(inputs: NormalizedAnalyticsInputs) -> list[list[Any]]:
    report_data = _build_department_course_report_data(inputs)
    sections = [
        {
            "title": "TOP 10 CẶP PHÒNG BAN - KHÓA THEO LƯỢT HỌC",
            "header": ["Phòng ban", "Mã khóa học", "Tên khóa học", "Số lượt học viên"],
            "rows": _frame_to_rows(
                report_data["rankings"]["top_participants"],
                ["department", "course_id", "course_name", "participant_count"],
            ),
        },
        {
            "title": "TOP 10 CẶP PHÒNG BAN - KHÓA THEO TỔNG CHI PHÍ",
            "header": ["Phòng ban", "Mã khóa học", "Tên khóa học", "Tổng chi phí"],
            "rows": _frame_to_rows(
                report_data["rankings"]["top_cost"],
                ["department", "course_id", "course_name", "total_cost"],
            ),
        },
        {
            "title": "TOP 10 CẶP PHÒNG BAN - KHÓA THEO TỶ LỆ PHỦ",
            "header": ["Phòng ban", "Mã khóa học", "Tên khóa học", "Tỷ lệ phủ đào tạo"],
            "rows": _frame_to_rows(
                report_data["rankings"]["top_coverage"],
                ["department", "course_id", "course_name", "coverage_rate"],
            ),
        },
        {
            "title": "CHI TIẾT PHÒNG BAN THEO KHÓA",
            "header": [
                "Phòng ban",
                "Tổng nhân sự",
                "Mã khóa học",
                "Tên khóa học",
                "Số lớp",
                "Số lượt học viên",
                "Số có mặt",
                "Số nhân viên đã đào tạo",
                "Tỷ lệ phủ đào tạo",
                "Điểm trung bình",
                "Hài lòng trung bình",
                "Tổng giờ đào tạo",
                "Tổng chi phí",
                "Tỷ lệ tham gia",
                "Giờ / nhân viên",
                "Chi phí / nhân viên",
                "Xếp hạng lượt học",
            ],
            "rows": _frame_to_rows(
                report_data["detail_rows"],
                [
                    "department",
                    "total_employees",
                    "course_id",
                    "course_name",
                    "session_count",
                    "participant_count",
                    "attended_count",
                    "trained_employees",
                    "coverage_rate",
                    "avg_score",
                    "avg_satisfaction",
                    "total_hours",
                    "total_cost",
                    "participation_rate",
                    "hours_per_employee",
                    "cost_per_employee",
                    "participant_rank",
                ],
            ),
        },
    ]
    return _build_standard_report_canvas(
        title="BÁO CÁO PHÒNG BAN THEO KHÓA",
        meta_pairs=[
            ("Năm báo cáo", report_data["fiscal_year"]),
            ("Làm mới lúc", _format_timestamp_display(report_data["refreshed_at"])),
        ],
        kpi_items=[
            ("Số cặp phòng ban - khóa", report_data["kpis"]["department_course_count"]),
            ("Số lớp", report_data["kpis"]["session_count"]),
            ("Số lượt học viên", report_data["kpis"]["participant_count"]),
            ("Số nhân viên đã đào tạo", report_data["kpis"]["trained_employees"]),
            ("Tổng giờ đào tạo", report_data["kpis"]["total_hours"]),
            ("Tổng chi phí", report_data["kpis"]["total_cost"]),
        ],
        sections=sections,
        width=17,
    )


def build_trend_canvas(inputs: NormalizedAnalyticsInputs) -> list[list[Any]]:
    report_data = _build_trend_report_data(inputs)
    highlight_rows = [
        [
            "Tháng có nhiều lượt học nhất",
            report_data["highlights"]["highest_participants"].get("month_key", ""),
            report_data["highlights"]["highest_participants"].get("participant_count", ""),
        ],
        [
            "Tháng có nhiều chi phí nhất",
            report_data["highlights"]["highest_cost"].get("month_key", ""),
            report_data["highlights"]["highest_cost"].get("total_cost", ""),
        ],
    ]
    sections = [
        {
            "title": "SO SÁNH NHANH THÁNG GẦN NHẤT",
            "header": [
                "Tháng gần nhất",
                "Tháng trước",
                "Chênh lượt đào tạo",
                "Chênh lượt học viên",
                "Chênh tổng giờ",
                "Chênh tổng chi phí",
            ],
            "rows": [[
                report_data["comparison"]["latest_month"],
                report_data["comparison"]["previous_month"],
                report_data["comparison"]["training_delta"],
                report_data["comparison"]["participant_delta"],
                report_data["comparison"]["total_hours_delta"],
                report_data["comparison"]["total_cost_delta"],
            ]],
        },
        {
            "title": "THÁNG NỔI BẬT",
            "header": ["Chỉ tiêu", "Kỳ tháng", "Giá trị"],
            "rows": highlight_rows,
        },
        {
            "title": "CHI TIẾT XU HƯỚNG THEO THÁNG",
            "header": [
                "Kỳ tháng",
                "Số lượt đào tạo",
                "Số nhân viên",
                "Tỷ lệ tham gia",
                "Tỷ lệ đạt",
                "Điểm trung bình",
                "Hài lòng trung bình",
                "Tổng giờ đào tạo",
                "Tổng chi phí",
                "Số lớp",
                "Số lượt học viên",
                "Giờ / học viên",
                "Chi phí / học viên",
                "Chênh lượt đào tạo",
                "Chênh lượt học viên",
                "Chênh chi phí",
                "Chênh số lớp",
            ],
            "rows": _frame_to_rows(
                report_data["detail_rows"],
                [
                    "month_key",
                    "training_count",
                    "unique_employees",
                    "attendance_rate",
                    "pass_rate",
                    "avg_score",
                    "avg_satisfaction",
                    "total_hours",
                    "total_cost",
                    "session_count",
                    "participant_count",
                    "hours_per_participant",
                    "cost_per_participant",
                    "training_delta_vs_previous",
                    "participant_delta_vs_previous",
                    "cost_delta_vs_previous",
                    "session_delta_vs_previous",
                ],
            ),
        },
    ]
    return _build_standard_report_canvas(
        title="BÁO CÁO PHÂN TÍCH XU HƯỚNG",
        meta_pairs=[
            ("Năm báo cáo", report_data["fiscal_year"]),
            ("Làm mới lúc", _format_timestamp_display(report_data["refreshed_at"])),
        ],
        kpi_items=[
            ("Số tháng có dữ liệu", report_data["kpis"]["month_count"]),
            ("Tổng lớp", report_data["kpis"]["session_count"]),
            ("Tổng lượt học viên", report_data["kpis"]["participant_count"]),
            ("Tổng nhân viên", report_data["kpis"]["unique_employees"]),
            ("Tổng giờ đào tạo", report_data["kpis"]["total_hours"]),
            ("Tổng chi phí", report_data["kpis"]["total_cost"]),
        ],
        sections=sections,
        width=17,
    )


def _build_department_report_data(inputs: NormalizedAnalyticsInputs) -> dict[str, Any]:
    records = inputs.active_training_records
    grouped = records.groupby("department_name", sort=True, dropna=False)
    rows = grouped.agg(
        department=("department_name", _first_nonblank_value),
        training_count=("department_name", "size"),
        unique_employees=("emp_id", _unique_nonblank_count),
        unique_courses=("course_group_key", _unique_nonblank_count),
        avg_score=("score", _mean_or_blank),
        avg_satisfaction=("satisfaction", _mean_or_blank),
        total_hours=("duration_hours", _sum_or_zero),
        total_cost=("cost_per_pax", _sum_or_zero),
        unique_sessions=("session_key", _unique_nonblank_count),
        applied_count=("applied_is_yes", "sum"),
    ).reset_index(drop=True)

    if rows.empty:
        detail_rows = pd.DataFrame(
            columns=[
                "department",
                "training_count",
                "unique_employees",
                "unique_courses",
                "avg_score",
                "avg_satisfaction",
                "applied_rate",
                "total_hours",
                "total_cost",
                "total_employees",
                "trained_employees",
                "coverage_rate",
                "unique_sessions",
                "hours_per_employee",
                "cost_per_employee",
                "coverage_rank",
                "cost_rank",
            ]
        )
    else:
        rows["applied_rate"] = _ratio_series(rows["applied_count"], rows["training_count"])
        rows["total_employees"] = rows["department"].map(inputs.department_headcount).fillna(0).astype(int)
        rows["trained_employees"] = rows["unique_employees"]
        rows["coverage_rate"] = _ratio_series(rows["trained_employees"], rows["total_employees"])
        rows["hours_per_employee"] = _safe_divide_series(rows["total_hours"], rows["unique_employees"])
        rows["cost_per_employee"] = _safe_divide_series(rows["total_cost"], rows["unique_employees"])
        rows = _assign_rank(rows, "coverage_rate", "department", "coverage_rank")
        rows = _assign_rank(rows, "total_cost", "department", "cost_rank")
        detail_rows = rows.sort_values("coverage_rank").reset_index(drop=True)

    total_employees = int(detail_rows["total_employees"].sum()) if not detail_rows.empty else 0
    trained_employees = int(detail_rows["trained_employees"].sum()) if not detail_rows.empty else 0
    total_hours = _sum_or_zero(detail_rows["total_hours"]) if not detail_rows.empty else 0
    total_cost = _sum_or_zero(detail_rows["total_cost"]) if not detail_rows.empty else 0

    return {
        "fiscal_year": inputs.fiscal_year,
        "refreshed_at": inputs.last_refreshed,
        "kpis": {
            "department_count": int(len(detail_rows)),
            "total_employees": total_employees,
            "trained_employees": trained_employees,
            "company_coverage_rate": _ratio_value(trained_employees, total_employees),
            "total_hours": total_hours,
            "total_cost": total_cost,
        },
        "rankings": {
            "top_coverage": _sort_desc(detail_rows, "coverage_rate", "department").head(5),
            "top_hours": _sort_desc(detail_rows, "total_hours", "department").head(5),
            "top_cost": _sort_desc(detail_rows, "total_cost", "department").head(5),
            "watchlist": detail_rows.loc[detail_rows["total_employees"].gt(0)]
            .sort_values(["coverage_rate", "department"], ascending=[True, True], na_position="last")
            .head(5)
            .reset_index(drop=True),
        },
        "detail_rows": detail_rows[
            [
                "department",
                "training_count",
                "unique_employees",
                "unique_courses",
                "avg_score",
                "avg_satisfaction",
                "applied_rate",
                "total_hours",
                "total_cost",
                "total_employees",
                "trained_employees",
                "coverage_rate",
                "unique_sessions",
                "hours_per_employee",
                "cost_per_employee",
                "coverage_rank",
                "cost_rank",
            ]
        ].copy(),
    }


def _build_department_course_report_data(inputs: NormalizedAnalyticsInputs) -> dict[str, Any]:
    records = inputs.active_training_records
    grouped = records.groupby(["department_name", "course_group_key"], sort=True, dropna=False)
    rows = grouped.agg(
        department=("department_name", _first_nonblank_value),
        course_id=("course_id", _first_nonblank_value),
        course_name=("course_name", _first_nonblank_value),
        session_count=("session_key", _unique_nonblank_count),
        participant_count=("course_group_key", "size"),
        attended_count=("attendance_is_present", "sum"),
        trained_employees=("emp_id", _unique_nonblank_count),
        avg_score=("score", _mean_or_blank),
        avg_satisfaction=("satisfaction", _mean_or_blank),
        total_hours=("duration_hours", _sum_or_zero),
        total_cost=("cost_per_pax", _sum_or_zero),
    ).reset_index(drop=True)

    if rows.empty:
        detail_rows = pd.DataFrame(
            columns=[
                "department",
                "total_employees",
                "course_id",
                "course_name",
                "session_count",
                "participant_count",
                "attended_count",
                "trained_employees",
                "coverage_rate",
                "avg_score",
                "avg_satisfaction",
                "total_hours",
                "total_cost",
                "participation_rate",
                "hours_per_employee",
                "cost_per_employee",
                "participant_rank",
            ]
        )
    else:
        rows["total_employees"] = rows["department"].map(inputs.department_headcount).fillna(0).astype(int)
        rows["coverage_rate"] = _ratio_series(rows["trained_employees"], rows["total_employees"])
        rows["participation_rate"] = _ratio_series(rows["attended_count"], rows["participant_count"])
        rows["hours_per_employee"] = _safe_divide_series(rows["total_hours"], rows["trained_employees"])
        rows["cost_per_employee"] = _safe_divide_series(rows["total_cost"], rows["trained_employees"])
        rows = _assign_rank(rows, "participant_count", "course_name", "participant_rank")
        detail_rows = rows.sort_values("participant_rank").reset_index(drop=True)

    return {
        "fiscal_year": inputs.fiscal_year,
        "refreshed_at": inputs.last_refreshed,
        "kpis": {
            "department_course_count": int(len(detail_rows)),
            "session_count": int(detail_rows["session_count"].sum()) if not detail_rows.empty else 0,
            "participant_count": int(detail_rows["participant_count"].sum()) if not detail_rows.empty else 0,
            "trained_employees": int(detail_rows["trained_employees"].sum()) if not detail_rows.empty else 0,
            "total_hours": _sum_or_zero(detail_rows["total_hours"]) if not detail_rows.empty else 0,
            "total_cost": _sum_or_zero(detail_rows["total_cost"]) if not detail_rows.empty else 0,
        },
        "rankings": {
            "top_participants": _sort_desc(detail_rows, "participant_count", "course_name").head(10),
            "top_cost": _sort_desc(detail_rows, "total_cost", "course_name").head(10),
            "top_coverage": _sort_desc(detail_rows, "coverage_rate", "course_name").head(10),
        },
        "detail_rows": detail_rows[
            [
                "department",
                "total_employees",
                "course_id",
                "course_name",
                "session_count",
                "participant_count",
                "attended_count",
                "trained_employees",
                "coverage_rate",
                "avg_score",
                "avg_satisfaction",
                "total_hours",
                "total_cost",
                "participation_rate",
                "hours_per_employee",
                "cost_per_employee",
                "participant_rank",
            ]
        ].copy(),
    }


def _build_trend_report_data(inputs: NormalizedAnalyticsInputs) -> dict[str, Any]:
    records = inputs.active_training_records
    grouped = records.groupby("training_month_key", sort=True, dropna=False)
    rows = grouped.agg(
        month_key=("training_month_key", _first_nonblank_value),
        training_count=("training_month_key", "size"),
        unique_employees=("emp_id", _unique_nonblank_count),
        attended_count=("attendance_is_present", "sum"),
        scored_count=("score", lambda series: int(series.notna().sum())),
        passed_count=("score", lambda series: int(series.ge(PASS_SCORE).sum())),
        avg_score=("score", _mean_or_blank),
        avg_satisfaction=("satisfaction", _mean_or_blank),
        total_hours=("duration_hours", _sum_or_zero),
        total_cost=("cost_per_pax", _sum_or_zero),
        session_count=("session_key", _unique_nonblank_count),
        participant_count=("training_month_key", "size"),
    ).reset_index(drop=True)

    if rows.empty:
        detail_rows = pd.DataFrame(
            columns=[
                "month_key",
                "training_count",
                "unique_employees",
                "attendance_rate",
                "pass_rate",
                "avg_score",
                "avg_satisfaction",
                "total_hours",
                "total_cost",
                "session_count",
                "participant_count",
                "hours_per_participant",
                "cost_per_participant",
                "training_delta_vs_previous",
                "participant_delta_vs_previous",
                "cost_delta_vs_previous",
                "session_delta_vs_previous",
            ]
        )
    else:
        rows["attendance_rate"] = _ratio_series(rows["attended_count"], rows["training_count"])
        rows["pass_rate"] = _ratio_series(rows["passed_count"], rows["scored_count"])
        rows["hours_per_participant"] = _safe_divide_series(rows["total_hours"], rows["participant_count"])
        rows["cost_per_participant"] = _safe_divide_series(rows["total_cost"], rows["participant_count"])
        rows = rows.sort_values("month_key").reset_index(drop=True)
        rows["training_delta_vs_previous"] = rows["training_count"].diff().round(2)
        rows["participant_delta_vs_previous"] = rows["participant_count"].diff().round(2)
        rows["cost_delta_vs_previous"] = rows["total_cost"].diff().round(2)
        rows["session_delta_vs_previous"] = rows["session_count"].diff().round(2)
        rows.loc[
            rows.index == 0,
            [
                "training_delta_vs_previous",
                "participant_delta_vs_previous",
                "cost_delta_vs_previous",
                "session_delta_vs_previous",
            ],
        ] = np.nan
        detail_rows = rows[
            [
                "month_key",
                "training_count",
                "unique_employees",
                "attendance_rate",
                "pass_rate",
                "avg_score",
                "avg_satisfaction",
                "total_hours",
                "total_cost",
                "session_count",
                "participant_count",
                "hours_per_participant",
                "cost_per_participant",
                "training_delta_vs_previous",
                "participant_delta_vs_previous",
                "cost_delta_vs_previous",
                "session_delta_vs_previous",
            ]
        ].copy()

    latest = detail_rows.iloc[-1] if not detail_rows.empty else None
    previous = detail_rows.iloc[-2] if len(detail_rows) > 1 else None
    highlights_by_participants = _sort_desc(detail_rows, "participant_count", "month_key")
    highlights_by_cost = _sort_desc(detail_rows, "total_cost", "month_key")

    return {
        "fiscal_year": inputs.fiscal_year,
        "refreshed_at": inputs.last_refreshed,
        "kpis": {
            "month_count": int(len(detail_rows)),
            "session_count": int(detail_rows["session_count"].sum()) if not detail_rows.empty else 0,
            "participant_count": int(detail_rows["participant_count"].sum()) if not detail_rows.empty else 0,
            "unique_employees": int(detail_rows["unique_employees"].sum()) if not detail_rows.empty else 0,
            "total_hours": _sum_or_zero(detail_rows["total_hours"]) if not detail_rows.empty else 0,
            "total_cost": _sum_or_zero(detail_rows["total_cost"]) if not detail_rows.empty else 0,
        },
        "comparison": {
            "latest_month": _to_python_cell(latest["month_key"]) if latest is not None else "",
            "previous_month": _to_python_cell(previous["month_key"]) if previous is not None else "",
            "training_delta": _to_python_cell(latest["training_delta_vs_previous"]) if latest is not None else "",
            "participant_delta": _to_python_cell(latest["participant_delta_vs_previous"]) if latest is not None else "",
            "total_hours_delta": _to_python_cell(latest["total_hours"] - previous["total_hours"])
            if latest is not None and previous is not None
            else "",
            "total_cost_delta": _to_python_cell(latest["cost_delta_vs_previous"]) if latest is not None else "",
        },
        "highlights": {
            "highest_participants": (
                highlights_by_participants.iloc[0].to_dict() if not highlights_by_participants.empty else {}
            ),
            "highest_cost": highlights_by_cost.iloc[0].to_dict() if not highlights_by_cost.empty else {},
        },
        "detail_rows": detail_rows,
    }


def build_external_assignment_canvas(inputs: NormalizedAnalyticsInputs) -> list[list[Any]]:
    report_data = _build_external_assignment_report_data(inputs)
    width = 22
    canvas: list[list[Any]] = []
    canvas.extend(_pad_rows([["BÁO CÁO CỬ ĐI HỌC BÊN NGOÀI"]], width))
    canvas.extend(_pad_rows([[
        "Năm báo cáo",
        report_data["fiscal_year"],
        "Điều kiện lọc",
        report_data["filters"]["delivery_type"] + " + " + report_data["filters"]["training_format"],
        "Làm mới lúc",
        _format_timestamp_display(report_data["refreshed_at"]),
    ]], width))
    canvas.append(_blank_row(width))
    canvas.extend(_pad_rows([["KPI"]], width))
    canvas.extend(_pad_rows([[
        "Số lớp bên ngoài", "", "", "",
        "Số lượt học viên", "", "", "",
        "Số nhân viên unique", "", "", "",
        "Tổng chi phí ước tính", "", "", "",
    ]], width))
    canvas.extend(_pad_rows([[
        report_data["kpis"]["session_count"], "", "", "",
        report_data["kpis"]["participant_count"], "", "", "",
        report_data["kpis"]["unique_employee_count"], "", "", "",
        report_data["kpis"]["total_cost"], "", "", "",
    ]], width))
    canvas.append(_blank_row(width))
    canvas.extend(_pad_rows([["TỔNG HỢP THEO LỚP"]], width))
    canvas.extend(_pad_rows([[
        "Kỳ tháng", "Ngày đào tạo", "Mã lớp", "Mã khóa học", "Tên khóa học",
        "Đơn vị đào tạo", "Địa điểm", "Số học viên", "Số nhân viên unique",
        "Phòng ban tham gia", "Chi phí lớp", "Chi phí bình quân/người", "Cơ sở tính chi phí",
    ]], width))
    session_rows = _frame_to_rows(
        report_data["session_rows"],
        [
            "month_key",
            "training_date",
            "session_code",
            "course_id",
            "course_name",
            "training_unit",
            "location",
            "participant_count",
            "unique_employee_count",
            "departments_participating",
            "session_total_cost",
            "avg_cost_per_pax",
            "cost_basis_label",
        ],
    )
    canvas.extend(_pad_rows(session_rows or [["Không có dữ liệu phù hợp"]], width))
    canvas.append(_blank_row(width))
    canvas.extend(_pad_rows([["TỔNG HỢP THEO PHÒNG BAN"]], width))
    canvas.extend(_pad_rows([[
        "Phòng ban",
        "Số lượt học viên",
        "Số nhân viên unique",
        "Số lớp tham gia",
        "Tổng chi phí phân bổ",
        "Chi phí bình quân/người",
    ]], width))
    department_rows = _frame_to_rows(
        report_data["department_rows"],
        [
            "department",
            "participant_count",
            "unique_employee_count",
            "session_count",
            "total_allocated_cost",
            "avg_cost_per_pax",
        ],
    )
    canvas.extend(_pad_rows(department_rows or [["Không có dữ liệu phù hợp"]], width))
    canvas.append(_blank_row(width))
    canvas.extend(_pad_rows([["CHI TIẾT TỪNG HỌC VIÊN"]], width))
    canvas.extend(_pad_rows([[
        "Năm", "Kỳ tháng", "Ngày đào tạo", "Mã lớp", "Mã khóa học", "Tên khóa học",
        "Đơn vị đào tạo", "Địa điểm", "Mã NV", "Họ tên", "Phòng ban", "Khối/Division",
        "Chức danh", "Trạng thái điểm danh", "Trong/ngoài kế hoạch", "Loại chương trình",
        "Loại hình đào tạo", "Hình thức đào tạo", "Chi phí/người", "Tổng chi phí lớp",
        "Cơ sở tính chi phí",
    ]], width))
    detail_rows = _frame_to_rows(
        report_data["detail_rows"],
        [
            "year",
            "month_key",
            "training_date",
            "session_code",
            "course_id",
            "course_name",
            "training_unit",
            "location",
            "emp_id",
            "full_name",
            "department",
            "division",
            "job_title",
            "attendance_status",
            "plan_scope",
            "program_type",
            "delivery_type",
            "training_format",
            "participant_cost",
            "session_total_cost",
            "cost_basis_label",
        ],
    )
    canvas.extend(_pad_rows(detail_rows or [["Không có dữ liệu phù hợp"]], width))
    return canvas


def build_session_reconciliation_canvas(inputs: NormalizedAnalyticsInputs) -> list[list[Any]]:
    report_data = _build_session_reconciliation_report_data(inputs)
    sections = [
        {
            "title": "TOP LỚP LỆCH NHIỀU NHẤT THEO ĐĂNG KÝ",
            "header": ["Mã lớp", "Tên khóa học", "Đăng ký nhập tay", "Raw unique", "Chênh lệch"],
            "rows": _frame_to_rows(
                report_data["rankings"]["top_registered_delta"],
                ["session_code", "course_name", "registered_count", "raw_unique_count", "registered_vs_raw_delta"],
            ),
        },
        {
            "title": "TOP LỚP LỆCH NHIỀU NHẤT THEO THỰC TẾ",
            "header": ["Mã lớp", "Tên khóa học", "Thực tế nhập tay", "Raw có mặt", "Chênh lệch"],
            "rows": _frame_to_rows(
                report_data["rankings"]["top_actual_delta"],
                ["session_code", "course_name", "actual_count", "raw_present_count", "actual_vs_raw_present_delta"],
            ),
        },
        {
            "title": "CHI TIẾT ĐỐI CHIẾU LỚP ĐÀO TẠO",
            "header": [
                "Kỳ tháng",
                "Ngày đào tạo",
                "Mã lớp đào tạo",
                "Mã lớp",
                "Mã khóa học",
                "Tên khóa học",
                "Đơn vị đào tạo",
                "Loại hình đào tạo",
                "Hình thức đào tạo",
                "Đăng ký nhập tay",
                "Thực tế nhập tay",
                "Tỷ lệ tham gia nhập tay",
                "Raw: số dòng",
                "Raw: số học viên unique",
                "Raw: có mặt",
                "Raw: vắng có phép",
                "Raw: vắng không phép",
                "Raw: đi muộn",
                "Fact: số bản ghi",
                "Fact: số nhân viên unique",
                "Fact: có mặt",
                "Chênh đăng ký vs raw unique",
                "Chênh thực tế vs raw có mặt",
                "Chênh raw unique vs fact unique",
                "Trạng thái đối chiếu",
                "Gợi ý xử lý",
                "Chi phí dự kiến lớp",
                "Tổng giờ đào tạo",
            ],
            "rows": _frame_to_rows(
                report_data["detail_rows"],
                [
                    "month_key",
                    "training_date",
                    "session_id",
                    "session_code",
                    "course_id",
                    "course_name",
                    "training_unit",
                    "delivery_type",
                    "training_format",
                    "registered_count",
                    "actual_count",
                    "manual_attendance_rate",
                    "raw_row_count",
                    "raw_unique_count",
                    "raw_present_count",
                    "raw_excused_absent_count",
                    "raw_unexcused_absent_count",
                    "raw_late_count",
                    "fact_record_count",
                    "fact_unique_count",
                    "fact_present_count",
                    "registered_vs_raw_delta",
                    "actual_vs_raw_present_delta",
                    "raw_vs_fact_unique_delta",
                    "reconciliation_status",
                    "suggested_action",
                    "estimated_cost",
                    "total_hours",
                ],
            ),
        },
    ]
    return _build_standard_report_canvas(
        title="BÁO CÁO ĐỐI CHIẾU LỚP ĐÀO TẠO",
        meta_pairs=[
            ("Năm báo cáo", report_data["fiscal_year"]),
            ("Làm mới lúc", _format_timestamp_display(report_data["refreshed_at"])),
        ],
        kpi_items=[
            ("Tổng số lớp", report_data["kpis"]["total_sessions"]),
            ("Lớp đã có raw", report_data["kpis"]["sessions_with_raw"]),
            ("Lớp chưa có raw", report_data["kpis"]["sessions_without_raw"]),
            ("Lớp khớp", report_data["kpis"]["matched_sessions"]),
            ("Lớp lệch số liệu", report_data["kpis"]["mismatched_sessions"]),
        ],
        sections=sections,
        width=28,
    )


def _build_external_assignment_report_data(inputs: NormalizedAnalyticsInputs) -> dict[str, Any]:
    records = inputs.active_training_records.copy()
    filtered = records.loc[
        records["delivery_type_key"].eq(_normalize_key("Đào tạo bên ngoài"))
        & records["training_format_key"].eq(_normalize_key("Cử nhân sự đi học"))
        & records["year_key"].eq(inputs.fiscal_year)
    ].copy()
    filtered = filtered.sort_values(["training_date_iso", "session_code", "full_name"]).reset_index(drop=True)

    session_lookup = _build_external_assignment_session_lookup(inputs.active_training_sessions)
    filtered = _attach_external_assignment_costs(filtered, session_lookup, inputs.fiscal_year)

    if filtered.empty:
        empty = pd.DataFrame()
        return {
            "fiscal_year": inputs.fiscal_year,
            "refreshed_at": inputs.last_refreshed,
            "filters": {
                "delivery_type": "Đào tạo bên ngoài",
                "training_format": "Cử nhân sự đi học",
                "participant_scope": "ALL_RECORDED",
            },
            "kpis": {"session_count": 0, "participant_count": 0, "unique_employee_count": 0, "total_cost": ""},
            "session_rows": empty,
            "department_rows": empty,
            "detail_rows": empty,
        }

    session_rows = filtered.groupby("session_group_key", sort=True, dropna=False).agg(
        month_key=("training_month_key", _first_nonblank_value),
        training_date_iso=("training_date_iso", _first_nonblank_value),
        training_date=("training_date_display", _first_nonblank_value),
        session_code=("session_code", _first_nonblank_value),
        course_id=("course_id", _first_nonblank_value),
        course_name=("course_name", _first_nonblank_value),
        training_unit=("training_unit", _first_nonblank_value),
        location=("location", _first_nonblank_value),
        participant_count=("session_group_key", "size"),
        unique_employee_count=("emp_id", _unique_nonblank_count),
        departments_participating=("department_name", lambda series: ", ".join(sorted(pd.unique(series.loc[series.ne("")])))),
        session_total_cost=("session_total_cost", _first_nonblank_numeric_or_blank),
        avg_cost_per_pax=("participant_cost", _first_nonblank_numeric_or_blank),
        cost_basis=("cost_basis", _first_nonblank_value),
        cost_basis_label=("cost_basis_label", _first_nonblank_value),
    ).reset_index(drop=True)
    session_rows = session_rows.sort_values(["training_date_iso", "session_code"]).reset_index(drop=True)

    detail_rows = filtered[
        [
            "year",
            "training_month_key",
            "training_date_display",
            "session_code",
            "course_id",
            "course_name",
            "training_unit",
            "location",
            "emp_id",
            "full_name",
            "department_name",
            "division",
            "job_title",
            "attendance_status",
            "plan_scope",
            "program_type",
            "delivery_type",
            "training_format",
            "participant_cost",
            "session_total_cost",
            "cost_basis_label",
        ]
    ].rename(columns={"training_month_key": "month_key", "training_date_display": "training_date", "department_name": "department"})

    department_rows = detail_rows.groupby("department", sort=True, dropna=False).agg(
        participant_count=("department", "size"),
        unique_employee_count=("emp_id", _unique_nonblank_count),
        session_count=("session_code", _unique_nonblank_count),
        total_allocated_cost=("participant_cost", lambda series: _sum_or_blank_if_all_missing(series)),
    ).reset_index()
    if not department_rows.empty:
        department_rows["avg_cost_per_pax"] = np.where(
            department_rows["total_allocated_cost"].isna(),
            np.nan,
            (department_rows["total_allocated_cost"] / department_rows["participant_count"]).round(2),
        )

    session_total_costs = session_rows["session_total_cost"]
    total_cost = _sum_or_blank_if_all_missing(session_total_costs)

    return {
        "fiscal_year": inputs.fiscal_year,
        "refreshed_at": inputs.last_refreshed,
        "filters": {
            "delivery_type": "Đào tạo bên ngoài",
            "training_format": "Cử nhân sự đi học",
            "participant_scope": "ALL_RECORDED",
        },
        "kpis": {
            "session_count": int(len(session_rows)),
            "participant_count": int(len(detail_rows)),
            "unique_employee_count": _unique_nonblank_count(detail_rows["emp_id"]),
            "total_cost": _to_python_cell(total_cost),
        },
        "session_rows": session_rows.drop(columns=["training_date_iso"]),
        "department_rows": department_rows,
        "detail_rows": detail_rows,
    }


def _build_session_reconciliation_report_data(inputs: NormalizedAnalyticsInputs) -> dict[str, Any]:
    sessions = inputs.active_training_sessions.loc[
        inputs.active_training_sessions["year_key"].eq(inputs.fiscal_year)
    ].copy()
    sessions = sessions.sort_values(["training_date_iso", "session_code"]).reset_index(drop=True)
    raw_summary = _build_session_reconciliation_summary(inputs.active_raw_participants, sessions, "raw")
    fact_summary = _build_session_reconciliation_summary(inputs.active_training_records, sessions, "fact")

    detail_rows = sessions[
        [
            "training_month_key",
            "training_date_iso",
            "session_id",
            "session_code",
            "course_id",
            "course_name",
            "training_unit",
            "delivery_type",
            "training_format",
            "registered_count",
            "actual_count",
            "attendance_rate",
            "estimated_cost",
            "total_hours",
            "session_group_key",
        ]
    ].rename(columns={"training_month_key": "month_key", "attendance_rate": "manual_attendance_rate"})
    detail_rows["training_date"] = _format_date_series(detail_rows["training_date_iso"])
    detail_rows = detail_rows.merge(raw_summary, on="session_group_key", how="left")
    detail_rows = detail_rows.merge(fact_summary, on="session_group_key", how="left")
    for column in [
        "raw_row_count",
        "raw_unique_count",
        "raw_present_count",
        "raw_excused_absent_count",
        "raw_unexcused_absent_count",
        "raw_late_count",
        "fact_record_count",
        "fact_unique_count",
        "fact_present_count",
    ]:
        detail_rows[column] = detail_rows[column].fillna(0).astype(int)

    detail_rows["registered_vs_raw_delta"] = np.where(
        detail_rows["registered_count"].notna(),
        (detail_rows["registered_count"] - detail_rows["raw_unique_count"]).round(2),
        np.nan,
    )
    detail_rows["actual_vs_raw_present_delta"] = np.where(
        detail_rows["actual_count"].notna(),
        (detail_rows["actual_count"] - detail_rows["raw_present_count"]).round(2),
        np.nan,
    )
    detail_rows["raw_vs_fact_unique_delta"] = (
        detail_rows["raw_unique_count"] - detail_rows["fact_unique_count"]
    ).round(2)
    detail_rows["reconciliation_status"] = _derive_reconciliation_status(detail_rows)
    detail_rows["suggested_action"] = _derive_reconciliation_suggested_action(detail_rows)

    ranked_registered = detail_rows.assign(
        abs_registered_delta=detail_rows["registered_vs_raw_delta"].fillna(0).abs()
    ).sort_values(["abs_registered_delta"], ascending=[False]).head(5)
    ranked_actual = detail_rows.assign(
        abs_actual_delta=detail_rows["actual_vs_raw_present_delta"].fillna(0).abs()
    ).sort_values(["abs_actual_delta"], ascending=[False]).head(5)

    final_detail_rows = detail_rows[
        [
            "month_key",
            "training_date",
            "session_id",
            "session_code",
            "course_id",
            "course_name",
            "training_unit",
            "delivery_type",
            "training_format",
            "registered_count",
            "actual_count",
            "manual_attendance_rate",
            "raw_row_count",
            "raw_unique_count",
            "raw_present_count",
            "raw_excused_absent_count",
            "raw_unexcused_absent_count",
            "raw_late_count",
            "fact_record_count",
            "fact_unique_count",
            "fact_present_count",
            "registered_vs_raw_delta",
            "actual_vs_raw_present_delta",
            "raw_vs_fact_unique_delta",
            "reconciliation_status",
            "suggested_action",
            "estimated_cost",
            "total_hours",
        ]
    ].copy()

    return {
        "fiscal_year": inputs.fiscal_year,
        "refreshed_at": inputs.last_refreshed,
        "kpis": {
            "total_sessions": int(len(final_detail_rows)),
            "sessions_with_raw": int(final_detail_rows["raw_row_count"].gt(0).sum()),
            "sessions_without_raw": int(final_detail_rows["raw_row_count"].eq(0).sum()),
            "matched_sessions": int(final_detail_rows["reconciliation_status"].eq("Khớp").sum()),
            "mismatched_sessions": int(final_detail_rows["reconciliation_status"].ne("Khớp").sum()),
        },
        "rankings": {
            "top_registered_delta": ranked_registered,
            "top_actual_delta": ranked_actual,
        },
        "detail_rows": final_detail_rows,
    }


def _build_external_assignment_session_lookup(training_sessions: pd.DataFrame) -> pd.DataFrame:
    if training_sessions.empty:
        return pd.DataFrame(columns=["lookup_key", "estimated_cost", "avg_cost_per_pax", "cost_per_pax"])
    lookup = pd.concat(
        [
            training_sessions.assign(lookup_key=_normalize_key_series(training_sessions["session_id"])),
            training_sessions.assign(lookup_key=_normalize_key_series(training_sessions["session_code"])),
        ],
        ignore_index=True,
    )
    lookup = lookup.loc[lookup["lookup_key"].ne("")].drop_duplicates("lookup_key", keep="last")
    return lookup[["lookup_key", "estimated_cost", "avg_cost_per_pax", "cost_per_pax"]].copy()


def _attach_external_assignment_costs(
    filtered: pd.DataFrame,
    session_lookup: pd.DataFrame,
    fiscal_year: str,
) -> pd.DataFrame:
    if filtered.empty:
        return filtered
    lookup_keys = set(session_lookup["lookup_key"]) if not session_lookup.empty else set()
    key1 = _normalize_key_series(filtered["session_id"])
    key2 = _normalize_key_series(filtered["session_code"])
    match_key = pd.Series(
        np.select([key1.isin(lookup_keys), key2.isin(lookup_keys)], [key1, key2], default=""),
        index=filtered.index,
    )
    merged = filtered.assign(session_match_key=match_key).merge(
        session_lookup.add_prefix("snapshot_"),
        left_on="session_match_key",
        right_on="snapshot_lookup_key",
        how="left",
    )
    session_costs = merged.groupby("session_group_key", sort=False).agg(
        estimated_cost=("estimated_cost", lambda series: _first_nonblank_numeric_or_blank(series.where(series.gt(0)))),
        avg_cost_per_pax=("avg_cost_per_pax", lambda series: _first_nonblank_numeric_or_blank(series.where(series.gt(0)))),
        cost_per_pax=("cost_per_pax", lambda series: _first_nonblank_numeric_or_blank(series.where(series.gt(0)))),
        snapshot_estimated_cost=("snapshot_estimated_cost", _first_nonblank_numeric_or_blank),
        snapshot_avg_cost_per_pax=("snapshot_avg_cost_per_pax", _first_nonblank_numeric_or_blank),
        snapshot_cost_per_pax=("snapshot_cost_per_pax", _first_nonblank_numeric_or_blank),
        participant_count=("session_group_key", "size"),
    ).reset_index()

    session_costs["estimated_cost_final"] = session_costs["estimated_cost"].fillna(
        session_costs["snapshot_estimated_cost"]
    )
    session_costs["avg_cost_per_pax_final"] = session_costs["avg_cost_per_pax"].fillna(
        session_costs["snapshot_avg_cost_per_pax"]
    )
    session_costs["cost_per_pax_final"] = session_costs["cost_per_pax"].fillna(
        session_costs["snapshot_cost_per_pax"]
    )
    session_costs["cost_basis"] = np.select(
        [
            session_costs["estimated_cost_final"].notna(),
            session_costs["avg_cost_per_pax_final"].notna(),
            session_costs["cost_per_pax_final"].notna(),
        ],
        ["ESTIMATED_COST", "AVG_COST_PER_PAX", "COST_PER_PAX"],
        default="MISSING",
    )
    session_costs["cost_basis_label"] = session_costs["cost_basis"].map(
        {
            "ESTIMATED_COST": "Chi phí dự kiến lớp",
            "AVG_COST_PER_PAX": "Chi phí bình quân/người",
            "COST_PER_PAX": "Chi phí/người",
            "MISSING": "Chưa có dữ liệu chi phí",
        }
    )
    session_costs["session_total_cost"] = np.select(
        [
            session_costs["cost_basis"].eq("ESTIMATED_COST"),
            session_costs["cost_basis"].eq("AVG_COST_PER_PAX"),
            session_costs["cost_basis"].eq("COST_PER_PAX"),
        ],
        [
            session_costs["estimated_cost_final"],
            (session_costs["avg_cost_per_pax_final"] * session_costs["participant_count"]).round(2),
            (session_costs["cost_per_pax_final"] * session_costs["participant_count"]).round(2),
        ],
        default=np.nan,
    )
    session_costs["participant_cost"] = np.select(
        [
            session_costs["cost_basis"].eq("ESTIMATED_COST"),
            session_costs["cost_basis"].eq("AVG_COST_PER_PAX"),
            session_costs["cost_basis"].eq("COST_PER_PAX"),
        ],
        [
            (session_costs["estimated_cost_final"] / session_costs["participant_count"].clip(lower=1)).round(2),
            session_costs["avg_cost_per_pax_final"],
            session_costs["cost_per_pax_final"],
        ],
        default=np.nan,
    )
    merged = merged.merge(
        session_costs[
            ["session_group_key", "session_total_cost", "participant_cost", "cost_basis", "cost_basis_label"]
        ],
        on="session_group_key",
        how="left",
    )
    merged["year"] = fiscal_year
    merged["training_date_display"] = _format_date_series(merged["training_date_iso"])
    return merged


def _build_session_reconciliation_summary(
    source: pd.DataFrame,
    sessions: pd.DataFrame,
    prefix: str,
) -> pd.DataFrame:
    if source.empty or sessions.empty:
        if prefix == "raw":
            return pd.DataFrame(
                columns=[
                    "session_group_key",
                    "raw_row_count",
                    "raw_unique_count",
                    "raw_present_count",
                    "raw_excused_absent_count",
                    "raw_unexcused_absent_count",
                    "raw_late_count",
                ]
            )
        return pd.DataFrame(columns=["session_group_key", "fact_record_count", "fact_unique_count", "fact_present_count"])

    lookup = _build_reconciliation_lookup(sessions)
    resolved = _resolve_reconciliation_scope(source, lookup)
    resolved = resolved.loc[resolved["resolved_year_key"].eq(sessions["year_key"].iloc[0])].copy()
    if resolved.empty:
        return _build_session_reconciliation_summary(pd.DataFrame(), pd.DataFrame(), prefix)

    grouped = resolved.groupby("resolved_group_key", sort=False, dropna=False)
    if prefix == "raw":
        summary = grouped.agg(
            raw_row_count=("resolved_group_key", "size"),
            raw_unique_count=("participant_identity_key", _unique_nonblank_count),
            raw_present_count=("attendance_is_present", "sum"),
            raw_excused_absent_count=("attendance_is_excused_absent", "sum"),
            raw_unexcused_absent_count=("attendance_is_unexcused_absent", "sum"),
            raw_late_count=("attendance_is_late", "sum"),
        ).reset_index()
    else:
        summary = grouped.agg(
            fact_record_count=("resolved_group_key", "size"),
            fact_unique_count=("participant_identity_key", _unique_nonblank_count),
            fact_present_count=("attendance_is_present", "sum"),
        ).reset_index()
    return summary.rename(columns={"resolved_group_key": "session_group_key"})


def _build_reconciliation_lookup(sessions: pd.DataFrame) -> pd.DataFrame:
    if sessions.empty:
        return pd.DataFrame(columns=["lookup_key", "resolved_group_key", "resolved_year_key"])
    course_or_name = _first_nonblank([sessions["course_id"], sessions["course_name"]], "")
    lookup = pd.concat(
        [
            pd.DataFrame({"lookup_key": _normalize_key_series(sessions["session_id"]), "resolved_group_key": sessions["session_group_key"], "resolved_year_key": sessions["year_key"]}),
            pd.DataFrame({"lookup_key": _normalize_key_series(sessions["session_code"]), "resolved_group_key": sessions["session_group_key"], "resolved_year_key": sessions["year_key"]}),
            pd.DataFrame({"lookup_key": _normalize_key_series(course_or_name + "::" + sessions["training_date_iso"]), "resolved_group_key": sessions["session_group_key"], "resolved_year_key": sessions["year_key"]}),
            pd.DataFrame({"lookup_key": _normalize_key_series(sessions["course_name"] + "::" + sessions["training_date_iso"]), "resolved_group_key": sessions["session_group_key"], "resolved_year_key": sessions["year_key"]}),
        ],
        ignore_index=True,
    )
    lookup = lookup.loc[lookup["lookup_key"].ne("")].drop_duplicates("lookup_key", keep="last")
    return lookup


def _resolve_reconciliation_scope(source: pd.DataFrame, lookup: pd.DataFrame) -> pd.DataFrame:
    if source.empty:
        return source.assign(resolved_group_key=pd.Series(dtype="object"), resolved_year_key=pd.Series(dtype="object"))
    course_or_name = _first_nonblank([source["course_id"], source["course_name"]], "")
    key1 = _normalize_key_series(source["session_id"])
    key2 = _normalize_key_series(source["session_code"])
    key3 = _normalize_key_series(course_or_name + "::" + source["training_date_iso"])
    key4 = _normalize_key_series(source["course_name"] + "::" + source["training_date_iso"])
    lookup_keys = set(lookup["lookup_key"]) if not lookup.empty else set()
    match_key = pd.Series(
        np.select(
            [key1.isin(lookup_keys), key2.isin(lookup_keys), key3.isin(lookup_keys), key4.isin(lookup_keys)],
            [key1, key2, key3, key4],
            default="",
        ),
        index=source.index,
    )
    merged = source.assign(_match_key=match_key).merge(lookup, left_on="_match_key", right_on="lookup_key", how="left")
    fallback_group_key = _first_nonblank([source["session_id"], source["session_code"], source["course_id"], source["course_name"]], "")
    merged["resolved_group_key"] = _string_series_from(merged["resolved_group_key"]).replace("", pd.NA).fillna(fallback_group_key)
    merged["resolved_year_key"] = _string_series_from(merged["resolved_year_key"]).replace("", pd.NA).fillna(source["year_key"])
    return merged


def _derive_reconciliation_status(detail_rows: pd.DataFrame) -> pd.Series:
    registered_matches = detail_rows["registered_count"].isna() | detail_rows["registered_count"].eq(detail_rows["raw_unique_count"])
    actual_matches = detail_rows["actual_count"].isna() | detail_rows["actual_count"].eq(detail_rows["raw_present_count"])
    fact_unique_matches = detail_rows["raw_unique_count"].eq(detail_rows["fact_unique_count"])
    fact_present_matches = detail_rows["raw_present_count"].eq(detail_rows["fact_present_count"])
    return pd.Series(
        np.select(
            [
                detail_rows["raw_row_count"].eq(0),
                detail_rows["raw_row_count"].gt(0) & detail_rows["fact_record_count"].eq(0),
                registered_matches & actual_matches & fact_unique_matches & fact_present_matches,
            ],
            ["Chưa có data raw", "Đã có raw, chưa đồng bộ", "Khớp"],
            default="Lệch số liệu",
        ),
        index=detail_rows.index,
    )


def _derive_reconciliation_suggested_action(detail_rows: pd.DataFrame) -> pd.Series:
    return pd.Series(
        np.select(
            [
                detail_rows["raw_row_count"].eq(0),
                detail_rows["raw_row_count"].gt(0) & detail_rows["fact_record_count"].eq(0),
                detail_rows["registered_count"].notna() & detail_rows["registered_count"].ne(detail_rows["raw_unique_count"]),
                detail_rows["actual_count"].notna() & detail_rows["actual_count"].ne(detail_rows["raw_present_count"]),
                detail_rows["raw_unique_count"].ne(detail_rows["fact_unique_count"]) | detail_rows["raw_present_count"].ne(detail_rows["fact_present_count"]),
            ],
            [
                "Bổ sung data raw cho lớp",
                "Chạy Đồng bộ dữ liệu đào tạo",
                "Kiểm tra số lượng HV đăng ký ở Lớp đào tạo",
                "Kiểm tra số lượng HV thực tế / điểm danh",
                "Kiểm tra raw rồi đồng bộ lại",
            ],
            default="Không cần xử lý",
        ),
        index=detail_rows.index,
    )


def build_grade_level_canvas(inputs: NormalizedAnalyticsInputs) -> list[list[Any]]:
    report_data = _build_grade_level_report_data(inputs)
    sections = [
        {
            "title": "TỔNG HỢP THEO NGẠCH",
            "header": ["Ngạch", "Số nhân sự", "Số lượt học viên", "Tổng giờ đào tạo", "Giờ TB / người", "Điểm trung bình", "Hài lòng trung bình"],
            "rows": _frame_to_rows(
                report_data["by_grade"],
                ["grade_name", "unique_employees", "training_count", "total_hours", "hours_per_employee", "avg_score", "avg_satisfaction"],
            ),
        },
        {
            "title": "TỔNG HỢP THEO CẤP BẬC",
            "header": ["Cấp bậc", "Số nhân sự", "Số lượt học viên", "Tổng giờ đào tạo", "Giờ TB / người", "Điểm trung bình", "Hài lòng trung bình"],
            "rows": _frame_to_rows(
                report_data["by_level"],
                ["level_name", "unique_employees", "training_count", "total_hours", "hours_per_employee", "avg_score", "avg_satisfaction"],
            ),
        },
        {
            "title": "CHI TIẾT PHÒNG BAN - NGẠCH",
            "header": ["Phòng ban", "Ngạch", "Số nhân sự", "Số lượt học viên", "Tổng giờ đào tạo", "Giờ TB / người", "Điểm trung bình", "Hài lòng trung bình", "Xếp hạng lượt học"],
            "rows": _frame_to_rows(
                report_data["by_dept_grade"],
                ["department", "grade_name", "unique_employees", "training_count", "total_hours", "hours_per_employee", "avg_score", "avg_satisfaction", "rank"],
            ),
        },
        {
            "title": "CHI TIẾT PHÒNG BAN - CẤP BẬC",
            "header": ["Phòng ban", "Cấp bậc", "Số nhân sự", "Số lượt học viên", "Tổng giờ đào tạo", "Giờ TB / người", "Điểm trung bình", "Hài lòng trung bình", "Xếp hạng lượt học"],
            "rows": _frame_to_rows(
                report_data["by_dept_level"],
                ["department", "level_name", "unique_employees", "training_count", "total_hours", "hours_per_employee", "avg_score", "avg_satisfaction", "rank"],
            ),
        },
        {
            "title": "DANH SÁCH HỌC VIÊN CHI TIẾT (Top 50 theo giờ đào tạo)",
            "header": ["Phòng ban", "Ngạch", "Cấp bậc", "Chức danh", "Mã NV", "Họ tên", "Số lượt học viên", "Tổng giờ đào tạo", "Điểm trung bình"],
            "rows": _frame_to_rows(
                report_data["top_learners"],
                ["department", "grade_name", "level_name", "job_title", "emp_id", "full_name", "training_count", "total_hours", "avg_score"],
            ),
        },
    ]
    return _build_standard_report_canvas(
        title="BÁO CÁO PHÂN TÍCH THEO NGẠCH VÀ CẤP BẬC",
        meta_pairs=[
            ("Năm báo cáo", report_data["fiscal_year"]),
            ("Làm mới lúc", _format_timestamp_display(report_data["refreshed_at"])),
        ],
        kpi_items=[
            ("Tổng giờ đào tạo", report_data["kpis"]["total_hours"]),
            ("Số ngạch tham gia", report_data["kpis"]["grade_count"]),
            ("Số cấp bậc tham gia", report_data["kpis"]["level_count"]),
            ("Tổng nhân sự đã ĐT", report_data["kpis"]["trained_employees"]),
            ("Giờ ĐT / người", report_data["kpis"]["hours_per_employee"]),
        ],
        sections=sections,
        width=12,
    )


def _build_grade_level_report_data(inputs: NormalizedAnalyticsInputs) -> dict[str, Any]:
    records = inputs.active_training_records
    
    def _aggregate(group_cols: list[str], result_cols: list[str]) -> pd.DataFrame:
        grouped = records.groupby(group_cols, sort=True, dropna=False)
        rows = grouped.agg(
            training_count=(group_cols[0], "size"),
            unique_employees=("emp_id", _unique_nonblank_count),
            avg_score=("score", _mean_or_blank),
            avg_satisfaction=("satisfaction", _mean_or_blank),
            total_hours=("duration_hours", _sum_or_zero),
        ).reset_index()
        for idx, col in enumerate(group_cols):
            rows[result_cols[idx]] = rows[col]
            rows[result_cols[idx]] = rows[result_cols[idx]].astype(object).apply(
                lambda v: _first_nonblank_value(pd.Series([v])) if pd.notna(v) else ""
            )
        
        if rows.empty:
            result_df = pd.DataFrame(columns=result_cols + ["training_count", "unique_employees", "avg_score", "avg_satisfaction", "total_hours", "hours_per_employee"])
        else:
            rows["hours_per_employee"] = _safe_divide_series(rows["total_hours"], rows["unique_employees"])
            result_df = rows.copy()
        return result_df

    by_grade = _aggregate(["grade_name"], ["grade_name"])
    by_level = _aggregate(["level_name"], ["level_name"])
    
    by_dept_grade = _aggregate(["department_name", "grade_name"], ["department", "grade_name"])
    if not by_dept_grade.empty:
        by_dept_grade = _assign_rank(by_dept_grade, "training_count", "grade_name", "rank")
        by_dept_grade = by_dept_grade.sort_values(["department", "rank"]).reset_index(drop=True)
    else:
        by_dept_grade["rank"] = pd.Series(dtype="int64")

    by_dept_level = _aggregate(["department_name", "level_name"], ["department", "level_name"])
    if not by_dept_level.empty:
        by_dept_level = _assign_rank(by_dept_level, "training_count", "level_name", "rank")
        by_dept_level = by_dept_level.sort_values(["department", "rank"]).reset_index(drop=True)
    else:
        by_dept_level["rank"] = pd.Series(dtype="int64")

    learners_grouped = records.groupby(["emp_id"], sort=True, dropna=False)
    learners = learners_grouped.agg(
        emp_id=("emp_id", lambda s: _first_nonblank_value(s)),
        full_name=("full_name", lambda s: _first_nonblank_value(s)),
        department=("department_name", lambda s: _first_nonblank_value(s)),
        grade_name=("grade_name", lambda s: _first_nonblank_value(s)),
        level_name=("level_name", lambda s: _first_nonblank_value(s)),
        job_title=("job_title", lambda s: _first_nonblank_value(s)),
        training_count=("emp_id", "size"),
        total_hours=("duration_hours", _sum_or_zero),
        avg_score=("score", _mean_or_blank),
    ).reset_index(drop=True)
    
    if not learners.empty:
        top_learners = learners.sort_values(["total_hours", "training_count"], ascending=[False, False]).head(50).reset_index(drop=True)
    else:
        top_learners = pd.DataFrame(columns=["emp_id", "full_name", "department", "grade_name", "level_name", "job_title", "training_count", "total_hours", "avg_score"])

    total_hours = _sum_or_zero(records["duration_hours"])
    trained_employees = _unique_nonblank_count(records["emp_id"])

    return {
        "fiscal_year": inputs.fiscal_year,
        "refreshed_at": inputs.last_refreshed,
        "kpis": {
            "total_hours": total_hours,
            "trained_employees": trained_employees,
            "grade_count": len(by_grade.loc[by_grade["grade_name"].ne("")]) if not by_grade.empty else 0,
            "level_count": len(by_level.loc[by_level["level_name"].ne("")]) if not by_level.empty else 0,
            "hours_per_employee": round(float(total_hours) / float(trained_employees), 2) if trained_employees > 0 else 0,
        },
        "by_grade": by_grade,
        "by_level": by_level,
        "by_dept_grade": by_dept_grade,
        "by_dept_level": by_dept_level,
        "top_learners": top_learners,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# Looker Studio Flat Data Sheet
# ═══════════════════════════════════════════════════════════════════════════════

_LOOKER_COLUMNS = [
    "emp_id", "full_name", "department", "division", "grade", "level",
    "course_id", "course_name", "course_category", "training_format",
    "delivery_type", "training_date", "training_month", "duration_hours",
    "score", "satisfaction", "attendance_status",
]

_LOOKER_HEADERS = [
    "Mã NV", "Họ tên", "Phòng ban", "Khối", "Ngạch", "Cấp bậc",
    "Mã khóa học", "Tên khóa học", "Nhóm khóa học", "Hình thức ĐT",
    "Loại hình ĐT", "Ngày đào tạo", "Tháng đào tạo", "Thời lượng (giờ)",
    "Điểm", "Hài lòng", "Trạng thái điểm danh",
]


def build_looker_flat_matrix(inputs: NormalizedAnalyticsInputs) -> list[list[Any]]:
    """Build a flat, denormalized table from active training records for Looker Studio.

    Each row = one training record. All dimensions and metrics are included as columns.
    This sheet is designed to be the single data source for Looker Studio dashboards.
    """
    records = inputs.active_training_records
    if records.empty:
        return [_LOOKER_HEADERS]

    rows: list[list[Any]] = [_LOOKER_HEADERS]
    for _, rec in records.iterrows():
        row: list[Any] = []
        for col in _LOOKER_COLUMNS:
            val = rec.get(col)
            if pd.isna(val) if isinstance(val, float) else val is None:
                row.append("")
            elif isinstance(val, pd.Timestamp):
                row.append(val.strftime("%Y-%m-%d"))
            else:
                row.append(val)
        rows.append(row)
    return rows


__all__ = [
    "AnalyticsInputs",
    "NormalizedAnalyticsInputs",
    "fetch_data",
    "normalize_inputs",
    "transform_data",
    "build_dashboard_exec_matrix",
    "build_dashboard_operations_matrix",
    "build_course_matrix",
    "build_department_canvas",
    "build_department_course_canvas",
    "build_trend_canvas",
    "build_external_assignment_canvas",
    "build_session_reconciliation_canvas",
    "build_grade_level_canvas",
    "build_looker_flat_matrix",
    "write_data",
]
