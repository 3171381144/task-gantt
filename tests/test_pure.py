"""Unit tests for pure helpers: scheduling, estimation, parsing, normalization."""
import app


# --- small numeric / parsing helpers -------------------------------------------------

def test_clamp_bounds():
    assert app.clamp(5, 0, 10) == 5
    assert app.clamp(-1, 0, 10) == 0
    assert app.clamp(99, 0, 10) == 10


def test_parse_int_and_float_fallbacks():
    assert app.parse_int("12") == 12
    assert app.parse_int("", 7) == 7
    assert app.parse_int(None, 3) == 3
    assert app.parse_int("abc", 1) == 1
    assert app.parse_float("1.5") == 1.5
    assert app.parse_float("x", 2.0) == 2.0


def test_parse_bool_variants():
    assert app.parse_bool("yes") is True
    assert app.parse_bool("off") is False
    assert app.parse_bool("", True) is True
    assert app.parse_bool(True) is True


def test_parse_bool_env(monkeypatch):
    monkeypatch.delenv("FOO_FLAG", raising=False)
    assert app.parse_bool_env("FOO_FLAG", True) is True
    monkeypatch.setenv("FOO_FLAG", "false")
    assert app.parse_bool_env("FOO_FLAG", True) is False
    monkeypatch.setenv("FOO_FLAG", "1")
    assert app.parse_bool_env("FOO_FLAG", False) is True


def test_add_days_and_duration():
    assert app.add_days("2026-01-01", 5) == "2026-01-06"
    # 6 hours == 1 working day; 7 hours rolls into a second day.
    assert app.task_duration_days(6) == 1
    assert app.task_duration_days(7) == 2
    assert app.task_duration_days(0) == 1


def test_infer_category():
    assert app.infer_category("Web 工具", "前端 dashboard") == "web_tool"
    assert app.infer_category("数据分析", "etl pipeline") == "data_project"
    assert app.infer_category("课程脚本", "内容创作") == "content_project"
    assert app.infer_category("随便", "什么") == "general"


# --- status / progress / dependency normalization -----------------------------------

def test_normalize_progress():
    assert app.normalize_progress("done") == 100
    assert app.normalize_progress("planned") == 0
    assert app.normalize_progress("in_progress", 30) == 30
    assert app.normalize_progress("planned", 150) == 100  # clamped


def test_infer_status_from_progress():
    assert app.infer_status_from_progress(0) == "planned"
    assert app.infer_status_from_progress(100) == "done"
    assert app.infer_status_from_progress(50) == "in_progress"
    assert app.infer_status_from_progress(50, "blocked") == "blocked"


def test_normalize_dependencies():
    assert app.normalize_dependencies("1, 2、3") == [1, 2, 3]
    assert app.normalize_dependencies([1, "2", 0, -4]) == [1, 2]
    assert app.normalize_dependencies(None) == []


# --- JSON extraction -----------------------------------------------------------------

def test_extract_json_object_from_fenced_text():
    raw = "noise before ```json\n{\"a\": 1, \"b\": [2,3]}\n``` trailing"
    assert app.extract_json_object(raw) == {"a": 1, "b": [2, 3]}


def test_extract_json_object_rejects_non_json():
    try:
        app.extract_json_object("no json here")
    except ValueError:
        return
    raise AssertionError("expected ValueError")


# --- estimation ----------------------------------------------------------------------

def test_estimate_task_hours_shape():
    hours, confidence, reason = app.estimate_task_hours("写接口", "实现 API", 3, 0, "web_tool")
    assert hours > 0
    assert 0.45 <= confidence <= 0.92
    assert isinstance(reason, str) and reason


# --- scheduling ----------------------------------------------------------------------

def test_schedule_task_batch_respects_dependencies():
    tasks = [
        {"key": "a", "estimate_hours": 6, "dependency_keys": []},
        {"key": "b", "estimate_hours": 6, "dependency_keys": ["a"]},
    ]
    app.schedule_task_batch(tasks, "2026-01-01")
    assert tasks[0]["start_date"] == "2026-01-01"
    # b depends on a, so b starts strictly after a ends.
    assert tasks[1]["start_date"] > tasks[0]["end_date"]


def test_schedule_task_batch_handles_dependency_cycle():
    # A <-> B cycle must not hang; it falls back to declaration order and still dates both.
    tasks = [
        {"key": "a", "estimate_hours": 6, "dependency_keys": ["b"]},
        {"key": "b", "estimate_hours": 6, "dependency_keys": ["a"]},
    ]
    app.schedule_task_batch(tasks, "2026-01-01")
    assert tasks[0]["start_date"] and tasks[1]["start_date"]


def test_build_rule_task_suggestions_offline():
    category, tasks = app.build_rule_task_suggestions(
        "示例项目", "实现登录\n实现导出\n编写文档", "2026-01-01", None
    )
    assert category
    assert len(tasks) >= 1
    assert all(task.get("title") for task in tasks)
