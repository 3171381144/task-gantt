"""Integration tests against a throwaway SQLite database (no network)."""
import app


def make_db(tmp_path):
    return app.TaskDatabase(tmp_path / "t.db", planner=None)


def test_connect_uses_wal(tmp_path):
    db = make_db(tmp_path)
    with db.connect() as connection:
        mode = connection.execute("PRAGMA journal_mode").fetchone()[0]
    assert str(mode).lower() == "wal"


def test_migrations_recorded(tmp_path):
    db = make_db(tmp_path)
    with db.connect() as connection:
        versions = {row[0] for row in connection.execute("SELECT version FROM schema_migrations")}
        columns = {row[1] for row in connection.execute("PRAGMA table_info(projects)")}
    assert 1 in versions
    # migration 1 ensures these columns exist
    assert {"analysis_source", "analysis_model", "analysis_note", "deleted_at"} <= columns


def test_migrations_are_idempotent(tmp_path):
    make_db(tmp_path)
    # Re-opening the same DB must not re-run or error on the already-applied migration.
    db2 = app.TaskDatabase(tmp_path / "t.db", planner=None)
    with db2.connect() as connection:
        count = connection.execute("SELECT COUNT(*) FROM schema_migrations WHERE version = 1").fetchone()[0]
    assert count == 1


def test_project_and_task_lifecycle(tmp_path):
    db = make_db(tmp_path)
    detail = db.create_project(
        {"name": "测试项目", "description": "用于测试", "start_date": "2026-01-01"},
        use_suggestions=False,
    )
    project_id = detail["project"]["id"]

    db.create_task({"project_id": project_id, "title": "任务一", "estimate_hours": 6})
    db.create_task({"project_id": project_id, "title": "任务二", "estimate_hours": 12})

    refreshed = db.get_project_detail(project_id)
    assert len(refreshed["tasks"]) == 2

    # recalculation should assign schedule dates without error
    db.recalculate_project(project_id)
    scheduled = db.get_project_detail(project_id)
    assert all(task["start_date"] for task in scheduled["tasks"])


def test_create_project_with_suggestions_offline(tmp_path):
    db = make_db(tmp_path)
    detail = db.create_project(
        {"name": "离线项目", "description": "实现登录\n实现导出\n编写文档", "start_date": "2026-01-01"},
        use_suggestions=True,
    )
    # planner is None -> rules fallback still produces tasks
    assert detail["imported_task_count"] >= 1
    assert detail["analysis"]["source"] in {"rules", "structured"}


def test_on_stage_callback_fires(tmp_path):
    db = make_db(tmp_path)
    seen = []
    db.create_project(
        {"name": "回调项目", "description": "实现登录\n编写文档", "start_date": "2026-01-01"},
        use_suggestions=True,
        on_stage=seen.append,
    )
    assert "llm" in seen and "persist" in seen


def test_snapshot_restore_roundtrip(tmp_path):
    db = make_db(tmp_path)
    detail = db.create_project(
        {"name": "快照项目", "description": "d", "start_date": "2026-01-01"},
        use_suggestions=False,
    )
    project_id = detail["project"]["id"]
    db.create_task({"project_id": project_id, "title": "唯一任务", "estimate_hours": 6})
    snapshots = db.list_project_snapshots(project_id)
    assert snapshots
    restored = db.restore_project_snapshot(project_id, snapshots[-1]["id"])
    assert "tasks" in restored
