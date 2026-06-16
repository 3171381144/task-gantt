"""Tests for the new runtime pieces: rate limiter, job registry, and HTTP layer."""
import base64
import http.client
import json
import os
import threading
import time

import pytest

import app


# --- RateLimiter ---------------------------------------------------------------------

def test_rate_limiter_allows_then_blocks():
    limiter = app.RateLimiter(max_calls=3, window_seconds=60)
    assert limiter.allow("ip") is True
    assert limiter.allow("ip") is True
    assert limiter.allow("ip") is True
    assert limiter.allow("ip") is False  # 4th within window
    # a different client is independent
    assert limiter.allow("other") is True


def test_rate_limiter_window_expiry():
    limiter = app.RateLimiter(max_calls=1, window_seconds=0.2)
    assert limiter.allow("ip") is True
    assert limiter.allow("ip") is False
    time.sleep(0.25)
    assert limiter.allow("ip") is True


# --- JobRegistry ---------------------------------------------------------------------

def test_job_registry_success_flow():
    registry = app.JobRegistry(max_workers=2)
    try:
        stages = []

        def worker(handle):
            handle.stage("llm")
            handle.stage("persist")
            stages.append("ran")
            return {"ok": True}

        job_id = registry.submit(worker)
        snapshot = _wait_for(registry, job_id, {"done", "error"})
        assert snapshot["status"] == "done"
        assert snapshot["detail"] == {"ok": True}
        assert snapshot["progress"] == 100
    finally:
        registry.shutdown()


def test_job_registry_error_flow():
    registry = app.JobRegistry(max_workers=1)
    try:
        def worker(handle):
            raise ValueError("boom")

        job_id = registry.submit(worker)
        snapshot = _wait_for(registry, job_id, {"done", "error"})
        assert snapshot["status"] == "error"
        assert "boom" in snapshot["error"]
    finally:
        registry.shutdown()


def _wait_for(registry, job_id, statuses, timeout=5.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        snapshot = registry.get(job_id)
        if snapshot and snapshot["status"] in statuses:
            return snapshot
        time.sleep(0.02)
    raise AssertionError("job did not reach a terminal state in time")


# --- HTTP server ---------------------------------------------------------------------

@pytest.fixture
def server_factory(tmp_path):
    started = []

    def factory(env=None):
        env = env or {}
        previous = {key: os.environ.get(key) for key in env}
        os.environ.update(env)
        try:
            db = app.TaskDatabase(tmp_path / f"srv{len(started)}.db", planner=None)
            srv = app.AppServer(("127.0.0.1", 0), app.TaskGanttHandler, db)
        finally:
            for key, value in previous.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value
        thread = threading.Thread(target=srv.serve_forever, daemon=True)
        thread.start()
        started.append(srv)
        return srv

    yield factory

    for srv in started:
        srv.shutdown()
        srv.jobs.shutdown()
        srv.server_close()


def _request(srv, method, path, body=None, headers=None, raw_content_length=None):
    conn = http.client.HTTPConnection("127.0.0.1", srv.server_address[1], timeout=5)
    headers = dict(headers or {})
    if raw_content_length is not None:
        # Fake an oversized Content-Length without sending the bytes.
        conn.putrequest(method, path)
        conn.putheader("Content-Length", str(raw_content_length))
        for key, value in headers.items():
            conn.putheader(key, value)
        conn.endheaders()
        conn.send(b"{}")
    else:
        payload = json.dumps(body).encode("utf-8") if body is not None else None
        if payload is not None:
            headers.setdefault("Content-Type", "application/json")
        conn.request(method, path, body=payload, headers=headers)
    response = conn.getresponse()
    data = response.read()
    conn.close()
    return response.status, data


def test_health_is_public_and_ok(server_factory):
    srv = server_factory()
    status, data = _request(srv, "GET", "/api/health")
    assert status == 200
    payload = json.loads(data)
    assert payload["status"] == "ok"
    assert payload["version"] == app.APP_VERSION


def test_auth_gate(server_factory):
    srv = server_factory({"TASK_GANTT_AUTH_TOKEN": "secret", "TASK_GANTT_AUTH_USER": "admin"})
    # health stays public even with auth enabled
    assert _request(srv, "GET", "/api/health")[0] == 200
    # unauthenticated app access is rejected
    assert _request(srv, "GET", "/")[0] == 401
    # correct Basic credentials pass
    token = base64.b64encode(b"admin:secret").decode()
    status, _ = _request(srv, "GET", "/", headers={"Authorization": f"Basic {token}"})
    assert status == 200


def test_body_too_large(server_factory):
    srv = server_factory()
    status, _ = _request(srv, "POST", "/api/import", raw_content_length=app.MAX_BODY_BYTES + 1)
    assert status == 413


def test_llm_endpoint_rate_limited(server_factory):
    srv = server_factory()
    body = {"name": "限流项目", "description": "实现登录\n编写文档", "start_date": "2026-01-01", "use_suggestions": True}
    statuses = [_request(srv, "POST", "/api/projects", body=body)[0] for _ in range(app.LLM_RATE_LIMIT_MAX + 1)]
    assert statuses[: app.LLM_RATE_LIMIT_MAX] == [202] * app.LLM_RATE_LIMIT_MAX
    assert statuses[-1] == 429


def test_async_job_end_to_end(server_factory):
    srv = server_factory()
    body = {"name": "异步项目", "description": "实现登录\n实现导出\n编写文档", "start_date": "2026-01-01", "use_suggestions": True}
    status, data = _request(srv, "POST", "/api/projects", body=body)
    assert status == 202
    job_id = json.loads(data)["job_id"]

    deadline = time.time() + 10
    detail = None
    while time.time() < deadline:
        code, payload = _request(srv, "GET", f"/api/jobs/{job_id}")
        assert code == 200
        snapshot = json.loads(payload)
        if snapshot["status"] == "done":
            detail = snapshot["detail"]
            break
        if snapshot["status"] == "error":
            raise AssertionError(f"job failed: {snapshot.get('error')}")
        time.sleep(0.1)

    assert detail is not None, "job did not complete in time"
    assert detail["imported_task_count"] >= 1
