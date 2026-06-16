# 变更记录

本文件用于记录每次代码或部署配置变更。后续只要修改代码、静态资源、Docker 配置或部署脚本，都要同步更新这里。

## 记录规则

每次变更至少写清楚：

- 日期：实际修改日期。
- 变更摘要：用户能理解的功能或修复说明。
- 涉及代码：列出关键文件和主要改动点。
- 验证方式：写明执行过的检查、测试或部署命令。
- 部署提醒：如果服务器需要 `git pull`、重建容器、迁移数据或修改配置，要明确写出来。

## 2026-06-16

### P0+P1 加固：安全、并发、异步任务与测试

- 变更摘要：
  - 安全（P0）：新增可选 Basic Auth 鉴权（`TASK_GANTT_AUTH_TOKEN`，未设则放行，`/api/health` 免鉴权）；请求体 8MB 上限（防内存型 DoS）；LLM 端点限流（默认 60s 内 5 次，防烧钱）；内部错误不再把异常详情返回客户端。
  - 并发（P0）：SQLite 开启 WAL + `busy_timeout`，消除多线程下 “database is locked”。
  - 异步（P1）：智能拆分创建 / 智能追加 / 会议更新改为后台线程池任务，接口返回 `job_id`，前端轮询 `GET /api/jobs/{id}` 显示真实阶段进度（替换原匀速假进度条）；LLM 调用加一次性网络重试。
  - 可维护（P1）：Schema 迁移版本化（`schema_migrations` 表）；新增 `GET /api/health`（含版本号）并据此调整容器健康检查；新增 `pytest` 测试层（纯函数 / 数据库闭环 / 鉴权·限流·请求体上限·异步任务）。
- 涉及代码：
  - `app.py`：新增 `RateLimiter`、`JobRegistry`/`JobHandle`、Basic Auth 与请求体上限校验、`/api/health` 与 `/api/jobs/{id}` 路由、`_run_migrations`、planner `_post_chat` 重试；三个 LLM 端点改异步并增加 `on_stage` 回调。
  - `static/app.js`：新增 `runLlmJob` 轮询与 `applyLlmJobProgress`/`applyImportJobProgress`，三个提交流程改为真实进度轮询。
  - `static/index.html`：静态资源缓存版本号更新为 `p0p1-20260616`。
  - `docker-compose.yml`：健康检查改打 `/api/health`。
  - `.env.example`：新增鉴权 / 代理信任 / 并发数环境变量说明。
  - `tests/`、`requirements-dev.txt`：新增测试套件与开发依赖。
  - `README.md`、`.gitignore`：补充鉴权 / 异步 / 测试说明，忽略 `.pytest_cache/`。
- 验证方式：
  - `python -m py_compile app.py`。
  - `pip install -r requirements-dev.txt && pytest -q` → 31 passed（含 WAL、迁移、401 鉴权、413 超限、429 限流、异步任务端到端）。
  - 本地启动后浏览器验证：`/api/health` 返回 `200`；智能创建走 `202 + job_id` → 轮询 `/api/jobs/{id}` 至 `done`，生成 5 个任务；控制台无报错。
- 部署提醒：服务器拉取最新分支后执行 `docker compose up -d --build`。如需对公网开启登录，在 `.env` 设置 `TASK_GANTT_AUTH_TOKEN`（并按需设 `TASK_GANTT_AUTH_USER`）。WAL 会在 `data/` 下生成 `task_gantt.db-wal`/`-shm`（已被 `.gitignore` 覆盖）。保留服务器本地 `data/`、`.env` 和 `deploy/frpc.toml`。

## 2026-06-03

### 智能创建项目概览摘要修复

- 变更摘要：修复 LLM 智能拆分创建项目时，项目概览可能保存整段粘贴材料的问题；现在会优先提取“一句话介绍/项目简介”等摘要字段，并过滤 Markdown 标题、文件路径、状态清单和超长原文。
- 涉及代码：
  - `app.py`：新增项目概览摘要清洗与兜底生成逻辑，收紧 LLM `project_description` 输出要求，并在 rules fallback、LLM fallback、项目落库路径统一使用短摘要。
- 验证方式：
  - 本地执行 `python -m py_compile app.py`。
  - 使用类似高考志愿项目交接材料的输入验证摘要结果，确认不会保存 `README`、Markdown 标题或整段原文。
  - 本地请求 `http://127.0.0.1:8010/api/projects` 返回 `200`。
- 部署提醒：服务器需要拉取最新 `main` 后执行 `docker compose up -d --build`；保留服务器本地 `data/`、`.env` 和 `deploy/frpc.toml`。

## 2026-05-28

### 初始化 GitHub 与服务器部署

- 变更摘要：把任务甘特图工具整理为可部署项目，推送到 GitHub，并接入 Tool-Nexus 的 FRP/Portal 链路。
- 涉及代码：
  - `app.py`：支持通过环境变量配置监听地址、端口和数据目录。
  - `Dockerfile`：新增应用容器构建配置。
  - `docker-compose.yml`：新增 `task-gantt` 与 `frpc` 常驻服务，并使用服务器项目目录下的 `data/` 持久化数据库。
  - `deploy/frpc.toml.example`：新增 FRP 客户端示例配置。
  - `.gitignore`、`.dockerignore`、`.gitattributes`：排除密钥、数据库、日志、本地工具和生成素材。
  - `DEPLOY.md`：记录服务器部署、更新和备份流程。
- 验证方式：
  - 本地执行 `python -m py_compile app.py`。
  - 本地执行 `docker compose config --services`。
  - 服务器执行 `docker compose up -d --build`，确认 `task-gantt` 与 `task-gantt-frpc` 容器启动。
- 部署提醒：服务器后续更新使用 `git pull` 后执行 `docker compose up -d --build`；`data/` 和 `deploy/frpc.toml` 留在服务器本地，不进入 GitHub。

### 文档维护约定

- 变更摘要：新增维护规则，要求后续每次改代码时同步更新 Markdown 变更记录。
- 涉及代码：
  - `CHANGELOG.md`：新增变更记录模板和当前部署记录。
  - `README.md`：新增“维护约定”章节。
  - `DEPLOY.md`：新增“后续代码变更记录”章节。
- 验证方式：确认 Markdown 文件按 UTF-8 写入，并检查 Git 差异。
- 部署提醒：本次仅更新文档；服务器可在下次代码更新时一起 `git pull`。

### 首次访问新手指引

- 变更摘要：新增首次进入任务甘特图时的 6 步半透明新手引导，并提供顶部“新手指引”按钮供用户随时重看。
- 涉及代码：
  - `static/index.html`：新增新手指引入口按钮和导览浮层结构，并更新 JS/CSS 版本号避免缓存旧资源。
  - `static/app.js`：新增导览步骤、首次访问检测、localStorage 完成状态、目标高亮定位、键盘操作和重新查看逻辑。
  - `static/styles.css`：新增半透明遮罩、高亮目标、导览卡片、进度条和移动端适配样式。
  - `.gitignore`：忽略本地验证产生的临时文件和浏览器调试日志。
- 验证方式：
  - 本地执行 `python -m py_compile app.py`。
  - 本地执行 `node --check static/app.js`。
  - 使用 Playwright 启动本地页面，验证首次导览出现、6 步可切换、完成后刷新不再自动弹出、顶部按钮可重新打开。
- 部署提醒：服务器执行 `git pull` 后运行 `docker compose up -d --build`，浏览器会通过新的资源版本号加载最新前端文件。
