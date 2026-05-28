# 变更记录

本文件用于记录每次代码或部署配置变更。后续只要修改代码、静态资源、Docker 配置或部署脚本，都要同步更新这里。

## 记录规则

每次变更至少写清楚：

- 日期：实际修改日期。
- 变更摘要：用户能理解的功能或修复说明。
- 涉及代码：列出关键文件和主要改动点。
- 验证方式：写明执行过的检查、测试或部署命令。
- 部署提醒：如果服务器需要 `git pull`、重建容器、迁移数据或修改配置，要明确写出来。

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
