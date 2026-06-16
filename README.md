# 任务甘特图工具

一个本地优先的任务规划小工具：
- 用 SQLite 做表存储
- 用网页管理项目与任务
- 支持按项目描述自动拆分子任务并估时
- 支持依赖关系排期和甘特图展示
- 支持 CSV / JSON / XLSX 导入导出
- 支持每周会议纪要解析后自动更新当前项目任务进度
- 已接入 SiliconFlow LLM，可用 `deepseek-ai/DeepSeek-V3.2` 做任务拆解与估时

## 启动方式

```bash
python app.py
```

启动后打开：

```text
http://127.0.0.1:8010
```

## LLM 配置

项目根目录已经提供 `.env`，你只需要填写：

```env
SILICONFLOW_API_KEY=你的_key
```

默认配置：
- `SILICONFLOW_BASE_URL=https://api.siliconflow.cn/v1/chat/completions`
- `SILICONFLOW_MODEL=deepseek-ai/DeepSeek-V3.2`
- `SILICONFLOW_TIMEOUT_SECONDS=45`

填写后重启服务即可。

会议更新说明：
- 顶部新增“会议更新”按钮，粘贴周会纪要后可自动更新当前项目任务进度
- 仅会更新已存在任务，不会凭空创建新任务

点击“LLM 智能拆分创建”时：
- 已配置 key：优先调用 SiliconFlow + DeepSeek
- 未配置 key 或调用失败：自动回退到内置规则拆分，不会阻塞创建项目

## 当前第一版能力

- 新建空项目
- LLM 智能拆分创建项目
- 任务增删改查
- 按依赖重算排期
- 甘特图查看
- 导入 CSV / JSON / XLSX
- 导出 CSV / JSON / XLSX

## 导入字段建议

支持中英文字段名，常用列如下：

- `title` / `任务` / `任务名`
- `description` / `描述`
- `status` / `状态`
- `owner` / `负责人`
- `priority` / `优先级`
- `complexity` / `复杂度`
- `estimate_hours` / `预估工时`
- `actual_hours` / `实际工时`
- `start_date` / `开始日期`
- `end_date` / `结束日期`
- `depends_on` / `依赖` / `前置任务`
- `parent` / `父任务`
- `notes` / `备注`

## 访问鉴权（可选，公网部署建议开启）

应用本身默认不鉴权，方便本地使用。若通过隧道/反代对公网暴露，建议在 `.env` 设置：

```env
TASK_GANTT_AUTH_TOKEN=一个足够长的随机串
TASK_GANTT_AUTH_USER=admin   # 可选，默认 admin
```

设置后浏览器会弹出 Basic Auth 登录框；未设置则完全放行。`/api/health` 始终免鉴权，供容器健康检查使用。

其它可选环境变量见 `.env.example`：`TASK_GANTT_TRUST_PROXY`（是否信任反代的 `X-Forwarded-For`，用于限流识别 IP）、`TASK_GANTT_LLM_WORKERS`（LLM 后台任务并发数）。

## LLM 任务为后台异步执行

智能拆分创建、智能追加导入、会议更新都在后台线程池执行：接口立即返回 `job_id`，前端轮询 `GET /api/jobs/{id}` 显示**真实阶段进度**（提交→调用模型→整理任务→写入），不再是匀速假进度。LLM 端点带有限流，避免被刷爆 API 额度。

## 运行测试

```bash
pip install -r requirements-dev.txt
pytest -q
```

覆盖排期/估时/解析等纯函数、数据库闭环、以及鉴权/限流/请求体上限/异步任务等运行时行为。

## 说明

- 数据库文件默认保存在 `data/task_gantt.db`（已开启 WAL，会附带 `-wal`/`-shm` 文件）
- 首次启动会自动生成一个示例项目
- `XLSX` 功能依赖 `openpyxl`

## 维护约定

后续只要修改代码、静态资源、Docker 配置或部署脚本，都要同步更新 `CHANGELOG.md`。

每次记录至少包含：变更摘要、涉及代码文件、验证方式和服务器部署提醒。这样服务器端用 GitHub 更新时，可以清楚知道这次改动影响了什么、需要执行哪些命令。
