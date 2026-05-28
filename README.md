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

## 说明

- 数据库文件默认保存在 `data/task_gantt.db`
- 首次启动会自动生成一个示例项目
- `XLSX` 功能依赖 `openpyxl`


