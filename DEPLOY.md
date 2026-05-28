# 任务甘特图工具服务器部署

这套部署把当前工具作为一个长期在线服务运行，并通过 Tool-Nexus 的 `frps/frpc` 接入 `*.aim888888.xyz` 门户。

## 推荐结构

- `task-gantt`：当前应用容器，监听容器内 `8010`。
- `task-gantt-frpc`：FRP 客户端容器，连接 Tool-Nexus 的 `frp.aim888888.xyz:7000`。
- `data/`：服务器项目目录下的持久化数据目录，保存 `task_gantt.db`。
- Tool-Nexus Portal：负责登录态、白名单和统一入口。

## 1. 在 Tool-Nexus 里创建网页工具

登录：`https://portal.aim888888.xyz`

新增网页工具建议填写：

- 工具名称：`任务甘特图`
- 子域名前缀：`gantt`
- 默认打开后缀：`/`
- 访问范围：按需要选择私有或公开；私有时把可访问成员加入白名单

这个子域名前缀必须和 `deploy/frpc.toml` 里的 `customDomains` 保持一致。

## 2. 上传项目到服务器

如果这个项目还没有 Git 仓库，可以先用压缩包或 `scp` 上传到服务器，例如放到：

```bash
/opt/task-gantt
```

推荐后续把它放进 Git，这样更新时只需要 `git pull`。

## 3. 配置 FRP 客户端

在服务器项目目录执行：

```bash
cd /opt/task-gantt
cp deploy/frpc.toml.example deploy/frpc.toml
nano deploy/frpc.toml
```

需要修改：

```toml
auth.token = "Tool-Nexus 服务器里的真实 FRP_TOKEN"
customDomains = ["gantt.aim888888.xyz"]
```

如果你换了子域名，比如 `plan`，Portal 里也要创建 `plan`，这里也要改成 `plan.aim888888.xyz`。

## 4. 配置 LLM 环境变量

可选。需要 SiliconFlow 时，在服务器项目目录创建 `.env`：

```bash
cp .env.example .env
nano .env
```

填入：

```env
SILICONFLOW_API_KEY=你的_key
```

没有 key 也能运行，只是智能拆分会回退到内置规则。

## 5. 启动并保持常驻

```bash
cd /opt/task-gantt
docker compose up -d --build
```

查看状态：

```bash
docker compose ps
docker compose logs -f task-gantt
docker compose logs -f frpc
```

`restart: unless-stopped` 会让应用和 frpc 在容器异常退出或服务器重启后自动恢复。

## 6. 访问

浏览器打开：

```text
https://gantt.aim888888.xyz
```

如果该项目在 Tool-Nexus 中设为私有，未登录或未授权用户会先跳转到 Portal 登录页。

## 7. 更新代码

```bash
cd /opt/task-gantt
git pull
docker compose up -d --build
```

如果不是 Git 部署，就上传新文件后执行同样的 `docker compose up -d --build`。

## 8. 数据备份

数据库在服务器项目目录的 `data/` 中。导出备份示例：

```bash
tar czf task-gantt-data.tar.gz -C data .
```

恢复前先停止服务，再把归档解回 `data/` 目录。
