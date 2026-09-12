# 供电可靠性智能分析智能体

供电可靠性日报与统计分析服务。当前代码来自已部署的新版服务，包含日常统计、历史区间统计、外力破坏剔除、分片上传、结果下载和 Aliyun/NAS 上传网关。

## 主要模块

| 路径 | 说明 |
| --- | --- |
| `web/backend/main.py` | FastAPI 服务入口和 API 路由 |
| `web/backend/job_manager.py` | 后台任务、日志流、任务状态和多页面任务管理 |
| `web/backend/historical_utils.py` | 历史区间统计配置与数据选择 |
| `web/backend/exclude_utils.py` | 表 11 重大事件日文件管理 |
| `web/backend/log_parser.py` | 统计脚本输出解析 |
| `web/backend/database.py` | PostgreSQL 批次、版本、差异预览、激活与回滚 |
| `web/backend/migrate_history.py` | 历史文件清点、哈希与顺序回放工具 |
| `web/backend/reconcile_snapshot.py` | 数据库有效数据与历史快照无序行哈希比对 |
| `web/frontend/` | 日常统计、历史区间统计、外力破坏页面 |
| `统计材料/运行脚本/process_outage_tables_v1.0.8.py` | 日常统计核心算法 |
| `统计材料/运行脚本/report_io.py` | 常量内存、原子发布的 XLSX 写出器 |
| `统计材料/运行脚本/postgres_source.py` | 把数据库版本载入现有 pandas 规则 |
| `统计材料/运行脚本/process_outage_tables_historical_v1.0.0.py` | 历史区间统计 |
| `统计材料/运行脚本/exclude_external_damage_v1.0.0.py` | 外力破坏剔除 |
| `web/deploy/aliyun/upload_gateway.py` | Aliyun 本地上传接收及 NAS 拉取网关 |

## 功能流程

1. 用户在页面上传 Excel；大文件通过 4 MB 分片上传，避免单次请求阻塞代理。
2. 服务按页面将数据写入 `input/Newdata/daily`、`historical` 或 `exclude`。
3. 日常统计脚本读取最新数据、历史结果、表 11 剔除配置和国家能源局台账。
4. 算法进行窗口替换、数据清洗、去重、停电统计、频繁停电/预警识别和线路累计状态合并。
5. 服务通过 SSE 推送脚本日志，最终提供处理结果、发出版、统计表及摘要下载。

## 日常统计规则概要

- 剔除重大事件日、停电时长小于 5 分钟、单用户工单和同一用户/馈线 6 小时内重复记录。
- 统计粒度为“用户编码 + 所属馈线编码”。
- 频繁停电：年内超过 5 次、连续 60 天超过 3 次、年内预安排停电超过 3 次。
- 停电预警：年内 4–5 次、近 50 天 3 次、近 30 天 2 次。
- 线路累计预警保留历史线路；状态变化或新线路使用当前数据截止日更新预警时间。

## 运行环境

- Python 3.11+
- FastAPI / Uvicorn
- pandas / openpyxl
- 依赖见 `web/backend/requirements.txt`

### 本地启动

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r web/backend/requirements.txt
$env:STATS_BASE = "$PWD\统计材料"
uvicorn main:app --app-dir web/backend --host 127.0.0.1 --port 22329
```

访问 <http://127.0.0.1:22329/>。

运行数据不提交到仓库，应由部署环境挂载或单独准备：

```text
统计材料/
├── input/Newdata/daily/
├── input/Newdata/historical/
├── input/Newdata/exclude/
├── input/exclude/2025/
├── input/exclude/2026/
├── output/
├── 国家能源局台账.xlsx
├── 统计表（模板）.xlsx
└── input/Annual Summary/2025/2025年停电用户数据.xlsx
```

### Docker / Nginx

NAS Docker 示例在 `web/deploy/nas/outage-stats/`。Aliyun、Nginx 和 FRP 示例中的路径、域名、证书和端口均需按部署环境调整，敏感值通过环境变量或本机配置注入。

## API 摘要

- `GET /api/health`：健康状态和脚本版本
- `GET /api/exclude/years`：表 11 配置年份和最新文件
- `POST /api/upload`：上传数据文件
- `POST /api/upload/chunk`：分片上传
- `GET /api/imports/{batch_id}/preview`：查看纠错批次差异和缺失日期
- `POST /api/imports/{batch_id}/activate`：确认日期范围并事务激活
- `POST /api/imports/{batch_id}/reject`：拒绝待确认批次
- `POST /api/imports/{batch_id}/rollback`：回滚最近一次激活
- `POST /api/run`：启动日常统计
- `POST /api/historical/run`：启动历史区间统计
- `POST /api/exclude/run`：启动外力破坏剔除
- `GET /api/jobs/{job_id}/stream`：SSE 日志流
- `GET /api/download/{job_id}/{file_key}`：下载结果文件或全部 ZIP

## 性能与可靠性改造

- 两个大结果工作簿使用 XlsxWriter `constant_memory` 逐行写入，不再建立完整 openpyxl 单元格树，也不再写后逐格格式化。
- 标识符在写入时规范化，文本列和整数列按整列设置格式；列宽只扫描前 2,000 行。
- 每个结果先写入同目录临时文件，成功关闭后原子替换；异常不会暴露半成品。
- 全部下载先生成磁盘临时 ZIP，再由文件响应发送；XLSX 不重复压缩，响应结束自动清理。
- 日志包含每个 Sheet 和文件的行数、列数、耗时、大小及可用平台上的峰值内存。

合成基准 `tests/benchmark_export.py` 在 254,683 行 × 48 列下完成单个工作簿写出约 63 秒（机器与数据压缩率会影响结果）。最终生产验收仍以真实三文件全流程不超过 30 分钟、峰值内存低于 3 GB 为准。

## PostgreSQL 批次与纠错工作流

`DATA_BACKEND` 支持三种模式：

- `excel`：完全沿用文件历史，数据库代码不参与。
- `shadow`：上传时同步建立数据库待确认批次，但统计仍读 Excel；用于迁移和影子比对。
- `postgres`：统计读取数据库当前有效版本；可在 `/api/run` 请求体传 `activation_id` 重跑历史版本。

上传批次先校验必需字段、日期和重复记录键。记录键优先使用“停电记录id”，其次“事件id + 用户id”，最后使用“工单号 + 用户编码 + 馈线编码 + 去重后停电开始时间”。批次默认是 `pending_confirmation`，页面展示增删改、每日行数、缺失日期和地市影响；确认后才在单个事务中替换指定日期窗口。最近一次激活可直接回滚，更早历史通过新纠错批次恢复。

数据库容器配置位于 `web/deploy/postgres/`：仅监听 `127.0.0.1:25432`，限制 768 MB 内存、2 核 CPU、20 个连接，数据目录为 `/mnt/data-disk/outage-stats/postgres-data`。实际 `.env` 与 `DATABASE_URL` 必须只保存在服务器，不能提交 Git。

历史迁移先清点并哈希，再按文件名导出时间回放：

```bash
cd /mnt/data-disk/outage-stats/web/backend
python migrate_history.py \
  --stats-base /mnt/data-disk/outage-stats/统计材料 \
  --report /mnt/data-disk/outage-stats/archive/migration-inventory.json

# 人工处理报告中的 ambiguous_inputs 后才执行：
python migrate_history.py \
  --stats-base /mnt/data-disk/outage-stats/统计材料 \
  --report /mnt/data-disk/outage-stats/archive/migration-replay.json \
  --apply
```

使用 `reconcile_snapshot.py` 将每个可对应节点与历史快照做无序行哈希比对。连续 7 次生产运行对比通过前，不应把 `DATA_BACKEND` 从 `shadow` 切到 `postgres`。

数据库备份和 90 天原始文件归档脚本位于同一目录；示例定时任务每天执行，并把磁盘 70%/80%/90% 记录为 NOTICE/WARNING/CRITICAL。当前归档与数据库仍在同一块阿里云磁盘，不属于异地容灾。

## 安全说明

仓库只保存代码和部署模板，不保存业务 Excel、输出结果、上传分片、SSH 密钥、Basic Auth 密码、FRP token 或 NAS 密码。部署前请通过环境变量或服务器本地配置注入凭据，并定期轮换现有服务器和 FRP 凭据。

## 当前版本信息

- 日常算法：v1.0.8
- 历史算法：v1.0.0
- 外力破坏剔除：v1.0.0
- 前端资源版本：以 `web/frontend/ASSET_VERSION` 为准
