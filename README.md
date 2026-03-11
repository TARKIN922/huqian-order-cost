# 📊 订单处理平台

基于 Flask 的亚马逊多渠道订单自动处理平台，支持多用户并发使用，一键生成多渠道订单表与费用表。

---

## ✨ 功能特性

- **多渠道订单表生成** — 自动解析后台订单 CSV，按月份匹配货件表，分类整理多渠道订单与网红订单，并生成透析汇总
- **费用表生成** — 读取 FBA 发货记录与产品成本，自动计算每票产品成本，按月份分组输出带格式的发货成本表
- **多渠道订单费用明细** — 汇总多渠道与网红 FBA 配送费，追加写入费用表
- **多用户隔离** — 每次处理任务拥有独立工作目录，多人同时使用互不干扰
- **一键打包下载** — 结果文件打包为 ZIP 下载，下载完成后服务器自动清理临时文件
- **支持德文/英文表头** — 自动识别并转换德文亚马逊报表表头

---

## 📁 输入文件说明

上传时需要提供三类文件：

| 类型 | 说明 | 格式 |
|------|------|------|
| 后台订单原表 | 亚马逊后台导出的月度交易报表，文件名需含月份标识，如 `JanMonthlyTransaction...` | CSV |
| 后台货件表 | 与订单对应的货件信息表，文件名含月份，如 `2月货件.csv` | CSV / Excel |
| 成本文件 | 产品成本表（文件名含"产品成本"）+ FBA 发货表（文件名以 `fba` 开头） | Excel |

---

## 📦 输出文件说明

下载的 ZIP 包含以下文件夹：

```
result.zip
├── 后台订单原表/        ← 原始上传文件
├── 后台货件表/          ← 原始上传文件
├── 成本/               ← 原始上传文件
├── 多渠道订单表/        ← 生成：各月多渠道/网红订单 + 透析表
└── 费用表/             ← 生成：发货成本 + 多渠道订单费用明细
```

---

## 🚀 快速开始（Docker，推荐）

### 前提

安装 [Docker Desktop](https://www.docker.com/products/docker-desktop/)，启动后任务栏鲸鱼图标变绿即可。

### 启动

```bash
git clone https://github.com/你的用户名/order-processor.git
cd order-processor
docker compose up --build -d
```

浏览器访问 → **http://localhost:5000**

### 停止

```bash
docker compose down
```

### 更新代码后重新构建

```bash
docker compose up --build -d
```

> 国内网络拉取镜像失败时，将 `Dockerfile` 第一行改为：
> ```dockerfile
> FROM swr.cn-north-4.myhuaweicloud.com/ddn-k8s/docker.io/python:3.11-slim
> ```

---

## 🛠 本地运行（不用 Docker）

```bash
# 安装依赖
pip install -r requirements.txt

# 启动
python app.py
```

浏览器访问 → **http://localhost:5000**

---

## 🗂 项目结构

```
order-processor/
├── app.py                  # 后端主程序（Flask）
├── requirements.txt        # Python 依赖
├── Dockerfile              # Docker 镜像配置
├── docker-compose.yml      # 一键启动配置
├── .dockerignore
├── .gitignore
└── templates/
    └── index.html          # 前端页面
```

---

## 🔌 API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| `POST` | `/api/upload` | 上传文件（type: orders / shipments / cost） |
| `POST` | `/api/process` | 启动处理任务 |
| `GET` | `/api/task/{id}` | 查询任务进度 |
| `GET` | `/api/download/{id}` | 下载结果 ZIP（下载后自动清理） |
| `GET` | `/api/files/{id}` | 查看已上传文件列表 |

---

## 🧰 技术栈

- **后端** — Python 3.11 / Flask / Pandas / openpyxl
- **前端** — 原生 HTML / CSS / JavaScript
- **部署** — Docker / Gunicorn

---

## 📄 License

MIT
