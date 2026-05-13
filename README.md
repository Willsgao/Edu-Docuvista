# Edu-DocuVista · 教育文档智能解析引擎

> 从试卷、答题卡、作文纸等教育类 PDF 中自动识别题目、答案、批注区域，提取结构化数据，配合 LLM 生成批改评语与成绩分析。

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue)](https://www.python.org/)
[![Vue 3](https://img.shields.io/badge/Vue-3-green)](https://vuejs.org/)
[![Flask](https://img.shields.io/badge/Flask-3.1.1-gray)](https://flask.palletsprojects.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## 能做什么

| 输入 | 输出 |
|------|------|
| 标准化试卷 PDF | 题目 / 选项 / 答题区结构化数据 |
| 作文扫描件 | 全文文字 + 书写质量评估 |
| 答题卡图片 | 客观题自动判分、得分统计 |
| 批注卷（红笔） | 区分原卷内容与教师批注 |

**典型场景**：教师批量上传 → 系统自动切题、提取答案 → LLM 生成评语 → Excel 导出 / 在线查看批改结果。

---

## 技术架构

```
PDF / 图片
    │
    ▼
┌─────────────────────────────┐
│  GPU Layout Engine          │  ← PaddleOCR Layout 布局检测
│  识别题目区 / 答题区 / 装订线 │
└──────────────┬──────────────┘
               │
    ┌──────────┴──────────────┐
    ▼                         ▼
┌─────────────┐        ┌─────────────┐
│  题目裁剪    │        │  表格重构    │
│  (逐题切分)  │        │  (答题卡/表格式) │
└──────┬──────┘        └──────┬──────┘
       │                       │
       └──────────┬────────────┘
                  ▼
┌─────────────────────────────────┐
│  LLM 语义提取                    │
│  题目解析 + 答案抽取 + 评语生成    │
└──────────────┬──────────────────┘
               │
               ▼
        Excel / JSON / 可视化
```

---

## 目录结构

```
Edu-DocuVista/
├── backend/
│   ├── api/              # Flask REST API 蓝图
│   │   ├── upload.py     # 文件上传
│   │   ├── file.py       # 文件管理
│   │   ├── convert.py    # PDF 转换
│   │   ├── llm_routes.py # LLM 接口
│   │   └── websocket_routes.py  # 实时进度推送
│   ├── pipeline/         # 5 步解析管线
│   │   ├── pdf2png       # PDF → 图片
│   │   ├── gpu_layout    # GPU 布局检测
│   │   ├── crop_table    # 题目区域裁剪
│   │   ├── rebuild       # 表格 / 答题卡重构
│   │   └── llm_extract   # LLM 语义提取
│   ├── llm_services/     # LLM 调用封装（批量/单条/状态管理）
│   ├── models/           # SQLAlchemy 数据模型
│   ├── service/          # 业务服务层
│   └── app.py            # Flask 主入口
├── frontend/
│   ├── src/              # Vue 3 组件
│   └── public/          # 静态资源
├── data/                 # 上传文件 + 中间结果
├── test_codes/           # 离线测试脚本
├── test_data/            # 测试用 PDF 样本
└── docs/
    └── screenshots/     # 界面截图
```

---

## 核心技术细节

### 1. 5 步解析管线

```
Step 1: pdf2png        → PyMuPDF 高清渲染 (DPI=150)
Step 2: gpu_layout     → GPU 服务器远程调用 PaddleOCR Layout
Step 3: crop_table     → 根据布局坐标逐区域裁剪
Step 4: rebuild        → 合并相邻答题区 / 重构表格结构
Step 5: llm_extract    → LLM 解析题目、提取答案、生成评语
```

每步独立运行、结果缓存，任意步骤失败可从断点重试。

### 2. WebSocket 实时推送

后端处理全程通过 WebSocket 向前端推送进度：

```json
{ "step": "gpu_layout", "progress": 67, "message": "正在识别第 3/10 页布局" }
```

### 3. 多级缓存

| 数据类型 | 缓存策略 | 失效条件 |
|----------|----------|----------|
| PNG 渲染图 | 磁盘文件 | PDF 文件哈希变化 |
| Layout 检测结果 | SQLite | 同路径 PDF 重新上传 |
| LLM 解析结果 | SQLite (JSON 字段) | 手动清除 / 重新解析 |

### 4. LLM 批量处理

支持单文档和多文档批量解析，内部自动：

- 请求限流 + 指数退避重试（2s → 4s → 8s）
- 4 种 JSON 响应格式自动兼容（数组 / 对象 / 嵌套 / 深度搜索）
- 任务队列状态机：pending → running → success / failed

### 5. 批注卷区分（待扩展）

通过颜色分析区分：
- **蓝色/黑色** → 学生原卷内容
- **红色** → 教师批注 / 得分

---

## 技术栈

| 层 | 技术 | 说明 |
|----|------|------|
| 前端 | Vue 3 + Element Plus | 响应式教育管理界面 |
| 后端 | Python 3 + Flask 3 + SQLAlchemy | REST API + WebSocket |
| PDF 解析 | PyMuPDF + PyPDF2 | 文本型/扫描件双模式 |
| 布局检测 | PaddleOCR Layout (GPU) | 题目/答题区/装订线识别 |
| LLM | 火山引擎 / OpenAI 兼容接口 | 语义提取 + 评语生成 |
| 实时通信 | Flask-SocketIO | 任务进度推送 |
| 数据库 | SQLite | 结构化数据持久化 |

---

## 快速启动

```bash
# 1. 后端
cd backend
pip install -r requirements.txt
python app.py

# 2. 前端（新窗口）
cd frontend
npm install
npm run serve

# 3. 访问
http://localhost:8080
```

**首次使用**：上传一份标准试卷 PDF，系统自动完成全流程解析。

---

## 示例数据

`test_data/` 目录包含匿名化示例 PDF，可直接用于功能演示，无需额外准备数据。

---

## License

MIT License · 高玉伟
