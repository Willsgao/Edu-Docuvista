# Edu-DocuVista

> 财务报告 PDF 表格提取工具，基于 PyQt5 GUI 实现，支持文本型与图片型 PDF 双模式解析，提供表格对比、历史记录与 Excel 导出功能。

---

## 🎯 核心功能

- **文本型 PDF 解析** — 基于 PyMuPDF (fitz) 坐标定位，无外部依赖
- **图片型 PDF 解析** — 集成豆包 (Doubao) 视觉识别 API，精准提取扫描件表格
- **表格对比** — 两份 PDF 同页面表格并排对比，高亮差异区域
- **历史记录** — 完整保留每次解析结果，支持回溯与筛选
- **页级预览** — 分页渲染预览图，可按页筛选处理范围
- **Excel 导出** — 解析结果一键导出为结构化 Excel 数据
- **优雅降级** — 无 API Key 时自动降级，缓存预览图并生成空 Excel，不报错退出

---

## 🏗️ 技术架构

```
文本型 PDF
    │
    ▼
┌──────────────────────┐
│  PyMuPDF (fitz)      │  PDF 渲染 + 坐标处理 + 文字/线条/色块提取
│  word级坐标定位       │
└──────────────────────┘
    │
    ▼
┌──────────────────────┐
│  表格结构解析         │  列边界检测 + 行阈值分离 + 单元格合并
│  (coords-based)       │
└──────────────────────┘

图片型 PDF
    │
    ▼
┌──────────────────────┐
│  豆包 API 视觉识别     │  表格区域检测 + OCR + 结构化输出
│  Doubao Vision API    │
└──────────────────────┘
    │
    ▼
┌──────────────────────┐
│  解析结果统一输出      │
│  → Excel / 预览图     │
└──────────────────────┘
```

---

## 📂 项目结构

```
Edu-DocuVista/
├── backend/                 # 后端核心
│   ├── app.py              # Flask API 入口
│   ├── config/             # 配置文件
│   │   └── settings.json   # API Key 与路径配置
│   ├── processor.py        # 文本型 PDF 解析器
│   ├── processing_manager.py  # 处理状态管理器
│   ├── ui/                 # PyQt5 UI 层
│   │   ├── main_window.py     # 主窗口
│   │   └── table_compare_manager.py  # 表格对比 UI
│   └── utils/              # 工具函数
├── backend_run.py          # 后端启动脚本
├── data/                   # 数据存储
│   └── mid_cache/          # 中间缓存（预览图、解析结果）
└── frontend/               # Vue 前端（可选）
```

---

## 🚀 快速开始

### 环境要求

- Python >= 3.9
- PyQt5
- PyMuPDF (`fitz`)
- Redis（可选，用于缓存加速）

### 安装依赖

```bash
cd backend
pip install -r requirements.txt
```

### 配置文件

编辑 `backend/config/settings.json`，填入豆包 API Key：

```json
{
    "doubao_api_key": "your_api_key_here",
    "cache_dir": "data/mid_cache",
    "output_dir": "data/output"
}
```

> 无 API Key 时，系统自动降级：缓存 PDF 预览图并生成空 Excel，不报错退出。

### 启动后端

```bash
python backend_run.py
```

### 启动前端（可选）

```bash
cd frontend
npm install
npm run serve
```

---

## 📊 核心模块说明

### processor.py — 文本型 PDF 解析

基于 PyMuPDF word 级坐标实现表格提取：

1. 提取页面所有 word 的 `(x0, y0, x1, y1, text)` 坐标
2. 收集全页 y 坐标分布，统计最小典型间距
3. 取 **中位间距的一半** 作为动态行阈值（替代硬编码 5pt）
4. 基于填充色块 + 线条信号辅助列边界检测
5. 重叠面积单元格分配处理复杂合并单元格

### processing_manager.py — 状态管理

管理 PDF 切换时的状态清理，防止跨文档数据污染：

- `_save_previous_page_data()` — 切换 PDF 前保存当前页数据
- 状态重置与缓存清理

### 豆包 API 集成 — 图片型 PDF

```python
# 无 API Key 优雅降级
if not api_key:
    # 缓存页面拆分图 → data/mid_cache/
    # 生成空 Excel → 不报错退出
    return {"status": "degraded", "output": None}
```

---

## ⚙️ 配置说明

| 配置项 | 说明 |
|--------|------|
| `doubao_api_key` | 豆包视觉识别 API Key（可选） |
| `cache_dir` | 预览图与中间数据缓存目录 |
| `output_dir` | Excel 输出目录 |
| `row_threshold_mode` | `dynamic`（推荐）/ `fixed`（硬编码 5pt） |

---

## 📝 License

MIT License
