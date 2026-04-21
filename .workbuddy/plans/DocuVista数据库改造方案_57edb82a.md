---
name: DocuVista数据库改造方案
overview: 为DocuVista设计分阶段数据库改造方案，支持2000+银行数据管理、版本控制、数据溯源，为二期商业化会员系统预留扩展空间。
todos:
  - id: design-bank-schema
    content: 设计银行数据仓库数据库表结构
    status: completed
  - id: create-migration-script
    content: 创建数据库迁移脚本
    status: completed
    dependencies:
      - design-bank-schema
  - id: modify-table-worker
    content: 修改table_worker.py支持数据库写入
    status: completed
    dependencies:
      - design-bank-schema
  - id: create-data-service
    content: 创建银行数据服务层
    status: completed
    dependencies:
      - create-migration-script
  - id: create-data-api
    content: 创建银行数据查询API
    status: completed
    dependencies:
      - create-data-service
  - id: test-data-flow
    content: 测试数据流转（PDF→数据库→前端）
    status: completed
    dependencies:
      - create-data-api
---

## 项目背景

DocuVista是一个银行PDF数据处理系统，当前为一期开发（内部使用），未来将商业化给会员使用。

## 核心需求

1. **数据规模**：每年处理2000个银行的PDF数据
2. **数据展示**：在线浏览 + 多银行横向对比分析
3. **版本管理**：支持数据历史版本保留与回滚
4. **数据溯源**：每条数据可追溯到原始PDF文件名和页码
5. **商业化扩展**：二期会员权限管理、付费功能

## 现状分析

- **数据存储**：Excel文件存储在 `data/backend/static/excel_data/{pdf_id}/xxx_合并.xlsx`
- **数据库**：SQLite仅存储元数据（files表、texts表、table_processing_records表）
- **问题**：无法支撑2000+银行的多维度查询、横向对比、版本管理

## 改造目标

将系统从"文件驱动"改造为"数据库驱动"的数据仓库，支持结构化存储、灵活查询、版本管理和数据溯源。

## 技术选型

### 数据库选择

- **当前**：SQLite（已有）
- **推荐**：保持SQLite用于一期内部使用，迁移到云端数据库（如MySQL/PostgreSQL）用于商业化二期
- **架构设计**：预留ORM层（SQLAlchemy），支持多数据库切换

### 技术栈

- **后端**：Python + Flask（已有）
- **数据库**：SQLite（当前）→ MySQL/PostgreSQL（未来）
- **ORM**：SQLAlchemy（支持多数据库）
- **数据迁移**：Alembic版本管理

## 数据库设计

### 核心表结构

#### 1. 银行基础信息表 (banks)

```sql
CREATE TABLE banks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bank_code VARCHAR(20) UNIQUE,      -- 银行统一代码
    bank_name VARCHAR(200),              -- 银行名称
    bank_type VARCHAR(50),               -- 国有/股份制/城商/农商等
    description TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);
```

#### 2. 报告记录表 (reports)

```sql
CREATE TABLE reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bank_id INTEGER NOT NULL,
    report_type VARCHAR(20),            -- 年报/季报/半年报
    period VARCHAR(20),                  -- 2024A, 2025Q1
    report_date DATE,                   -- 报告发布日期
    pdf_filename VARCHAR(500),          -- 原始PDF文件名
    pdf_path VARCHAR(1000),             -- PDF存储路径
    pdf_hash VARCHAR(64),               -- PDF文件哈希（防篡改）
    status VARCHAR(20),                 -- pending/processing/completed/failed
    excel_output_path VARCHAR(1000),    -- 生成的Excel路径
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (bank_id) REFERENCES banks(id)
);
```

#### 3. 原始表格数据表 (table_data)

```sql
CREATE TABLE table_data (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    report_id INTEGER NOT NULL,
    table_name VARCHAR(200),            -- 表格名称（利润表/资产负债表等）
    page_number INTEGER,                -- 原始PDF页码
    row_index INTEGER,                  -- 行号
    indicator_name VARCHAR(500),         -- 指标名称
    -- 年份数据列（动态扩展）
    value_2020 DECIMAL(20,4),
    value_2021 DECIMAL(20,4),
    value_2022 DECIMAL(20,4),
    value_2023 DECIMAL(20,4),
    value_2024 DECIMAL(20,4),
    value_2025 DECIMAL(20,4),
    -- 数据状态
    is_adjusted BOOLEAN DEFAULT 0,      -- 是否经过人工调整
    adjusted_value DECIMAL(20,4),       -- 调整后值
    notes TEXT,                         -- 备注说明
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (report_id) REFERENCES reports(id)
);
```

#### 4. 数据溯源表 (data_sources)

```sql
CREATE TABLE data_sources (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    table_data_id INTEGER NOT NULL,
    pdf_path VARCHAR(1000),             -- 原始PDF路径
    page_number INTEGER,                 -- 页码
    image_path VARCHAR(1000),            -- 对应图片路径
    ocr_cache_path VARCHAR(1000),       -- OCR缓存路径
    llm_cache_path VARCHAR(1000),        -- LLM缓存路径
    llm_response TEXT,                  -- LLM原始响应
    confidence_score DECIMAL(5,4),      -- LLM识别置信度
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (table_data_id) REFERENCES table_data(id)
);
```

#### 5. 数据版本历史表 (data_versions)

```sql
CREATE TABLE data_versions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    table_data_id INTEGER NOT NULL,
    version INTEGER NOT NULL,           -- 版本号
    old_value TEXT,                     -- 修改前的值
    new_value TEXT,                     -- 修改后的值
    change_type VARCHAR(20),            -- manual_edit/auto_correct/import
    changed_by VARCHAR(100),             -- 修改人
    change_reason TEXT,                 -- 修改原因
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (table_data_id) REFERENCES table_data(id)
);
```

#### 6. 任务处理记录表 (processing_jobs) - 扩展现有表

```sql
CREATE TABLE processing_jobs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id VARCHAR(100) UNIQUE,         -- 任务唯一ID
    report_id INTEGER,
    bank_name VARCHAR(200),
    status VARCHAR(20),                 -- pending/queued/processing/completed/failed
    stage VARCHAR(50),                   -- 当前阶段
    progress INTEGER,                   -- 进度百分比
    total_images INTEGER,
    processed_images INTEGER,
    success_count INTEGER DEFAULT 0,
    failed_count INTEGER DEFAULT 0,
    error_message TEXT,
    start_time DATETIME,
    end_time DATETIME,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (report_id) REFERENCES reports(id)
);
```

#### 7. 会员表 (members) - 二期预留

```sql
CREATE TABLE members (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username VARCHAR(100) UNIQUE,
    password_hash VARCHAR(255),
    member_level VARCHAR(20),           -- free/premium/enterprise
    allowed_banks TEXT,                -- JSON数组，可访问的银行ID列表
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);
```

## 目录结构

```
backend/
├── models/
│   ├── unified_db.py              # [MODIFY] 扩展现有数据库管理器
│   └── bank_warehouse.py          # [NEW] 银行数据仓库模型
├── database/
│   ├── migrations/                # [NEW] Alembic迁移脚本
│   │   ├── versions/
│   │   │   └── 001_initial_schema.py
│   │   └── env.py
│   └── bank_schema.py             # [NEW] 银行数据仓库表定义
├── services/
│   └── bank_data_service.py       # [NEW] 银行数据服务层
├── api/
│   └── bank_data_api.py           # [NEW] 银行数据API
└── workers/
    └── table_worker.py             # [MODIFY] 修改为写入数据库
```

## 实施策略

### 一期（当前）- 内部使用

1. 在SQLite中创建新表
2. 修改table_worker.py，数据同时写入数据库
3. 保留Excel文件作为备份/导出
4. 开发基础查询API

### 二期（商业化）

1. 迁移到MySQL/PostgreSQL
2. 实现会员权限管理
3. 数据分析功能
4. Excel导出优化