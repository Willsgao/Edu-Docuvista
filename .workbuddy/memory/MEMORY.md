# 工作记忆

## 2026-04-05 银行数据仓库改造

### 新增模块
- `backend/database/bank_warehouse/` - 银行数据仓库模块
  - `bank_schema.py` - 表结构定义（banks, reports, table_data, data_sources, data_versions, processing_jobs, members）
  - `bank_warehouse.py` - 数据仓库管理器（CRUD操作）
- `backend/tests/harness/test_bank_warehouse.py` - Harness测试套件（16个测试全部通过）

### 表结构
- `banks` - 银行基础信息表（支持预留国际字段：swift_code, isin, country_code）
- `reports` - 报告记录表
- `table_data` - 核心数据表（使用JSON存储多年数据）
- `data_sources` - 数据溯源表
- `data_versions` - 数据版本历史表
- `processing_jobs` - 处理任务记录表
- `members` - 会员表（二期预留）

### 使用方法
```python
from backend.database.bank_warehouse import BankWarehouseManager

warehouse = BankWarehouseManager()
warehouse.init_database()  # 初始化表

# 保存银行
bank_id = warehouse.save_bank({
    'bank_code': 'ICBC',
    'bank_name': '中国工商银行',
    'bank_type': '国有大型银行'
})

# 保存报告
report_id = warehouse.save_report({
    'bank_id': bank_id,
    'report_type': 'annual',
    'period': '2024A'
})

# 批量保存表格数据
warehouse.save_batch_table_data(
    report_id=report_id,
    table_name='利润表',
    rows=[...]
)
```

### table_worker.py 集成
- 添加了 `ENABLE_BANK_WAREHOUSE` Feature Flag
- 在任务完成后自动保存到数据仓库（阶段7）
- 支持从文件名解析银行信息和报告期间
- 支持从Excel读取表格数据

### 启用方式
```bash
# Linux/Mac
export ENABLE_BANK_WAREHOUSE=true

# Windows
set ENABLE_BANK_WAREHOUSE=true

# 然后启动Worker
python backend/workers/table_worker.py
```

---

## 2026-04-05 修复记录

### 问题：重新解析功能失效
**现象**：点击重新解析后，82张图片被跳过，使用旧的 Excel 文件而不是重新处理。

**根本原因（第一层）**：
- `check_existing_table_task` 函数只检查 Redis 中的任务记录
- Redis 数据有 TTL，过期后任务记录消失
- 但 Excel 文件还在文件夹里
- 前端检查时发现 `has_existing=False` → `shouldRerun=False` → 不清除旧数据

**修复（第一层）**：
- 修改 `backend/api/convert/table_processor.py` 中的 `check_existing_table_task` 函数
- 同时检查 Redis 和 `EXCEL_DATA_DIR` 文件夹
- 即使 Redis 记录过期，只要 Excel 文件存在就返回 `has_existing=True`

---

**根本原因（第二层）**：
- `_clear_processed_images` 只清除了 `processing_records.json`
- 但 `is_image_processed()` **不依赖这个文件**
- 它直接检查 **LLM 缓存文件** (`obj_cache/{pdf_folder}/llm/*.json.gz`) 是否存在
- 即使清除了记录，LLM 缓存还在，图片仍然被跳过

**修复（第二层）**：
- 修改 `backend/core/incremental_processor/simple_incremental_processor.py` 中的 `clear_pdf_records` 函数
- 同时清除 LLM 缓存目录 (`obj_cache/{pdf_folder}/llm/`)

---

**关键文件**：
- `backend/api/convert/table_processor.py` - `check_existing_table_task` (第1495行)
- `backend/workers/table_worker.py` - `_clear_processed_images` (第2003行)
- `backend/core/incremental_processor/simple_incremental_processor.py` - `clear_pdf_records` (第261行)

**线上 vs 本地差异**：
- 本地正常：处理任务少，Redis TTL 还没过期，LLM 缓存也不多
- 线上异常：任务多，Redis 数据过期，LLM 缓存大量存在

**部署步骤**：
1. 提交代码到线上
2. 重启后端服务（Worker 是常驻进程，必须重启）
3. 清除问题 PDF 的 LLM 缓存目录（如果还有问题）

---

## 2026-04-05 晚间修复

### 修复内容
1. **bank_warehouse/__init__.py** - 添加 `BankWarehouseManager` 导出
2. **bank_warehouse.py** - 替换所有 emoji 字符为 ASCII 避免 Windows GBK 编码错误
3. **test_bank_warehouse.py** - 替换所有 emoji 字符，修复 `test_04_search_banks` 测试逻辑

### 测试结果
- 16 个 Harness 测试全部通过
- 关键模块导入正常

### Git 状态
- 修改文件：`app_factory.py`, `database/__init__.py`, `table_worker.py`
- 新增文件：`bank_data_api.py`, `bank_warehouse/`, `bank_data_service.py`, `tests/`
- 分支：`refactor-branch`（领先 origin 17 commits）
