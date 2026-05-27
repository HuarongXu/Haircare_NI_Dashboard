# NI Dashboard Wave2 工具说明

## 1. 工具结构

- `app_config.py`: 统一路径配置文件（后续改路径只改这里）
- `extract_ni_code_cn.py`: 提取 NI Code 文件，生成 CN/HKTW/WIP/code list
- `sap_m_ld.py`: 执行 SAP M_LD 并导出报表
- `sap_mm60.py`: 执行 SAP MM60 并导出报表
- `sap_vb03.py`: 执行 SAP VB03 并导出报表（自动读取 code list 中最大 SOS Date 作为查询截止日期）
- `run_sap_reports.py`: 一键按顺序执行 3 个 SAP 脚本

## 2. 路径统一配置

所有路径放在 `app_config.py`:

- `NI_SOURCE_DIR`: NI Code List 源目录（OneDrive）
- `OUTPUT_DIR`: 输出目录（默认当前项目目录）
- `CODE_LIST_FILENAME`: code list 文件名
- `CODE_LIST_FILE`: code list 完整路径
- `SAP_SAVE_PATH`: SAP 导出保存目录

后续如需修改路径，仅编辑 `app_config.py`，无需改其他脚本。

## 3. 推荐执行流程

1. 先执行 `extract_ni_code_cn.py`，生成最新 `code list.xlsx`
2. 打开 SAP GUI 并保持登录会话
3. 执行 `run_sap_reports.py`，自动顺序跑:
   - M_LD
   - MM60
   - VB03

## 4. 运行方式（VS Code）

- 打开目标脚本后直接点击右上角 Run
- 或在终端执行:

```powershell
python extract_ni_code_cn.py
python run_sap_reports.py
```

## 5. 注意事项

- 运行 SAP 脚本前，必须先打开 SAP GUI 并登录。
- SAP 界面元素 ID 如因系统版本变化失效，需要在对应脚本中更新 `findById(...)`。
- 若导出文件同名，脚本会先删除旧文件再保存新文件。

## 6. 常见问题

1. 报错找不到 `code list.xlsx`
- 先执行 `extract_ni_code_cn.py`
- 或检查 `app_config.py` 里的 `CODE_LIST_FILE` 是否正确

2. 报错 SAP 连接失败
- 确认 SAP GUI 已打开且至少有一个活动会话
- 确认本机已开启 SAP GUI Scripting

3. 仅某一个 SAP 报表失败
- `run_sap_reports.py` 会继续执行后续步骤
- 最后汇总会标记失败步骤，可单独运行该脚本排查
