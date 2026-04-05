# docx-text-replacer

![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python&logoColor=white)
![python-docx](https://img.shields.io/badge/python--docx-1.1.0+-green)
![License](https://img.shields.io/badge/license-MIT-lightgrey)

一個簡單的命令列工具，用於在 `.docx`（Microsoft Word）文件中進行文字替換，並完整保留原始格式。

---

## 功能特色

- 支援段落與表格內的全文替換
- 保留原始格式（粗體、字型、顏色等）
- 支援中文、英文及 Unicode 字元
- 不覆蓋原始檔案，自動產生新的 `_modified.docx`

---

## 安裝依賴

```bash
pip install python-docx
```

---

## 使用方式

```bash
python main.py <輸入檔案> <目標字串> <替換字串>
```

### 範例

```bash
python main.py report.docx 的 之
```

```
Replaced 5 occurrence(s) of '的' with '之'
Output saved to: report_modified.docx
```

輸出檔案會自動存為 `<原檔名>_modified.docx`，儲存於相同目錄下。

---

## 專案結構

```
docx-text-replacer/
├── main.py                 # CLI 入口
├── replacer/
│   ├── __init__.py
│   └── docx_replacer.py    # 核心替換邏輯
├── tests/
│   └── test_basic.py       # 單元測試
├── requirements.txt
└── spec.md
```

---

## 執行測試

```bash
pip install pytest python-docx
python -m pytest tests/ -v
```

---

## 已知限制

- 若目標字串因格式不同被拆成多個 run，可能無法被替換
- 目前不支援頁首與頁尾的文字替換
