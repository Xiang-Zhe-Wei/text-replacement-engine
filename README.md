# docx-text-replacer

![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python&logoColor=white)
![python-docx](https://img.shields.io/badge/python--docx-1.1.0+-green)
![License](https://img.shields.io/badge/license-MIT-lightgrey)

一個簡單的命令列工具，用於在 `.docx`（Microsoft Word）文件中進行文字替換，並完整保留原始格式。

支援兩種模式：
1. **批次模式**：自動套用 `replacements.md` 中定義的 105 條口語→正式用字規則
2. **單一模式**：從命令列指定單一替換對

---

## 功能特色

- 支援段落與表格內的全文替換
- 保留原始格式（粗體、字型、顏色等）
- 支援中文、英文及 Unicode 字元
- 不覆蓋原始檔案，自動產生新的 `_modified.docx`
- 批次模式顯示每條規則的替換次數統計

---

## 安裝依賴

```bash
pip install python-docx
```

---

## 使用方式

### 圖形介面（GUI）

適合沒有終端機的 Windows / Mac 一般用戶：

```bash
python gui.py
```

功能：
- 點選按鈕選擇 `.docx` 檔案
- 選擇替換模式：**批次**（套用 replacements.md 所有規則）或**單一**（手動輸入）
- 自訂輸出檔案名稱（留空則自動命名 `原檔名_modified.docx`）
- 執行後顯示每條規則的替換次數統計

---

### 批次模式（套用 replacements.md 所有規則）

```bash
python main.py <輸入檔案>
```

**範例：**

```bash
python main.py thesis.docx
```

```
  '的' → '之': 12 處
  '在' → '於': 5 處
  '由於' → '因': 3 處
  ...
共替換 87 處，輸出至：thesis_modified.docx
```

### 單一模式（指定單一替換對）

```bash
python main.py <輸入檔案> <目標字串> <替換字串>
```

**範例：**

```bash
python main.py report.docx 的 之
```

```
Replaced 5 occurrence(s) of '的' with '之'
Output saved to: report_modified.docx
```

輸出檔案會自動存為 `<原檔名>_modified.docx`，儲存於相同目錄下。

---

## 新增或修改替換規則

開啟 `replacements.md`，在表格末尾新增一行即可：

```
| 106 | 口語用字 | 修正用字 |
```

> 若修正用字有多個選項或含不定長度樣式（如 `有...問題`），請在修正用字欄加上 `*(需依文意選擇)*`，工具會自動跳過該條規則，須手動處理。

---

## 專案結構

```
docx-text-replacer/
├── gui.py                  # 圖形介面入口
├── main.py                 # CLI 入口
├── replacements.md         # 口語用字與修正用字對照表（可自行擴充）
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

## 技術說明

### 單次掃描替換（避免連鎖替換問題）

批次模式採用 **單次 regex 掃描**，而非逐條套用 `str.replace()`。

**問題背景：** 若逐條替換，較短的規則會對前一條規則的輸出再次作用。例如：

1. No.95 將「A、B和C」→「A、B**及**C」
2. No.12 再把「**及**」→「與」
3. 最終錯誤輸出「A、B**與**C」

**解法：** 將所有規則合併成一個 regex pattern，較長的目標字串排在前面優先匹配。掃描時每個位置只會被替換一次，已替換的內容不會再被後續規則處理。

### 跳過無法自動處理的規則

修正用字欄標有 `*(需依文意選擇)*` 的規則（含多個候選詞或 `...` 不定長度樣式）會在載入時自動跳過，需使用者手動替換。

---

## 已知限制

- 若目標字串因格式不同被拆成多個 run，可能無法被替換
- 目前不支援頁首與頁尾的文字替換
- 含 `...` 樣式的規則（如 `有...問題`）需手動替換
