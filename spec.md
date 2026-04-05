# 專案名稱：docx-text-replacer

## 一、專案概述（Overview）

本專案旨在開發一個簡單且可擴充的工具，用於處理 Microsoft Word (`.docx`) 文件中的文字替換。

使用者可以輸入一個 Word 檔案，指定欲替換的目標字串與替換後字串，程式將自動完成全文替換並輸出為新的檔案。

### 使用範例

* 輸入檔案：`test.docx`
* 替換內容：「的」→「之」
* 輸出檔案：`test_modified.docx`

---

## 二、專案目標（Goals）

1. 提供一個簡單易用的 Word 文字替換工具
2. 保留原始文件格式（字型、粗體、段落等）
3. 支援中英文及 Unicode 字元
4. 建立良好架構，方便未來擴充（如 GUI、批次處理）

---

## 三、技術選型（Tech Stack）

* 程式語言：Python 3.x
* 套件：`python-docx`

---

## 四、功能需求（Functional Requirements）

### 4.1 檔案輸入

* 接收 `.docx` 檔案路徑
* 檢查檔案是否存在
* 檢查副檔名是否為 `.docx`

### 4.2 文字替換

* 讀取 Word 文件內容
* 逐段落（paragraph）處理
* 逐 run（文字片段）進行替換
* 將所有符合的字串替換為指定字串

### 4.3 檔案輸出

* 不覆蓋原始檔案
* 自動產生新檔名：

  ```
  原檔名_modified.docx
  ```
* 確保輸出檔案可正常開啟

---

## 五、CLI 介面設計（Command Line Interface）

### 指令格式

```bash
python main.py <輸入檔案> <目標字串> <替換字串>
```

### 範例

```bash
python main.py test.docx 的 之
```

---

## 六、專案架構（Project Structure）

```
docx-text-replacer/
│
├── spec.md
├── main.py
├── replacer/
│   ├── __init__.py
│   └── docx_replacer.py
│
├── tests/
│   └── test_basic.py
│
├── requirements.txt
└── README.md
```

---

## 七、核心邏輯設計（Core Logic）

### 7.1 基本流程

1. 載入 `.docx` 文件
2. 取得所有段落（paragraphs）
3. 對每個段落：

   * 取得 runs
   * 在每個 run 中進行字串替換
4. 儲存新文件

---

### 7.2 偽程式碼（Pseudo Code）

```
讀取輸入參數（檔案路徑、target、replacement）

載入 docx 文件

for 每個段落 paragraph:
    for 每個 run:
        run.text = run.text.replace(target, replacement)

輸出新檔案
```

---

## 八、python-docx 注意事項

1. Word 文件中的文字不是一整段字串，而是由多個 `run` 組成
2. 不同格式（粗體、顏色）會導致文字被拆成不同 run
3. 若直接操作 paragraph.text 可能會破壞格式
4. 建議在 run 層級進行替換

---

## 九、邊界情況（Edge Cases）

1. 檔案不存在
2. 非 `.docx` 檔案
3. 替換字串不存在於文件中
4. 空文件
5. 中文 / Unicode 字元處理
6. 一個詞被拆成多個 run（進階問題）

---

## 十、測試需求（Testing）

應包含以下測試：

1. 基本替換（英文）
2. 中文替換（例如「的」→「之」）
3. 無匹配字串情況
4. 多次出現替換
5. 文件格式是否維持正常

---

## 十一、未來擴充（Future Work）

1. 批次處理多個檔案
2. 支援 Regex 替換
3. GUI 介面（Tkinter / PyQt）
4. 支援表格（table）內容替換
5. 支援頁首頁尾（header/footer）
6. Web 版本（Flask / FastAPI）

---

## 十二、成功條件（Success Criteria）

1. 所有目標字串成功被替換
2. 輸出文件格式未損壞
3. 程式執行穩定無錯誤
4. CLI 操作簡單直觀

---
