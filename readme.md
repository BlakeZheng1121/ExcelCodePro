# ExcelCode Pro

ExcelCode Pro 可將 Excel 資料轉換成 C/C++ 程式碼，支援圖形介面與命令列操作，特別適合批次產生程式碼或整理大量表格資料。

## 目錄

1. [功能概述](#功能概述)
2. [程式架構](#程式架構)
3. [安裝與執行](#安裝與執行)
4. [基本使用流程](#基本使用流程)
5. [範本系統](#範本系統)
   - [範本基礎](#範本基礎)
   - [範本標記語法](#範本標記語法)
   - [橫向和直向讀取模式](#橫向和直向讀取模式)
   - [命名範圍](#命名範圍)
   - [參數區塊功能](#參數區塊功能)
   - [多維陣列處理](#多維陣列處理)
   - [預設範本](#預設範本)
6. [範本範例](#範本範例)
7. [命令行模式](#命令行模式)
8. [進階功能](#進階功能)
9. [疑難排解](#疑難排解)
10. [技術支援](#技術支援)

## 功能概述

- **多檔案處理**：可同時轉換多個 Excel 檔案
- **工作表選擇**：自由指定要讀取的工作表
- **範圍管理**：支援手動輸入或命名範圍
- **可自訂範本**：透過範本決定程式碼格式
- **命令行模式**：便於自動化與批次處理

## 程式架構

1. **main.py** – 程式入口
2. **gui.py** – 圖形介面
3. **console.py** – 命令行介面
4. **excel_handler.py** – 讀取與整理 Excel
5. **code_generator.py** – 範本解析與程式碼產生
6. **utils.py** – 共用工具

## 安裝與執行

### 取得程式

- 從 GitHub 下載原始碼或使用提供的執行檔

### 安裝相依套件

```bash
pip install -r requirements.txt
```

### 啟動程式

#### 圖形介面

```bash
python main.py
```
或直接啟動提供的可執行檔。

#### 命令列

```bash
python console.py --config config.json --output out.c
```

## 基本使用流程

1. **選擇 Excel 檔案**：可多選
2. **選擇工作表**：於下拉選單指定
3. **設定範圍**：使用範圍管理器輸入如 `A1:D10`，或新增命名範圍
4. **選擇或撰寫範本**：可從預設範本下拉選單挑選，或自訂內容
5. **產生程式碼**：按下「生成程式碼」後即可在右側看到結果

## 範本系統

### 範本基礎

範本是一段文字，內含 `{{標記}}` 作為佔位符，生成程式碼時會以對應的 Excel 資料取代。可自由撰寫 C/C++ 或其他語言的輸出格式。

### 範本標記語法

| 標記 | 說明 |
|------|------|
| `{{LOOP_START}}` / `{{LOOP_END}}` | 包覆需要重複的區塊 |
| `{{VALUE}}` | 目前儲存格內容 |
| `{{ROW_INDEX}}`、`{{COL_INDEX}}` | 目前的行與列索引(從0開始) |
| `{{ALL_COLUMNS}}` | 該行所有欄位資料(以逗號分隔) |
| `{{ALL_ROWS}}` | 該列所有行資料 |
| `{{COL:n}}` | 取得此行第 n 欄的值 |
| `{{ROW:n}}` | 取得此列第 n 行的值 |

#### 檔案與範圍相關標記

| 標記 | 說明 |
|------|------|
| `{{FILE_NAME}}`、`{{FILE_INDEX}}`、`{{FILE_COUNT}}` | 多檔案資訊 |
| `{{RANGE[名稱]_LOOP_START}}` / `{{RANGE[名稱]_LOOP_END}}` | 指定命名範圍的迴圈 |
| `{{RANGE[名稱]_ROW_COUNT}}`、`{{RANGE[名稱]_COL_COUNT}}` | 範圍尺寸 |
| `{{RANGE[名稱]_VALUE[i,j]}}` | 直接取得範圍內的值 |

### 橫向和直向讀取模式

預設以橫向(ROW)方式讀取 Excel。若欲改為直向(COLUMN)讀取，可在範本首行加入：

```
{{DIRECTION:COLUMN}}
```

### 命名範圍

在範圍管理器中可以為任何選取區塊設定名稱。於範本內便能透過 `{{RANGE[名稱]_...}}` 標記來引用該區塊，方便處理多個不連續的資料區域。

### 參數區塊功能

可定義重複使用的區塊，例如：

```
{{ARGUMENT_START:TABLE}}
// 範圍名稱=左上,右上
unsigned int {{RANGE[name]_NAME}}[] = {
{{RANGE[name]_LOOP_START}}
    {{ALL_COLUMNS}},
{{RANGE[name]_LOOP_END}}
};
{{ARGUMENT_END:TABLE}}
```

在程式碼其他位置可透過 `{{ARGUMENT_START:...}}` 與 `{{ARGUMENT_END:...}}` 包覆，方便產生多個相同結構的區塊。

### 多維陣列處理

配合 `FILES_LOOP_START`、`RANGES_LOOP_START` 等標記，可一次處理多檔案及多範圍，生成三維或四維陣列。

### 預設範本

程式內建多種範本可直接套用：

- `陣列初始化`
- `簡化權重表設定`
- `二維陣列`
- `二維陣列-直向讀取`
- `三維陣列`
- `三維陣列-直向讀取`
- `三維多範圍陣列`
- `權重表設定`
- `權重表設定-直向讀取`
- `多範圍處理`
- `命名範圍處理`

選擇任何一項即可取得對應的範本文字，之後仍可自行調整。

## 範本範例

以下示範如何利用參數區塊定義多個表格：

```c
// 使用參數區塊定義多個表格
{{ARGUMENT_START:WEIGHT_TABLES}}
// 範圍名稱=Normal,Free,Skill
static unsigned int {{RANGE[name]_NAME}}_table[{{RANGE[name]_ROW_COUNT}}][{{RANGE[name]_COL_COUNT}}] = {
{{RANGE[name]_LOOP_START}}
    { {{ALL_COLUMNS}} },
{{RANGE[name]_LOOP_END}}
};
{{ARGUMENT_END:WEIGHT_TABLES}}
```

## 命令行模式

利用 JSON 配置檔描述檔案、範圍與範本，即可批次產生程式碼。

```bash
python console.py --config config.json --output output.c
```

範例 `config.json`：

```json
{
  "excel_files": ["path/to/file.xlsx"],
  "selected_sheet": "Sheet1",
  "selected_ranges": [{"range_str": "A1:D10"}],
  "template_type": "preset",
  "preset_template": "二維陣列",
  "template_direction": "row"
}
```

## 進階功能

- **設定檔儲存**：可將目前設定存成 JSON 方便下次載入
- **範本管理**：自訂範本可匯入/匯出並集中管理
- **收納式側邊欄**：界面更清爽

## 疑難排解

1. **標記未被取代**：檢查範圍是否正確、標記拼寫是否無誤
2. **資料型別錯誤**：預設會自動判斷數值或字串，可透過範本自行包覆格式

## 技術支援

若有問題或建議，歡迎在 GitHub 回報或來信 [wu0851202@gmail.com]

---

ExcelCode Pro v1.0.0
© 2025 WWKing - Alphabet Studio 版權所有
