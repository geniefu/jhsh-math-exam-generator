---
name: jhsh-math-exam-generator
description: >-
  專業生成符合「新北市立錦和高中」及台灣中學段考規範之數學科試題卷、答案卷與詳解卷。
  完整融合教育部《十二年國民基本教育課程綱要—國民中小學暨普通型高級中等學校—數學領域（108課綱）》
  （涵蓋核心素養 數-J/數S-U、學習表現 n/s/g/a/f/d、學習內容 N/S/G/A/F/D、7~12年級嚴格防超綱紅線與學習評量實施要點）、
  國教院數學素養評量要素與 PISA 數學推論架構。
  內建「Word 原生可編輯 OMML 數學方程式引擎」、「SymPy 符號運算 100% 自動驗算防錯機制」及「座標先行 300 DPI 幾何與函數繪圖引擎」，
  支援「單選題＋填充題＋非選計算證明題」三大題型與錦和高中標準排版。
---

# 國高中數學科 108 課綱素養命題規範與自動化產卷技能 (jhsh-math-exam-generator)

本 Skill 專門用於指導 Antigravity 代理人自動化命題，生成完全符應「新北市立錦和高級中學」教務處/試務組規範，且深度結合**教育部《十二年國民基本教育課程綱要—數學領域（108 課綱）》**、國教院素養導向紙筆測驗要素與國際 PISA 數學評量架構之正式數學段考試題卷、答案與解析卷，以及命題審題雙向細目檢核表。

---

## 快速開始：使用者提詞標準模板 (User Prompt Template)

> [!TIP]
> **若使用者初次使用或發起新段考任務，請引導使用者複製以下標準提詞模板：**

```text
請使用 /jhsh-math-exam-generator 技能為我生成一份數學科段考試卷，規格參數如下：
1、標題：【例如：新北市立錦和高級中學115學年度第一學期 國中部八年級 數學科第一次段考試題卷】（年級請依出題範圍調整為七年級、八年級或九年級）。
2、年級與範圍：【請填寫年級與章節，例如：八年級下學期 第 1 章「等差數列與等差級數」到 第 2 章「一次函數及其圖形」】。
3、題型與配分：總分 100 分
   - 第一部分：單一選擇題【例如：12 題，每題 4 分，共 48 分】
   - 第二部分：填充題【例如：10 格，每格 4 分，共 40 分，請於卷末附填充題標準答案欄】
   - 第三部分：非選擇題（計算與證明題）【例如：2 題，每題 6 分，共 12 分，附作答區方框】
4、難易程度：【例如：難易適中（預估平均通過率約 60%～65%）】。
5、圖形與公式要求：所有代數分式、根號、聯立方程與幾何線段符號請使用 Word 原生可編輯 OMML 方程式；幾何題與座標題請繪製 300 DPI 高清黑白向量圖並採右側浮動文繞圖排版。
6、其它：【選填，例如：嚴格遵守 108 數學課綱防超綱限制，並以 Python SymPy 100% 自動驗算確認條件無矛盾且答案唯一】。
```

---

## 核心工作架構 (Modular Progressive Disclosure)

本技能採用精簡三層架構：核心引導、執行腳本庫與法規參考文獻：

```mermaid
flowchart TD
    A[確認段考參數：年級/分軌、範圍、題型配分] --> B[檢視 references/curriculum_guidelines_7_9.md 或 10_12.md 防超綱紅線]
    B --> C[呼叫 scripts/math_sympy_verifier.py: 100% 符號運算先行驗算 (防無解/矛盾/重解)]
    C --> D[呼叫 scripts/math_geometry_plotter.py: 300 DPI 座標先行幾何與十字座標圖]
    D --> E[呼叫 scripts/math_omml_builder.py: 分式/根式/聯立方程轉 Word 原生 OMML]
    E --> F[呼叫 scripts/docx_math_builder.py: 產出雙份 Word 試題與解析卷 (含填充欄/非選框)]
    F --> G[檢核 references/rubric_non_multiple_choice.md: 附非選 0~3 級分完整評分規準]
```

### 1. 執行輔助腳本庫 (`scripts/`)
- [`docx_math_builder.py`](file:///C:/Users/genie/.gemini/config/skills/jhsh-math-exam-generator/scripts/docx_math_builder.py)：封裝 1cm 邊界、全卷 11 點標楷體、粗體底線抬頭、紅字扣 5 分警語、題號凸排、選擇題智慧並排、填充題作答表格、非選題作答方框與動態頁尾代碼 `〔第 X 頁，共 Y 頁〕`。
- [`math_omml_builder.py`](file:///C:/Users/genie/.gemini/config/skills/jhsh-math-exam-generator/scripts/math_omml_builder.py)：Word 原生可編輯 OMML `<m:oMath>` 建構器（分式、根式、上下標、線段頂標 $\overline{AB}$、向量 $\vec{u}$、聯立方程大括號與矩陣）。
- [`math_geometry_plotter.py`](file:///C:/Users/genie/.gemini/config/skills/jhsh-math-exam-generator/scripts/math_geometry_plotter.py)：「座標先行」300 DPI 幾何與函數繪圖引擎（強制等比例 `equal`、直角小方框、角度弧、等長記號、頂點字母外推防壓線、標準十字直角坐標系）。
- [`math_sympy_verifier.py`](file:///C:/Users/genie/.gemini/config/skills/jhsh-math-exam-generator/scripts/math_sympy_verifier.py)：SymPy 符號運算 100% 自動驗算防錯（代數解驗證、三角不等式檢驗、四選項互異且唯一性檢驗、全卷配分加總 100 分檢驗）。

### 2. 課綱法規與評量參考庫 (`references/`)
- [`curriculum_guidelines_7_9.md`](file:///C:/Users/genie/.gemini/config/skills/jhsh-math-exam-generator/references/curriculum_guidelines_7_9.md)：國中 7～9 年級第四學習階段學習表現、學習內容與法定備註欄防超綱禁區（**7年級禁考絕對值方程、8年級禁考 $f(x)$ 符號與分離係數法、9年級禁考一般式配方法與舊課綱圓冪定理**）。
- [`curriculum_guidelines_10_12.md`](file:///C:/Users/genie/.gemini/config/skills/jhsh-math-exam-generator/references/curriculum_guidelines_10_12.md)：高中 10～12 年級必修與高二分軌（數 A vs. 數 B 邊界）、高三分軌（數甲 vs. 數乙邊界）。
- [`rubric_non_multiple_choice.md`](file:///C:/Users/genie/.gemini/config/skills/jhsh-math-exam-generator/references/rubric_non_multiple_choice.md)：國中教育會考（CAP）與高中段考非選擇題 0～3 級分評分規準與分段給分實務。

---

## 錦和高中數學科試務組排版標準規範

1. **試題抬頭（粗體加底線）**：
   - 格式：<u>**新北市立錦和高級中學 11X學年度第X學期 [國中/高中]部○年級數學科第○次段考試題**</u>
   - 包含：﹝命題範圍：...﹞、右上角標註「班級：____ 座號：__ 姓名：________」。
2. **必放注意事項（警語）**：
   - 答案卡警語（紅色粗體）：「**答案卷(卡)未寫班級、姓名、座號，或畫卡錯誤致電腦無法判讀考生身份者，一律扣 5 分**」。
   - 非選／填充警語：「**填充題與非選擇題請用黑色墨水筆於答案卷規定欄位內作答，違者扣該部分總分 5 分；非選擇題需寫出完整計算或推論過程才予計分**」。
3. **字體與邊界設定**：
   - 頁邊界：上、下、左、右各 **1.0 cm**（0.3937 英吋）。
   - 中文一律使用 **標楷體**，英數字使用 **Times New Roman**，原生數學方程式使用 **Cambria Math**。
   - **全卷正文一律維持 11 點字 (11 Pt)**，段落行距設為 1.15 倍行距（含分式根號段落不鎖死固定行高）。
   - 題號凸排：`left_indent = 0.28 inch, first_line_indent = -0.28 inch`。
4. **選擇題選項智慧橫排**：
   - 短選項（$\le 10$ 字元且無圖）：四選一列橫排（`...　　(B)...`）。
   - 中等長度（$11 \sim 22$ 字元且無圖）：兩選一列（2×2）。
   - 長句或**附有右側浮動圖形**：一律採單列排列（1 option per line）。
5. **填充題與非選擇題版面**：
   - 填充題：自動產生帶框線標準作答表格（每列 4～5 格，留白高度 $\ge 1.2\text{ cm}$）。
   - 非選題：自動繪製作答區方框（高度約 $4.5 \sim 6.0\text{ cm}$），供學生書寫完整步驟。
6. **動態頁碼與省紙控制**：
   - 頁尾置中加入動態欄位代碼：`〔第 X 頁，共 Y 頁〕`。最後一頁版面嚴格不可少於整頁的 1/3。

---

## 產卷完畢自檢清單 (Checklist)

每次完成試題產出後，必須在交付前逐一核對：
- [ ] 是否已逐題核對 **108 數學課綱「各年級嚴格防超綱紅線清單」**（例如：7 年級無絕對值方程、8 年級無 $f(x)$ 符號、9 年級無一般式配方法、10 年級無牛頓一次因式檢驗法、11B 跨軌禁區）？
- [ ] 全卷分數、根號、次方、聯立方程組與幾何線段頂標是否皆已使用 **Word 原生 `<m:oMath>` (OMML)** 而非純文字拼湊？
- [ ] 是否已透過 **Python `SymPy`** 100% 驗算過每一題的標準答案與幾何條件自洽性？
- [ ] 幾何圖形是否鎖定 `ax.set_aspect('equal')`、具備直角/角度弧/等長記號，且頂點字母已外推無壓線？
- [ ] 圖表字級是否符合 **+8 Pt 特大清晰標準（16.5～21 Pt）** 且採黑白灰階/斜線填充設計？
- [ ] 附圖試題是否採用 **右側浮動「矩形文繞圖 (`wp:anchor` + `wrapSquare`)」**（寬度 2.4～2.7 英吋）？
- [ ] 選擇題短選項是否已採用「四選一列」或「兩選一列」智慧橫排以節省紙張？
- [ ] 是否包含錦和高中標準粗體底線抬頭、紅字扣 5 分警語、全卷 11 Pt 標楷體/Times New Roman、1cm 邊界、題號凸排與動態頁尾 `〔第 X 頁，共 Y 頁〕`？
- [ ] 最後一頁版面是否大於 1/3 頁？
- [ ] 是否產出獨立三份檔案：試題卷、答案與解析卷（含非選 0~3 級分規準）、命題審題檢核表？
