# 國高中數學科 108 課綱素養命題與自動化產卷技能 (jhsh-math-exam-generator)

本技能 (Skill) 專門用於指導 Google Antigravity 等 AI 代理人進行國高中數學科段考自動化命題，生成符合「新北市立錦和高級中學」試務組規範、**教育部《十二年國民基本教育課程綱要—數學領域（108 課綱）》**與 PISA 數學推論架構的正式試題卷、答案解析卷與審題檢核表。

---

## ✨ 五大核心亮點

1. **完整內建「教育部 108 數學領域課綱」與「7～12 年級防超綱紅線系統」**：
   - 內建國中（`數-J`）與普通型高中（`數S-U`）三面九項核心素養、六大類學習表現（`n` 數與量、`s` 空間與形狀、`g` 坐標幾何、`a` 代數、`f` 函數、`d` 資料與不確定性）與學習內容（`N/S/G/A/F/D`）完整編碼體系。
   - **嚴格執行課綱「備註欄」防超綱紀律**：自動攔截舊課綱超綱題型（例如：7 年級不考絕對值方程式與聯立不等式、8 年級不出現抽象符號 $f(x)$ 與分離係數法、9 年級不考一般式配方法與舊圓冪定理、10 年級不考牛頓一次因式檢驗法、高二數 B 嚴格不跨數 A 空間向量外積與平面方程）。
2. **Word 原生可編輯數學方程式 (OMML `<m:oMath>`)**：
   - 自動將分式 $\frac{a}{b}$、根式 $\sqrt{x}$、次方、聯立方程組 $\begin{cases}...\end{cases}$ 與幾何線段頂標 $\overline{AB}$、弧 $\overset{\frown}{AB}$ 寫入為 **Word 原生方程式物件**。老師打開 Word 點兩下即可直接修改分子、分母與數字！
3. **「座標先行 (Coordinate-First)」300 DPI 幾何與函數繪圖引擎**：
   - 強制鎖定等比例 (`aspect='equal'`)，內建直角記號、角度圓弧、等長刻線、平行箭頭、陰影面積斜線填充、立體幾何虛線背面與標準十字直角座標系，頂點字母自動外推絕不壓線。
4. **SymPy 符號運算 100% 自動驗算防錯機制**：
   - 每道題目寫入考卷前，AI 強制先以 Python `SymPy` 驗算代數解、畢氏定理與三角不等式，杜絕無解題或條件矛盾。
5. **支援數學科三大題型與智慧省紙排版**：
   - 完整支援「單一選擇題（含短選項四選一列橫排）」、「填充題（自動生成標準作答欄表格）」、「非選擇題／計算證明題（附作答方框與國中會考 0~3 級分評分規準）」。

---

## 📦 如何安裝本技能

### 方式 A：全域安裝（推薦，所有專案與新對話皆可隨時調用）

#### 1. 在 Antigravity 對話框貼上一句話自動安裝（最推薦）：
```text
請幫我把這個 GitHub 倉庫安裝成全域 Skill：
https://github.com/geniefu/jhsh-math-exam-generator
並幫我檢查安裝所需的 Python 套件 (python-docx, matplotlib, numpy, sympy)
```

#### 2. 透過 Git 指令一鍵安裝（未來更新只需 `git pull`）：
- **Windows (PowerShell)**:
  ```powershell
  git clone https://github.com/geniefu/jhsh-math-exam-generator.git "$HOME\.gemini\config\skills\jhsh-math-exam-generator"
  ```
- **macOS / Linux**:
  ```bash
  git clone https://github.com/geniefu/jhsh-math-exam-generator.git ~/.gemini/config/skills/jhsh-math-exam-generator
  ```

#### 3. 手動解壓縮安裝：
將 `jhsh-math-exam-generator` 資料夾複製至個人全域設定目錄：
- **Windows**: `C:\Users\<使用者名稱>\.gemini\config\skills\jhsh-math-exam-generator\`
- **macOS / Linux**: `~/.gemini/config/skills/jhsh-math-exam-generator/`

---

## 🚀 快速開始：提詞範例 (Prompt)

安裝完成後，在對話框中輸入 `/jhsh-math-exam-generator` 並提供段考參數：

```text
請使用 /jhsh-math-exam-generator 技能為我生成一份數學科段考試卷，規格參數如下：
1、年級與範圍：【請填寫年級與章節，例如：八年級下學期 第 1 章「等差數列與等差級數」到 第 2 章「一次函數及其圖形」】。
2、題型與配分：總分 100 分
   - 第一部分：單一選擇題【例如：12 題，每題 4 分，共 48 分】
   - 第二部分：填充題【例如：10 格，每格 4 分，共 40 分，請於卷末附填充題標準答案欄】
   - 第三部分：非選擇題（計算與證明題）【例如：2 題，每題 6 分，共 12 分，附作答區方框】
3、難易程度：【例如：難易適中（預估平均通過率約 60%～65%，兼顧基本概念熟練與多步驟素養推論）】。
4、圖形與公式要求：所有代數分式、根號、聯立方程與幾何線段符號請使用 Word 原生可編輯 OMML 方程式；幾何題與座標題請繪製 300 DPI 高清黑白向量圖並採右側浮動文繞圖排版。
5、其它：【選填，例如：嚴格遵守 108 數學課綱防超綱限制，並以 Python SymPy 100% 自動驗算確認條件無矛盾且答案唯一】。
```

---

## 🛠️ 環境依賴需求 (Dependencies)

```bash
pip install python-docx matplotlib numpy sympy
```

