# CLAUDE

このアプリの設計全体をまとめたドキュメントです

---

## 目的:

Excel ワークブックの中から複数のセルを選択し、その内容（値または数式）をテキスト形式で相互変換・編集できる簡易的な一括編集ツールを提供する。変換したテキストはコピー・貼り付けによって編集でき、最終的にワークブックに反映させてダウンロードできる。

---

## 技術スタック:

* **Vite + Vue 3 + TypeScript + HTML/CSS**

  * UIは Composition API で構成
  * スタイルは Tailwind を使わず、class を手動指定

### Viteを使う理由

* 静的サイトとしてビルド・公開するため

---

## 依存関係:

* **xlsx-populate**: Excelファイル（`.xlsx`）の読み取りと編集のため
* **base64-js**: base64 文字列を扱うため

---

## データモデル:

* **selectedCells: `SelectedCell[]`**

  * セルの選択状態を表す。以下の型で構成される

    ```ts
    {
      sheet: string
      row: number
      col: number
      value?: string
      formula?: string
    }
    ```
  * `value` と `formula` は排他的だが両方が存在してもよい（直積で保持）
  * 表現形式:

    * 値: `Sheet1!(2,3):abc`
    * 数式: `Sheet1!(4,5):=SUM(A1:A2)`

* **validationErrors: `string[]`**

  * テキストエリアの変換時に発生したパースエラー等を格納する

---

## UI状態:

* **workbook: Workbook**

  * 開いている Excel ファイルのオブジェクト

* **filename: string**

  * 開かれた Excel ファイルの元のファイル名（保存時に使用）

* **selectedSheet: string**

  * 現在表示・操作しているシート名

* **visibleRows: Cell\[]\[]**

  * 表示中シートのセル内容。2次元配列で構成され、値または数式が含まれる

* **selectedCells: SelectedCell\[]**

  * 現在選択されているセルの一覧（複数可）

* **text: string**

  * `selectedCells` をテキスト形式に変換した文字列（textareaに表示される）

---

## UI計算値:

* **sheets: string\[]**

  * ワークブック内のシート名一覧

* **validationErrors: string\[]**

  * テキスト表現の入力に対する変換エラーがあればここに格納され、UI上に表示される

* **parsed: { values: any\[]\[]; validationError: string | null }**

  * textarea への入力をパースしたもの

---

## UI操作:

* **ワークブックを開く**

  * ユーザーが `.xlsx` ファイルを input 要素から選択すると読み込まれ、最初のシートが表示される

* **シートの選択を切り替える**

  * タブでシートを切り替え、対応するセルデータを表示

* **セルを選択または選択解除する**

  * テーブル上で Ctrl+右クリックすることで、セルの選択状態をトグルできる
  * 選択されたセルは背景色で強調表示される

* **テキストエリアで編集する**

  * `selectedCells` の状態を textarea でテキストとして表示・編集できる
  * 編集された内容はバリデーションされ、正しければ `selectedCells` を更新。失敗すればエラー表示

* **ワークブックに変更を反映してダウンロード**

  * "Download" ボタンをクリックすると、選択されたセルに対する変更がワークブックに適用され、保存可能な `.xlsx` ファイルとしてダウンロードされる

* **バリデーションエラーの表示**

  * テキストエリアの解析に失敗した場合、ページ上部にエラー内容を表示するカードが出現する
