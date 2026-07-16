# 議事録作成ツール v1.0.0

AIで作成したテキスト議事録を貼り付け、CTM指定のExcelテンプレートへ自動入力するGUIツールです。
WindowsとMacの両方で利用できます。

## バージョン

- リリース版: v1.0.0

## できること

- 議事録テキストを画面に貼り付けてExcelを作成
- Excel保存先を作成時に選択
- 作成完了後に保存先を表示
- テンプレートExcelが見つからない場合に選択ダイアログを表示
- 入力エラー、テンプレートエラー、保存エラーを日本語で表示
- 議題数やToDo件数の増減に対応
- テンプレートの罫線・書式・セル結合を維持

## ファイル構成

- `app.py`: GUI画面
- `parser.py`: 議事録テキストの解析
- `excel_writer.py`: Excelテンプレートへの書き込み
- `models.py`: データ定義
- `version.py`: バージョン情報
- `resources/会議議事録テンプレ.xlsx`: Excelテンプレート
- `resources/icon.icns`: Macアプリ用アイコン
- `output/会議議事録サンプル_v1.0.0.xlsx`: 実議事録サンプルから作成した動作確認用Excel
- `USER_GUIDE_A4.md`: 一般ユーザー向けのA4一枚程度の使い方
- `../run_minutes_tool.py`: 起動用ファイル
- `../requirements.txt`: 通常利用に必要なライブラリ
- `../requirements-dev.txt`: テスト用ライブラリ

## 通常の起動方法

1. Python 3.11以降をインストールします。
2. ターミナルまたはコマンドプロンプトで、このプロジェクトのフォルダへ移動します。

Macの例:

```bash
cd "/Users/ユーザー名/path/to/ctm-matsuzaki.github.io"
```

Windowsの例:

```bat
cd "C:\path\to\ctm-matsuzaki.github.io"
```

3. 必要ライブラリをインストールします。

```bash
pip install -r requirements.txt
```

Macで `pip` が使えない場合:

```bash
python3 -m pip install -r requirements.txt
```

4. アプリを起動します。

```bash
python3 run_minutes_tool.py
```

Windowsで `python3` が使えない場合:

```bat
python run_minutes_tool.py
```

## 基本操作

1. 画面中央のテキストボックスに、AIで作成した議事録テキストを貼り付けます。
2. `Excel作成` ボタンを押します。
3. Excelファイルの保存先を選択します。
4. 作成が完了すると、`作成完了` メッセージと保存先が表示されます。
5. 入力欄を空にする場合は `クリア`、アプリを閉じる場合は `終了` を押します。

## テンプレートExcelが見つからない場合

通常は `resources/会議議事録テンプレ.xlsx` を自動で読み込みます。

もしテンプレートが見つからない場合は、Excelテンプレートを選択する画面が表示されます。
その場合は、使用する `会議議事録テンプレ.xlsx` を選択してください。

## エラーが出た場合

- `議事録テキストを貼り付けてください`: 入力欄が空です。
- `一部の項目を自動判定できませんでした。黄色セルを確認してください。`: Excelは作成できます。黄色セルに不足情報を追記してください。
- `Excelテンプレートが見つかりません`: テンプレート選択画面で `会議議事録テンプレ.xlsx` を選択してください。
- `Excelファイルを保存できませんでした`: 同名ファイルをExcelで開いている場合は閉じてください。保存先フォルダの権限も確認してください。

## 入力テキストの想定フォーマット

以下の見出しを含むテキストを想定しています。

```text
基本情報
会議名
...
日時
...
場所
...
出席者
...

議題① 議題名
内容
・内容1
決定事項
・決定事項1

ToDo
担当 期限 内容
担当者 不明 タスク内容

次回会議
日時
...
場所
...
```

ToDoは、次の形式に対応しています。

- `担当 期限 内容`
- `担当<TAB>期限<TAB>内容`
- `担当 | 期限 | 内容`

PDFからコピーした際にToDoの担当・期限・内容が複数行に折り返された場合も、できるだけ復元します。

## Windows用exeの作成手順

Pythonが入っていないWindows PCで使う場合は、開発用のWindows PCでexeを作成して配布します。

1. Windows PCでプロジェクトフォルダを開きます。
2. 必要ライブラリをインストールします。

```bat
pip install -r requirements.txt
```

3. PyInstallerでexeを作成します。

```bat
pyinstaller --onefile --windowed --name 議事録作成ツール --add-data "minutes_tool\resources\会議議事録テンプレ.xlsx;minutes_tool\resources" run_minutes_tool.py
```

4. 完成したexeを確認します。

```text
dist\議事録作成ツール.exe
```

5. `dist\議事録作成ツール.exe` を利用者のWindows PCへ配布します。

## Mac用アプリ（.app）の作成手順

Pythonが入っていないMacで使う場合は、開発用のMacで `.app` を作成して配布します。

1. Macでプロジェクトフォルダを開きます。

```bash
cd "/path/to/ctm-matsuzaki.github.io"
```

2. Python 3.13 と必要ライブラリを用意します。

```bash
python3 -m pip install -r requirements.txt
```

PyInstallerが未インストールの場合:

```bash
python3 -m pip install pyinstaller
```

3. PyInstallerで `.app` を作成します。通常はこちらを使います。

```bash
python3 -m PyInstaller CTM議事録作成ツール.spec
```

4. 完成したアプリを確認します。

```text
dist/議事録作成ツール.app
```

5. `dist/議事録作成ツール.app` を利用者のMacへ配布します。

Macのセキュリティ設定により初回起動できない場合は、Finderでアプリを右クリックして `開く` を選択してください。

### Macアプリの仕様

- アプリ名: `議事録作成ツール`
- 起動方法: `.app` をダブルクリック
- ターミナル表示: なし
- 同梱ファイル: `minutes_tool/resources/会議議事録テンプレ.xlsx`
- アイコン: `minutes_tool/resources/icon.icns` を使用します。アイコンを差し替える場合は、同じ場所に新しい `.icns` を保存してから再ビルドします。

### specを使わずに直接ビルドする場合

```bash
python3 -m PyInstaller \
  --windowed \
  --name "議事録作成ツール" \
  --icon "minutes_tool/resources/icon.icns" \
  --add-data "minutes_tool/resources/会議議事録テンプレ.xlsx:minutes_tool/resources" \
  run_minutes_tool.py
```

## 一般ユーザー向けの使い方

利用者に配布する簡易手順書は、次のファイルです。

```text
minutes_tool/USER_GUIDE_A4.md
```

A4一枚程度で、起動、貼り付け、Excel作成、よくあるエラーをまとめています。

## リリース確認項目

v1.0.0では、実際の議事録サンプルからExcelを1件作成し、以下を確認しています。

- 確認用Excel: `minutes_tool/output/会議議事録サンプル_v1.0.0.xlsx`
- 基本情報、議題、ToDo、次回会議がExcelへ入力されること
- テンプレートの主要セル結合が維持されること
- テンプレートの主要列幅が維持されること
- テンプレートの主要セル書式が維持されること

## 保守・拡張のポイント

- 入力フォーマットを増やす場合は `parser.py` に解析ルールを追加します。
- Excelの配置を変更する場合は `excel_writer.py` の行定義を変更します。
- GUI項目を増やす場合は `app.py` に画面部品を追加します。

## テスト

テストを実行する場合は、開発用ライブラリをインストールします。

```bash
pip install -r requirements-dev.txt
```

テスト実行:

```bash
python -m pytest tests
```
