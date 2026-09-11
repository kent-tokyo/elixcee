# elixcee playground

ブラウザーで`.xlsx`を開き、複数シートを切り替え、セルや範囲を編集できます。限定VBAをWeb Workerで実行し、Rust/WASMで再計算し、診断を確認して検証済みの`.xlsx`をダウンロードできます。ページはクライアントだけで動き、ワークブックのバイト列はアップロードされません。
`xl/externalLinks/`を含むワークブックは、外部リンクの完全保持を保証できないためブラウザー編集を拒否します。リンクを保持する場合はネイティブランタイムを使用してください。
PivotTable/PivotCache、埋め込みmedia、threaded comments、slicer、custom XMLを含むファイルも、ブラウザーWriterの完全保持対象外のため拒否します。これらを保持する場合はネイティブランタイムを使用してください。

[GitHub Pagesのplayground](https://kent-tokyo.github.io/elixcee/playground/)を開くか、先に[日本語クイックスタート](../docs/quickstart-ja.md)を読んでください。
画面内の言語セレクターで、UIとクイックスタートへのリンクを英語・日本語・簡体中文に切り替えられます。翻訳は併記しません。

ブラウザー版は非公開の`@elixcee/xlsx`パッケージによるメモリ上のXLSX読み書き、複数シート、型付き編集、数式再計算、診断を扱います。VBAは`Sub`、`Dim`、スカラー代入、`Cells(row,col).Value`、算術演算、限定的な`If`／`For`／`Do`／`Select Case`をWorker内で実行します。ファイル・ネットワーク・COM・Shell・UI・未対応構文は拒否します。完全なヘッドレスVBAデータ処理は[日本語チュートリアル](../docs/tutorial-beginners-ja.md)を参照してください。
「ピボット集計」は、カテゴリ別のSum／Count／Averageを通常ワークシートとして作成し、元データの変更時に更新します。OOXMLのPivotTable／PivotCache編集ではなく、ネイティブPivotTableを含むファイルはブラウザーの事前検査で拒否します。

## ローカルのPythonサンプル

```bash
python playground/run.py
python playground/run.py --output /tmp/elixcee-playground.xlsx
```

翻訳：[English](README.md) · [简体中文](README-zh.md)
