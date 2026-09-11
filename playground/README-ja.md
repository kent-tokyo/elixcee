# elixcee playground

ブラウザーで`.xlsx`を開き、複数シートを切り替え、セルや範囲を編集できます。限定VBAをWeb Workerで実行し、Rust/WASMで再計算し、診断を確認して検証済みの`.xlsx`をダウンロードできます。ページはクライアントだけで動き、ワークブックのバイト列はアップロードされません。

[GitHub Pagesのplayground](https://kent-tokyo.github.io/elixcee/playground/)を開くか、先に[日本語クイックスタート](../docs/quickstart-ja.md)を読んでください。
画面内の言語セレクターで、UIとクイックスタートへのリンクを英語・日本語・簡体中文に切り替えられます。翻訳は併記しません。

ブラウザー版は非公開の`@elixcee/xlsx`パッケージによるメモリ上のXLSX読み書き、複数シート、型付き編集、数式再計算、診断を扱います。VBAは`Sub`、`Dim`、スカラー代入、`Cells(row,col).Value`、算術演算だけをWorker内で実行します。ファイル・ネットワーク・COM・Shell・UI・制御構文は拒否します。完全なヘッドレスVBAデータ処理は[日本語チュートリアル](../docs/tutorial-beginners-ja.md)を参照してください。

## ローカルのPythonサンプル

```bash
python playground/run.py
python playground/run.py --output /tmp/elixcee-playground.xlsx
```

翻訳：[English](README.md) · [简体中文](README-zh.md)
