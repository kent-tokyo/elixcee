# elixcee playground

ブラウザーで小さなワークブックを編集し、Rust/WASMエンジンで数式を再計算し、結果を`.xlsx`としてダウンロードできます。ページはクライアントだけで動き、ワークブックのバイト列はアップロードされません。

[GitHub Pagesのplayground](https://kent-tokyo.github.io/elixcee/playground/)を開くか、先に[日本語クイックスタート](../docs/quickstart-ja.md)を読んでください。
画面内の言語セレクターで、UIとクイックスタートへのリンクを英語・日本語・簡体中文に切り替えられます。翻訳は併記しません。

ブラウザー版は非公開の`@elixcee/xlsx`パッケージによるメモリ上のXLSX読み書きと型付き編集を扱います。ブラウザーでPythonやVBAを実行するものではありません。ヘッドレスVBAデータ処理は[日本語チュートリアル](../docs/tutorial-beginners-ja.md)を参照してください。

## ローカルのPythonサンプル

```bash
python playground/run.py
python playground/run.py --output /tmp/elixcee-playground.xlsx
```

翻訳：[English](README.md) · [简体中文](README-zh.md)
