# elixcee playground

在浏览器中编辑一个小型工作簿，使用 Rust/WASM 引擎重新计算公式，并下载生成的 `.xlsx` 文件。页面只在客户端运行，不会上传工作簿字节数据。

打开 [GitHub Pages playground](https://kent-tokyo.github.io/elixcee/playground/)，或先阅读[简体中文快速开始](../docs/quickstart-zh.md)。
使用页面内的语言选择器可在英语、日语和简体中文之间切换界面及快速开始链接；翻译不会并排显示。

浏览器页面覆盖私有 `@elixcee/xlsx` 包提供的内存 XLSX 读写和类型化工作簿编辑，不会在浏览器中运行 Python 或 VBA。无头 VBA 数据处理请阅读[初学者教程](../docs/tutorial-beginners-zh.md)。

## 在本地运行 Python 示例

```bash
python playground/run.py
python playground/run.py --output /tmp/elixcee-playground.xlsx
```

翻译：[English](README.md) · [日本語](README-ja.md)
