# elixcee playground

在浏览器中打开 `.xlsx`，切换多个工作表，编辑单元格和区域；在 Web Worker 中运行受限 VBA 子集，使用 Rust/WASM 重新计算，查看诊断，并下载经过验证的 `.xlsx` 文件。页面只在客户端运行，不会上传工作簿字节数据。

打开 [GitHub Pages playground](https://kent-tokyo.github.io/elixcee/playground/)，或先阅读[简体中文快速开始](../docs/quickstart-zh.md)。
使用页面内的语言选择器可在英语、日语和简体中文之间切换界面及快速开始链接；翻译不会并排显示。

浏览器页面覆盖私有 `@elixcee/xlsx` 包提供的内存 XLSX 读写、多工作表、类型化编辑、公式重算和诊断。VBA 沙盒只允许 `Sub`、`Dim`、标量赋值、`Cells(row,col).Value` 和算术运算；文件、网络、COM、Shell、界面和控制流都会被拒绝。完整的无头 VBA 数据处理请阅读[初学者教程](../docs/tutorial-beginners-zh.md)。

## 在本地运行 Python 示例

```bash
python playground/run.py
python playground/run.py --output /tmp/elixcee-playground.xlsx
```

翻译：[English](README.md) · [日本語](README-ja.md)
