# elixcee playground

在浏览器中打开 `.xlsx`，切换多个工作表，编辑单元格和区域；在 Web Worker 中运行受限 VBA 子集，使用 Rust/WASM 重新计算，查看诊断，并下载经过验证的 `.xlsx` 文件。页面只在客户端运行，不会上传工作簿字节数据。
包含 `xl/externalLinks/` 部件的工作簿会被浏览器编辑器拒绝，因为当前不能保证外部链接无损保留。需要保留链接时请使用原生运行时。
包含 PivotTable/PivotCache、嵌入媒体、threaded comments、slicer 或 custom XML 的文件也会被拒绝，因为浏览器 Writer 尚不能保证这些部件无损保留。需要保留这些部件时请使用原生运行时。

打开 [GitHub Pages playground](https://kent-tokyo.github.io/elixcee/playground/)，或先阅读[简体中文快速开始](../docs/quickstart-zh.md)。
使用页面内的语言选择器可在英语、日语和简体中文之间切换界面及快速开始链接；翻译不会并排显示。

浏览器页面覆盖私有 `@elixcee/xlsx` 包提供的内存 XLSX 读写、多工作表、类型化编辑、公式重算和诊断。VBA 沙盒允许 `Sub`、`Dim`、标量赋值、`Cells(row,col).Value`、算术运算以及受限的 `If`／`For`／`Do`／`Select Case`；文件、网络、COM、Shell、界面和未支持的语法会被拒绝。完整的无头 VBA 数据处理请阅读[初学者教程](../docs/tutorial-beginners-zh.md)。
“透视汇总”按钮会将分类的 Sum／Count／Average 创建为普通工作表，并在源数据变化时刷新。它不是 OOXML PivotTable／PivotCache 编辑器；包含原生PivotTable的文件会在浏览器预检阶段被拒绝。

## 在本地运行 Python 示例

```bash
python playground/run.py
python playground/run.py --output /tmp/elixcee-playground.xlsx
```

翻译：[English](README.md) · [日本語](README-ja.md)
