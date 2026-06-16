# 写作进程工具

这个项目包含两个本地工具：

- 生成写作进程 Excel 表格：`tools/build_writing_progress.mjs`
- 自动比较写作前/写作后文档字数差异：`tools/count_writing_diff.cmd`

## 字数差异统计

把开始写作前的文件放入 `写作前/`，把写作后的文件放入 `写作后/`，然后双击或运行：

```powershell
.\tools\count_writing_diff.cmd
```

结果会显示在窗口里，并保存到 `outputs/latest_word_diff.txt`。

支持 `.docx`、`.pdf`、`.tex`、`.txt`、`.md`。

## Excel 生成

写作进程表格由 `tools/build_writing_progress.mjs` 生成，输出位于 `outputs/`。`outputs/` 是生成目录，默认不纳入 Git。

## Git 说明

仓库默认追踪工具脚本和说明文档，不追踪：

- `node_modules/`
- `outputs/` 中的生成文件
- `写作前/` 和 `写作后/` 中的私人写作文档
- Python 缓存和临时文件
