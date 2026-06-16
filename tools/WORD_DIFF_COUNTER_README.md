# 写作字数差异统计工具

这个工具用于自动比较“写作前”和“写作后”的文档，计算新增字数、删除字数和净增字数。

## 最简单用法

把开始写作前的文件放进：

```text
写作前
```

把写作后的文件放进：

```text
写作后
```

然后在项目根目录运行：

```powershell
.\tools\count_writing_diff.cmd
```

程序会自动比较两个文件夹里的同名文件。优先按相对路径匹配，例如：

```text
写作前\chapter1.tex
写作后\chapter1.tex
```

如果路径不一样，但文件名在两个文件夹中都是唯一的，也会自动匹配。

## 手动指定文件或文件夹

比较两个文件：

```powershell
.\tools\count_writing_diff.cmd ".\写作前\main.tex" ".\写作后\main.tex"
```

比较两个文件夹：

```powershell
.\tools\count_writing_diff.cmd ".\写作前" ".\写作后"
```

保存 CSV 结果：

```powershell
.\tools\count_writing_diff.cmd -Csv ".\outputs\writing_diff_results.csv"
```

输出 JSON：

```powershell
.\tools\count_writing_diff.cmd -Json
```

## 支持格式

- `.docx`
- `.pdf`
- `.tex`
- `.txt`
- `.md`

## 字数规则

- 中文字符：每个汉字算 1 字
- 英文和数字：连续的一段算 1 个词
- 标点、空格、换行不计入字数

## LaTeX 统计方式

`.tex` 文件会先做正文清洗，再统计字数。程序会尽量忽略：

- 注释 `% ...`
- 数学公式 `$...$`、`$$...$$`、`\[...\]`
- 常见公式环境，如 `equation`、`align`
- 引用命令，如 `\cite{}`、`\ref{}`、`\label{}`
- 图片、表格、宏包、文档类等非正文命令

它会保留常见正文命令里的文字，例如：

- `\section{标题}`
- `\subsection{标题}`
- `\textbf{正文}`
- `\emph{正文}`
- `\caption{说明}`
- `\footnote{脚注}`

如果你的 LaTeX 项目有很多章节文件，推荐把整个项目分别放进 `写作前` 和 `写作后`，用文件夹模式比较。

## PDF 统计方式

`.pdf` 文件会先尝试抽取 PDF 里的文本，再统计差异。这个方式适合：

- LaTeX 编译后的 PDF
- Word 导出的 PDF
- 普通可复制文字的 PDF

注意：扫描版 PDF、图片型 PDF、复杂双栏排版、页眉页脚和参考文献可能会影响统计。对于 LaTeX 写作，通常优先比较 `.tex` 源文件；如果你想看最终成稿变化，再比较 PDF。

## 输出字段

- `Before total words`：写作前总字数
- `After total words`：写作后总字数
- `Added words`：新增字数
- `Deleted words`：删除字数
- `Net words`：净增字数
- `Changed words`：改写涉及字数

日常填写 Excel 的“字数”栏，建议使用 `Added words`。如果你更关心文章最终篇幅变化，可以使用 `Net words`。
