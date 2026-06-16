#!/usr/bin/env python3
"""
Compare writing documents or two folders of document versions.

Default workflow:
- Put old versions in: 写作前
- Put new versions in: 写作后
- Run tools/count_writing_diff.cmd

Supported inputs:
- .docx
- .pdf
- .tex
- .txt
- .md

Counting rule:
- Each CJK character counts as 1.
- Each continuous Latin/number sequence counts as 1.
- Punctuation and whitespace count as 0.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import zipfile
from dataclasses import asdict, dataclass
from difflib import SequenceMatcher
from pathlib import Path
from xml.etree import ElementTree as ET


SUPPORTED_EXTENSIONS = {".docx", ".pdf", ".tex", ".txt", ".md"}
TOKEN_RE = re.compile(
    r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
    r"\U00020000-\U0002a6df\U0002a700-\U0002b73f"
    r"\U0002b740-\U0002b81f\U0002b820-\U0002ceaf]"
    r"|[A-Za-z0-9]+(?:[-_'][A-Za-z0-9]+)*"
)
WORD_XML_NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
DOCX_TEXT_PARTS = [
    "word/document.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
]
LATEX_TEXT_COMMANDS = {
    "title",
    "author",
    "date",
    "chapter",
    "section",
    "subsection",
    "subsubsection",
    "paragraph",
    "subparagraph",
    "caption",
    "textbf",
    "textit",
    "emph",
    "underline",
    "texttt",
    "textrm",
    "textsf",
    "footnote",
    "item",
}
LATEX_DROP_COMMANDS = {
    "label",
    "ref",
    "autoref",
    "cref",
    "Cref",
    "eqref",
    "cite",
    "citet",
    "citep",
    "parencite",
    "textcite",
    "url",
    "href",
    "includegraphics",
    "bibliography",
    "bibliographystyle",
    "addbibresource",
    "usepackage",
    "documentclass",
    "input",
    "include",
}
LATEX_DROP_ENVIRONMENTS = {
    "equation",
    "equation*",
    "align",
    "align*",
    "gather",
    "gather*",
    "multline",
    "multline*",
    "split",
    "cases",
    "figure",
    "figure*",
    "table",
    "table*",
    "tikzpicture",
    "lstlisting",
    "verbatim",
}


@dataclass(frozen=True)
class DiffStats:
    before_words: int
    after_words: int
    added_words: int
    deleted_words: int
    net_words: int
    changed_words: int


@dataclass(frozen=True)
class FileResult:
    status: str
    key: str
    before_file: str
    after_file: str
    before_words: int
    after_words: int
    added_words: int
    deleted_words: int
    net_words: int
    changed_words: int
    note: str = ""


def read_plain_text(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gb18030", "utf-16"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def read_docx_text(path: Path) -> str:
    chunks: list[str] = []
    with zipfile.ZipFile(path) as docx:
        for part in DOCX_TEXT_PARTS:
            if part not in docx.namelist():
                continue
            root = ET.fromstring(docx.read(part))
            for paragraph in root.findall(".//w:p", WORD_XML_NAMESPACES):
                pieces: list[str] = []
                for node in paragraph.iter():
                    tag = node.tag.rsplit("}", 1)[-1]
                    if tag == "t" and node.text:
                        pieces.append(node.text)
                    elif tag in {"tab", "br", "cr"}:
                        pieces.append(" ")
                if pieces:
                    chunks.append("".join(pieces))
    return "\n".join(chunks)


def read_pdf_text(path: Path) -> str:
    try:
        from pypdf import PdfReader

        reader = PdfReader(str(path))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception:
        try:
            import pdfplumber

            with pdfplumber.open(str(path)) as pdf:
                return "\n".join(page.extract_text() or "" for page in pdf.pages)
        except Exception as exc:
            raise ValueError(f"Could not extract PDF text from {path}: {exc}") from exc


def remove_latex_comments(text: str) -> str:
    lines: list[str] = []
    for line in text.splitlines():
        out: list[str] = []
        escaped = False
        for char in line:
            if char == "\\" and not escaped:
                escaped = True
                out.append(char)
                continue
            if char == "%" and not escaped:
                break
            out.append(char)
            escaped = False
        lines.append("".join(out))
    return "\n".join(lines)


def strip_balanced_brace(text: str, start: int) -> tuple[str, int]:
    if start >= len(text) or text[start] != "{":
        return "", start
    depth = 0
    content: list[str] = []
    i = start
    while i < len(text):
        char = text[i]
        if char == "{" and (i == 0 or text[i - 1] != "\\"):
            depth += 1
            if depth > 1:
                content.append(char)
        elif char == "}" and (i == 0 or text[i - 1] != "\\"):
            depth -= 1
            if depth == 0:
                return "".join(content), i + 1
            content.append(char)
        else:
            content.append(char)
        i += 1
    return "".join(content), i


def strip_latex(text: str) -> str:
    text = remove_latex_comments(text)
    text = re.sub(r"\\begin\{(" + "|".join(re.escape(e) for e in LATEX_DROP_ENVIRONMENTS) + r")\}.*?\\end\{\1\}", " ", text, flags=re.DOTALL)
    text = re.sub(r"\$\$.*?\$\$", " ", text, flags=re.DOTALL)
    text = re.sub(r"\\\[.*?\\\]", " ", text, flags=re.DOTALL)
    text = re.sub(r"\$[^$]*\$", " ", text)
    text = re.sub(r"\\\([^)]*\\\)", " ", text, flags=re.DOTALL)

    out: list[str] = []
    i = 0
    while i < len(text):
        char = text[i]
        if char != "\\":
            out.append(char)
            i += 1
            continue

        match = re.match(r"\\([A-Za-z]+|.)", text[i:])
        if not match:
            i += 1
            continue
        command = match.group(1)
        i += len(match.group(0))

        while i < len(text) and text[i].isspace():
            i += 1
        if i < len(text) and text[i] == "[":
            end = text.find("]", i + 1)
            i = len(text) if end == -1 else end + 1
            while i < len(text) and text[i].isspace():
                i += 1

        if command in LATEX_TEXT_COMMANDS and i < len(text) and text[i] == "{":
            content, i = strip_balanced_brace(text, i)
            out.append(" ")
            out.append(strip_latex(content))
            out.append(" ")
        elif command in LATEX_DROP_COMMANDS:
            if i < len(text) and text[i] == "{":
                _, i = strip_balanced_brace(text, i)
            out.append(" ")
        else:
            out.append(" ")

    text = "".join(out)
    text = re.sub(r"\\begin\{[^}]+\}|\\end\{[^}]+\}", " ", text)
    text = text.replace("~", " ")
    text = re.sub(r"[{}_^&#]", " ", text)
    return text


def read_tex_text(path: Path) -> str:
    return strip_latex(read_plain_text(path))


def read_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return read_docx_text(path)
    if suffix == ".pdf":
        return read_pdf_text(path)
    if suffix == ".tex":
        return read_tex_text(path)
    if suffix in {".txt", ".md"}:
        return read_plain_text(path)
    raise ValueError(f"Unsupported file type: {suffix}")


def tokenize(text: str) -> list[str]:
    return TOKEN_RE.findall(text)


def diff_stats(before_text: str, after_text: str) -> DiffStats:
    before_tokens = tokenize(before_text)
    after_tokens = tokenize(after_text)
    matcher = SequenceMatcher(a=before_tokens, b=after_tokens, autojunk=False)

    added = 0
    deleted = 0
    changed = 0
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "insert":
            added += j2 - j1
        elif tag == "delete":
            deleted += i2 - i1
        elif tag == "replace":
            deleted += i2 - i1
            added += j2 - j1
            changed += max(i2 - i1, j2 - j1)

    return DiffStats(
        before_words=len(before_tokens),
        after_words=len(after_tokens),
        added_words=added,
        deleted_words=deleted,
        net_words=len(after_tokens) - len(before_tokens),
        changed_words=changed,
    )


def collect_supported_files(root: Path) -> dict[str, Path]:
    files: dict[str, Path] = {}
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        rel = path.relative_to(root).as_posix()
        files[rel] = path
    return files


def normalized_match_name(path: Path) -> str:
    stem = path.stem.lower()
    stem = re.sub(r"(老版本|旧版本|写作前|写作后|before|after|old|new)", "", stem, flags=re.IGNORECASE)
    stem = re.sub(r"[\s_\-]*\(?\d+\)?$", "", stem)
    stem = re.sub(r"[\s_\-（）()]+", "", stem)
    return f"{stem}{path.suffix.lower()}"


def build_pairs(before_root: Path, after_root: Path) -> list[tuple[str, Path | None, Path | None, str]]:
    before_files = collect_supported_files(before_root)
    after_files = collect_supported_files(after_root)
    pairs: list[tuple[str, Path | None, Path | None, str]] = []
    used_before: set[str] = set()
    used_after: set[str] = set()

    for key in sorted(set(before_files) & set(after_files)):
        pairs.append((key, before_files[key], after_files[key], "relative path match"))
        used_before.add(key)
        used_after.add(key)

    unmatched_before = {k: v for k, v in before_files.items() if k not in used_before}
    unmatched_after = {k: v for k, v in after_files.items() if k not in used_after}
    before_by_name: dict[str, list[tuple[str, Path]]] = {}
    for key, path in unmatched_before.items():
        before_by_name.setdefault(path.name.lower(), []).append((key, path))

    for after_key, after_path in sorted(unmatched_after.items()):
        candidates = before_by_name.get(after_path.name.lower(), [])
        available = [(k, p) for k, p in candidates if k not in used_before]
        if len(available) == 1:
            before_key, before_path = available[0]
            pairs.append((after_key, before_path, after_path, "unique filename match"))
            used_before.add(before_key)
            used_after.add(after_key)

    unmatched_before = {k: v for k, v in before_files.items() if k not in used_before}
    unmatched_after = {k: v for k, v in after_files.items() if k not in used_after}
    before_by_normalized_name: dict[str, list[tuple[str, Path]]] = {}
    for key, path in unmatched_before.items():
        before_by_normalized_name.setdefault(normalized_match_name(path), []).append((key, path))

    for after_key, after_path in sorted(unmatched_after.items()):
        candidates = before_by_normalized_name.get(normalized_match_name(after_path), [])
        available = [(k, p) for k, p in candidates if k not in used_before]
        if len(available) == 1:
            before_key, before_path = available[0]
            pairs.append((after_key, before_path, after_path, "normalized filename match"))
            used_before.add(before_key)
            used_after.add(after_key)

    for key, path in sorted(before_files.items()):
        if key not in used_before:
            pairs.append((key, path, None, "only in before folder"))
    for key, path in sorted(after_files.items()):
        if key not in used_after:
            pairs.append((key, None, path, "only in after folder"))
    return pairs


def compare_pair(key: str, before: Path | None, after: Path | None, note: str = "") -> FileResult:
    before_text = read_text(before) if before else ""
    after_text = read_text(after) if after else ""
    stats = diff_stats(before_text, after_text)
    if before and after:
        status = "matched"
    elif after:
        status = "new_file"
    else:
        status = "deleted_file"
    return FileResult(
        status=status,
        key=key,
        before_file=str(before) if before else "",
        after_file=str(after) if after else "",
        before_words=stats.before_words,
        after_words=stats.after_words,
        added_words=stats.added_words,
        deleted_words=stats.deleted_words,
        net_words=stats.net_words,
        changed_words=stats.changed_words,
        note=note,
    )


def compare_inputs(before: Path, after: Path) -> list[FileResult]:
    if before.is_dir() and after.is_dir():
        return [compare_pair(key, b, a, note) for key, b, a, note in build_pairs(before, after)]
    if before.is_file() and after.is_file():
        return [compare_pair(before.name, before, after, "single file")]
    raise ValueError("Before and after must both be files or both be folders.")


def totals(results: list[FileResult]) -> dict[str, int]:
    fields = [
        "before_words",
        "after_words",
        "added_words",
        "deleted_words",
        "net_words",
        "changed_words",
    ]
    return {field: sum(getattr(result, field) for result in results) for field in fields}


def write_csv(path: Path, results: list[FileResult]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(asdict(results[0]).keys()) if results else list(FileResult.__dataclass_fields__))
        writer.writeheader()
        for result in results:
            writer.writerow(asdict(result))


def print_table(results: list[FileResult]) -> None:
    total = totals(results)
    print("写作字数差异统计")
    print(f"比较文件数: {len(results)}")
    print(f"写作前总字数: {total['before_words']}")
    print(f"写作后总字数: {total['after_words']}")
    print(f"新增字数: {total['added_words']}")
    print(f"删除字数: {total['deleted_words']}")
    print(f"净增字数: {total['net_words']}")
    print(f"改写涉及字数: {total['changed_words']}")
    print("")
    print("逐文件结果:")
    for result in results:
        print(
            f"- {result.key} | {result.status} | 新增 {result.added_words} | "
            f"删除 {result.deleted_words} | 净增 {result.net_words} | {result.note}"
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare two document versions or folders and calculate writing word changes."
    )
    parser.add_argument("before", nargs="?", type=Path, help="Before file/folder. Default: 写作前")
    parser.add_argument("after", nargs="?", type=Path, help="After file/folder. Default: 写作后")
    parser.add_argument("--json", action="store_true", help="Print JSON output.")
    parser.add_argument("--csv", type=Path, help="Write per-file results to CSV.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(__file__).resolve().parents[1]
    before = (args.before or project_root / "写作前").resolve()
    after = (args.after or project_root / "写作后").resolve()

    if not before.exists():
        print(f"Before path not found: {before}", file=sys.stderr)
        return 2
    if not after.exists():
        print(f"After path not found: {after}", file=sys.stderr)
        return 2

    try:
        results = compare_inputs(before, after)
    except Exception as exc:
        print(f"Failed to calculate word diff: {exc}", file=sys.stderr)
        return 1

    if args.csv:
        write_csv(args.csv.resolve(), results)

    payload = {
        "before": str(before),
        "after": str(after),
        "totals": totals(results),
        "results": [asdict(result) for result in results],
    }
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print_table(results)
        if args.csv:
            print("")
            print(f"CSV 已写入: {args.csv.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
