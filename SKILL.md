---
name: ops-docx-gen
description: "Word document (.docx) creation, editing, tracked changes (redlining), and text extraction. Use when: creating new Word reports/documents from scratch, editing existing .docx files, adding tracked changes or comments, extracting text content from Word files, or converting documents to markdown. Triggers on: Word文档, docx, 报告生成, Word报告, create document, edit document, tracked changes, redlining."
---

# DOCX Creation, Editing & Analysis

## Workflow Decision Tree

| Task | Approach |
|------|----------|
| **Read/extract text** | `pandoc file.docx -o output.md` |
| **Create new document** | JS: `docx` library → read `references/docx-js.md` |
| **Edit existing document** | Python: unpack → Document library → pack. Read `references/ooxml.md` |
| **Tracked changes (redline)** | Unpack → Python Document library with `<w:ins>`/`<w:del>` → pack |
| **Raw XML inspection** | `python scripts/unpack.py file.docx outdir/` |

## Creating New Documents (JavaScript)

1. **Read** `references/docx-js.md` completely before starting
2. Write JS using `Document`, `Paragraph`, `TextRun`, `Table` etc.
3. Export: `Packer.toBuffer(doc)` → `fs.writeFileSync("out.docx", buffer)`

```bash
# Dependency (installed globally or use NODE_PATH)
npm install -g docx
# Or: NODE_PATH=~/.openclaw/workspace/claude-office-skills/node_modules node script.js
```

## Editing Existing Documents (Python)

1. **Read** `references/ooxml.md` completely before starting
2. Unpack: `python scripts/unpack.py input.docx unpacked/`
3. Write Python script using Document library (see ooxml.md "Document Library" section)
4. Pack: `python scripts/pack.py unpacked/ output.docx`

## Redlining (Tracked Changes)

1. Get markdown: `pandoc --track-changes=all file.docx -o current.md`
2. Read `references/ooxml.md` — focus on "Tracked Changes" section
3. Unpack document
4. Batch changes (3-10 per batch), grep XML to map text → elements
5. **Minimal edits only**: mark only changed text, not entire sentences
6. Pack final document
7. Verify: `pandoc --track-changes=all output.docx -o verify.md`

## Key Rules

- Script paths are relative to this skill: `scripts/unpack.py`, `scripts/pack.py`
- `<w:pPr>` element order: `pStyle` → `numPr` → `spacing` → `ind` → `jc`
- Add `xml:space='preserve'` to `<w:t>` with leading/trailing spaces
- RSIDs must be 8-digit hex (0-9, A-F only)
- Images go in `word/media/`, referenced in `document.xml`

## Dependencies

| Tool | Purpose | Install |
|------|---------|---------|
| `pandoc` | Text extraction | `brew install pandoc` |
| `docx` (npm) | New document creation | `npm install -g docx` |
| `defusedxml` | Secure XML parsing | `pip install defusedxml` |
| `soffice` | PDF conversion (optional) | `brew install --cask libreoffice` |
