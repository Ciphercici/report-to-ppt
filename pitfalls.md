# Common Pitfalls

Mistakes that caused previous report-to-PPT conversions to fail,
and how to avoid them.

## 1. Using Marp for report documents

**Problem**: Marp is designed for markdown files that are *already written*
as presentations (with `---` slide separators). Report documents don't have
these separators. Force-adding them without content reorganization produces
visually correct but logically broken slides. Additionally, Marp's `--pptx`
output renders each slide as a **single background image** — text is not
selectable or editable.

**Fix**: Never use Marp for report-to-PPT conversion. Use the bundled
`build_ppt.py` script instead.

## 2. Using Pandoc with --slide-level

**Problem**: Pandoc creates slide breaks at heading levels (e.g., `##`).
Report documents often have one `##` section with 30+ images and paragraphs,
and another with 2 lines. Result: some slides have 107 shapes crammed in,
others are nearly empty.

**Fix**: Don't rely on heading-level auto-splitting. Manually design the
slide outline first, then implement each slide individually.

## 3. Regex capture groups with parentheses in filenames

**Problem**: File/directory names like `Ciro Santilli (三西猴).assets/`
contain `(` and `)`. Regex patterns like `[^)]+` stop at the first `)`:

```python
# WRONG
re.findall(r'!\[.*?\]\(\./([^)]+)\)', body)
# Captures: "Ciro Santilli (三西猴" — truncated at the first paren!
```

**Fix**: Use greedy match anchored to image file extension:

```python
# RIGHT
IMG_RE = re.compile(r'!\[.*?\]\(\./(.+\.(?:png|jpg|jpeg|gif|webp|svg))\)', re.I)
```

Or better: use `pathlib.Path` to construct paths and avoid regex for path
extraction altogether. List assets directory and match by filename.

## 4. Chinese quotation marks in Python strings

**Problem**: Chinese curly double quotes `""` (U+201C / U+201D) look like
ASCII double quotes `"` in some editors. When used inside Python double-quoted
strings, they cause SyntaxError:

```python
# BROKEN — "" are actually ASCII double quotes
bullet(tf, "使中国官方更难将其简单定性为"反共人士"")
```

**Fix**: Use Chinese angle quotes `「」` instead:

```python
# WORKS
bullet(tf, '使中国官方更难将其简单定性为「反共人士」')
```

Or use Python single quotes `'...'` if the content doesn't contain ASCII single quotes.

## 5. Not verifying image embedding

**Problem**: Script runs without errors but produces a 0.1 MB PPTX — images
were never embedded because paths didn't resolve. The failure was silent.

**Fix**: Always verify after generation:
- Check file size (should be > 1 MB if images are embedded)
- Check shape counts per slide
- Verify `shape_type == 13` (Picture) counts match expectations

## 6. Skipping the outline step

**Problem**: Jumping straight to code without planning the slide structure.
This produces a "format conversion" rather than a designed presentation.

**Fix**: Always present the slide outline to the user before writing any code.
12-18 slides, each with a clear information focus.

## 7. Dumping full paragraphs onto slides

**Problem**: Copying entire paragraphs from the source document onto slides.
Slides should contain key points, not full text.

**Fix**: Each bullet point should be ≤ 2 lines. Condense paragraphs to
3-5 key points. Detailed text belongs in the report, not the slides.
