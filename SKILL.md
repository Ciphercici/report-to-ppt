---
name: report-to-ppt
description: >
  Convert report-format markdown (or docx) documents into clean, editable,
  minimalist PowerPoint presentations. Use this skill whenever the user asks
  to "convert this to PPT", "make a presentation from this report", "turn this
  document into slides", "生成PPT", "转为幻灯片", "做成PPT", or any similar
  request involving document-to-presentation conversion. Do NOT use Marp or
  pandoc for this — they produce image-based or poorly structured output.
  Instead, use the bundled python-pptx script for editable, well-designed slides.
---

# Report to PPT

Convert report-format markdown documents into minimalist, editable
PowerPoint presentations.

## Core Principle

Report-format documents cannot be "format-converted" directly into good PPT.
You must first **extract and reorganize** content into a presentation narrative,
then lay out each slide individually.

The quality path: **Claude designs the outline → script generates base PPTX →
officecli tweaks and verifies.**

## Toolchain

| Tool | Role |
|------|------|
| `build_ppt.py` | Content structure → initial PPTX (the "rough cut") |
| `officecli` | Post-generation tweaks: fix text, adjust colors, verify quality |

### Installing officecli

If `officecli` is not already installed:

```bash
# macOS / Linux
curl -fsSL https://d.officecli.ai/install.sh | bash

# Windows (PowerShell)
irm https://d.officecli.ai/install.ps1 | iex
```

Verify: `officecli --version`

## Workflow

### Phase 1: Design the slide outline

1. Read the source document thoroughly
2. Design a **12-18 slide outline** organized as a narrative:
   Cover → Summary → Background → Analysis → Evidence → Conclusion
3. Write it as a `.slide.md` file in Marp format:
   - `---` separates slides
   - `<!-- _class: lead -->` marks cover and section dividers
   - `# Title` / `## Subtitle` for slide headings
   - `- ` for bullet points
   - `![bg right:40%](path)` for images positioned on the right
   - `**bold**` for key metrics that become KPI cards
4. Each slide: one information focus, 3-5 bullet points, max 1 image
5. Present the outline to the user for approval

### Phase 2: Generate base PPTX

```bash
python <skill-path>/scripts/build_ppt.py outline.slide.md -o output.pptx
```

The script auto-applies the design system: title bars, accent colors, KPI cards,
cover/conclusion dark backgrounds, image borders and scaling, sub-header styling.

### Phase 3: Fine-tune with officecli

Use officecli to make targeted adjustments **without touching Python code**.
Natural language → CLI commands — change text, fix colors, adjust layout.

**Inspect before editing:**
```bash
officecli view output.pptx outline          # Slide-by-slide structure
officecli view output.pptx issues           # Find formatting problems
officecli get output.pptx '/slide[3]' --depth 1    # List all shapes on slide 3
```

**Common tweaks:**
```bash
# Fix text
officecli set output.pptx '/slide[2]/shape[2]' --prop text="正确的标题"

# Change colors
officecli set output.pptx '/slide[1]/shape[1]' --prop fill=1E2A38

# Adjust font size
officecli set output.pptx '/slide[3]/shape[3]' --prop font.size=18pt

# Replace text across all slides
officecli set output.pptx / --find "旧文字" --replace "新文字"

# Add a shape
officecli add output.pptx '/slide[5]' --type shape --prop text="补充内容" --prop x=2cm --prop y=6cm --prop font.size=14pt

# Remove a shape
officecli remove output.pptx '/slide[4]/shape[2]'
```

Run `officecli help pptx` for the full element/property schema. Use `officecli help pptx shape` before setting shape properties.

### Phase 4: Verify

```bash
# Quick slide count + structure check
officecli view output.pptx stats

# Auto-detect formatting issues
officecli view output.pptx issues --type format

# Manual content spot-check
officecli view output.pptx outline
```

A good result: 12-18 slides, file > 1 MB, zero `issues` flagged.

### Quick Draft (Auto Mode)

For a quick first draft without writing an outline:

```bash
python <skill-path>/scripts/build_ppt.py report.md -o output.pptx
```

Auto-parses the markdown, chunks content, generates slides. Quality is lower
than the outline path — raw text becomes bullet points verbatim. Use officecli
to clean up afterwards.

## .slide.md Format Quick Reference

```markdown
---
marp: true
size: 16:9
---

<!-- _class: lead -->
# Report Title
## Subtitle

---

<!-- _class: lead -->
## Section Name
Paragraph text for summary/lead slides.

---

## Slide Title
- Bullet point one
- Bullet point two
- **Key metric:** 3,000+ stars

![bg right:40%](./assets/image.png)
```

Key rules:
- `---` = new slide
- `<!-- _class: lead -->` = dark background (cover or section divider)
- `![bg right:40%](path)` = image on right 40% of slide
- `![bg](path)` = full-slide background image
- `**text with numbers**` = auto-detected as KPI card
- `### Sub-header` = section header within a content slide

## Design System

See `references/design-system.md` for the complete visual specification.

Quick reference:
- Slide size: 16:9 (13.333" x 7.5")
- Title: 30pt, #1A1A1A, Bold, top-left with blue accent bar
- Body: 17pt, #444444, with 32pt line spacing
- Accent: #3B82C4 (blue), #E84D4D (red for emphasis)
- KPI cards: light gray background (#F0F3F7), large number + small label
- Cover/Conclusion: dark background (#1E2A38), white text

## Common Pitfalls

See `references/pitfalls.md` for detailed examples of what NOT to do.

## Bundled Resources

| Resource | Purpose |
|----------|---------|
| `scripts/build_ppt.py` | Generate base PPTX from .slide.md outline (outline mode) or raw .md (auto mode) |
| `references/design-system.md` | Full visual specification with code examples |
| `references/pitfalls.md` | Common mistakes and how to avoid them |
