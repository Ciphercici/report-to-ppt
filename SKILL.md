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

The quality path: **Claude designs the outline, script handles the formatting.**
Claude reads the report and writes a `.slide.md` outline (Marp format) with
proper content提炼 and narrative structure. The script reads that outline
and generates well-designed PPTX automatically.

## Workflow (Quality Path)

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

### Phase 2: Generate PPTX

Run the bundled script against the `.slide.md`:

```bash
python <skill-path>/scripts/build_ppt.py outline.slide.md -o output.pptx
```

The script will:
1. Parse the Marp-format outline
2. Apply the v2 design system (title bars, accent colors, KPI cards)
3. Render cover/section-dividers with dark backgrounds
4. Place images with borders and auto-scaling
5. Convert `**metrics**` to KPI cards on summary slides
6. Handle `###` sub-headers as section titles within slides
7. Clean up table markdown into readable key-value pairs

### Phase 3: Verify

```python
from pptx import Presentation
prs = Presentation('output.pptx')
for i, s in enumerate(prs.slides):
    texts = [sh.text_frame.text[:60] for sh in s.shapes
             if sh.has_text_frame and sh.text_frame.text.strip()]
    imgs = sum(1 for sh in s.shapes if sh.shape_type == 13)
    print(f'{i+1:2d}: {len(s.shapes)} shapes, {imgs} img | {texts[0] if texts else "(empty)"}')
```

A good result: 15±3 slides, 5-25 shapes per slide, images present where expected,
file size > 1 MB.

### Quick Draft (Auto Mode)

For a quick first draft without writing an outline, run auto mode directly
on the raw report:

```bash
python <skill-path>/scripts/build_ppt.py report.md -o output.pptx
```

This auto-parses the markdown, chunks content, and generates slides. Quality
will be lower than the outline path — raw text becomes bullet points
verbatim, without Claude-level提炼.

### Manual Fine-Tuning

If specific slides need pixel-level control:

1. Copy `build_ppt.py` to the target directory
2. Edit the slide functions under `# MANUAL MODE`
3. Use helpers: `title_bar()`, `body_box()`, `bullet()`, `add_image()`,
   `kpi_card()`, `page_number()`
4. Run: `python build_ppt.py`

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
| `scripts/build_ppt.py` | Main script — outline mode, auto mode, or manual mode |
| `references/design-system.md` | Full visual specification with code examples |
| `references/pitfalls.md` | Common mistakes and how to avoid them |
