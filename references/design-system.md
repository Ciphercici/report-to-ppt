# Design System

Complete visual specification for report-to-PPT slides.

## Canvas

| Property | Value |
|----------|-------|
| Slide size | 16:9 widescreen |
| Dimensions | 13.333" × 7.5" (12192000 × 6858000 EMU) |
| Background | #FFFFFF (white) |

## Color Palette

| Token | Hex | Usage |
|-------|-----|-------|
| `dark` | #1A1A1A | Slide titles, key text |
| `body` | #444444 | Body text, bullet points |
| `muted` | #888888 | Subtitles, page numbers, secondary info |
| `accent` | #3B82C4 | Title accent bar, KPI values, section headers |
| `accent2` | #E84D4D | Warnings, emphasis on negative findings |
| `light` | #F0F3F7 | Card backgrounds |
| `white` | #FFFFFF | Text on dark backgrounds |
| `cover_bg` | #1E2A38 | Cover and conclusion slide backgrounds |
| `cover_sub` | #8A9BAE | Subtitle text on dark backgrounds |
| `border` | #DDDDDD | Image borders, subtle separators |

## Typography

| Element | Size | Weight | Color | Notes |
|---------|------|--------|-------|-------|
| Slide title | 30pt | Bold | `dark` | Top of slide, with blue accent bar above |
| Section header | 20pt | Bold | `accent` | Within body, e.g. "1. 信息档案化" |
| Body text | 17pt | Regular | `body` | Line spacing 32pt |
| Bullet points | 17pt | Regular | `body` | Prefixed with "  " for visual indent |
| Secondary text | 14pt | Regular | `muted` | Subtitles, captions |
| KPI value | 36pt | Bold | `accent` | Inside KPI cards |
| KPI label | 11pt | Regular | `muted` | Below KPI value |
| Page number | 10pt | Regular | `muted` | Bottom-right corner |
| Cover title | 54pt | Bold | `white` | Main title on dark cover |
| Cover subtitle | 36pt | Bold | `accent` | Secondary title on cover |
| Cover meta | 22pt / 14pt | Regular | `cover_sub` | Report type and date |

## Layout Grid

All measurements in inches from top-left corner.

### Title bar (every content slide)
```
y=0.15  Title text (30pt Bold, #1A1A1A)
y=0.55  Blue accent bar (0.8" wide, 0.04" tall, #3B82C4)
y=0.70  Optional subtitle (14pt, #888888)
```

### Content area
```
x=0.8, y=1.2, width=11.5, height=5.5  (default)
x=0.8, y=1.2, width=5.5, height=5.5   (with right-side image)
```

### KPI cards
```
2.5" × 1.3" cards, #F0F3F7 background
Value: 36pt inside, Label: 11pt below
Spacing: 0.3" between cards
```

### Image placement
- Right side: x=7.0, y=1.3, max 5.6" × 5.2"
- Full width: x=0.8, y=1.3, max 11.7" × 5.8"
- Auto-scaled to fit bounding box with preserved aspect ratio
- 0.5pt border in `border` color

### Page number
```
x=12.3, y=7.05, width=0.8, height=0.35
10pt, #888888, right-aligned
```

## Slide Types

### Cover slide
- Dark background (#1E2A38)
- Vertical blue accent bar on left
- Large white title + blue subtitle
- Report type and date at bottom

### Summary slide
- 4 KPI cards in a row at top
- 3–5 bullet points below
- Standard title bar

### Content slide (text + image)
- Title bar
- Left: bullet points (width ~5.5")
- Right: single image with border

### Content slide (full text)
- Title bar
- Section headers (20pt, blue, bold)
- Bullet points under each section
- Spacers between sections

### Section divider
- Clean white background
- Blue vertical accent bar
- Large section title centered

### Conclusion slide
- Same dark background as cover
- Blue accent bar
- Key takeaway text in white
- Quoted statement in muted color

## Implementation

All helpers are defined in `scripts/build_ppt.py`. Key functions:

```python
title_bar(slide, title, subtitle=None)    # Standard title area
body_box(slide, top, left, width, height)  # Content text frame (returns tf)
bullet(tf, text, size, color)             # Add bullet point
para(tf, text, size, color, bold, ...)    # Add paragraph to text frame
kpi_card(slide, left, top, value, label)  # KPI stat card
add_image(slide, filename, l, t, w, h)    # Image with auto-scaling
page_number(slide)                        # Page number bottom-right
new_slide()                               # Create blank white slide
```
