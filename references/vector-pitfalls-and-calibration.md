# Vector PDF-to-PPTX Pitfalls & Calibration Reference

Technical lessons and exact calibrations verified against WPS Presentation vector PDFs re-exported through Microsoft PowerPoint (Office 16).

## 1. Duplicate WPS Border-Expansion Paths & Occluded Shapes

- **Symptom**: Every rounded rectangle, circle, or arrow in a WPS-exported PDF is immediately followed by a second path of the exact same fill color expanded by `0.5pt` on each side and tessellated into hundreds of tiny line segments (`ops={'l': 496}` for rounded rects, `{'l': 512}` for circles, `{'l': 62}` for arrows).
- **Fix**: Filter out any drawing path where `ops.get("l", 0) >= 50`.
- **Occluded duplicates**: Some slides contain an invisible lower-layer card at the exact same `(x0, y0, x1, y1)` bbox as a colored card drawn directly on top of it (e.g., white rounded rect covered 100% by a pale teal rounded rect). Drop any earlier shape whose bbox matches a later opaque shape within `0.5pt`.

## 2. Removing PowerPoint's Default Theme Drop Shadow

- **Symptom**: `slide.shapes.add_shape(...)` in `python-pptx` inserts `<p:style>` with `<a:effectRef idx="2">`, causing PowerPoint to render drop shadows under flat cards, pills, and circles.
- **Fix** (`remove_shape_shadow` in `pptx_helpers.py`):
  1. Set `shape.shadow.inherit = False`.
  2. Set `<p:style><a:effectRef idx="0">` (`0` = no theme effect).
  3. Ensure `<p:spPr><a:effectLst/>` exists and is empty.

## 3. East Asian Font Binding (`<a:ea>`) & Font Fallbacks

`python-pptx`'s `run.font.name = ...` only writes `<a:latin typeface="..."/>`. On Chinese Windows, PowerPoint ignores `<a:latin>` for CJK characters and falls back to the theme's East Asian font (`等线` or `宋体`), corrupting layout and glyph metrics.

Always populate `<a:latin>`, `<a:ea>`, and `<a:cs>` inside `run._r.get_or_add_rPr()`:

| PDF Embedded Font | Windows `<a:latin>` | Windows `<a:ea>` / `<a:cs>` | Bold | Notes |
|---|---|---|---|---|
| `MicrosoftYaHei` | `Microsoft YaHei` | `微软雅黑` | `False` | Round `±0.05pt` sizes to nearest `0.5pt` |
| `MicrosoftYaHei-Bold` | `Microsoft YaHei` | `微软雅黑` | `True` | Also check `span["flags"] & 16` |
| `FangSong_GB2312` | `FangSong` | `仿宋` | `False` | `FangSong_GB2312` is often absent; `simfang.ttf` registers as `FangSong` / `仿宋` |
| `DengXian` | `DengXian` | `等线` | `False` | — |
| `STKaiti` | `STKaiti` | `华文楷体` | `False` | Installed as `STKAITI.TTF` |
| `MS-UIGothic` | `MS UI Gothic` | `MS UI Gothic` | `False` | Frequently embedded for non-breaking hyphen `‑` (`U+2011`) |

## 4. Exact Point Line Spacing vs. Proportional `1.2`

- **Pitfall**: Setting `paragraph.line_spacing = 1.2` (float) in PowerPoint multiplies the font's *inherent OS/2 line-box height* (`~1.29 × size` for Microsoft YaHei, `1.20 × size` for FangSong) by `1.2`, inflating `16pt` YaHei pitch from `19.2pt` to `~24.8pt` and overflowing cards vertically!
- **Fix**: WPS "1.2×" line spacing in exported PDFs uses `1.20 × font_size_pt` directly. Set **explicit point line spacing** (`<a:spcPts>` via `paragraph.line_spacing = Pt(pitch_pt)`):
  - Compute `pitch_pt` directly from consecutive lines in multi-line PDF blocks (`lines[i+1].y0 - lines[i].y0`), or fall back to `round(font_size_pt * 1.2, 2)`.
  - For bullet lists using `Pt(font_size_pt * 1.2)`, `space_after = Pt(8.0)` reproduces WPS spacing across `14pt`, `15pt`, and `16pt` lists.

## 5. Vertical Offset (`y_offset`) Calibration Table

When `margin_top = 0` and fixed `Pt(pitch_pt)` line spacing is used, PowerPoint shifts the visual top of the first line slightly depending on font metrics and `pitch_pt` vs. natural line height:

| Font & Context | Observed Drift Without Offset | Recommended `y_offset` |
|---|---|---|
| `Microsoft YaHei Bold` `48pt`, `Pt(57.6)` | `~7.2pt` too high | `+7.2pt` |
| `Microsoft YaHei Bold` `32pt`, `Pt(38.4)` | `~3.5pt` too high | `+3.5pt` |
| `Microsoft YaHei Bold` `28pt` single-line (`word_wrap=False`) | `~2.7pt` too high | `+2.7pt` |
| `Microsoft YaHei` `14–16pt` bullets / single lines | `~1.4–1.5pt` too high | `+1.4pt` to `+1.5pt` |
| `FangSong` `16pt`, `Pt(19.2)` | `~3.2pt` too low | `-3.2pt` |
| `FangSong` `12pt`, `Pt(14.4)` | `~2.6pt` too low | `-2.6pt` |
| `DengXian` `18pt` single-line | `~2.2pt` too low | `-2.2pt` |
| `STKaiti` `14pt`, `Pt(25.1)` (reference list) | `~4.9pt` too low | `-4.9pt` |
| `STKaiti` `16pt` single-line | `~1.5pt` too low | `-1.5pt` |

For multi-line `2-line header` blocks (`12pt` kicker + `28pt` title), use `p0.line_spacing = Pt(15.8)` and `p0.space_after = Pt(round(gap + 3.2, 1))` where `gap = title.y0 - kicker.y1`.

## 6. Native Hanging-Indent Bullets

Represent bullet points (`•`) inside ONE grouped textbox using `<a:pPr>` rather than separate `•` textboxes:

```xml
<a:pPr marL="342900" indent="-342900">
  <a:buClr><a:srgbClr val="222B3A"/></a:buClr>
  <a:buFont typeface="Microsoft YaHei"/>
  <a:buChar char="•"/>
</a:pPr>
```

where `marL = (body_x0 - bullet_x0) * 12700` EMU (`27pt = 342900 EMU`) and `indent = -marL`.

## 7. Preventing Unintended Line Wraps

- **Single-line blocks**: Use `word_wrap = False` and width `max_lw + first_sz * 1.5`.
- **Wrapped CJK paragraphs**: Use `word_wrap = True` and width `max_lw + first_sz * 0.45` (enough margin so the longest line does not prematurely wrap in PowerPoint, while staying under `+1.0 * first_sz` so lines break consistently).
