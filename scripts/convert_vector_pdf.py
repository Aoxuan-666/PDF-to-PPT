"""Generic vector PDF-to-PPTX converter preserving native shapes, <a:ea> fonts, and semantic text grouping."""

import argparse
import os
import re
import sys
import pymupdf
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from pptx_helpers import (
    add_runs_to_paragraph,
    add_textbox,
    add_vector_shapes,
    block_all_spans,
    default_font_y_offset,
    exact_pitch_pt,
    extract_clean_shapes,
    norm_size,
    place_badge_in_oval,
    place_bullet_group,
    place_header_block,
    place_simple_block,
    place_text_in_banner,
)


def line_text(line) -> str:
    return "".join(sp["text"] for sp in line["spans"])


def block_text(block) -> str:
    return "".join(line_text(l) for l in block["lines"])


def is_bullet_block(block) -> bool:
    lines = block["lines"]
    return len(lines) >= 2 and len(lines[0]["spans"]) == 1 and lines[0]["spans"][0]["text"].strip() == "•"


def split_paragraphs_in_block(block) -> list[list]:
    """Detect intentional paragraph breaks inside a multi-line block vs automatic word wrapping."""
    lines = block["lines"]
    if len(lines) <= 1:
        return [lines]
    max_lw = max(l["bbox"][2] - l["bbox"][0] for l in lines)
    first_sz = norm_size(lines[0]["spans"][0]["size"])
    groups = []
    cur = [lines[0]]
    for i in range(1, len(lines)):
        prev = lines[i - 1]
        prev_w = prev["bbox"][2] - prev["bbox"][0]
        next_txt = line_text(lines[i]).lstrip()
        # A short line whose next line's first character would easily fit on it indicates a hard break
        if prev_w < max_lw - first_sz * 1.8 and next_txt and (prev_w + first_sz * 1.1) < max_lw:
            groups.append(cur)
            cur = [lines[i]]
        else:
            cur.append(lines[i])
    groups.append(cur)
    return groups


def point_in_rect(cx: float, cy: float, rect: tuple[float, float, float, float], pad: float = 1.0) -> bool:
    x0, y0, x1, y1 = rect
    return (x0 - pad) <= cx <= (x1 + pad) and (y0 - pad) <= cy <= (y1 + pad)


def convert_pdf_to_pptx(pdf_path: str, out_pptx_path: str):
    doc = pymupdf.open(pdf_path)
    prs = Presentation()
    r0 = doc[0].rect
    slide_w, slide_h = float(r0.width), float(r0.height)
    prs.slide_width = Pt(slide_w)
    prs.slide_height = Pt(slide_h)
    blank_layout = prs.slide_layouts[6]

    for page in doc:
        slide = prs.slides.add_slide(blank_layout)
        pw, ph = float(page.rect.width), float(page.rect.height)
        bg_rgb, shapes = extract_clean_shapes(page, pw, ph)

        bg_fill = slide.background.fill
        bg_fill.solid()
        bg_fill.fore_color.rgb = RGBColor(*bg_rgb)

        pptx_shapes = add_vector_shapes(slide, shapes)
        blocks = [b for b in page.get_text("dict")["blocks"] if b["type"] == 0]

        # Track consumed lines as (block_idx, line_idx)
        consumed_lines: set[tuple[int, int]] = set()
        consumed_blocks: set[int] = set()

        # 1. Match short badge lines centered inside OVAL shapes
        for si, s in enumerate(shapes):
            if s["kind"] != "circle" or pptx_shapes[si] is None:
                continue
            candidates = []
            for bi, b in enumerate(blocks):
                for li, l in enumerate(b["lines"]):
                    if (bi, li) in consumed_lines:
                        continue
                    txt = line_text(l).strip()
                    if not txt or len(txt) > 4:
                        continue
                    lx0, ly0, lx1, ly1 = l["bbox"]
                    cx, cy = (lx0 + lx1) / 2.0, (ly0 + ly1) / 2.0
                    if point_in_rect(cx, cy, s["rect"], pad=1.0):
                        candidates.append((bi, li, l))
            if len(candidates) == 1:
                bi, li, l = candidates[0]
                place_badge_in_oval(pptx_shapes[si], l)
                consumed_lines.add((bi, li))

        # 2. Match single-line banners/pills vertically centered inside otherwise-empty ROUNDED_RECTANGLE shapes
        for si, s in enumerate(shapes):
            if s["kind"] != "rounded_rect" or pptx_shapes[si] is None or s["h"] > 80.0:
                continue
            rx0, ry0, rx1, ry1 = s["rect"]
            rcx, rcy = (rx0 + rx1) / 2.0, (ry0 + ry1) / 2.0
            inside = []
            for bi, b in enumerate(blocks):
                for li, l in enumerate(b["lines"]):
                    if (bi, li) in consumed_lines:
                        continue
                    lx0, ly0, lx1, ly1 = l["bbox"]
                    cx, cy = (lx0 + lx1) / 2.0, (ly0 + ly1) / 2.0
                    if point_in_rect(cx, cy, s["rect"], pad=0.5):
                        inside.append((bi, li, l))
            if len(inside) == 1:
                bi, li, l = inside[0]
                lx0, ly0, lx1, ly1 = l["bbox"]
                lcx, lcy = (lx0 + lx1) / 2.0, (ly0 + ly1) / 2.0
                if abs(lcy - rcy) <= 4.5:
                    if abs(lcx - rcx) <= 12.0:
                        place_text_in_banner(pptx_shapes[si], l["spans"], ml_pt=10.0, align=PP_ALIGN.CENTER)
                    else:
                        ml = max(0.0, round(lx0 - rx0, 1))
                        place_text_in_banner(pptx_shapes[si], l["spans"], ml_pt=ml, align=PP_ALIGN.LEFT)
                    consumed_lines.add((bi, li))

        # Mark blocks as fully consumed if all their lines were consumed
        for bi, b in enumerate(blocks):
            if all((bi, li) in consumed_lines for li in range(len(b["lines"]))):
                consumed_blocks.add(bi)

        # Helper to check if block has zero consumed lines
        def is_clean_block(bi: int) -> bool:
            return bi not in consumed_blocks and not any((bi, li) in consumed_lines for li in range(len(blocks[bi]["lines"])))

        # 3. Detect consecutive bullet blocks in the same column
        bi = 0
        while bi < len(blocks):
            if is_clean_block(bi) and is_bullet_block(blocks[bi]):
                grp = [bi]
                bx0 = blocks[bi]["lines"][0]["bbox"][0]
                bsz = norm_size(blocks[bi]["lines"][1]["spans"][0]["size"])
                bj = bi + 1
                while bj < len(blocks):
                    if (
                        is_clean_block(bj)
                        and is_bullet_block(blocks[bj])
                        and abs(blocks[bj]["lines"][0]["bbox"][0] - bx0) < 2.0
                        and abs(norm_size(blocks[bj]["lines"][1]["spans"][0]["size"]) - bsz) < 1.0
                        and (blocks[bj]["bbox"][1] - blocks[grp[-1]]["bbox"][3]) < 25.0
                    ):
                        grp.append(bj)
                        bj += 1
                    else:
                        break
                # Measure gap between bullets
                if len(grp) >= 2:
                    b0_lines = blocks[grp[0]]["lines"][1:]
                    b1_first_y = blocks[grp[1]]["lines"][1]["bbox"][1]
                    pitch_pt = round(bsz * 1.2, 2)
                    expected_next_y = b0_lines[0]["bbox"][1] + len(b0_lines) * pitch_pt
                    gap = max(0.0, round(b1_first_y - expected_next_y, 1))
                else:
                    gap = 8.0
                place_bullet_group(slide, [blocks[k] for k in grp], space_after_pt=gap)
                for k in grp:
                    consumed_blocks.add(k)
                bi = bj
                continue
            bi += 1

        # 4. Detect 2-line headers (small kicker + large title near top)
        for bi, b in enumerate(blocks):
            if not is_clean_block(bi):
                continue
            lines = b["lines"]
            if len(lines) == 2 and b["bbox"][1] < 60.0:
                sz0 = norm_size(lines[0]["spans"][0]["size"])
                sz1 = norm_size(lines[1]["spans"][0]["size"])
                if sz1 >= sz0 * 1.8:
                    place_header_block(slide, b)
                    consumed_blocks.add(bi)

        # 5. Detect consecutive centered cover-style title lines (large font >= 30pt, same center X)
        bi = 0
        while bi < len(blocks) - 1:
            b1, b2 = blocks[bi], blocks[bi + 1]
            if (
                is_clean_block(bi)
                and is_clean_block(bi + 1)
                and len(b1["lines"]) == 1
                and len(b2["lines"]) == 1
            ):
                sz1 = norm_size(b1["lines"][0]["spans"][0]["size"])
                sz2 = norm_size(b2["lines"][0]["spans"][0]["size"])
                cx1 = (b1["bbox"][0] + b1["bbox"][2]) / 2.0
                cx2 = (b2["bbox"][0] + b2["bbox"][2]) / 2.0
                dy = b2["bbox"][1] - b1["bbox"][1]
                if sz1 >= 30.0 and abs(sz1 - sz2) < 1.0 and abs(cx1 - cx2) < 5.0 and abs(dy - sz1 * 1.2) < 6.0:
                    pitch_pt = round(dy, 2)
                    left_margin = min(36.0, b1["bbox"][0], b2["bbox"][0])
                    box_w = max(b1["bbox"][2] - b1["bbox"][0], b2["bbox"][2] - b2["bbox"][0])
                    # Center textbox horizontally around cx1
                    half_w = max(cx1 - left_margin, box_w / 2.0 + 20.0)
                    tb_x = cx1 - half_w
                    tb_w = half_w * 2.0
                    y_off = default_font_y_offset(b1["lines"][0]["spans"][0]["font"], sz1, True, pitch_pt)
                    tb = add_textbox(slide, tb_x, b1["bbox"][1] + y_off, tb_w, (b2["bbox"][3] - b1["bbox"][1]) + 14.0)
                    tf = tb.text_frame
                    for idx_l, blk in enumerate((b1, b2)):
                        p = tf.paragraphs[0] if idx_l == 0 else tf.add_paragraph()
                        p.alignment = PP_ALIGN.CENTER
                        p.line_spacing = Pt(pitch_pt)
                        add_runs_to_paragraph(p, block_all_spans(blk))
                    consumed_blocks.add(bi)
                    consumed_blocks.add(bi + 1)
                    bi += 2
                    continue
            bi += 1

        # 6. Detect consecutive left-aligned reference/list paragraph blocks (>= 4 single-line blocks with steady pitch)
        bi = 0
        while bi < len(blocks):
            if is_clean_block(bi) and len(blocks[bi]["lines"]) == 1 and re.match(r"^\[\d+\]", block_text(blocks[bi]).strip()):
                grp = [bi]
                x0_ref = blocks[bi]["bbox"][0]
                bj = bi + 1
                while bj < len(blocks):
                    if (
                        is_clean_block(bj)
                        and len(blocks[bj]["lines"]) == 1
                        and abs(blocks[bj]["bbox"][0] - x0_ref) < 3.0
                        and 15.0 <= (blocks[bj]["bbox"][1] - blocks[grp[-1]]["bbox"][1]) <= 35.0
                    ):
                        grp.append(bj)
                        bj += 1
                    else:
                        break
                if len(grp) >= 3:
                    pitches = [
                        blocks[grp[k + 1]]["bbox"][1] - blocks[grp[k]]["bbox"][1]
                        for k in range(len(grp) - 1)
                    ]
                    avg_pitch = round(sum(pitches) / len(pitches), 1)
                    first_b = blocks[grp[0]]
                    last_b = blocks[grp[-1]]
                    first_sp = first_b["lines"][0]["spans"][0]
                    y_off = default_font_y_offset(first_sp["font"], first_sp["size"], True, avg_pitch)
                    max_rw = max(blocks[k]["bbox"][2] - blocks[k]["bbox"][0] for k in grp)
                    tb = add_textbox(
                        slide,
                        first_b["bbox"][0],
                        first_b["bbox"][1] + y_off,
                        max(max_rw + 10.0, 630.0),
                        (last_b["bbox"][3] - first_b["bbox"][1]) + 20.0,
                        word_wrap=True,
                    )
                    tf = tb.text_frame
                    para_idx = -1
                    for k in grp:
                        txt_k = block_text(blocks[k]).strip()
                        if para_idx < 0 or re.match(r"^\[\d+\]", txt_k):
                            para_idx += 1
                            p = tf.paragraphs[0] if para_idx == 0 else tf.add_paragraph()
                            p.line_spacing = Pt(avg_pitch)
                            p.space_after = Pt(0)
                        add_runs_to_paragraph(p, block_all_spans(blocks[k]))
                        consumed_blocks.add(k)
                    bi = bj
                    continue
            bi += 1

        # 7. Place remaining clean blocks and unconsumed lines from partially-consumed blocks
        for bi, b in enumerate(blocks):
            if bi in consumed_blocks:
                continue
            if is_clean_block(bi):
                lines = b["lines"]
                if len(lines) == 1:
                    place_simple_block(slide, b, word_wrap=False)
                else:
                    para_groups = split_paragraphs_in_block(b)
                    if len(para_groups) == 1:
                        place_simple_block(slide, b, join_lines=True)
                    else:
                        first_sp = lines[0]["spans"][0]
                        first_sz = norm_size(first_sp["size"])
                        pitch_pt = exact_pitch_pt(lines, first_sz)
                        max_lw = max(l["bbox"][2] - l["bbox"][0] for l in lines)
                        y_off = default_font_y_offset(first_sp["font"], first_sz, True, pitch_pt)
                        tb = add_textbox(
                            slide,
                            b["bbox"][0],
                            b["bbox"][1] + y_off,
                            max_lw + first_sz * 0.45,
                            (b["bbox"][3] - b["bbox"][1]) + 10.0,
                            word_wrap=True,
                        )
                        tf = tb.text_frame
                        for pi, pg in enumerate(para_groups):
                            p = tf.paragraphs[0] if pi == 0 else tf.add_paragraph()
                            p.line_spacing = Pt(pitch_pt)
                            add_runs_to_paragraph(p, [sp for l in pg for sp in l["spans"]])
            else:
                for li, l in enumerate(b["lines"]):
                    if (bi, li) in consumed_lines:
                        continue
                    lx0, ly0, lx1, ly1 = l["bbox"]
                    sp0 = l["spans"][0]
                    lsz = norm_size(sp0["size"])
                    y_off = default_font_y_offset(sp0["font"], lsz, False)
                    tb = add_textbox(
                        slide, lx0, ly0 + y_off, (lx1 - lx0) + lsz * 1.5, (ly1 - ly0) + 6.0, word_wrap=False
                    )
                    add_runs_to_paragraph(tb.text_frame.paragraphs[0], l["spans"])

    os.makedirs(os.path.dirname(os.path.abspath(out_pptx_path)), exist_ok=True)
    prs.save(out_pptx_path)
    print(f"Saved PPTX to: {out_pptx_path}")


def main():
    parser = argparse.ArgumentParser(description="Convert a vector slide PDF to an editable PPTX.")
    parser.add_argument("pdf_path", help="Source PDF path.")
    parser.add_argument("pptx_path", nargs="?", help="Output PPTX path (defaults to same stem .pptx).")
    args = parser.parse_args()
    out_path = args.pptx_path or (os.path.splitext(args.pdf_path)[0] + ".pptx")
    convert_pdf_to_pptx(args.pdf_path, out_path)


if __name__ == "__main__":
    main()
