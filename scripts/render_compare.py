"""Render original PDF vs PowerPoint-exported PDF side-by-side contact sheets and compute pixel diffs."""

import argparse
import os
import pymupdf
from PIL import Image, ImageDraw, ImageChops


def render_page(page, dpi: int = 120) -> Image.Image:
    pix = page.get_pixmap(dpi=dpi, alpha=False)
    return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)


def compare_pdfs(orig_pdf: str, exp_pdf: str, out_dir: str, dpi: int = 120, per_sheet: int = 4):
    os.makedirs(out_dir, exist_ok=True)
    doc1 = pymupdf.open(orig_pdf)
    doc2 = pymupdf.open(exp_pdf)
    print(f"Orig pages: {len(doc1)}, Exported pages: {len(doc2)}")
    assert len(doc1) == len(doc2), f"Page count mismatch: {len(doc1)} vs {len(doc2)}"

    pairs = []
    for i in range(len(doc1)):
        im1 = render_page(doc1[i], dpi=dpi)
        im2 = render_page(doc2[i], dpi=dpi)
        im2_cmp = im2.resize(im1.size, Image.Resampling.BILINEAR) if im2.size != im1.size else im2

        diff = ImageChops.difference(im1, im2_cmp)
        raw = diff.tobytes()
        n_pix = len(raw) // 3
        mad = sum(raw) / len(raw)
        diff_px = sum(1 for j in range(0, len(raw), 3) if raw[j] > 15 or raw[j + 1] > 15 or raw[j + 2] > 15)
        pct_diff = 100.0 * diff_px / n_pix
        print(f"Slide {i + 1:2d}: MAD={mad:.2f}/255, diff_px(>15)={pct_diff:.2f}%")
        pairs.append((i + 1, im1, im2_cmp))

    pad = 16
    w, h = pairs[0][1].size
    sheet_w = w * 2 + pad * 3

    for start in range(0, len(pairs), per_sheet):
        chunk = pairs[start:start + per_sheet]
        sheet_h = (h + pad) * len(chunk) + pad
        canvas = Image.new("RGB", (sheet_w, sheet_h), (235, 235, 235))
        draw = ImageDraw.Draw(canvas)
        for r, (pno, im1, im2) in enumerate(chunk):
            y = pad + r * (h + pad)
            canvas.paste(im1, (pad, y))
            canvas.paste(im2, (pad * 2 + w, y))
            draw.rectangle([pad, y, pad + w - 1, y + h - 1], outline=(160, 160, 160), width=1)
            draw.rectangle([pad * 2 + w, y, pad * 2 + w * 2 - 1, y + h - 1], outline=(160, 160, 160), width=1)
            draw.text((pad + 4, y + 4), f"ORIG #{pno}", fill=(255, 0, 0))
            draw.text((pad * 2 + w + 4, y + 4), f"PPTX #{pno}", fill=(0, 128, 0))
        out_path = os.path.join(out_dir, f"compare_{chunk[0][0]}_{chunk[-1][0]}.png")
        canvas.save(out_path)
        print("Wrote", out_path)


def main():
    parser = argparse.ArgumentParser(description="Compare original PDF with PowerPoint-exported verification PDF.")
    parser.add_argument("orig_pdf", help="Original PDF path.")
    parser.add_argument("exp_pdf", help="PowerPoint-exported verification PDF path.")
    parser.add_argument("out_dir", help="Output directory for comparison PNG sheets.")
    parser.add_argument("--dpi", type=int, default=120, help="Render DPI (default: 120).")
    args = parser.parse_args()
    compare_pdfs(args.orig_pdf, args.exp_pdf, args.out_dir, dpi=args.dpi)


if __name__ == "__main__":
    main()
