"""Inspect a PDF's page dimensions, vector shapes, text blocks, and installed Windows CJK fonts."""

import argparse
import collections
import sys
import pymupdf


def list_windows_cjk_fonts():
    """Return installed Windows fonts matching common CJK and UI families."""
    if sys.platform != "win32":
        return []
    import winreg
    keywords = ("yahei", "微软雅黑", "fangsong", "仿宋", "dengxian", "等线", "kaiti", "楷体", "simsun", "宋体", "simhei", "黑体", "gothic", "noto")
    matches = []
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
        except OSError:
            continue
        i = 0
        while True:
            try:
                name, val, _ = winreg.EnumValue(key, i)
                low = name.lower()
                if any(k in low for k in keywords):
                    matches.append((name, val))
                i += 1
            except OSError:
                break
    return sorted(set(matches))


def inspect_pdf(pdf_path: str, detail: bool = False):
    doc = pymupdf.open(pdf_path)
    print(f"File: {pdf_path}")
    print(f"Pages: {len(doc)}")
    print(f"Metadata: {doc.metadata}")

    all_fonts = set()
    total_text_chars = 0
    total_images = 0
    total_drawings = 0

    for idx, page in enumerate(doc):
        r = page.rect
        imgs = page.get_images(full=True)
        drawings = page.get_drawings()
        text = page.get_text("text").strip()
        fonts = page.get_fonts()
        for f in fonts:
            all_fonts.add((f[3], f[1]))
        total_text_chars += len(text)
        total_images += len(imgs)
        total_drawings += len(drawings)
        ratio = r.width / r.height if r.height else 0.0
        print(
            f"Page {idx + 1:2d}: {r.width:.1f}x{r.height:.1f} pt (ratio {ratio:.3f}), "
            f"text_len={len(text)}, imgs={len(imgs)}, drawings={len(drawings)}, fonts={len(fonts)}"
        )

        if detail:
            for di, d in enumerate(drawings):
                ops = collections.Counter(it[0] for it in d["items"])
                if ops.get("l", 0) >= 50:
                    continue
                dr = tuple(round(x, 2) for x in d["rect"])
                fcol = tuple(round(c * 255) for c in d["fill"]) if d.get("fill") else None
                scol = tuple(round(c * 255) for c in d["color"]) if d.get("color") else None
                print(f"  D{di}: type={d.get('type')} rect={dr} fill={fcol} stroke={scol} ops={dict(ops)}")

            tb = page.get_text("dict")
            for bi, b in enumerate(tb["blocks"]):
                if b["type"] != 0:
                    print(f"  B{bi}: non-text type={b['type']}")
                    continue
                bb = tuple(round(x, 2) for x in b["bbox"])
                print(f"  B{bi} bbox={bb}:")
                for li, l in enumerate(b["lines"]):
                    lbb = tuple(round(x, 2) for x in l["bbox"])
                    spans_desc = []
                    for s in l["spans"]:
                        c = s["color"]
                        rgb = ((c >> 16) & 255, (c >> 8) & 255, c & 255)
                        fn = s["font"].split("+")[-1]
                        spans_desc.append(f"[{fn} {s['size']:.2f}pt rgb{rgb} f={s['flags']} {s['text']!r}]")
                    print(f"    L{li} {lbb}: " + " ".join(spans_desc))

    print("\nEmbedded PDF fonts:", sorted(all_fonts))
    if total_text_chars == 0 and total_images > 0:
        print("Route recommendation: nature-image2ppt (scanned/image-only PDF detected)")
    else:
        print("Route recommendation: pdf-to-ppt (vector/selectable-text PDF detected)")

    cjk_fonts = list_windows_cjk_fonts()
    if cjk_fonts:
        print("\nInstalled Windows CJK/UI fonts:")
        for name, val in cjk_fonts:
            print(f"  {name} -> {val}")


def main():
    parser = argparse.ArgumentParser(description="Inspect PDF structure for PPTX conversion triage.")
    parser.add_argument("pdf_path", help="Path to source PDF file.")
    parser.add_argument("--detail", action="store_true", help="Print per-page drawings and text spans.")
    args = parser.parse_args()
    inspect_pdf(args.pdf_path, detail=args.detail)


if __name__ == "__main__":
    main()
