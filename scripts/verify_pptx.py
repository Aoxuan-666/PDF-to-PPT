"""Programmatic verification gate for PDF-to-PPTX text parity and <a:ea> East Asian font tags."""

import argparse
import re
from collections import Counter
import pymupdf
from pptx import Presentation

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def norm(s: str) -> str:
    return re.sub(r"[\s•]+", "", s)


def verify(pdf_path: str, pptx_path: str):
    doc = pymupdf.open(pdf_path)
    prs = Presentation(pptx_path)

    print(f"PDF pages: {len(doc)}, PPTX slides: {len(prs.slides)}")
    print(f"PPTX size: {prs.slide_width.pt:.2f} x {prs.slide_height.pt:.2f} pt")
    assert len(doc) == len(prs.slides), f"Slide count mismatch: {len(doc)} vs {len(prs.slides)}"

    r0 = doc[0].rect
    assert abs(prs.slide_width.pt - r0.width) < 0.5, f"Width mismatch: {prs.slide_width.pt} vs {r0.width}"
    assert abs(prs.slide_height.pt - r0.height) < 0.5, f"Height mismatch: {prs.slide_height.pt} vs {r0.height}"

    total_runs = 0
    missing_ea = 0
    all_ok = True

    for i in range(len(doc)):
        pdf_norm = norm(doc[i].get_text("text"))
        pptx_para_norms = []
        pptx_all_parts = []

        for shp in prs.slides[i].shapes:
            if shp.has_text_frame:
                for p in shp.text_frame.paragraphs:
                    p_parts = []
                    for r in p.runs:
                        total_runs += 1
                        rPr = r._r.find(f"{{{A_NS}}}rPr")
                        ea = rPr.find(f"{{{A_NS}}}ea") if rPr is not None else None
                        if ea is None or not ea.get("typeface"):
                            missing_ea += 1
                        p_parts.append(r.text)
                        pptx_all_parts.append(r.text)
                    pn = norm("".join(p_parts))
                    if pn:
                        pptx_para_norms.append(pn)

        pptx_norm = norm("".join(pptx_all_parts))
        c_pdf = Counter(pdf_norm)
        c_pptx = Counter(pptx_norm)
        multiset_ok = (c_pdf == c_pptx)

        pdf_lines_missing = []
        for b in doc[i].get_text("dict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                lt = norm("".join(sp["text"] for sp in l["spans"]))
                if not lt:
                    continue
                if not any(lt in pn for pn in pptx_para_norms):
                    pdf_lines_missing.append(lt)

        if not multiset_ok or pdf_lines_missing:
            all_ok = False
            print(f"Slide {i + 1}: FAIL")
            if not multiset_ok:
                print("  Diff PDF - PPTX:", c_pdf - c_pptx)
                print("  Diff PPTX - PDF:", c_pptx - c_pdf)
            if pdf_lines_missing:
                print("  Missing lines:", pdf_lines_missing)
        else:
            print(f"Slide {i + 1:2d}: PASS (chars={len(pptx_norm)}, shapes={len(prs.slides[i].shapes)})")

    print(f"\nTotal runs: {total_runs}, missing <a:ea>: {missing_ea}")
    assert all_ok, "Text parity verification failed!"
    assert missing_ea == 0, f"{missing_ea} runs missing <a:ea>!"
    print("ALL VERIFICATION GATES PASSED.")


def main():
    parser = argparse.ArgumentParser(description="Verify PDF-to-PPTX text parity and <a:ea> XML font coverage.")
    parser.add_argument("pdf_path", help="Source PDF path.")
    parser.add_argument("pptx_path", help="Generated PPTX path.")
    args = parser.parse_args()
    verify(args.pdf_path, args.pptx_path)


if __name__ == "__main__":
    main()
