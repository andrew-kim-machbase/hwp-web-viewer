#!/usr/bin/env python3
"""Compare preview page PNGs against PDF-rendered pages and emit JSON metrics."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

try:
    from PIL import Image, ImageChops
except ImportError as exc:  # pragma: no cover
    raise SystemExit("Pillow is required. Install with: .venv/bin/pip install pillow") from exc

try:
    import pypdfium2 as pdfium
except ImportError as exc:  # pragma: no cover
    raise SystemExit("pypdfium2 is required. Install with: .venv/bin/pip install pypdfium2") from exc


def compute_metrics(preview: Image.Image, reference: Image.Image) -> dict[str, float]:
    diff = ImageChops.difference(preview, reference)
    hist = diff.histogram()
    total_channels = preview.size[0] * preview.size[1] * 3
    sq_sum = 0
    abs_sum = 0
    for idx, count in enumerate(hist):
        val = idx % 256
        sq_sum += count * (val**2)
        abs_sum += count * val
    rms = (sq_sum / total_channels) ** 0.5
    mae = abs_sum / total_channels
    return {"rms": rms, "mae": mae}


def render_pdf_page(pdf: pdfium.PdfDocument, page_index: int, target_size: tuple[int, int]) -> Image.Image:
    page = pdf[page_index]
    bitmap = page.render(scale=2.0)
    image = bitmap.to_pil().convert("RGB")
    return image.resize(target_size, Image.Resampling.LANCZOS)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--preview-dir", required=True, help="Directory containing preview-page-*.png files")
    parser.add_argument("--pdf", required=True, help="Path to reference PDF")
    parser.add_argument("--out-json", required=True, help="Output JSON report path")
    parser.add_argument("--max-pages", type=int, default=20)
    parser.add_argument("--save-rendered", action="store_true", help="Save resized rendered PDF pages")
    args = parser.parse_args()

    preview_dir = Path(args.preview_dir)
    out_json = Path(args.out_json)
    out_json.parent.mkdir(parents=True, exist_ok=True)

    preview_pages = sorted(preview_dir.glob("preview-page-*.png"))
    if not preview_pages:
        raise SystemExit(f"No preview page images found in: {preview_dir}")

    pdf = pdfium.PdfDocument(str(args.pdf))
    pdf_page_count = len(pdf)
    compare_count = min(len(preview_pages), pdf_page_count, max(1, args.max_pages))

    results = []
    rms_sum = 0.0
    mae_sum = 0.0

    rendered_dir = preview_dir.parent / "pdf-rendered"
    if args.save_rendered:
        rendered_dir.mkdir(parents=True, exist_ok=True)

    for i in range(compare_count):
        preview_img = Image.open(preview_pages[i]).convert("RGB")
        pdf_img = render_pdf_page(pdf, i, preview_img.size)
        if args.save_rendered:
            pdf_img.save(rendered_dir / f"pdf-page-{i + 1:03d}.png")

        metrics = compute_metrics(preview_img, pdf_img)
        rms_sum += metrics["rms"]
        mae_sum += metrics["mae"]
        results.append(
            {
                "page": i + 1,
                "preview_image": preview_pages[i].name,
                "rms": round(metrics["rms"], 4),
                "mae": round(metrics["mae"], 4),
            }
        )

    report = {
        "pdf": str(Path(args.pdf)),
        "preview_dir": str(preview_dir),
        "preview_pages_total": len(preview_pages),
        "pdf_pages_total": pdf_page_count,
        "pages_compared": compare_count,
        "avg_rms": round(rms_sum / compare_count, 4),
        "avg_mae": round(mae_sum / compare_count, 4),
        "pages": results,
    }

    out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
