from __future__ import annotations

import argparse
import os
from pathlib import Path

import mammoth


def _guess_ext(content_type: str) -> str:
    content_type = (content_type or "").lower().strip()
    if content_type == "image/jpeg":
        return ".jpg"
    if content_type == "image/png":
        return ".png"
    if content_type == "image/gif":
        return ".gif"
    if content_type == "image/bmp":
        return ".bmp"
    if content_type == "image/tiff":
        return ".tiff"
    if content_type == "image/webp":
        return ".webp"
    return ""


def convert(docx_path: Path, output_html_path: Path, images_dir: Path) -> None:
    docx_path = docx_path.resolve()
    output_html_path = output_html_path.resolve()
    images_dir = images_dir.resolve()

    images_dir.mkdir(parents=True, exist_ok=True)
    output_html_path.parent.mkdir(parents=True, exist_ok=True)

    image_index = 0

    def convert_image(image: mammoth.images.Image):
        nonlocal image_index
        image_index += 1
        ext = _guess_ext(getattr(image, "content_type", "") or "")
        if not ext:
            ext = ""
        filename = f"word-{image_index:03d}{ext}"
        output_path = images_dir / filename
        with output_path.open("wb") as f:
            with image.open() as image_bytes:
                f.write(image_bytes.read())
        rel_path = os.path.relpath(output_path, output_html_path.parent)
        rel_path = rel_path.replace("\\", "/")
        return {"src": rel_path}

    with docx_path.open("rb") as docx_file:
        result = mammoth.convert_to_html(
            docx_file,
            convert_image=mammoth.images.inline(convert_image),
        )

    html_fragment = result.value
    html_fragment = html_fragment.replace("><", ">\n<")
    html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{docx_path.stem}</title>
  <style>
    html, body {{
      margin: 0;
      padding: 0;
      background: #fff;
      color: #000;
      font-family: \"Microsoft YaHei\", \"PingFang SC\", Arial, sans-serif;
      font-size: 16px;
      line-height: 1.75;
    }}
    .page {{
      max-width: 900px;
      margin: 0 auto;
      padding: 48px 24px 80px;
    }}
    p {{
      margin: 0 0 12px;
      white-space: pre-wrap;
    }}
    img {{
      max-width: 100%;
      height: auto;
    }}
    table {{
      border-collapse: collapse;
      width: 100%;
    }}
    td, th {{
      border: 1px solid #ddd;
      padding: 6px 10px;
      vertical-align: top;
    }}
    ul, ol {{
      margin: 0 0 12px 24px;
      padding: 0;
    }}
  </style>
</head>
<body>
  <main class="page">
    {html_fragment}
  </main>
</body>
</html>
"""

    output_html_path.write_text(html, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True)
    parser.add_argument("--out-html", required=True)
    parser.add_argument("--images-dir", required=True)
    args = parser.parse_args()

    convert(
        docx_path=Path(args.docx),
        output_html_path=Path(args.out_html),
        images_dir=Path(args.images_dir),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

