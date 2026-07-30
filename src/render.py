"""Render a .pptx to per-slide PNGs plus contact sheets, for visual QA."""

import os
import subprocess
import sys
import glob

SOFFICE = "/opt/libreoffice26.2/program/soffice"


def to_pdf(pptx, outdir):
    os.makedirs(outdir, exist_ok=True)
    env = dict(os.environ, HOME="/root")
    subprocess.run([SOFFICE, "--headless", "--norestore", "--convert-to", "pdf",
                    "--outdir", outdir, pptx], check=True, env=env,
                   stdout=subprocess.DEVNULL, timeout=600)
    return os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")


def to_pngs(pdf, outdir, dpi=110, prefix="slide"):
    import fitz
    os.makedirs(outdir, exist_ok=True)
    doc = fitz.open(pdf)
    paths = []
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=dpi)
        p = os.path.join(outdir, f"{prefix}_{i:02d}.png")
        pix.save(p)
        paths.append(p)
    return paths


def contact_sheet(pngs, out, cols=3, thumb_w=620, label=True):
    from PIL import Image, ImageDraw
    ims = [Image.open(p).convert("RGB") for p in pngs]
    w, h = ims[0].size
    tw = thumb_w
    th = int(h * tw / w)
    pad, top = 12, 22 if label else 0
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * (tw + pad) + pad,
                             rows * (th + pad + top) + pad), (235, 238, 243))
    dr = ImageDraw.Draw(sheet)
    for i, im in enumerate(ims):
        r, c = divmod(i, cols)
        x = pad + c * (tw + pad)
        y = pad + r * (th + pad + top)
        if label:
            dr.text((x + 3, y + 4), f"Slide {i + 1}", fill=(40, 50, 70))
        sheet.paste(im.resize((tw, th), Image.LANCZOS), (x, y + top))
        dr.rectangle([x, y + top, x + tw, y + top + th], outline=(190, 198, 210))
    sheet.save(out)
    return out


def render(pptx, workdir, dpi=110, sheet_cols=3, per_sheet=9):
    name = os.path.splitext(os.path.basename(pptx))[0]
    base = os.path.join(workdir, name.replace(" ", "_"))
    pdf = to_pdf(pptx, base)
    pngs = to_pngs(pdf, base, dpi=dpi)
    sheets = []
    for i in range(0, len(pngs), per_sheet):
        out = f"{base}/sheet_{i // per_sheet + 1}.png"
        contact_sheet(pngs[i:i + per_sheet], out, cols=sheet_cols)
        sheets.append(out)
    return pdf, pngs, sheets


if __name__ == "__main__":
    pptx = sys.argv[1]
    workdir = sys.argv[2] if len(sys.argv) > 2 else "/projects/sandbox/_render"
    pdf, pngs, sheets = render(pptx, workdir)
    print(f"pdf: {pdf}\npages: {len(pngs)}")
    for s in sheets:
        print("sheet:", s)
