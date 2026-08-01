"""
Embed the theme fonts into a generated .pptx.

python-pptx cannot embed fonts, so this post-processes the package the same way
PowerPoint does: font data goes in ppt/fonts/*.fntdata, gets a relationship
from presentation.xml, and is declared in <p:embeddedFontLst>. The result opens
with the correct typefaces on a machine that has never seen these fonts.

Both families are SIL Open Font License, which permits embedding.
"""

import os
import re
import shutil
import zipfile

FONT_REL = ("http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/font")
CT_ENTRY = '<Default Extension="fntdata" ContentType="application/x-fontdata"/>'
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# typeface -> {style: fntdata filename in assets/fonts}
FAMILIES = [
    ("Quattrocento Sans", {
        "regular": "quattrocentosans-regular.fntdata",
        "bold": "quattrocentosans-bold.fntdata",
        "italic": "quattrocentosans-italic.fntdata",
    }),
    ("Sorts Mill Goudy", {
        "regular": "sortsmillgoudy-regular.fntdata",
        "italic": "sortsmillgoudy-italic.fntdata",
    }),
]


def _fonts_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                        "assets", "fonts")


def embed(pptx_path, fonts_dir=None, quiet=True):
    fonts_dir = fonts_dir or _fonts_dir()
    available = []
    for typeface, styles in FAMILIES:
        got = {st: os.path.join(fonts_dir, fn) for st, fn in styles.items()
               if os.path.exists(os.path.join(fonts_dir, fn))}
        if got:
            available.append((typeface, got))
    if not available:
        if not quiet:
            print("no font data found; skipping embed")
        return False

    src = pptx_path + ".tmp"
    shutil.move(pptx_path, src)
    zin = zipfile.ZipFile(src, "r")
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()

    existing = {name for name, _ in [(i.filename, b) for i, b in items]}
    rels_name = "ppt/_rels/presentation.xml.rels"
    pres_name = "ppt/presentation.xml"
    ct_name = "[Content_Types].xml"

    rels = next(b for i, b in items if i.filename == rels_name).decode("utf-8")
    pres = next(b for i, b in items if i.filename == pres_name).decode("utf-8")
    ct = next(b for i, b in items if i.filename == ct_name).decode("utf-8")

    used = {int(m) for m in re.findall(r'Id="rId(\d+)"', rels)}
    next_id = (max(used) if used else 0) + 1

    new_files = []
    font_xml = []
    idx = 1
    for typeface, styles in available:
        entry = [f'<p:embeddedFont><p:font typeface="{typeface}" '
                 f'charset="-122" pitchFamily="34"/>']
        for style in ("regular", "bold", "italic", "boldItalic"):
            if style not in styles:
                continue
            target = f"fonts/font{idx}.fntdata"
            arc = f"ppt/{target}"
            idx += 1
            with open(styles[style], "rb") as fh:
                new_files.append((arc, fh.read()))
            rid = f"rId{next_id}"
            next_id += 1
            rels = rels.replace(
                "</Relationships>",
                f'<Relationship Id="{rid}" Type="{FONT_REL}" '
                f'Target="{target}"/></Relationships>')
            entry.append(f'<p:{style} r:id="{rid}"/>')
        entry.append("</p:embeddedFont>")
        font_xml.append("".join(entry))

    # declare the part type once
    if 'Extension="fntdata"' not in ct:
        ct = ct.replace("<Override", CT_ENTRY + "<Override", 1) \
            if "<Override" in ct else ct.replace("</Types>", CT_ENTRY + "</Types>")

    # <p:embeddedFontLst> must follow notesSz in the schema's element order
    block = "<p:embeddedFontLst>" + "".join(font_xml) + "</p:embeddedFontLst>"
    if "<p:embeddedFontLst>" not in pres:
        m = re.search(r"<p:notesSz[^>]*/>", pres)
        if m:
            pres = pres[:m.end()] + block + pres[m.end():]
        else:
            pres = pres.replace("</p:presentation>", block + "</p:presentation>")
    # tell PowerPoint the fonts travel with the file, without duplicating
    # attributes the template already declares
    for attr in ("embedTrueTypeFonts", "saveSubsetFonts"):
        if re.search(rf'<p:presentation[^>]*\b{attr}=', pres):
            continue
        pres = pres.replace("<p:presentation ",
                            f'<p:presentation {attr}="1" ', 1)

    replacements = {ct_name: ct.encode("utf-8"),
                    rels_name: rels.encode("utf-8"),
                    pres_name: pres.encode("utf-8")}

    zout = zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED)
    for info, data in items:
        zout.writestr(info, replacements.get(info.filename, data))
    for name, data in new_files:
        if name not in existing:
            zout.writestr(name, data)
    zout.close()
    os.remove(src)
    if not quiet:
        print(f"embedded {len(new_files)} font file(s) into "
              f"{os.path.basename(pptx_path)}")
    return True


if __name__ == "__main__":
    import sys
    for p in sys.argv[1:]:
        embed(p, quiet=False)
