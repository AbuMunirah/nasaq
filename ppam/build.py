"""Builds the VBA edition of Nasaq (.ppam).

  python3 ppam/build.py vba
      Writes ppam/out/Nasaq.bas: the VBA module, with English/Arabic text taken
      from manifest.xml so both editions share the same labels.

  python3 ppam/build.py ppam <base.ppam|base.pptm>
      Takes a .ppam or macro-enabled .pptm saved by PowerPoint (containing the
      imported Nasaq module) and writes ppam/out/Nasaq.ppam with the ribbon
      (customUI14.xml) and icons added. A .pptm is converted to an add-in.
"""

import re
import shutil
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import unescape

ROOT = Path(__file__).resolve().parent.parent
PPAM = ROOT / "ppam"
OUT = PPAM / "out"

# Strings used only by the VBA edition. {0}/{1} are placeholders.
EXTRA = {
    "App.Name": ("Nasaq", "نَسَق"),
    "MatchSize.Name": ("Match Size", "مطابقة الحجم"),
    "MatchWidth.Name": ("Match Width", "مطابقة العرض"),
    "MatchHeight.Name": ("Match Height", "مطابقة الارتفاع"),
    "Align.Name": ("Align", "المحاذاة"),
    "Message.NeedsTwo": (
        "{0} needs at least 2 selected shapes ({1} selected).",
        "يتطلب أمر «{0}» تحديد شكلين على الأقل (عدد الأشكال المحددة: {1}).",
    ),
    "Message.SwapNeedsTwo": (
        "Swap needs exactly 2 selected shapes ({0} selected).",
        "يتطلب التبديل تحديد شكلين بالضبط (عدد الأشكال المحددة: {0}).",
    ),
    "Message.Unexpected": ("Something went wrong: {0}", "حدث خطأ: {0}"),
}

# Main part content types: a .ppam is a macro-enabled presentation marked as an add-in.
PPTM_MAIN = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
PPAM_MAIN = "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"

# Ribbon image id -> icon in assets/ (64 px scales down cleanly to 32 and 16 px).
IMAGES = {
    "align_left": "align-left",
    "align_center": "align-center",
    "align_right": "align-right",
    "align_top": "align-top",
    "align_middle": "align-middle",
    "align_bottom": "align-bottom",
    "match_size": "match-size",
    "match_width": "match-width",
    "match_height": "match-height",
    "swap": "swap",
    "text_align_left": "text-align-left",
}


def manifest_strings():
    """Reads English defaults and Arabic overrides for every bt:String in manifest.xml."""
    xml = (ROOT / "manifest.xml").read_text(encoding="utf-8")
    strings = {}
    pattern = r'<bt:String id="([\w.]+)" DefaultValue="([^"]*)">\s*<bt:Override Locale="ar-SA" Value="([^"]*)"/>'
    for key, en, ar in re.findall(pattern, xml):
        strings[key] = (unescape(en, {"&quot;": '"'}), unescape(ar, {"&quot;": '"'}))
    return strings


def vba_literal(text):
    """ASCII text as a VBA string literal; anything else via U("hex hex ...")."""
    if text.isascii():
        return '"' + text.replace('"', '""') + '"'
    return 'U("' + " ".join(f"{ord(c):X}" for c in text) + '")'


def text_function(name, strings, index):
    lines = [f"Private Function {name}(ByVal key As String) As String", "    Select Case key"]
    for key in sorted(strings):
        lines.append(f'        Case "{key}": {name} = {vba_literal(strings[key][index])}')
    lines += ["    End Select", "End Function"]
    return lines


def build_vba():
    strings = {**manifest_strings(), **EXTRA}
    body = text_function("TextEn", strings, 0) + [""] + text_function("TextAr", strings, 1)
    template = (PPAM / "Nasaq.template.bas").read_text(encoding="utf-8")
    module = template.replace("' {{STRINGS}}", "\n".join(body))
    if not module.isascii():
        raise SystemExit("Generated VBA must be ASCII-only")
    OUT.mkdir(exist_ok=True)
    target = OUT / "Nasaq.bas"
    # The VBA editor expects Windows line endings.
    target.write_bytes(module.replace("\r\n", "\n").replace("\n", "\r\n").encode("ascii"))
    print(f"Wrote {target.relative_to(ROOT)} ({len(strings)} strings)")


def build_ppam(base):
    base = Path(base)
    target = OUT / "Nasaq.ppam"
    OUT.mkdir(exist_ok=True)
    ui_rel = (
        '<Relationship Id="nasaqUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" '
        'Target="customUI/customUI14.xml"/>'
    )
    image_rels = "".join(
        f'<Relationship Id="{image_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="images/{image_id}.png"/>'
        for image_id in IMAGES
    )

    with zipfile.ZipFile(base) as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst:
        names = src.namelist()
        if "ppt/vbaProject.bin" not in names:
            raise SystemExit(f"{base} has no VBA project - import Nasaq.bas and save as .pptm or .ppam first")
        for name in names:
            if name.startswith("customUI/"):
                continue  # replaced below
            data = src.read(name)
            if name == "_rels/.rels":
                rels = re.sub(r'<Relationship [^>]*ui/extensibility"[^>]*/>', "", data.decode("utf-8"))
                data = rels.replace("</Relationships>", ui_rel + "</Relationships>").encode("utf-8")
            elif name == "docProps/core.xml":
                # Don't publish the name of whoever saved the base file.
                core = data.decode("utf-8")
                for tag in ("dc:creator", "cp:lastModifiedBy"):
                    core = re.sub(rf"<{tag}>[^<]*</{tag}>", f"<{tag}>Nasaq</{tag}>", core)
                data = core.encode("utf-8")
            elif name == "[Content_Types].xml":
                types = data.decode("utf-8").replace(PPTM_MAIN, PPAM_MAIN)
                if PPAM_MAIN not in types:
                    raise SystemExit(f"{base} is not a macro-enabled presentation or add-in")
                for ext, ctype in (("png", "image/png"), ("xml", "application/xml")):
                    if f'Extension="{ext}"' not in types:
                        # Insert right after the opening <Types ...> tag.
                        types = re.sub(
                            r"(<Types\b[^>]*>)",
                            rf'\1<Default Extension="{ext}" ContentType="{ctype}"/>',
                            types,
                            count=1,
                        )
                # Declare the ribbon part explicitly so it never inherits another default type.
                types = re.sub(r'<Override PartName="/customUI/customUI14.xml"[^>]*/>', "", types)
                types = types.replace(
                    "</Types>",
                    '<Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/></Types>',
                )
                data = types.encode("utf-8")
            dst.writestr(src.getinfo(name), data)

        dst.write(PPAM / "customUI14.xml", "customUI/customUI14.xml")
        dst.writestr(
            "customUI/_rels/customUI14.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + image_rels
            + "</Relationships>",
        )
        for image_id, asset in IMAGES.items():
            dst.write(ROOT / "assets" / f"{asset}-64.png", f"customUI/images/{image_id}.png")
    print(f"Wrote {target.relative_to(ROOT)}")


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "vba":
        build_vba()
    elif len(sys.argv) >= 3 and sys.argv[1] == "ppam":
        build_ppam(sys.argv[2])
    else:
        print(__doc__)
        sys.exit(1)
