"""Build the pls,fix PowerPoint demo deck: eight slides, one landing spot for each Excel link kind
(table, chart, picture, text), the Slide and Where pickers and shapes for the object tools. The
slides live in slides_pickers.py and slides_objects.py, the helpers in deck_kit.py; this file owns
the Highlighter theme, the properties and the byte-reproducible save. Python: no Rust PPTX crate."""

import os
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

from deck_kit import new_presentation
from slides_objects import build_slide_6, build_slide_7, build_slide_8
from slides_pickers import build_slide_1, build_slide_2, build_slide_3, build_slide_4, build_slide_5

# The deck's colour theme: the Highlighter palette (the defaults of src/settings.ts
# and the series tints of src/chart-colors.ts) instead of python-pptx's default
# Office blue. A native table the pane inserts takes the deck theme's Accent 1
# (src/ppt/table-style.ts writes Medium Style 2 Accent 1, so a user's own theme
# decides): ink here, with the olive as Accent 2 and the tints after it. Slots in
# the theme's own order: the two darks and the two lights, accents 1-6, the
# hyperlink colours (a followed link is dk2).
THEME_NAME = "Highlighter"
THEME_COLORS = (
    ("dk1", "111111"),
    ("lt1", "FFFFFF"),
    ("dk2", "575752"),
    ("lt2", "F7F7F2"),
    ("accent1", "111111"),
    ("accent2", "6D7E12"),
    ("accent3", "949494"),
    ("accent4", "AFB87D"),
    ("accent5", "CBCBCB"),
    ("accent6", "D3D8B8"),
    ("hlink", "4D5B00"),
    ("folHlink", "575752"),
)

OUT_PATH = Path(__file__).parent / "pls,fix Demo Deck.pptx"

# Reproducible builds: SOURCE_DATE_EPOCH wins when set; otherwise this
# fixed date. Bump BUILD_STAMP by hand when the deck's content changes -
# never datetime.now(), or every rebuild dirties the tracked .pptx even
# with no content change.
BUILD_STAMP = "2026-10-06"


def _build_timestamp() -> datetime:
    """SOURCE_DATE_EPOCH (the reproducible-builds convention) when set,
    else BUILD_STAMP - never datetime.now(), so two rebuilds of the same
    content are byte-identical."""
    epoch = os.environ.get("SOURCE_DATE_EPOCH")
    if epoch:
        return datetime.fromtimestamp(int(epoch), tz=timezone.utc)
    return datetime.strptime(BUILD_STAMP, "%Y-%m-%d").replace(tzinfo=timezone.utc)


def set_core_properties(prs: Presentation) -> None:
    """Overwrite the default template's stale docProps/core.xml: its author,
    description and 2013 timestamps are python-pptx's own template, not this
    project or this build. Timestamps come from _build_timestamp(), never
    datetime.now(), so the tracked .pptx stays reproducible."""
    props = prs.core_properties
    props.title = "pls,fix Demo Deck"
    props.author = "pls,fix"
    props.last_modified_by = "pls,fix"
    props.comments = ""
    stamp = _build_timestamp()
    props.created = stamp
    props.modified = stamp


def apply_theme_colors(prs: Presentation) -> None:
    """Give the deck the Highlighter colour scheme. python-pptx's default
    template carries two theme parts, the slide master's and the notes master's,
    each the Office palette: both get THEME_COLORS, so no Office blue is left in
    the file. Only the clrScheme element is rewritten, in place; the fonts and
    the format scheme of the theme stay as the template has them. The notes
    master exists by now (every slide has notes), so nothing is created here."""
    for master in (prs.slide_master, prs.notes_master):
        theme = master.part.part_related_by(RT.THEME)
        root = parse_xml(theme.blob)
        scheme = root.find(f"{qn('a:themeElements')}/{qn('a:clrScheme')}")
        if scheme is None:
            raise ValueError(f"{theme.partname} has no colour scheme to replace")
        scheme.set("name", THEME_NAME)
        for old in list(scheme):
            scheme.remove(old)
        for slot, value in THEME_COLORS:
            color = etree.SubElement(etree.SubElement(scheme, qn(f"a:{slot}")), qn("a:srgbClr"))
            color.set("val", value)
        # The slide master's theme is a blob part, the notes master's an XML part.
        if isinstance(theme, XmlPart):
            theme._element = root
        else:
            theme.blob = etree.tostring(root, encoding="UTF-8", standalone=True)


def drop_template_baggage(prs: Presentation) -> None:
    """Remove docProps/thumbnail.jpeg (a stale 4:3 preview) and
    ppt/printerSettings/printerSettings1.bin (Windows print-driver settings),
    both dead weight from python-pptx's template: drop the one relationship
    each has, so OpcPackage.iter_parts()'s reachability walk excludes the
    part - and its [Content_Types].xml entry - from save() on its own. No
    zip surgery; proved clean by unzip -t and a reload round-trip."""
    package = prs.part.package
    pkg_rels = package._rels  # OpcPackage has no public alias for this
    thumbnail_rid = next(rid for rid in pkg_rels if pkg_rels[rid].reltype == RT.THUMBNAIL)
    package.drop_rel(thumbnail_rid)

    pres_rels = prs.part.rels
    printer_rid = next(rid for rid in pres_rels if pres_rels[rid].reltype == RT.PRINTER_SETTINGS)
    prs.part.drop_rel(printer_rid)


def build() -> Presentation:
    prs = new_presentation()
    for slide in (
        build_slide_1, build_slide_2, build_slide_3, build_slide_4,
        build_slide_5, build_slide_6, build_slide_7, build_slide_8,
    ):
        slide(prs)
    apply_theme_colors(prs)
    set_core_properties(prs)
    drop_template_baggage(prs)
    return prs


def normalize_zip_timestamps(path: Path, stamp: datetime) -> None:
    """Rewrite every zip entry's stored date_time to `stamp`.

    python-pptx's writer (pptx.opc.serialized._ZipPkgWriter) calls
    zipfile.writestr() with a plain filename, which stamps each entry with
    time.localtime() at save() time - unrelated to docProps/core.xml and
    otherwise enough on its own to make two rebuilds of identical content
    differ byte for byte.
    """
    date_time = stamp.timetuple()[:6]
    with zipfile.ZipFile(path, "r") as zf:
        entries = [(info, zf.read(info.filename)) for info in zf.infolist()]

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for info, data in entries:
            new_info = zipfile.ZipInfo(info.filename, date_time=date_time)
            new_info.compress_type = zipfile.ZIP_DEFLATED
            new_info.external_attr = info.external_attr
            zf.writestr(new_info, data)


def main() -> None:
    prs = build()
    prs.save(OUT_PATH)
    normalize_zip_timestamps(OUT_PATH, _build_timestamp())
    print(f"wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
