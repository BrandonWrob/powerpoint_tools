"""pptx_reader - turn a PowerPoint file into text + images for the Claude Messages API.

    from pptx_reader import read_pptx, to_content_blocks, save_images

    result = read_pptx("deck.pptx")              # never modifies the file, never fetches anything
    blocks = to_content_blocks(result)           # list of {"type": "text"|"image", ...} for a user message
    save_images(result, "out/images")            # optional: image files + images.json (where each one sits)

What comes out
- result.text    every slide in order: title, text (bullets keep their level), tables, chart data,
                 SmartArt text, connectors, comments, speaker notes. Each picture is a line such as
                 [image img_3 | x=52% y=18% w=42% h=60% | alt: Plant photo]
- result.images  the pictures, already converted to what the API accepts (PNG/JPEG/GIF/WebP), resized
                 to fit a token budget, de-duplicated (a logo used on 40 slides is sent once).
- result.tables  every slide table and chart as raw rows, so they can also go into SQL.

Needs python-pptx, Pillow, lxml. Optional: LibreOffice (EMF/WMF/SVG pictures), matplotlib (render_charts=True).
"""
from __future__ import annotations

import base64
import csv
import datetime as _dt
import hashlib
import html
import io
import json
import math
import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

from lxml import etree

__all__ = ["read_pptx", "to_content_blocks", "blocks_from_text", "to_prompt_text", "save_images",
           "ConversionError", "PptxResult", "SlideImage", "SlideInfo", "RawTable"]

# ---------------------------------------------------------------- limits

MAX_FILE_BYTES = 100 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 400 * 1024 * 1024      # zip-bomb guard
MAX_ZIP_ENTRIES = 20_000
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # legacy .ppt or an encrypted .pptx
LIBREOFFICE_TIMEOUT_S = 120

# Messages API image rules (https://platform.claude.com/docs/en/build-with-claude/vision):
# JPEG/PNG/GIF/WebP only; base64 size limit per image (5 MB is the strictest platform limit);
# 32 MB per request; 100 images per request on 200k-context models; with more than 20 images
# each must be at most 2000 px per side. Staying at or below 1568 px / ~1568 visual tokens per
# image keeps every one of those satisfied and avoids paying for pixels the model would drop.
API_MEDIA_TYPES = {"PNG": "image/png", "JPEG": "image/jpeg", "GIF": "image/gif", "WEBP": "image/webp"}
MAX_IMAGE_RAW_BYTES = 5 * 1024 * 1024 * 3 // 4 - 1024   # stays under 5 MB once base64-encoded
MAX_TOTAL_IMAGE_RAW_BYTES = 20 * 1024 * 1024            # ~27 MB as base64, leaves room for text
MAX_EDGE_PX = 1568
MAX_CHART_POINTS = 50_000                               # per series; numbers in the file are never trusted as loop sizes
MAX_CHART_CELLS = 500_000                               # per chart
MAX_SOURCE_PIXELS = 60_000_000                          # larger pictures are not decoded (memory guard)
PATCH_PX = 28                                           # one visual token = one 28x28 px patch
TOKEN_TIERS = (1568, 1000, 600, 350)                    # per-image token caps, tried in this order
DEFAULT_IMAGE_TOKEN_BUDGET = 40_000
DEFAULT_MAX_IMAGES = 100

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "p188": "http://schemas.microsoft.com/office/powerpoint/2018/8/main",
}
_R_EMBED = f"{{{NS['r']}}}embed"
_R_LINK = f"{{{NS['r']}}}link"
_R_ID = f"{{{NS['r']}}}id"
_R_DM = f"{{{NS['r']}}}dm"
# Entities are never resolved and nothing is fetched while parsing parts ourselves.
_SAFE_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False, remove_comments=True)


# ---------------------------------------------------------------- result types

class ConversionError(Exception):
    """The file couldn't be read. The message is safe to show the user."""


@dataclass
class SlideImage:
    id: str                      # "img_3", the id used in the [image img_3 | ...] line
    media_type: str              # image/png | image/jpeg | image/gif | image/webp
    data: bytes                  # the bytes to send (converted/resized when needed)
    width: int                   # pixels, as sent
    height: int
    est_tokens: int              # ceil(w/28) * ceil(h/28)
    source: str                  # picture | shape fill | background | video poster | object preview | chart render | template
    original_format: str         # e.g. "TIFF", "EMF"
    original_size: tuple[int, int]
    sha1: str                    # of the original bytes in the file
    locations: list[dict] = field(default_factory=list)   # every place it appears, see _location()

    def block(self) -> dict:
        return {"type": "image", "source": {"type": "base64", "media_type": self.media_type,
                                            "data": base64.standard_b64encode(self.data).decode("ascii")}}

    def manifest(self) -> dict:
        ext = self.media_type.split("/")[1].replace("jpeg", "jpg")
        return {"id": self.id, "file": f"{self.id}.{ext}", "media_type": self.media_type,
                "width": self.width, "height": self.height, "bytes": len(self.data),
                "est_tokens": self.est_tokens, "source": self.source,
                "original_format": self.original_format, "original_size": list(self.original_size),
                "sha1": self.sha1, "locations": self.locations}


@dataclass
class RawTable:
    """A slide table or a chart's data as plain rows (first row = header)."""
    name: str                    # slide3_table, slide5_chart, slide5_chart_2 ...
    rows: list[list[str]]
    source: str                  # "slide 3 table", "slide 5 chart 'Revenue by region'"


@dataclass
class SlideInfo:
    """One slide on its own, for sending slides on demand (see deck_store.py)."""
    number: int
    title: str
    hidden: bool
    text: str                    # exactly this slide's part of PptxResult.text
    outline: str                 # one line: what is on the slide, without the content itself
    image_ids: list[str] = field(default_factory=list)   # attached pictures that appear on it


@dataclass
class PptxResult:
    filename: str
    text: str
    kind: str = "slides"
    warnings: list[str] = field(default_factory=list)
    images: list[SlideImage] = field(default_factory=list)
    tables: list[RawTable] = field(default_factory=list)
    meta: dict = field(default_factory=dict)
    slides: list[SlideInfo] = field(default_factory=list)
    header: str = ""             # the first paragraph of text: deck size and how to read positions

    @property
    def est_tokens(self) -> int:
        """Rough total: text (characters / 3.2) plus the attached images."""
        return _text_tokens(self.text) + sum(i.est_tokens for i in self.images)

    def image_manifest(self) -> list[dict]:
        return [i.manifest() for i in self.images]


def _text_tokens(text: str) -> int:
    return int(len(text) / 3.2) + 1


def _img_tokens(w: int, h: int) -> int:
    return math.ceil(w / PATCH_PX) * math.ceil(h / PATCH_PX)


# ---------------------------------------------------------------- small xml helpers

@lru_cache(maxsize=256)
def _compiled(expr: str):
    return etree.XPath(expr, namespaces=NS)


def _xp(el, expr: str):
    return _compiled(expr)(el)


def _find(el, path: str):
    return el.find(path, NS)


def _local(el) -> str:
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def _int(el, name: str, default: int = 0) -> int:
    try:
        return int(el.get(name, default))
    except (TypeError, ValueError):
        return default


_CTRL = re.compile(r"[\x00-\x08\x0c\x0e-\x1f\x7f]")


def _clean(s: str) -> str:
    """Document text -> safe to place in our output: no control characters, and it can't
    imitate one of our own [image img_N | ...] lines."""
    s = (s or "").replace("\x0b", "\n").replace("\r", "\n")
    return _CTRL.sub("", s).replace("[image img_", "[image_img_")


def _one_line(s: str, limit: int = 0) -> str:
    s = re.sub(r"\s+", " ", _clean(s)).strip()
    return s if not limit or len(s) <= limit else s[: limit - 1].rstrip() + "…"


def _num(v) -> str:
    """Chart cache value -> short text ('12', '0.3', '1250000')."""
    if v is None or v == "":
        return ""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return _one_line(str(v))
    if f != f or f in (float("inf"), float("-inf")):
        return str(f)
    if f.is_integer() and abs(f) < 1e15:
        return str(int(f))
    return format(f, ".10g")


def _csv(rows: list[list[str]]) -> str:
    buf = io.StringIO()
    csv.writer(buf, lineterminator="\n").writerows(rows)
    return buf.getvalue().rstrip("\n")


def _md_table(rows: list[list[str]]) -> str:
    width = max(len(r) for r in rows)
    rows = [[c.replace("|", "\\|").replace("\n", " ").strip() for c in r] + [""] * (width - len(r)) for r in rows]
    lines = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * width]
    lines += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(lines)


# ---------------------------------------------------------------- internal records

@dataclass
class _Img:
    """One distinct picture in the file (same bytes + same crop)."""
    blob: bytes
    content_type: str
    ext: str
    crop: tuple | None
    source: str
    sha1: str
    uses: list = field(default_factory=list)
    status: str = "ok"           # ok | off | small | vector | unreadable | budget
    fmt: str = ""
    size: tuple[int, int] = (0, 0)          # pixels after crop, before resizing
    orig_size: tuple[int, int] = (0, 0)
    out: SlideImage | None = None


@dataclass
class _Use:
    img: _Img | None             # None = linked/missing picture
    slide: int
    box: tuple | None
    z: float
    alt: str = ""
    note: str = ""               # "slide background", "fills a shape", ...
    order: int = 0               # position in the rendered text
    problem: str = ""            # set when img is None


@dataclass
class _Item:
    kind: str                    # title | text | table | chart | diagram | object | image
    box: tuple | None
    z: float
    head: str = ""
    lines: list[str] = field(default_factory=list)
    use: _Use | None = None


@dataclass
class _Slide:
    number: int
    title: str = ""
    hidden: bool = False
    items: list[_Item] = field(default_factory=list)
    connectors: list[tuple] = field(default_factory=list)    # (start_id, end_id, start_arrow, end_arrow)
    labels: dict = field(default_factory=dict)                # shape id -> short label
    comments: list[str] = field(default_factory=list)
    notes: str = ""


class _Ctx:
    def __init__(self, opts: dict, slide_w: int, slide_h: int):
        self.opts = opts
        self.w, self.h = max(slide_w, 1), max(slide_h, 1)
        self.images: dict[tuple, _Img] = {}
        self.tables: list[RawTable] = []
        self.warnings: list[str] = []
        self.counts = {"charts": 0, "tables": 0, "diagrams": 0, "objects": 0}
        self.slide: _Slide | None = None
        self.part = None
        self.ph_boxes: dict = {}                 # placeholder element -> box inherited from the layout
        self.title_el = None
        self.z = 0

    def warn(self, msg: str):
        if msg not in self.warnings:
            self.warnings.append(msg)

    def next_z(self) -> int:
        self.z += 1
        return self.z


# ---------------------------------------------------------------- opening the file

def _open(source, filename: str | None):
    from pptx.package import Package

    if isinstance(source, (str, Path)):
        path = Path(source)
        filename = filename or path.name
        if not path.is_file():
            raise ConversionError(f"File not found: {filename}")
        data = path.read_bytes() if path.stat().st_size <= MAX_FILE_BYTES else None
        if data is None:
            raise ConversionError(f"{filename} is larger than {MAX_FILE_BYTES // 2**20} MB.")
    else:
        data = source if isinstance(source, (bytes, bytearray)) else source.read()
        filename = filename or "presentation.pptx"
        if len(data) > MAX_FILE_BYTES:
            raise ConversionError(f"{filename} is larger than {MAX_FILE_BYTES // 2**20} MB.")
    if not data:
        raise ConversionError(f"{filename} is empty.")
    if data[:8] == OLE_MAGIC:
        raise ConversionError(f"{filename} is password-protected/encrypted or an old .ppt file. "
                              "Remove the password or save it as .pptx and upload again.")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            infos = z.infolist()
            names = {i.filename for i in infos}
    except zipfile.BadZipFile as e:
        raise ConversionError(f"{filename} is corrupt or not really a PowerPoint file.") from e
    if len(infos) > MAX_ZIP_ENTRIES or sum(i.file_size for i in infos) > MAX_UNCOMPRESSED_BYTES:
        raise ConversionError(f"{filename} expands to more than {MAX_UNCOMPRESSED_BYTES // 2**20} MB "
                              "or has too many parts.")
    try:
        # Package.open (not pptx.Presentation) so .pptm/.ppsx/.potx open too. Nothing is executed.
        prs = Package.open(io.BytesIO(data)).presentation_part.presentation
        prs.slides  # noqa: B018 - forces the slide list to load so errors surface here
    except Exception as e:  # noqa: BLE001
        raise ConversionError(f"Could not read {filename} as a PowerPoint file.") from e
    return prs, filename, any(n.lower().endswith("vbaproject.bin") for n in names)


# ---------------------------------------------------------------- geometry

def _xfrm_box(xf):
    if xf is None:
        return None
    off, ext = _find(xf, "a:off"), _find(xf, "a:ext")
    if off is None or ext is None:
        return None
    return (_int(off, "x"), _int(off, "y"), _int(ext, "cx"), _int(ext, "cy"))


def _group_tf(grp, parent_tf):
    """Children of a group are in the group's own coordinate space; map them to the slide."""
    xf = _find(grp, "p:grpSpPr/a:xfrm")
    box = _xfrm_box(xf)
    if box is None:
        return parent_tf
    ch_off, ch_ext = _find(xf, "a:chOff"), _find(xf, "a:chExt")
    cx0, cy0 = (_int(ch_off, "x"), _int(ch_off, "y")) if ch_off is not None else (box[0], box[1])
    cw, ch = (_int(ch_ext, "cx"), _int(ch_ext, "cy")) if ch_ext is not None else (box[2], box[3])
    sx = box[2] / cw if cw else 1.0
    sy = box[3] / ch if ch else 1.0

    def tf(b):
        x, y, w, h = b
        return parent_tf((box[0] + (x - cx0) * sx, box[1] + (y - cy0) * sy, w * sx, h * sy))
    return tf


def _box(ctx: _Ctx, el, xfrm_path: str, tf):
    box = _xfrm_box(_find(el, xfrm_path))
    if box is not None:
        return tf(box)
    return ctx.ph_boxes.get(el)          # placeholder that takes its position from the layout


def _pos(ctx: _Ctx, box) -> str:
    if box is None:
        return "position set by the slide layout"
    x, y, w, h = box
    return (f"x={round(100 * x / ctx.w)}% y={round(100 * y / ctx.h)}% "
            f"w={round(100 * w / ctx.w)}% h={round(100 * h / ctx.h)}%")


def _location(ctx: _Ctx, use: _Use) -> dict:
    loc = {"slide": use.slide, "z_order": int(use.z), "alt_text": use.alt, "note": use.note}
    if use.box is not None:
        x, y, w, h = use.box
        emu_in = 914400
        loc.update({"position": _pos(ctx, use.box),            # the same wording as in the slide text
                    "x_pct": round(100 * x / ctx.w, 2), "y_pct": round(100 * y / ctx.h, 2),
                    "w_pct": round(100 * w / ctx.w, 2), "h_pct": round(100 * h / ctx.h, 2),
                    "left_in": round(x / emu_in, 3), "top_in": round(y / emu_in, 3),
                    "width_in": round(w / emu_in, 3), "height_in": round(h / emu_in, 3)})
    return loc


# ---------------------------------------------------------------- text

def _link_target(part, r_pr) -> str:
    if r_pr is None:
        return ""
    h = _find(r_pr, "a:hlinkClick")
    rid = h.get(_R_ID) if h is not None else None
    if not rid:
        return ""
    try:
        rel = part.rels[rid]
        return rel.target_ref if rel.is_external else ""
    except Exception:  # noqa: BLE001
        return ""


def _paragraphs(tx_body, part) -> list[tuple[int, str]]:
    """[(level, text)] for every non-empty paragraph. Links are written as 'text (url)'."""
    out = []
    if tx_body is None:
        return out
    for p in tx_body.findall("a:p", NS):
        segs: list[list[str]] = []          # [text, url]
        for ch in p:
            name = _local(ch)
            if name in ("r", "fld"):
                t = ch.findtext("a:t", default="", namespaces=NS)
                url = _link_target(part, _find(ch, "a:rPr")) if name == "r" else ""
                if segs and segs[-1][1] == url:
                    segs[-1][0] += t
                else:
                    segs.append([t, url])
            elif name == "br":
                if segs:
                    segs[-1][0] += "\n"
                else:
                    segs.append(["\n", ""])
        text = ""
        for t, url in segs:
            text += t
            if url and url.strip() not in t and not url.lower().startswith(("javascript:", "file:")):
                text = text.rstrip() + f" ({_one_line(url, 300)})"
        text = "\n".join(line.strip() for line in _clean(text).split("\n")).strip()
        if text:
            p_pr = _find(p, "a:pPr")
            out.append((min(max(_int(p_pr, "lvl"), 0), 8) if p_pr is not None else 0, text))
    return out


def _para_lines(paras: list[tuple[int, str]]) -> list[str]:
    if len(paras) == 1 and paras[0][0] == 0:
        return paras[0][1].split("\n")
    lines = []
    for lvl, text in paras:
        first, *rest = text.split("\n")
        lines.append("  " * lvl + "- " + first)
        lines += ["  " * lvl + "  " + r for r in rest]
    return lines


# ---------------------------------------------------------------- pictures

def _register_image(ctx: _Ctx, part, blip, box, *, alt: str = "", note: str = "", source: str = "picture",
                    crop=None, z: float | None = None) -> _Item:
    alt = _one_line(alt, 200).translate({ord("["): "(", ord("]"): ")", ord("|"): "/"})
    use = _Use(None, ctx.slide.number, box, ctx.next_z() if z is None else z, alt, note)
    rid = blip.get(_R_EMBED) if blip is not None else None
    if not rid:
        use.problem = ("linked to an external file, which is never fetched"
                       if blip is not None and blip.get(_R_LINK) else "no image data in the file")
    else:
        try:
            img_part = part.rels[rid].target_part
            blob = img_part.blob
        except Exception:  # noqa: BLE001
            blob = None
            use.problem = "image data missing from the file"
        if blob:
            sha1 = hashlib.sha1(blob).hexdigest()
            key = (sha1, crop)
            img = ctx.images.get(key)
            if img is None:
                ext = str(img_part.partname).rsplit(".", 1)[-1].lower()
                img = ctx.images[key] = _Img(blob, img_part.content_type or "", ext, crop, source, sha1)
            img.uses.append(use)
            use.img = img
    return _Item("image", box, use.z, use=use)


def _crop_of(blip_fill):
    r = _find(blip_fill, "a:srcRect") if blip_fill is not None else None
    if r is None:
        return None
    vals = tuple(min(max(_int(r, k), 0), 100000) for k in ("l", "t", "r", "b"))
    return vals if any(vals) and vals[0] + vals[2] < 100000 and vals[1] + vals[3] < 100000 else None


def _c_nv_pr(el):
    found = _xp(el, "p:nvSpPr/p:cNvPr | p:nvPicPr/p:cNvPr | p:nvGraphicFramePr/p:cNvPr | p:nvCxnSpPr/p:cNvPr")
    return found[0] if found else None


def _picture(ctx: _Ctx, el, tf):
    nv = _c_nv_pr(el)
    box = _box(ctx, el, "p:spPr/a:xfrm", tf)
    alt = (nv.get("descr") or nv.get("title") or "") if nv is not None else ""
    media = _xp(el, "p:nvPicPr/p:nvPr/a:videoFile | p:nvPicPr/p:nvPr/a:audioFile")
    blip_fill = _find(el, "p:blipFill")
    blip = _find(blip_fill, "a:blip") if blip_fill is not None else None
    if media:
        kind = "video" if _local(media[0]) == "videoFile" else "audio clip"
        ctx.slide.items.append(_Item("object", box, ctx.next_z(), head=f"{kind} (not included"
                                     + ("; its poster frame follows)" if kind == "video" else ")")))
        if kind != "video":
            return
        ctx.slide.items.append(_register_image(ctx, ctx.part, blip, box, alt=alt, note="video poster frame",
                                               source="video poster"))
        return
    ctx.slide.items.append(_register_image(ctx, ctx.part, blip, box, alt=alt, crop=_crop_of(blip_fill)))


# ---------------------------------------------------------------- shapes with text

_SKIP_PLACEHOLDERS = {"sldNum", "dt", "ftr", "hdr"}


def _shape(ctx: _Ctx, el, tf):
    nv = _c_nv_pr(el)
    ph = _find(el, "p:nvSpPr/p:nvPr/p:ph")
    if ph is not None and ph.get("type") in _SKIP_PLACEHOLDERS:
        return
    box = _box(ctx, el, "p:spPr/a:xfrm", tf)
    paras = _paragraphs(_find(el, "p:txBody"), ctx.part)
    if nv is not None:
        label = _one_line(paras[0][1], 40) if paras else _one_line(nv.get("name") or "", 40)
        ctx.slide.labels[nv.get("id")] = label
    fill = _find(el, "p:spPr/a:blipFill")
    if fill is not None and _find(fill, "a:blip") is not None:
        alt = (nv.get("descr") or "") if nv is not None else ""
        ctx.slide.items.append(_register_image(ctx, ctx.part, _find(fill, "a:blip"), box, alt=alt,
                                               note="picture filling a shape", source="shape fill",
                                               crop=_crop_of(fill)))
    if not paras:
        return
    if el is ctx.title_el and not ctx.slide.title:
        ctx.slide.title = _one_line(" ".join(t for _, t in paras))
        return
    kind = "subtitle" if ph is not None and ph.get("type") == "subTitle" else "text"
    ctx.slide.items.append(_Item("text", box, ctx.next_z(), head=kind, lines=_para_lines(paras)))


def _connector(ctx: _Ctx, el):
    st, en = _xp(el, "p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn/@id"), _xp(el, "p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn/@id")
    if not (st and en):
        return

    def arrow(tag):
        e = _find(el, f"p:spPr/a:ln/a:{tag}")
        return e is not None and e.get("type", "none") != "none"
    ctx.slide.connectors.append((st[0], en[0], arrow("headEnd"), arrow("tailEnd")))


# ---------------------------------------------------------------- tables

def _table(ctx: _Ctx, tbl, box):
    rows, origin_above = [], {}
    for tr in tbl.findall("a:tr", NS):
        row, last = [], ""
        for ci, tc in enumerate(tr.findall("a:tc", NS)):
            if tc.get("vMerge") in ("1", "true"):
                text = origin_above.get(ci, "")
            elif tc.get("hMerge") in ("1", "true"):
                text = last
            else:
                text = " / ".join(t.replace("\n", " ") for _, t in _paragraphs(_find(tc, "a:txBody"), ctx.part))
                origin_above[ci] = text
            last = text
            row.append(text)
        rows.append(row)
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        return
    merged = bool(_xp(tbl, "a:tr/a:tc[@gridSpan or @rowSpan]"))
    ctx.counts["tables"] += 1
    n = sum(1 for t in ctx.tables if t.name.startswith(f"slide{ctx.slide.number}_table")) + 1
    ctx.tables.append(RawTable(f"slide{ctx.slide.number}_table" + (f"_{n}" if n > 1 else ""), rows,
                               f"slide {ctx.slide.number} table" + (f" {n}" if n > 1 else "")))
    head = f"table | {len(rows)} rows x {max(len(r) for r in rows)} columns"
    if merged:
        head += " | merged cells repeat their value"
    ctx.slide.items.append(_Item("table", box, ctx.next_z(), head=head, lines=_md_table(rows).split("\n")))


# ---------------------------------------------------------------- charts

@dataclass
class _Series:
    name: str
    cats: list[str] = field(default_factory=list)
    vals: list[str] = field(default_factory=list)
    xs: list[str] = field(default_factory=list)
    sizes: list[str] = field(default_factory=list)


@dataclass
class _Plot:
    base: str                    # bar | line | pie | doughnut | area | scatter | bubble | radar | stock | surface | ofPie
    kind: str                    # "clustered column", "100% stacked bar", "line", ...
    horizontal: bool = False
    stacked: bool = False
    series: list[_Series] = field(default_factory=list)
    x_title: str = ""
    y_title: str = ""
    secondary: bool = False


@dataclass
class _Chart:
    title: str = ""
    plots: list[_Plot] = field(default_factory=list)


def _rich_text(el) -> str:
    if el is None:
        return ""
    return _one_line(" ".join("".join(_xp(p, ".//a:t/text()")) for p in _xp(el, ".//a:p")), 200)


_DATE_FMT = re.compile(r"[dy]|(?<![a-z])m{1,4}(?![a-z])", re.I)


def _is_date_format(fmt: str) -> bool:
    fmt = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", fmt or "")
    return bool(fmt) and fmt.lower() != "general" and not re.search(r"[0#]", fmt) and bool(_DATE_FMT.search(fmt))


def _serial_to_date(v: str) -> str:
    try:
        d = _dt.datetime(1899, 12, 30) + _dt.timedelta(days=float(v))
    except (ValueError, OverflowError):
        return v
    return d.date().isoformat() if d.time() == _dt.time(0) else d.isoformat(sep=" ", timespec="minutes")


def _cache_values(parent, *, as_number: bool) -> list[str]:
    """Saved values under c:cat / c:val / c:xVal / c:yVal / c:bubbleSize (index order, gaps kept)."""
    if parent is None:
        return []
    levels = _xp(parent, "c:multiLvlStrRef/c:multiLvlStrCache/c:lvl")
    if levels:                              # outer levels only label the first item of each group
        pc = _find(parent, "c:multiLvlStrRef/c:multiLvlStrCache/c:ptCount")
        cols = [{_int(pt, "idx"): pt.findtext("c:v", default="", namespaces=NS) for pt in lvl.findall("c:pt", NS)}
                for lvl in levels]
        count = max([_int(pc, "val") if pc is not None else 0] + [max(c, default=-1) + 1 for c in cols])
        cols.reverse()                      # outermost first
        last = [""] * len(cols)
        out = []
        for i in range(min(count, MAX_CHART_POINTS)):
            parts = []
            for li, pts in enumerate(cols):
                v = pts.get(i)
                if li < len(cols) - 1:      # carry a group label down to the rest of its group
                    if v is None:
                        v = last[li]
                    else:
                        last[li] = v
                parts.append(_one_line(v or ""))
            out.append(" / ".join(p for p in parts if p))
        return out
    caches = _xp(parent, "c:numRef/c:numCache | c:strRef/c:strCache | c:numLit | c:strLit")
    if not caches:
        return []
    cache = caches[0]
    numeric = _local(cache).startswith("num")
    pts = {_int(pt, "idx"): pt.findtext("c:v", default="", namespaces=NS) for pt in cache.findall("c:pt", NS)}
    pc = _find(cache, "c:ptCount")
    count = max(_int(pc, "val") if pc is not None else 0, max(pts, default=-1) + 1)
    count = min(count, MAX_CHART_POINTS)
    is_date = numeric and not as_number and _is_date_format(cache.findtext("c:formatCode", default="", namespaces=NS))
    out = []
    for i in range(count):
        v = pts.get(i)
        if v is None:
            out.append("")
        elif is_date:
            out.append(_serial_to_date(v))
        elif numeric:
            out.append(_num(v))
        else:
            out.append(_one_line(v))
    return out


def _plot_kind(plot) -> tuple[str, str, bool, bool]:
    name = _local(plot)
    three_d = "3D" in name
    base = name.replace("3D", "")[: -len("Chart")]
    grouping = (_xp(plot, "c:grouping/@val") or ["standard"])[0]
    stacked = grouping in ("stacked", "percentStacked")
    horizontal = False
    if base == "bar":
        horizontal = (_xp(plot, "c:barDir/@val") or ["col"])[0] == "bar"
        word = "bar" if horizontal else "column"
        kind = {"stacked": f"stacked {word}", "percentStacked": f"100% stacked {word}"}.get(grouping, f"clustered {word}")
    elif base in ("line", "area"):
        kind = {"stacked": f"stacked {base}", "percentStacked": f"100% stacked {base}"}.get(grouping, base)
    elif base == "ofPie":
        kind = "bar-of-pie" if (_xp(plot, "c:ofPieType/@val") or ["pie"])[0] == "bar" else "pie-of-pie"
    else:
        kind = {"scatter": "scatter (XY)"}.get(base, base)
    return base, kind + (" (3-D)" if three_d else ""), horizontal, stacked


def _parse_chart(space) -> _Chart:
    """c:chartSpace -> the data PowerPoint saved with the chart (no link to a workbook is followed)."""
    ch = _Chart()
    chart = _find(space, "c:chart")
    if chart is None:
        return ch
    ch.title = _rich_text(_find(chart, "c:title"))
    area = _find(chart, "c:plotArea")
    if area is None:
        return ch
    axes = {}
    for ax in area:
        if _local(ax) in ("catAx", "valAx", "dateAx", "serAx"):
            ax_id = (_xp(ax, "c:axId/@val") or [""])[0]
            axes[ax_id] = (_local(ax), _rich_text(_find(ax, "c:title")), (_xp(ax, "c:axPos/@val") or [""])[0])
    first_value_axis, cells = None, 0
    for plot in area:
        if not _local(plot).endswith("Chart"):
            continue
        base, kind, horizontal, stacked = _plot_kind(plot)
        pl = _Plot(base, kind, horizontal, stacked)
        mine = [(i, axes[i]) for i in _xp(plot, "c:axId/@val") if i in axes]
        if base in ("scatter", "bubble"):
            x_ax = [a for a in mine if a[1][2] in ("b", "t")] or mine[:1]
            y_ax = [a for a in mine if a not in x_ax]
        else:
            x_ax = [a for a in mine if a[1][0] in ("catAx", "dateAx")]
            y_ax = [a for a in mine if a[1][0] == "valAx"]
        pl.x_title = x_ax[0][1][1] if x_ax else ""
        pl.y_title = y_ax[0][1][1] if y_ax else ""
        if y_ax:
            if first_value_axis is None:
                first_value_axis = y_ax[0][0]
            pl.secondary = y_ax[0][0] != first_value_axis
        for i, ser in enumerate(plot.findall("c:ser", NS), 1):
            tx = _find(ser, "c:tx")
            name = ""
            if tx is not None:
                name = _one_line(" ".join(_xp(tx, "c:strRef/c:strCache/c:pt/c:v/text() | c:v/text()")), 120)
            s = _Series(name or f"Series {i}")
            if base in ("scatter", "bubble"):
                s.xs = _cache_values(_find(ser, "c:xVal"), as_number=False)
                s.vals = _cache_values(_find(ser, "c:yVal"), as_number=True)
                s.sizes = _cache_values(_find(ser, "c:bubbleSize"), as_number=True)
            else:
                s.cats = _cache_values(_find(ser, "c:cat"), as_number=False)
                s.vals = _cache_values(_find(ser, "c:val"), as_number=True)
            pl.series.append(s)
            cells += len(s.cats) + len(s.vals) + len(s.xs) + len(s.sizes)
            if cells > MAX_CHART_CELLS:
                break
        ch.plots.append(pl)
        if cells > MAX_CHART_CELLS:
            break
    return ch


def _chart_tables(ch: _Chart) -> list[list[list[str]]]:
    """One or more row-lists (header first) covering every plot of the chart."""
    tables = []
    cat_plots = [p for p in ch.plots if p.base not in ("scatter", "bubble") and p.series]
    xy_plots = [p for p in ch.plots if p.base in ("scatter", "bubble") and p.series]

    def cats_of(p):
        longest = max(p.series, key=lambda s: len(s.cats))
        n = max(len(s.vals) for s in p.series)
        return longest.cats + [str(i + 1) for i in range(len(longest.cats), n)]

    groups: list[list[_Plot]] = []
    for p in cat_plots:                      # plots that share categories go in one table (combo charts)
        for g in groups:
            if cats_of(g[0]) == cats_of(p):
                g.append(p)
                break
        else:
            groups.append([p])
    several = len(ch.plots) > 1
    for g in groups:
        cats = cats_of(g[0])
        header = [g[0].x_title or "category"]
        cols = []
        for p in g:
            for s in p.series:
                tag = f" ({p.kind}{', secondary axis' if p.secondary else ''})" if several else ""
                header.append(s.name + tag)
                cols.append(s.vals)
        rows = [header] + [[c] + [col[i] if i < len(col) else "" for col in cols] for i, c in enumerate(cats)]
        tables.append(rows)
    for p in xy_plots:
        bubble = p.base == "bubble"
        rows = [["series", p.x_title or "x", p.y_title or "y"] + (["size"] if bubble else [])]
        for s in p.series:
            for i in range(max(len(s.xs), len(s.vals))):
                row = [s.name, s.xs[i] if i < len(s.xs) else "", s.vals[i] if i < len(s.vals) else ""]
                if bubble:
                    row.append(s.sizes[i] if i < len(s.sizes) else "")
                rows.append(row)
        tables.append(rows)
    return [t for t in tables if len(t) > 1]


def _add_chart_item(ctx: _Ctx, box, kind: str, title: str, axes: list[str], tables, chart: _Chart | None = None):
    ctx.counts["charts"] += 1
    n = sum(1 for t in ctx.tables if re.match(rf"slide{ctx.slide.number}_chart(_\d+)?$", t.name))
    head = " | ".join(["chart", kind or "unknown type"] + ([f"title: {title}"] if title else []) + axes)
    lines = []
    limit = ctx.opts["max_chart_rows"]
    for rows in tables:
        n += 1
        ctx.tables.append(RawTable(f"slide{ctx.slide.number}_chart" + (f"_{n}" if n > 1 else ""), rows,
                                   f"slide {ctx.slide.number} chart" + (f" '{title}'" if title else "")))
        lines += _csv(rows[: limit + 1]).split("\n")
        if len(rows) - 1 > limit:
            lines.append(f"(first {limit} of {len(rows) - 1} rows shown)")
    if not tables:
        lines.append("(no data is saved with this chart)")
    item = _Item("chart", box, ctx.next_z(), head=head, lines=lines)
    ctx.slide.items.append(item)
    if chart is not None and tables and ctx.opts["render_charts"] and ctx.opts["include_images"]:
        png = _render_chart_png(chart, box, ctx)
        if png:
            sha1 = hashlib.sha1(png).hexdigest()
            img = ctx.images[(sha1, None)] = _Img(png, "image/png", "png", None, "chart render", sha1)
            use = _Use(img, ctx.slide.number, box, item.z + 0.5, "",
                       "drawn from the chart data above; the original styling is not kept")
            img.uses.append(use)
            ctx.slide.items.append(_Item("image", box, use.z, use=use))


def _chart(ctx: _Ctx, rid: str, box):
    try:
        part = ctx.part.rels[rid].target_part
        space = getattr(part, "_element", None)
        if space is None:
            space = etree.fromstring(part.blob, _SAFE_PARSER)
        ch = _parse_chart(space)
    except Exception:  # noqa: BLE001 - chart XML varies a lot; never fail the whole deck for one chart
        ctx.counts["charts"] += 1
        ctx.slide.items.append(_Item("chart", box, ctx.next_z(), head="chart", lines=["(chart data could not be read)"]))
        return
    kinds = []
    for p in ch.plots:
        if p.kind not in kinds:
            kinds.append(p.kind)
    axes = []
    x_t = next((p.x_title for p in ch.plots if p.x_title), "")
    y_t = next((p.y_title for p in ch.plots if p.y_title and not p.secondary), "")
    y2_t = next((p.y_title for p in ch.plots if p.y_title and p.secondary), "")
    horizontal = any(p.horizontal for p in ch.plots)
    if x_t:
        axes.append(f"{'vertical' if horizontal else 'horizontal'} axis: {x_t}")
    if y_t:
        axes.append(f"{'horizontal' if horizontal else 'vertical'} axis: {y_t}")
    if y2_t:
        axes.append(f"secondary axis: {y2_t}")
    _add_chart_item(ctx, box, " + ".join(kinds), ch.title, axes, _chart_tables(ch), ch)


def _chartex(ctx: _Ctx, rid: str, box):
    """Newer chart types (waterfall, treemap, sunburst, funnel, histogram, box & whisker)."""
    try:
        root = etree.fromstring(ctx.part.rels[rid].target_part.blob, _SAFE_PARSER)
    except Exception:  # noqa: BLE001
        ctx.counts["charts"] += 1
        ctx.slide.items.append(_Item("chart", box, ctx.next_z(), head="chart", lines=["(chart data could not be read)"]))
        return
    data = {}
    for d in _xp(root, "cx:chartData/cx:data"):
        dims = {}
        for dim in d:
            if _local(dim) not in ("strDim", "numDim"):
                continue
            lvls = dim.findall("cx:lvl", NS)
            if not lvls:
                continue
            vals = []
            count = max((_int(lv, "ptCount") for lv in lvls), default=0)
            cols = [{_int(pt, "idx"): (pt.text or "") for pt in lv.findall("cx:pt", NS)} for lv in lvls]
            count = min(max([count] + [max(c, default=-1) + 1 for c in cols]), MAX_CHART_POINTS)
            for i in range(count):
                if _local(dim) == "numDim":
                    vals.append(_num(cols[0].get(i)))
                else:
                    vals.append(" / ".join(_one_line(c[i]) for c in reversed(cols) if c.get(i)))
            dims[dim.get("type", "val")] = vals
        data[d.get("id")] = dims
    kinds, tables = [], []
    for i, ser in enumerate(_xp(root, "cx:chart/cx:plotArea/cx:plotAreaRegion/cx:series"), 1):
        kind = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", ser.get("layoutId", "")).lower()
        if kind and kind not in kinds:
            kinds.append(kind)
        name = _one_line(" ".join(_xp(ser, "cx:tx/cx:txData/cx:v/text()")), 120) or f"Series {i}"
        dims = data.get((_xp(ser, "cx:dataId/@val") or [None])[0], {})
        cats = dims.get("cat", [])
        num_cols = [(k, v) for k, v in dims.items() if k != "cat"]
        if not num_cols:
            continue
        n = max(len(v) for _, v in num_cols)
        header = ["category"] + [name if len(num_cols) == 1 else f"{name} ({k})" for k, _ in num_cols]
        rows = [header] + [[cats[j] if j < len(cats) else str(j + 1)] + [v[j] if j < len(v) else "" for _, v in num_cols]
                           for j in range(n)]
        tables.append(rows)
    title = _rich_text(_find(root, "cx:chart/cx:title")) or \
        _one_line(" ".join(_xp(root, "cx:chart/cx:title/cx:tx/cx:txData/cx:v/text()")), 200)
    _add_chart_item(ctx, box, " + ".join(kinds), title, [], tables)


def _render_chart_png(ch: _Chart, box, ctx: _Ctx) -> bytes | None:
    """Optional picture of a chart, redrawn from its data with matplotlib."""
    try:
        from matplotlib.backends.backend_agg import FigureCanvasAgg
        from matplotlib.figure import Figure
    except ImportError:
        ctx.warn("render_charts needs matplotlib, which is not installed; charts are sent as data only.")
        return None

    def nums(vals):
        out = []
        for v in vals:
            try:
                out.append(float(v))
            except (TypeError, ValueError):
                out.append(float("nan"))
        return out

    try:
        aspect = (box[2] / box[3]) if box and box[3] else 1.6
        fig = Figure(figsize=(8, min(max(8 / max(aspect, 0.2), 3.5), 8)), dpi=100)
        FigureCanvasAgg(fig)
        ax = fig.add_subplot()
        ax2 = None
        tick_labels = None
        colors = (f"C{i % 10}" for i in range(10_000))      # one colour sequence across all plots and both axes
        for p in ch.plots:
            target = ax
            if p.secondary:
                if ax2 is None:
                    ax2 = ax.twinx()
                    ax2.set_zorder(ax.get_zorder() + 1)     # lines on the secondary axis stay in front of the bars
                    ax2.patch.set_visible(False)
                target = ax2
            if p.base in ("pie", "doughnut", "ofPie"):
                s = p.series[0]
                vals = [0 if v != v else max(v, 0) for v in nums(s.vals)]
                labels = (s.cats + [""] * len(vals))[: len(vals)]
                target.pie(vals, labels=labels, autopct="%1.0f%%",
                           wedgeprops={"width": 0.45} if p.base == "doughnut" else None)
                continue
            if p.base in ("scatter", "bubble"):
                for s in p.series:
                    xs, ys = nums(s.xs), nums(s.vals)
                    n = min(len(xs), len(ys))
                    if p.base == "bubble" and s.sizes:
                        sz = nums(s.sizes)[:n]
                        top = max([v for v in sz if v == v] or [1]) or 1
                        target.scatter(xs[:n], ys[:n], s=[1500 * (v if v == v else 0) / top for v in sz],
                                       alpha=0.6, label=s.name, color=next(colors))
                        target.margins(0.25)
                    else:
                        target.scatter(xs[:n], ys[:n], label=s.name, color=next(colors))
                continue
            n = max(len(s.vals) for s in p.series)
            idx = list(range(n))
            cats = max((s.cats for s in p.series), key=len)
            tick_labels = (cats + [str(i + 1) for i in range(len(cats), n)])[:n]
            bottoms = [0.0] * n
            for k, s in enumerate(p.series):
                vals = (nums(s.vals) + [float("nan")] * n)[:n]
                color = next(colors)
                if p.base == "bar":
                    width = 0.8 if p.stacked else 0.8 / len(p.series)
                    at = idx if p.stacked else [i - 0.4 + width * (k + 0.5) for i in idx]
                    base = bottoms if p.stacked else None
                    if p.horizontal:
                        target.barh(at, vals, height=width, left=base, label=s.name, color=color)
                    else:
                        target.bar(at, vals, width=width, bottom=base, label=s.name, color=color)
                elif p.base == "area":
                    low = list(bottoms) if p.stacked else [0.0] * n
                    high = [a + (0 if v != v else v) for a, v in zip(low, vals)]
                    target.fill_between(idx, low, high, alpha=0.6, label=s.name, color=color)
                else:
                    target.plot(idx, vals, marker="o", label=s.name, color=color)
                if p.stacked:
                    bottoms = [b + (0 if v != v else v) for b, v in zip(bottoms, vals)]
            if p.horizontal:
                target.set_yticks(idx, tick_labels)
                target.invert_yaxis()
            else:
                long = sum(len(t) for t in tick_labels) > 60
                target.set_xticks(idx, tick_labels, rotation=45 if long else 0, ha="right" if long else "center")
            if p.x_title:
                (target.set_ylabel if p.horizontal else target.set_xlabel)(p.x_title)
            if p.y_title:
                (target.set_xlabel if p.horizontal else target.set_ylabel)(p.y_title)
        if ch.title:
            ax.set_title(ch.title)
        handles, labels = ax.get_legend_handles_labels()
        if ax2 is not None:
            h2, l2 = ax2.get_legend_handles_labels()
            handles, labels = handles + h2, labels + l2
        if len(labels) > 1 or (labels and not ch.title):
            ax.legend(handles, labels, fontsize=8)
        fig.tight_layout()
        buf = io.BytesIO()
        fig.savefig(buf, format="png")
        return buf.getvalue()
    except Exception:  # noqa: BLE001 - a drawing failure must not lose the data table
        ctx.warn("Some charts could not be drawn; their data is still included as text.")
        return None


# ---------------------------------------------------------------- SmartArt, embedded objects, frames

def _diagram(ctx: _Ctx, rid: str, box):
    ctx.counts["diagrams"] += 1
    lines = []
    try:
        root = etree.fromstring(ctx.part.rels[rid].target_part.blob, _SAFE_PARSER)
        pts, root_id = {}, None
        for pt in _xp(root, "dgm:ptLst/dgm:pt"):
            typ = pt.get("type", "node")
            if typ == "doc":
                root_id = pt.get("modelId")
            if typ in ("node", "asst", "doc"):
                pts[pt.get("modelId")] = _one_line(" ".join(_xp(pt, "dgm:t//a:t/text()")))
        kids: dict[str, list] = {}
        for cxn in _xp(root, "dgm:cxnLst/dgm:cxn"):
            if cxn.get("type", "parOf") == "parOf":
                kids.setdefault(cxn.get("srcId"), []).append((_int(cxn, "srcOrd"), cxn.get("destId")))
        seen = set()

        def walk(node, depth):
            for _, kid in sorted(kids.get(node, [])):
                if kid in seen or kid not in pts or depth > 12:
                    continue
                seen.add(kid)
                if pts[kid]:
                    lines.append("  " * depth + "- " + pts[kid])
                walk(kid, depth + 1)
        if root_id is not None:
            walk(root_id, 0)
        for mid, text in pts.items():        # anything not reachable from the root
            if text and mid not in seen and mid != root_id:
                lines.append("- " + text)
    except Exception:  # noqa: BLE001
        lines = []
    ctx.slide.items.append(_Item("diagram", box, ctx.next_z(), head="SmartArt diagram (items indented under their parent)",
                                 lines=lines or ["(diagram text could not be read)"]))


def _frame(ctx: _Ctx, el, tf):
    box = _box(ctx, el, "p:xfrm", tf)
    nv = _c_nv_pr(el)
    alt = (nv.get("descr") or "") if nv is not None else ""
    tbl = _xp(el, "a:graphic/a:graphicData/a:tbl")
    if tbl:
        return _table(ctx, tbl[0], box)
    rid = _xp(el, "a:graphic/a:graphicData/c:chart/@r:id")
    if rid:
        return _chart(ctx, rid[0], box)
    rid = _xp(el, "a:graphic/a:graphicData/cx:chart/@r:id")
    if rid:
        return _chartex(ctx, rid[0], box)
    rid = _xp(el, "a:graphic/a:graphicData/dgm:relIds/@r:dm")
    if rid:
        return _diagram(ctx, rid[0], box)
    ole = _xp(el, "a:graphic/a:graphicData//p:oleObj")
    if ole:
        ctx.counts["objects"] += 1
        prog = _one_line(ole[0].get("progId") or "", 60)
        name = {"Excel": "Excel worksheet", "Word": "Word document", "PowerPoint": "PowerPoint slide",
                "Acro": "PDF", "Package": "attached file"}
        label = next((v for k, v in name.items() if prog.startswith(k)), prog or "object")
        blips = _xp(el, "a:graphic/a:graphicData//p:oleObj//p:pic/p:blipFill/a:blip")
        ctx.slide.items.append(_Item("object", box, ctx.next_z(), head=f"embedded {label} (contents not read"
                                     + ("; its preview picture follows)" if blips else ")")))
        if blips:
            ctx.slide.items.append(_register_image(ctx, ctx.part, blips[0], box, alt=alt, note="preview of the embedded object",
                                                   source="object preview"))
        ctx.warn("Embedded objects (e.g. Excel worksheets) are shown by their preview picture only; "
                 "upload the original file if its contents matter.")


def _walk(ctx: _Ctx, parent, tf, depth: int = 0):
    for el in parent:
        if not isinstance(el.tag, str):
            continue
        q = etree.QName(el)
        if q.namespace == NS["mc"] and q.localname == "AlternateContent":
            before = len(ctx.slide.items)
            choice, fallback = _find(el, "mc:Choice"), _find(el, "mc:Fallback")
            if choice is not None:
                _walk(ctx, choice, tf, depth)
            if len(ctx.slide.items) == before and fallback is not None:
                _walk(ctx, fallback, tf, depth)
            continue
        if q.namespace != NS["p"]:
            continue
        nv = _c_nv_pr(el)
        if nv is not None and nv.get("hidden") in ("1", "true"):
            continue
        if q.localname == "grpSp":
            if depth < 25:
                _walk(ctx, el, _group_tf(el, tf), depth + 1)
        elif q.localname == "sp":
            _shape(ctx, el, tf)
        elif q.localname == "pic":
            _picture(ctx, el, tf)
        elif q.localname == "graphicFrame":
            _frame(ctx, el, tf)
        elif q.localname == "cxnSp":
            _connector(ctx, el)


# ---------------------------------------------------------------- one slide

def _comments(part) -> list[str]:
    out = []
    try:
        for rel in part.rels.values():
            if rel.is_external or not rel.reltype.endswith("/comments"):
                continue
            root = etree.fromstring(rel.target_part.blob, _SAFE_PARSER)
            for t in _xp(root, "//p:cm/p:text"):
                if (t.text or "").strip():
                    out.append(_one_line(t.text, 500))
            for body in _xp(root, "//p188:txBody"):
                text = _one_line(" ".join(_xp(body, ".//a:t/text()")), 500)
                if text:
                    out.append(text)
    except Exception:  # noqa: BLE001
        pass
    return out


def _read_slide(ctx: _Ctx, slide, number: int) -> _Slide:
    s = ctx.slide = _Slide(number, hidden=slide._element.get("show") in ("0", "false"))
    ctx.part, ctx.z = slide.part, 0
    ctx.ph_boxes, ctx.title_el = {}, None
    try:
        for ph in slide.placeholders:
            if None not in (ph.left, ph.top, ph.width, ph.height):
                ctx.ph_boxes[ph._element] = (ph.left, ph.top, ph.width, ph.height)
        title = slide.shapes.title
        ctx.title_el = title._element if title is not None else None
    except Exception:  # noqa: BLE001 - odd placeholders only cost us an inherited position
        pass
    full = (0, 0, ctx.w, ctx.h)
    bg = _xp(slide._element, "p:cSld/p:bg/p:bgPr/a:blipFill/a:blip")
    if bg:
        s.items.append(_register_image(ctx, ctx.part, bg[0], full, note="slide background", source="background", z=0))
    if ctx.opts["include_template_images"]:
        for owner in (slide.slide_layout.slide_master, slide.slide_layout):
            for pic in _xp(owner._element, "p:cSld/p:spTree//p:pic[not(p:nvPicPr/p:nvPr/p:ph)]"):
                blip_fill = _find(pic, "p:blipFill")
                blip = _find(blip_fill, "a:blip") if blip_fill is not None else None
                box = _xfrm_box(_find(pic, "p:spPr/a:xfrm"))
                nv = _c_nv_pr(pic)
                s.items.append(_register_image(ctx, owner.part, blip, box, note="from the slide template",
                                               alt=(nv.get("descr") or "") if nv is not None else "",
                                               source="template", crop=_crop_of(blip_fill), z=0))
    tree = _xp(slide._element, "p:cSld/p:spTree")
    if tree:
        _walk(ctx, tree[0], lambda b: b)
    if ctx.opts["include_notes"]:
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
                s.notes = _clean(slide.notes_slide.notes_text_frame.text).strip()
        except Exception:  # noqa: BLE001
            pass
    s.comments = _comments(slide.part)
    return s


def _reading_order(ctx: _Ctx, items: list[_Item]) -> list[_Item]:
    """Top to bottom, then left to right; shapes whose tops are within 4% of the slide height share
    a row. A text box sitting directly under a picture, chart, table or diagram is treated as its
    caption and listed right after it."""
    placed = sorted((i for i in items if i.box is not None), key=lambda i: (i.box[1], i.box[0], i.z))
    tol = ctx.h * 0.04
    rows: list[list[_Item]] = []
    for it in placed:
        if rows and it.box[1] - rows[-1][0].box[1] <= tol:
            rows[-1].append(it)
        else:
            rows.append([it])
    ordered = [i for row in rows for i in sorted(row, key=lambda i: (i.box[0], i.z))]
    media = [i for i in ordered if i.kind != "text"]
    captions: dict[int, list[_Item]] = {}
    captioned = set()
    for t in (i for i in ordered if i.kind == "text"):
        best = None
        for m in media:
            gap = t.box[1] - (m.box[1] + m.box[3])
            overlap = min(t.box[0] + t.box[2], m.box[0] + m.box[2]) - max(t.box[0], m.box[0])
            if -0.02 * ctx.h <= gap <= 0.08 * ctx.h and overlap >= 0.5 * min(t.box[2], m.box[2]) \
                    and (best is None or gap <= best[0]):
                best = (gap, m)
        if best:
            captions.setdefault(id(best[1]), []).append(t)
            captioned.add(id(t))
    out = []
    for it in ordered:
        if id(it) not in captioned:
            out.append(it)
            out += captions.get(id(it), [])
    return out + sorted((i for i in items if i.box is None), key=lambda i: i.z)


# ---------------------------------------------------------------- image selection + encoding

def _vector_kind(img: _Img) -> str:
    b = img.blob
    if b[40:44] == b" EMF" or img.ext == "emf":
        return "EMF"
    if b[:4] == b"\xd7\xcd\xc6\x9a" or b[:4] == b"\x01\x00\x09\x00" or img.ext == "wmf":
        return "WMF"
    if "svg" in img.content_type or img.ext == "svg" or b.lstrip()[:5].lower() in (b"<svg ", b"<?xml"):
        return "SVG"
    return ""


def _trim_page(png: bytes) -> bytes:
    """LibreOffice draws a metafile on an empty page; cut the page margins away."""
    from PIL import Image, ImageChops
    with Image.open(io.BytesIO(png)) as im:
        rgb = Image.new("RGB", im.size, (255, 255, 255))
        rgb.paste(im, mask=im.getchannel("A") if im.mode == "RGBA" else None)
    box = ImageChops.difference(rgb, Image.new("RGB", rgb.size, rgb.getpixel((0, 0)))).getbbox()
    if not box or box == (0, 0, *rgb.size):
        return png
    buf = io.BytesIO()
    rgb.crop(box).save(buf, "PNG", optimize=True)
    return buf.getvalue()


def _vectors_to_png(imgs: list[_Img]) -> None:
    """EMF/WMF/SVG -> PNG with headless LibreOffice. Leaves status='vector' when that fails.

    First pass: plain export (the drawing on a page at 96 dpi), which also tells us the page size.
    Second pass: the same page about 3000 px tall so small print stays readable. Margins are trimmed."""
    from PIL import Image
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice or not imgs:
        return
    with tempfile.TemporaryDirectory() as tmp:
        tmp_p = Path(tmp)
        srcs = []
        for i, img in enumerate(imgs):
            src = tmp_p / f"v{i}.{img.fmt.lower()}"
            src.write_bytes(img.blob)
            srcs.append(src)

        def export(fmt: str, outdir: Path) -> bool:
            cmd = [soffice, f"-env:UserInstallation=file://{tmp_p / 'profile'}", "--headless", "--convert-to", fmt,
                   "--outdir", str(outdir), *map(str, srcs)]
            try:
                subprocess.run(cmd, check=True, capture_output=True, timeout=LIBREOFFICE_TIMEOUT_S)
                return True
            except (subprocess.SubprocessError, OSError):
                return False

        if not export("png", tmp_p / "lo"):
            return
        pages = {}
        for i in range(len(imgs)):
            try:
                with Image.open(tmp_p / "lo" / f"v{i}.png") as im:
                    pages[i] = im.size
            except Exception:  # noqa: BLE001
                pass
        if not pages:
            return
        page = max(set(pages.values()), key=list(pages.values()).count)
        scale = 3000 / max(page)
        hi_ok = scale > 1.2 and export(
            'png:draw_png_Export:{"PixelWidth":{"type":"long","value":"%d"},"PixelHeight":{"type":"long","value":"%d"}}'
            % (round(page[0] * scale), round(page[1] * scale)), tmp_p / "hi")
        for i, img in enumerate(imgs):
            if i not in pages:
                continue
            hi = tmp_p / "hi" / f"v{i}.png"
            best = hi if hi_ok and pages[i] == page and hi.exists() and hi.stat().st_size else tmp_p / "lo" / f"v{i}.png"
            try:
                img.blob = _trim_page(best.read_bytes())
                img.status = "ok"
            except Exception:  # noqa: BLE001
                pass


def _crop_px(crop, w: int, h: int):
    if not crop:
        return None
    l, t, r, b = crop
    box = (round(w * l / 100000), round(h * t / 100000), round(w * (1 - r / 100000)), round(h * (1 - b / 100000)))
    return box if box[2] - box[0] >= 1 and box[3] - box[1] >= 1 else None


def _probe(img: _Img, min_px: int) -> None:
    from PIL import Image
    try:
        with Image.open(io.BytesIO(img.blob)) as im:
            if not img.fmt:
                img.fmt = im.format or "?"
            img.orig_size = im.size
            box = _crop_px(img.crop, *im.size)
            img.size = (box[2] - box[0], box[3] - box[1]) if box else im.size
    except Exception:  # noqa: BLE001 - includes Pillow's decompression-bomb guard
        img.status = "unreadable"
        return
    if img.orig_size[0] * img.orig_size[1] > MAX_SOURCE_PIXELS:
        img.status = "huge"
    elif min(img.size) < min_px:
        img.status = "small"


def _fit(w: int, h: int, cap: int) -> tuple[int, int]:
    """Largest size (never larger than w x h) with long edge <= 1568 px and at most `cap` visual tokens."""
    s = min(1.0, MAX_EDGE_PX / max(w, h), math.sqrt(cap * PATCH_PX * PATCH_PX / (w * h)))
    while True:
        nw, nh = max(1, int(w * s)), max(1, int(h * s))
        if _img_tokens(nw, nh) <= cap or max(nw, nh) <= PATCH_PX:
            return nw, nh
        s *= 0.97


def _encode(img: _Img, cap: int) -> tuple[bytes, str, int, int]:
    from PIL import Image
    with Image.open(io.BytesIO(img.blob)) as im:
        fmt, (w, h) = im.format, im.size
        box = _crop_px(img.crop, w, h)
        tw, th = _fit(*img.size, cap)
        if (fmt in API_MEDIA_TYPES and box is None and (tw, th) == (w, h) and len(img.blob) <= MAX_IMAGE_RAW_BYTES
                and im.mode != "CMYK" and not getattr(im, "is_animated", False)):
            im.load()                                              # a damaged file fails here, not at the API
            return img.blob, API_MEDIA_TYPES[fmt], w, h            # send the original bytes untouched
        if fmt == "JPEG" and box is None:
            im.draft(None, (tw, th))                                # decode big JPEGs at reduced size
        im.seek(0)
        im.load()
        frame = im.crop(box) if box else im.copy()
    alpha = frame.mode in ("RGBA", "LA", "PA") or "transparency" in frame.info
    if frame.mode in ("I", "F") or frame.mode.startswith("I;"):
        frame = frame.convert("L") if frame.mode == "F" else frame.convert("I").point(lambda v: v * (1 / 256)).convert("L")
    frame = frame.convert("RGBA" if alpha else "RGB")
    if frame.size != (tw, th):
        frame = frame.resize((tw, th), Image.LANCZOS)

    def png(f):
        b = io.BytesIO()
        f.save(b, "PNG", optimize=True)
        return b.getvalue()

    def jpeg(f, quality):
        if f.mode == "RGBA":
            bg = Image.new("RGB", f.size, (255, 255, 255))
            bg.paste(f, mask=f.getchannel("A"))
            f = bg
        b = io.BytesIO()
        f.save(b, "JPEG", quality=quality, optimize=True)
        return b.getvalue()

    if fmt in ("JPEG", "MPO") and not alpha:
        data, media = jpeg(frame, 85), "image/jpeg"
    else:
        data, media = png(frame), "image/png"
        if not alpha and len(data) > 400_000:              # photo saved as BMP/TIFF/PNG: JPEG is far smaller
            j = jpeg(frame, 85)
            if len(j) < len(data) * 0.6:
                data, media = j, "image/jpeg"
    quality = 80
    while len(data) > MAX_IMAGE_RAW_BYTES:
        if quality < 50:
            frame = frame.resize((max(1, frame.width * 3 // 4), max(1, frame.height * 3 // 4)), Image.LANCZOS)
            quality = 80
        data, media = jpeg(frame, quality), "image/jpeg"
        quality -= 15
    return data, media, frame.width, frame.height


def _keep_rank(ctx: _Ctx, img: _Img) -> tuple:
    """Sort key for what to give up first when over budget: backgrounds and template pictures,
    then whatever is smallest on its slide; between equals, later slides go before earlier ones."""
    area = max((u.box[2] * u.box[3] / (ctx.w * ctx.h) for u in img.uses if u.box), default=0.0)
    return (img.source not in ("background", "template"), round(area, 4), -img.uses[0].order)


def _finalize_images(ctx: _Ctx) -> list[SlideImage]:
    o = ctx.opts
    imgs = list(ctx.images.values())
    if not o["include_images"]:
        for img in imgs:
            img.status = "off"
        return []
    for img in imgs:
        img.fmt = _vector_kind(img)
        if img.fmt:
            img.status = "vector"
    if o["use_libreoffice"]:
        _vectors_to_png([i for i in imgs if i.status == "vector"])
    for img in imgs:
        if img.status == "ok":
            _probe(img, o["min_image_px"])
    ok = [i for i in imgs if i.status == "ok"]
    by_area = sorted(ok, key=lambda i: _keep_rank(ctx, i))
    if o["max_images"] is not None:
        for img in by_area[: max(0, len(ok) - o["max_images"])]:
            img.status = "budget"
    ok = [i for i in ok if i.status == "ok"]

    def cost(cap):
        return sum(_img_tokens(*_fit(*i.size, cap)) for i in ok)

    unlimited = o["image_token_budget"] is None
    tier = 0 if unlimited else next((k for k, cap in enumerate(TOKEN_TIERS) if cost(cap) <= o["image_token_budget"]), None)
    if tier is None:
        tier = len(TOKEN_TIERS) - 1
        for img in sorted(ok, key=lambda i: _keep_rank(ctx, i)):
            if cost(TOKEN_TIERS[tier]) <= o["image_token_budget"]:
                break
            img.status = "budget"
            ok.remove(img)
    while True:
        encoded, total = {}, 0
        for img in list(ok):
            try:
                encoded[id(img)] = _encode(img, TOKEN_TIERS[tier])
                total += len(encoded[id(img)][0])
            except Exception:  # noqa: BLE001
                img.status = "unreadable"
                ok.remove(img)
        if unlimited or total <= MAX_TOTAL_IMAGE_RAW_BYTES:
            break
        if tier < len(TOKEN_TIERS) - 1:
            tier += 1
        else:                                                   # still too many bytes: drop the heaviest
            worst = max(ok, key=lambda i: len(encoded[id(i)][0]))
            worst.status = "budget"
            ok.remove(worst)
    out = []
    for img in sorted(imgs, key=lambda i: i.uses[0].order):     # ids follow the order pictures appear in the text
        if img.status != "ok":
            continue
        data, media, w, h = encoded[id(img)]
        img.out = SlideImage(f"img_{len(out) + 1}", media, data, w, h, _img_tokens(w, h), img.source,
                             img.fmt, img.orig_size, img.sha1, [_location(ctx, u) for u in img.uses])
        out.append(img.out)
    ctx.image_tier = TOKEN_TIERS[tier]
    return out


# ---------------------------------------------------------------- rendering text

_REASONS = {"off": "pictures are turned off", "small": "tiny icon or decoration",
            "budget": "over the image budget for one message",
            "huge": "the picture is too large to process", "unreadable": "the image data could not be decoded"}


def _image_line(ctx: _Ctx, use: _Use, positions: bool) -> str:
    img = use.img
    where = [_pos(ctx, use.box)] if positions else []
    if img is not None and img.out is not None and img.uses[0] is not use:
        return "[" + " | ".join([f"{img.out.id} again"] + where) + "]"      # described where it first appears
    details = where + ([use.note] if use.note else [])
    if img is not None and img.crop:
        details.append("cropped as on the slide")
    if use.alt:
        details.append(f"alt: {use.alt}")
    if img is None:
        head = f"picture not attached ({use.problem})"
    elif img.out is None:
        reason = _REASONS.get(img.status) or f"{img.fmt} vector picture; converting it needs LibreOffice"
        head = f"picture not attached ({reason})"
    else:
        head = f"image {img.out.id}"
    return "[" + " | ".join([head] + details) + "]"


def _slide_text(ctx: _Ctx, s: _Slide) -> str:
    mode = ctx.opts["positions"]
    head = f"## Slide {s.number}" + (f": {s.title}" if s.title else "") + (" (hidden slide)" if s.hidden else "")
    lines = [head]
    for it in s.items:
        if it.kind == "image":
            lines.append(_image_line(ctx, it.use, mode != "none"))
            continue
        if it.kind == "text":
            if mode == "all":
                lines.append(f"[{it.head} | {_pos(ctx, it.box)}]")
        elif mode == "none":
            lines.append(f"[{it.head}]")
        else:
            kind, _, rest = it.head.partition(" | ")
            lines.append("[" + " | ".join([kind, _pos(ctx, it.box)] + ([rest] if rest else [])) + "]")
        lines += it.lines
    if s.connectors:
        parts = []
        for a, b, head_arrow, tail_arrow in s.connectors:
            la, lb = s.labels.get(a) or f"shape {a}", s.labels.get(b) or f"shape {b}"
            if head_arrow and not tail_arrow:
                la, lb = lb, la
            parts.append(f'"{la}" {"<->" if head_arrow and tail_arrow else "->" if head_arrow or tail_arrow else "--"} "{lb}"')
        lines.append("Connectors: " + "; ".join(parts))
    if s.comments:
        lines.append("Review comments: " + " | ".join(s.comments))
    if s.notes:
        lines.append("Speaker notes: " + s.notes.replace("\n", "\n  "))
    if len(lines) == 1:
        lines.append("(no content)")
    return "\n".join(lines)


def _slide_outline(s: _Slide) -> str:
    """One line saying what a slide holds: enough to decide whether to open it."""
    bits = []
    body = _one_line(" ".join(line.lstrip(" -") for it in s.items if it.kind == "text" for line in it.lines), 90)
    if body:
        bits.append(f'text: "{body}"')
    pictures, missing = [], 0
    for it in s.items:
        parts = it.head.split(" | ")
        if it.kind == "chart":
            title = next((p[7:] for p in parts if p.startswith("title: ")), "")
            bits.append(f"chart: {parts[1] if len(parts) > 1 else 'chart'}" + (f' "{_one_line(title, 60)}"' if title else ""))
        elif it.kind == "table":
            bits.append(parts[1] if len(parts) > 1 else "table")
            bits[-1] = "table: " + bits[-1].replace(" rows x ", "x").replace(" columns", "")
        elif it.kind == "diagram":
            bits.append("SmartArt diagram")
        elif it.kind == "object":
            bits.append(it.head.split(" (")[0])
        elif it.kind == "image":
            img = it.use.img
            if img is None or img.out is None:
                missing += 1
            elif img.uses[0] is it.use:
                label = it.use.alt or it.use.note
                pictures.append(img.out.id + (f' "{_one_line(label, 50)}"' if label else ""))
    if pictures:
        bits.append("pictures: " + ", ".join(pictures))
    if missing:
        bits.append(f"{missing} picture{'s' if missing != 1 else ''} not available")
    extras = [name for name, has in (("connectors", s.connectors), ("comments", s.comments), ("notes", s.notes)) if has]
    if extras:
        bits.append("+ " + ", ".join(extras))
    return (f"{s.number}. {s.title or '(no title)'}" + (" (hidden)" if s.hidden else "")
            + (" | " + " | ".join(bits) if bits else " | (no content)"))


# ---------------------------------------------------------------- public API

def read_pptx(source, *, filename: str | None = None, include_images: bool = True,
              image_token_budget: int | None = DEFAULT_IMAGE_TOKEN_BUDGET,
              max_images: int | None = DEFAULT_MAX_IMAGES,
              min_image_px: int = 24, positions: str = "media", include_hidden_slides: bool = True,
              include_notes: bool = True, include_template_images: bool = False, render_charts: bool = False,
              max_chart_rows: int = 60, use_libreoffice: bool = True) -> PptxResult:
    """Read a .pptx (also .pptm/.ppsx/.potx). `source` is a path, bytes, or a binary file object.

    include_images      False -> text only; pictures stay as "[picture not attached ...]" lines with position + alt text.
    image_token_budget  Cap on the estimated tokens of all attached pictures. Pictures are first shrunk
                        (1568 -> 1000 -> 600 -> 350 tokens each); only if that isn't enough are pictures left out:
                        backgrounds first, then the ones that are smallest on their slide.
                        None = no budget: every picture at full size (up to 1568 tokens each). Use that only
                        when pictures are sent a few at a time, as deck_store.py does.
    max_images          Hard cap on attached pictures (the API allows 100 per request on 200k-context models).
                        None = no cap.
    min_image_px        Pictures with a side under this many pixels are treated as icons and not attached.
    positions           "media" (default): position on pictures, charts, tables, diagrams. "all": on text too.
                        "none": no positions.
    include_template_images  Also include pictures that come from the slide master/layout (logos).
    render_charts       Also attach a picture of each chart, redrawn from its data (needs matplotlib).
                        Off by default: the data table already says everything the redrawn picture would.
    use_libreoffice     Convert EMF/WMF/SVG pictures with LibreOffice when it is installed.

    Raises ConversionError with a message that is safe to show the user.
    """
    if positions not in ("media", "all", "none"):
        raise ValueError("positions must be 'media', 'all' or 'none'")
    opts = dict(include_images=include_images, image_token_budget=image_token_budget, max_images=max_images,
                min_image_px=min_image_px, positions=positions, include_notes=include_notes,
                include_template_images=include_template_images, render_charts=render_charts,
                max_chart_rows=max_chart_rows, use_libreoffice=use_libreoffice)
    prs, filename, has_macros = _open(source, filename)
    ctx = _Ctx(opts, prs.slide_width or 9144000, prs.slide_height or 6858000)
    slides, hidden = [], 0
    try:
        for n, slide in enumerate(prs.slides, 1):
            is_hidden = slide._element.get("show") in ("0", "false")
            hidden += is_hidden
            if is_hidden and not include_hidden_slides:
                continue
            try:
                slides.append(_read_slide(ctx, slide, n))
            except Exception:  # noqa: BLE001 - one broken slide must not lose the rest of the deck
                bad = _Slide(n, hidden=is_hidden)
                bad.items.append(_Item("text", None, 0, head="text", lines=["(this slide could not be read)"]))
                slides.append(bad)
                ctx.warn(f"Slide {n} could not be read.")
                for key in [k for k, img in ctx.images.items() if any(u.slide == n for u in img.uses)]:
                    ctx.images[key].uses = [u for u in ctx.images[key].uses if u.slide != n]
                    if not ctx.images[key].uses:
                        del ctx.images[key]
        total = len(prs.slides)
    except ConversionError:
        raise
    except Exception as e:  # noqa: BLE001
        raise ConversionError(f"Could not read {filename} as a PowerPoint file.") from e

    order = 0
    for s in slides:
        s.items = _reading_order(ctx, s.items)
        for it in s.items:
            if it.use is not None:
                order += 1
                it.use.order = order
    for img in ctx.images.values():
        img.uses.sort(key=lambda u: u.order)
    images = _finalize_images(ctx)

    w_in, h_in = ctx.w / 914400, ctx.h / 914400
    ratio = {1.78: "16:9", 1.33: "4:3", 1.6: "16:10"}.get(round(w_in / h_in, 2), "")
    meta_line = (f"PowerPoint deck: {total} slide{'s' if total != 1 else ''}"
                 + (f" ({hidden} hidden{'' if include_hidden_slides else ', left out'})" if hidden else "")
                 + f", {w_in:.2f} x {h_in:.2f} in" + (f" ({ratio})" if ratio else "") + ". "
                 "Each slide's shapes are listed top to bottom, then left to right.")
    if positions != "none":
        meta_line += (" Positions are percentages of the slide: x,y = the shape's top-left corner measured "
                      "from the slide's top-left, w,h = its size.")
    if images:
        meta_line += (" A line like [image img_1 | ...] marks a picture and the picture itself follows that line. "
                      "A picture used more than once is attached once; later uses read [img_1 again | ...].")
    infos = []
    for s in slides:
        ids = []
        for it in s.items:
            if it.use is not None and it.use.img is not None and it.use.img.out is not None \
                    and it.use.img.out.id not in ids:
                ids.append(it.use.img.out.id)
        infos.append(SlideInfo(s.number, s.title, s.hidden, _slide_text(ctx, s), _slide_outline(s), ids))
    text = "\n\n".join([meta_line] + [i.text for i in infos])

    statuses = [i.status for i in ctx.images.values()]
    if include_images:
        if "vector" in statuses:
            ctx.warn(f"{statuses.count('vector')} vector picture(s) (EMF/WMF/SVG) were not attached because "
                     "LibreOffice is not available to convert them.")
        if "budget" in statuses:
            ctx.warn(f"{statuses.count('budget')} picture(s) were left out to stay within the image budget "
                     "(backgrounds first, then the ones that are smallest on their slides).")
        if images and ctx.image_tier != TOKEN_TIERS[0]:
            ctx.warn(f"Pictures were shrunk to about {ctx.image_tier} tokens each to fit the image budget; "
                     "small print inside them may be hard to read.")
        if "unreadable" in statuses:
            ctx.warn(f"{statuses.count('unreadable')} picture(s) could not be decoded.")
    elif statuses:
        ctx.warn(f"{len(statuses)} picture(s) were not attached (text only).")
    if has_macros:
        ctx.warn("The file contains VBA macros; they were not run.")
    if not include_template_images:
        try:
            layouts = {id(s.slide_layout): s.slide_layout for s in prs.slides}.values()
            if any(_xp(o._element, "p:cSld/p:spTree//p:pic[not(p:nvPicPr/p:nvPr/p:ph)]")
                   for lay in layouts for o in (lay, lay.slide_master)):
                ctx.warn("Pictures that belong to the slide template (e.g. a logo on the master) were not included.")
        except Exception:  # noqa: BLE001
            pass

    meta = {"slides": total, "hidden_slides": hidden, "slide_width_in": round(w_in, 3),
            "slide_height_in": round(h_in, 3), **ctx.counts,
            "pictures_in_file": sum(len(i.uses) for i in ctx.images.values()),
            "images_attached": len(images), "est_image_tokens": sum(i.est_tokens for i in images),
            "est_text_tokens": _text_tokens(text), "image_bytes": sum(len(i.data) for i in images)}
    return PptxResult(filename, text, warnings=ctx.warnings, images=images, tables=ctx.tables, meta=meta,
                      slides=infos, header=meta_line)


_TAG = re.compile(r"</?\s*(document|content|conversion_notes)\b", re.I)
_MARKER = re.compile(r"^\[image (img_\d+)(?: \|[^\n]*)?\]$", re.M)


def to_prompt_text(result: PptxResult) -> str:
    """The deck wrapped in <document> tags. Tag look-alikes inside the file are escaped so slide
    text can't close the wrapper and pose as instructions outside it."""
    def esc(s: str) -> str:
        return _TAG.sub(lambda m: html.escape(m.group(0)), s)

    parts = [f'<document name="{html.escape(result.filename, quote=True)}" kind="{result.kind}">']
    if result.warnings:
        parts.append("<conversion_notes>")
        parts += [f"- {esc(w)}" for w in result.warnings]
        parts.append("</conversion_notes>")
    parts += ["<content>", esc(result.text), "</content>", "</document>"]
    return "\n".join(parts)


def blocks_from_text(text: str, images: list[SlideImage], *, cache: bool = True) -> list[dict]:
    """Split `text` at its [image img_N | ...] lines and put each picture right after its line.

    Works on any text that still contains the marker lines, so it can be applied after your own
    wrapping or truncation: a picture whose line is no longer in the text is simply not sent.
    """
    by_id = {i.id: i for i in images}
    blocks, pos, sent = [], 0, set()
    for m in _MARKER.finditer(text):
        img = by_id.get(m.group(1))
        if img is None or img.id in sent:
            continue
        sent.add(img.id)
        blocks.append({"type": "text", "text": text[pos:m.end()]})
        blocks.append(img.block())
        pos = m.end()
    tail = text[pos:].lstrip("\n") if blocks else text[pos:]
    if tail.strip() or not blocks:
        blocks.append({"type": "text", "text": tail})
    if cache:
        # One cache breakpoint at the end: follow-up questions re-read the deck at the cached price.
        blocks[-1] = {**blocks[-1], "cache_control": {"type": "ephemeral"}}
    return blocks


def to_content_blocks(result: PptxResult, *, cache: bool = True) -> list[dict]:
    """Content blocks for one user message: the wrapped text with every picture placed after its line.
    Append your own {"type": "text", "text": question} block after these."""
    return blocks_from_text(to_prompt_text(result), result.images, cache=cache)


def save_images(result: PptxResult, out_dir: str | Path) -> Path:
    """Write every attached picture plus images.json (id, file, size, and each slide position)."""
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    manifest = result.image_manifest()
    for img, entry in zip(result.images, manifest):
        (out / entry["file"]).write_bytes(img.data)
    path = out / "images.json"
    path.write_text(json.dumps({"filename": result.filename, "slide_width_in": result.meta.get("slide_width_in"),
                                "slide_height_in": result.meta.get("slide_height_in"), "images": manifest}, indent=2),
                    encoding="utf-8")
    return path


def _main() -> int:
    import argparse
    import sys

    ap = argparse.ArgumentParser(prog="pptx_reader", description="PowerPoint -> text + images for the Claude Messages API")
    ap.add_argument("file")
    ap.add_argument("--out", help="folder for text.txt, content_blocks.json, images/ and images.json")
    ap.add_argument("--no-images", action="store_true")
    ap.add_argument("--render-charts", action="store_true")
    ap.add_argument("--template-images", action="store_true")
    ap.add_argument("--budget", type=int, default=DEFAULT_IMAGE_TOKEN_BUDGET, help="image token budget")
    ap.add_argument("--positions", choices=["media", "all", "none"], default="media")
    args = ap.parse_args()
    try:
        r = read_pptx(args.file, include_images=not args.no_images, render_charts=args.render_charts,
                      include_template_images=args.template_images, image_token_budget=args.budget,
                      positions=args.positions)
    except ConversionError as e:
        print(f"!! {e}", file=sys.stderr)
        return 1
    print(f"===== {r.filename}: {r.meta['slides']} slides, ~{r.meta['est_text_tokens']:,} text tokens, "
          f"{len(r.images)} image(s) ~{r.meta['est_image_tokens']:,} tokens", file=sys.stderr)
    if args.out:
        out = Path(args.out)
        out.mkdir(parents=True, exist_ok=True)
        (out / "text.txt").write_text(to_prompt_text(r), encoding="utf-8")
        (out / "content_blocks.json").write_text(json.dumps(to_content_blocks(r)), encoding="utf-8")
        save_images(r, out / "images")
        print(f"written to {out}", file=sys.stderr)
    else:
        print(to_prompt_text(r))
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
