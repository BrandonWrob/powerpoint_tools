"""deck_store - send a deck's outline first; let the model pull slides and pictures on demand.

Same idea as the spreadsheet query tool: the prompt carries a short description of the file, and the
model calls tools for the parts it needs. For a deck that means a one-line-per-slide outline up front
(about 45 tokens a slide) instead of every slide's text and every picture.

    from deck_store import DeckStore, DECK_TOOLS

    store = DeckStore.from_file("upload.pptx", "upload.deck.sqlite")   # once per upload; keep next to the upload
    if store.inline_tokens <= 20_000:
        blocks = store.inline_blocks()        # small deck: just send all of it, no tools needed
    else:
        blocks = store.outline_blocks()       # big deck: outline + tools=DECK_TOOLS
    ...
    # for each tool_use block the model returns:
    result = store.tool_result(block.id, block.name, block.input)      # a tool_result block (text and/or pictures)

Later requests only need DeckStore("upload.deck.sqlite"); the .pptx is not read again.
The store is one SQLite file per upload, opened read-only, never queried with model-written SQL.
"""
from __future__ import annotations

import base64
import json
import re
import sqlite3
from pathlib import Path

from pptx_reader import PptxResult, _TAG, _text_tokens, blocks_from_text, read_pptx

__all__ = ["DeckStore", "DECK_TOOLS", "handle_deck_tool"]

MAX_READ_TOKENS = 6_000          # text returned by one read_slides call
MAX_PICTURES_PER_CALL = 6        # at most ~9,400 picture tokens per call
MAX_SEARCH_HITS = 15

# Generic on purpose: slide titles and file names can contain business data, so they stay in the
# message content (the outline), never in the tool schema.
_DOC = {"type": "string", "description": "File name of the deck, exactly as in its <document name=...> tag. "
                                          "Only needed when more than one deck is in the conversation."}
DECK_TOOLS = [
    {
        "name": "read_slides",
        "description": (
            "Get the full content of slides from a PowerPoint deck the user uploaded: all text, tables, chart "
            "data, SmartArt, connectors, comments and speaker notes, with the position of every picture, chart "
            "and table. The deck's <document> block only holds a one-line outline per slide, so call this "
            "before answering anything that depends on what a slide actually says. Ask for all the slides you "
            "need in one call. Set with_pictures to true only when you also need to see the pictures on those "
            "slides; pictures cost far more than text."),
        "input_schema": {
            "type": "object",
            "properties": {
                "slides": {"type": "array", "items": {"type": "integer"}, "description": "Slide numbers, e.g. [3, 4, 5]."},
                "with_pictures": {"type": "boolean", "description": "Also return the pictures on these slides. Default false."},
                "document": _DOC,
            },
            "required": ["slides"],
        },
    },
    {
        "name": "view_pictures",
        "description": (
            "Look at pictures from an uploaded deck by id (e.g. img_3; ids are in the outline and in read_slides "
            f"output). Use it when the answer may depend on what a picture shows: a screenshot, photo, diagram "
            f"or scanned table. A picture with no description in the outline has unknown content until you view "
            f"it. At most {MAX_PICTURES_PER_CALL} per call."),
        "input_schema": {
            "type": "object",
            "properties": {
                "pictures": {"type": "array", "items": {"type": "string"}, "description": 'Picture ids, e.g. ["img_3", "img_4"].'},
                "document": _DOC,
            },
            "required": ["pictures"],
        },
    },
    {
        "name": "search_slides",
        "description": (
            "Find which slides of an uploaded deck mention some words (case-insensitive, matched from the start of "
            "a word, numbers matched whole; searches slide text, "
            "tables, chart data, notes, comments and picture descriptions, not text inside pictures). Returns "
            "slide numbers with the matching lines. Use it to locate a topic in a long deck, then read_slides."),
        "input_schema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "A few words that should appear on the slide."},
                "document": _DOC,
            },
            "required": ["query"],
        },
    },
]
_TOOL_NAMES = {t["name"] for t in DECK_TOOLS}


def _esc(text: str) -> str:
    """Slide text must not be able to imitate the <document> wrapper."""
    return _TAG.sub(lambda m: m.group(0).replace("<", "&lt;"), text)


class DeckStore:
    """One deck, stored once, read many times. Open with DeckStore(path); create with build()/from_file()."""

    # ------------------------------------------------------------ create
    @classmethod
    def from_file(cls, source, store_path: str | Path, *, filename: str | None = None,
                  captions: dict[str, str] | None = None, **read_options) -> "DeckStore":
        """Read a .pptx and store it. Pictures are kept at full size (up to 1568 tokens each) because
        they are only ever sent a few at a time."""
        read_options.setdefault("image_token_budget", None)
        read_options.setdefault("max_images", None)
        return cls.build(read_pptx(source, filename=filename, **read_options), store_path, captions=captions)

    @classmethod
    def build(cls, result: PptxResult, store_path: str | Path, *, captions: dict[str, str] | None = None) -> "DeckStore":
        """Write a read_pptx() result to a new SQLite file.

        captions: optional {picture id: short description}. Most real decks have no alt text, so the
        model can't tell what a picture shows without viewing it. If you describe each picture once
        at upload time (one cheap model call per deck), pass the descriptions here: they appear in the
        outline and become searchable.
        """
        path = Path(store_path)
        if path.exists():
            path.unlink()
        captions = {k: re.sub(r"\s+", " ", v).strip()[:300].translate({ord("["): "(", ord("]"): ")"})
                    for k, v in (captions or {}).items() if v and v.strip()}
        conn = sqlite3.connect(path)
        try:
            conn.executescript(
                "CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT);"
                "CREATE TABLE slides (number INTEGER PRIMARY KEY, title TEXT, hidden INTEGER, text TEXT, outline TEXT, search TEXT);"
                "CREATE TABLE images (id TEXT PRIMARY KEY, media_type TEXT, width INTEGER, height INTEGER, est_tokens INTEGER,"
                " slide INTEGER, label TEXT, caption TEXT, data BLOB);"
                "CREATE TABLE slide_images (slide INTEGER, image_id TEXT, PRIMARY KEY (slide, image_id));")
            for img in result.images:
                first = img.locations[0] if img.locations else {}
                label = " | ".join(p for p in (f"slide {first.get('slide', '?')}", first.get("position") or "", first.get("note") or "",
                                               f"alt: {first['alt_text']}" if first.get("alt_text") else "") if p)
                conn.execute("INSERT INTO images VALUES (?,?,?,?,?,?,?,?,?)",
                             (img.id, img.media_type, img.width, img.height, img.est_tokens, first.get("slide"), label,
                              captions.get(img.id, ""), img.data))
            for s in result.slides:
                outline, extra = s.outline, []
                for image_id in s.image_ids:
                    conn.execute("INSERT OR IGNORE INTO slide_images VALUES (?,?)", (s.number, image_id))
                    cap = captions.get(image_id)
                    if cap:
                        extra.append(f"{image_id}: {cap}")
                        # in the outline, a described picture reads: img_3 "alt" -> img_3 "description"
                        outline = re.sub(rf'\b{image_id}\b(?: "[^"]*")?', lambda m: f'{image_id} "{cap[:80]}"', outline, count=1)
                conn.execute("INSERT INTO slides VALUES (?,?,?,?,?,?)",
                             (s.number, s.title, int(s.hidden), s.text, outline, (s.text + "\n" + "\n".join(extra)).lower()))
            meta = {"filename": result.filename, "header": result.header, "warnings": json.dumps(result.warnings),
                    "slide_count": str(result.meta.get("slides", len(result.slides))),
                    "text_tokens": str(_text_tokens(result.text)),
                    "image_tokens": str(sum(i.est_tokens for i in result.images)), "format": "1"}
            conn.executemany("INSERT INTO meta VALUES (?,?)", meta.items())
            conn.commit()
        finally:
            conn.close()
        return cls(path)

    # ------------------------------------------------------------ open
    def __init__(self, store_path: str | Path):
        self.path = Path(store_path)
        self.conn = sqlite3.connect(f"file:{self.path.resolve()}?mode=ro", uri=True, check_same_thread=False)
        self.conn.execute("PRAGMA query_only = ON")
        meta = dict(self.conn.execute("SELECT key, value FROM meta"))
        self.filename: str = meta["filename"]
        self.header: str = meta["header"]
        self.warnings: list[str] = json.loads(meta["warnings"])
        self.slide_count = int(meta["slide_count"])
        self.inline_tokens = int(meta["text_tokens"]) + int(meta["image_tokens"])   # cost of sending everything
        self.picture_count = self.conn.execute("SELECT COUNT(*) FROM images").fetchone()[0]

    def close(self):
        self.conn.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    # ------------------------------------------------------------ what goes in the user message
    def _wrap(self, strategy: str, body: str) -> str:
        name = self.filename.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")
        parts = [f'<document name="{name}" kind="slides" strategy="{strategy}">']
        if self.warnings:
            parts += ["<conversion_notes>"] + [f"- {_esc(w)}" for w in self.warnings] + ["</conversion_notes>"]
        return "\n".join(parts + ["<content>", _esc(body), "</content>", "</document>"])

    def outline_text(self) -> str:
        """The deck as one line per slide, with how to get the rest."""
        first = self.header.split(". ")[0] + "."
        how = (f"{first} Only an outline is included here: one line per slide with its title, the start of "
               "its text, and what else is on it. Call read_slides for a slide's full content (text, tables, "
               "chart data, notes, and where each picture sits), search_slides to find where something is "
               "mentioned, and view_pictures to look at a picture (ids like img_3). A picture's quoted "
               "description is its alt text or caption; a picture without one has unknown content until viewed.")
        lines = [row[0] for row in self.conn.execute("SELECT outline FROM slides ORDER BY number")]
        return self._wrap("outline", how + "\n\n" + "\n".join(lines))

    def outline_blocks(self, *, cache: bool = True) -> list[dict]:
        """Content for the user message in on-demand mode. Pass tools=DECK_TOOLS with the request."""
        block = {"type": "text", "text": self.outline_text()}
        if cache:
            block["cache_control"] = {"type": "ephemeral"}
        return [block]

    def inline_blocks(self, *, cache: bool = True) -> list[dict]:
        """Everything at once (all text, every picture at stored size). For small decks only: check
        inline_tokens first. Same layout as pptx_reader.to_content_blocks()."""
        texts = [row[0] for row in self.conn.execute("SELECT text FROM slides ORDER BY number")]
        return blocks_from_text(self._wrap("inline", "\n\n".join([self.header] + texts)), self._images(None), cache=cache)

    # ------------------------------------------------------------ tools
    def _images(self, ids: list[str] | None) -> list:
        from pptx_reader import SlideImage
        sql = "SELECT id, media_type, data, width, height, est_tokens FROM images"
        if ids is None:
            rows = self.conn.execute(sql).fetchall()
        else:
            marks = ",".join("?" for _ in ids)
            rows = self.conn.execute(f"{sql} WHERE id IN ({marks})", ids).fetchall()
        return [SlideImage(r[0], r[1], bytes(r[2]), r[3], r[4], r[5], "", "", (r[3], r[4]), "") for r in rows]

    def read_slides(self, slides, with_pictures: bool = False):
        """Full text of the given slides. Returns a string, or a list of text/image blocks with pictures."""
        if not isinstance(slides, (list, tuple)) or not slides:
            return f"Error: 'slides' must be a list of slide numbers between 1 and {self.slide_count}, e.g. [3, 4]."
        wanted, bad = [], []
        for n in slides:
            if isinstance(n, bool) or not isinstance(n, (int, float)) or int(n) != n or not 1 <= int(n) <= self.slide_count:
                bad.append(n)
            elif int(n) not in wanted:
                wanted.append(int(n))
        if not wanted:
            return f"Error: no such slide {bad[:5]}. The deck has slides 1 to {self.slide_count}."
        wanted.sort()
        marks = ",".join("?" for _ in wanted)
        rows = dict(self.conn.execute(f"SELECT number, text FROM slides WHERE number IN ({marks})", wanted))
        shown, left_out, used = [], [], 0
        for n in wanted:
            if n not in rows:
                bad.append(n)                 # e.g. a hidden slide that was left out at upload
                continue
            cost = _text_tokens(rows[n])
            if shown and used + cost > MAX_READ_TOKENS:
                left_out.append(n)
                continue
            used += cost
            shown.append(n)
        text = _esc("\n\n".join(rows[n] for n in shown))
        notes = []
        if left_out:
            notes.append(f"Not shown to keep this reply short; ask again for slides {left_out}.")
        if bad:
            notes.append(f"Not in the deck: {bad[:10]} (it has slides 1 to {self.slide_count}).")
        if not with_pictures:
            return text + ("\n\n" + " ".join(notes) if notes else "")
        marks = ",".join("?" for _ in shown)
        ids = [r[0] for r in self.conn.execute(
            f"SELECT image_id FROM slide_images JOIN images ON images.id = image_id WHERE slide_images.slide IN ({marks}) "
            "AND images.slide = slide_images.slide ORDER BY slide_images.slide, CAST(SUBSTR(image_id, 5) AS INTEGER)", shown)]
        if len(ids) > MAX_PICTURES_PER_CALL:
            notes.append(f"Only the first {MAX_PICTURES_PER_CALL} pictures are attached; use view_pictures for "
                         f"{', '.join(ids[MAX_PICTURES_PER_CALL:])}.")
            ids = ids[:MAX_PICTURES_PER_CALL]
        blocks = blocks_from_text(text, self._images(ids), cache=False)
        if notes:
            blocks.append({"type": "text", "text": " ".join(notes)})
        return blocks

    def view_pictures(self, pictures):
        """Pictures by id, each after a line saying where it sits. Returns a list of blocks (or an error string)."""
        if isinstance(pictures, str):
            pictures = [pictures]
        if not isinstance(pictures, (list, tuple)) or not pictures:
            return 'Error: \'pictures\' must be a list of picture ids such as ["img_3"].'
        ids = []
        for p in pictures:
            p = str(p).strip().lower()
            if re.fullmatch(r"img_\d+", p) and p not in ids:
                ids.append(p)
        marks = ",".join("?" for _ in ids)
        rows = {r[0]: r for r in self.conn.execute(
            f"SELECT id, media_type, label, caption, data FROM images WHERE id IN ({marks})", ids)} if ids else {}
        unknown = [str(p)[:20] for p in pictures if str(p).strip().lower() not in rows]
        found = [i for i in ids if i in rows]
        if not found:
            return (f"Error: no picture with id {unknown[:5]}. Picture ids look like img_3 and are listed in the "
                    f"outline; this deck has img_1 to img_{self.picture_count}." if self.picture_count
                    else "Error: this deck has no pictures that can be viewed.")
        blocks = []
        for i in found[:MAX_PICTURES_PER_CALL]:
            _, media, label, caption, data = rows[i]
            blocks.append({"type": "text", "text": f"{i} | {_esc(label)}" + (f" | description: {_esc(caption)}" if caption else "")})
            blocks.append({"type": "image", "source": {"type": "base64", "media_type": media,
                                                       "data": base64.standard_b64encode(bytes(data)).decode("ascii")}})
        notes = []
        if len(found) > MAX_PICTURES_PER_CALL:
            notes.append(f"Only {MAX_PICTURES_PER_CALL} pictures per call; ask again for {', '.join(found[MAX_PICTURES_PER_CALL:])}.")
        if unknown:
            notes.append(f"No such picture: {', '.join(unknown[:10])}.")
        if notes:
            blocks.append({"type": "text", "text": " ".join(notes)})
        return blocks

    def search(self, query) -> str:
        """Slides that mention the words in `query`, with the matching lines."""
        terms = [t for t in re.findall(r"[^\s,;:()\"']+", str(query or "").lower()) if len(t) > 1][:8]
        if not terms:
            return "Error: give one or more words to look for."
        # A word must start where a word starts ("order" finds "orders", not "backorder"); a number must
        # match whole ("17" does not find "1703").
        pats = [re.compile(r"(?<![a-z0-9])" + re.escape(t) + (r"(?![a-z0-9])" if t[-1].isdigit() else "")) for t in terms]
        scored = []
        for number, title, text, search in self.conn.execute("SELECT number, title, text, search FROM slides ORDER BY number"):
            hits = sum(1 for p in pats if p.search(search))
            if hits:
                in_title = sum(1 for p in pats if p.search((title or "").lower()))
                scored.append((hits, in_title, number, title, text, search))
        if not scored:
            return (f"No slide mentions {' / '.join(terms)}. Try other words or fewer of them. Text inside pictures "
                    "is not searchable; view the picture instead.")
        best = max(s[0] for s in scored)
        scored = sorted((s for s in scored if s[0] == best), key=lambda s: (-s[1], s[2]))     # title matches first
        out = [f"{len(scored)} slide{'s' if len(scored) != 1 else ''} mention "
               + (" and ".join(terms) if best == len(terms) else f"{best} of the {len(terms)} words (none has all)") + ":"]
        for _, _, number, title, text, search in scored[:MAX_SEARCH_HITS]:
            lines = []
            for line in (text + "\n" + search[len(text):]).split("\n")[1:]:      # [1:] skips the heading line
                if any(p.search(line.lower()) for p in pats) and len(lines) < 3:
                    line = line.strip()
                    lines.append(line if len(line) <= 160 else line[:159] + "…")
            out.append(f"Slide {number}: {title or '(no title)'}" + "".join(f"\n  {line}" for line in lines))
        if len(scored) > MAX_SEARCH_HITS:
            out.append(f"... and {len(scored) - MAX_SEARCH_HITS} more: slides {[s[2] for s in scored[MAX_SEARCH_HITS:]]}")
        return _esc("\n".join(out))

    def handle(self, name: str, tool_input: dict):
        """Run one tool call. Returns tool_result content: a string or a list of text/image blocks.
        Never raises for bad input; the message tells the model what to do instead."""
        tool_input = tool_input if isinstance(tool_input, dict) else {}
        try:
            if name == "read_slides":
                return self.read_slides(tool_input.get("slides"), bool(tool_input.get("with_pictures", False)))
            if name == "view_pictures":
                return self.view_pictures(tool_input.get("pictures"))
            if name == "search_slides":
                return self.search(tool_input.get("query"))
        except sqlite3.Error:
            return "Error: the stored deck could not be read. Tell the user the file needs to be uploaded again."
        return f"Error: unknown tool {str(name)[:40]!r}. Available: read_slides, view_pictures, search_slides."

    def tool_result(self, tool_use_id: str, name: str, tool_input: dict) -> dict:
        """A complete tool_result block for the next user message."""
        return _result_block(tool_use_id, self.handle(name, tool_input))


def _result_block(tool_use_id: str, content) -> dict:
    block = {"type": "tool_result", "tool_use_id": tool_use_id, "content": content}
    if isinstance(content, str) and content.startswith("Error:"):
        block["is_error"] = True
    return block


def handle_deck_tool(stores, tool_use_id: str, name: str, tool_input: dict) -> dict:
    """tool_result block when several decks are in one conversation. `stores` is a list of DeckStore.
    With one deck the 'document' input is optional; with several the model must name the file."""
    stores = [stores] if isinstance(stores, DeckStore) else list(stores)
    tool_input = tool_input if isinstance(tool_input, dict) else {}
    if name not in _TOOL_NAMES:
        return _result_block(tool_use_id, f"Error: unknown tool {str(name)[:40]!r}.")
    wanted = str(tool_input.get("document") or "").strip()
    match = [s for s in stores if s.filename == wanted] or [s for s in stores if s.filename.lower() == wanted.lower()]
    if not match and len(stores) == 1 and not wanted:
        match = stores
    if len(match) != 1:
        names = ", ".join(f'"{s.filename}"' for s in stores)
        return _result_block(tool_use_id, f"Error: set 'document' to one of these file names: {names}.")
    return match[0].tool_result(tool_use_id, name, tool_input)
