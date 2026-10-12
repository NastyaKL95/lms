#!/usr/bin/env python3
"""Convert prepared markdown presentations into styled .pptx files."""

from __future__ import annotations

import re
from pathlib import Path
from typing import List, Tuple

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


Heading = str
Bullet = Tuple[int, str]
Section = Tuple[Heading, List[Bullet]]


PRIMARY = RGBColor(15, 23, 42)
PRIMARY_SOFT = RGBColor(30, 64, 175)
ACCENT = RGBColor(14, 165, 233)
TEXT_MAIN = RGBColor(17, 24, 39)
TEXT_MUTED = RGBColor(75, 85, 99)
SURFACE = RGBColor(248, 250, 252)
BORDER = RGBColor(203, 213, 225)
WHITE = RGBColor(255, 255, 255)

FONT_FAMILY = "Calibri"


def clean_md_text(text: str) -> str:
    text = text.strip()
    text = text.replace("**", "")
    text = text.replace("`", "")
    text = re.sub(r"\[(.*?)\]\((.*?)\)", r"\1", text)
    return text.strip()


def parse_markdown(path: Path) -> Tuple[str, List[Section]]:
    lines = path.read_text(encoding="utf-8").splitlines()

    title = path.stem
    sections: List[Section] = []
    current_heading: str | None = None
    current_items: List[Bullet] = []

    def flush_current() -> None:
        nonlocal current_heading, current_items
        if current_heading is not None:
            sections.append((current_heading, current_items))
        current_heading = None
        current_items = []

    for raw_line in lines:
        line = raw_line.rstrip()
        stripped = line.strip()

        if not stripped or stripped == "---":
            continue

        if stripped.startswith("# "):
            title = clean_md_text(stripped[2:])
            continue

        if stripped.startswith("## ") or stripped.startswith("### "):
            flush_current()
            current_heading = clean_md_text(re.sub(r"^#{2,3}\s+", "", stripped))
            continue

        if current_heading is None:
            continue

        bullet_match = re.match(r"^(\s*)([-*])\s+(.*)$", line)
        number_match = re.match(r"^(\s*)\d+\.\s+(.*)$", line)

        if bullet_match:
            spaces, _, value = bullet_match.groups()
            level = max(0, len(spaces) // 2)
            current_items.append((min(level, 2), clean_md_text(value)))
        elif number_match:
            spaces, value = number_match.groups()
            level = max(0, len(spaces) // 2)
            current_items.append((min(level, 2), clean_md_text(value)))
        else:
            current_items.append((0, clean_md_text(stripped)))

    flush_current()
    return title, sections


def set_background(slide, color: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_footer(slide, deck_title: str, slide_number: int) -> None:
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.4),
        Inches(6.95),
        Inches(12.5),
        Inches(0.01),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = BORDER
    line.line.fill.background()

    footer_box = slide.shapes.add_textbox(Inches(0.45), Inches(7.0), Inches(10.6), Inches(0.3))
    footer_tf = footer_box.text_frame
    footer_tf.text = deck_title
    footer_paragraph = footer_tf.paragraphs[0]
    footer_paragraph.font.name = FONT_FAMILY
    footer_paragraph.font.size = Pt(11)
    footer_paragraph.font.color.rgb = TEXT_MUTED

    number_box = slide.shapes.add_textbox(Inches(11.2), Inches(6.98), Inches(1.6), Inches(0.35))
    number_tf = number_box.text_frame
    number_tf.text = f"{slide_number}"
    number_paragraph = number_tf.paragraphs[0]
    number_paragraph.font.name = FONT_FAMILY
    number_paragraph.font.bold = True
    number_paragraph.font.size = Pt(12)
    number_paragraph.font.color.rgb = PRIMARY_SOFT
    number_paragraph.alignment = PP_ALIGN.RIGHT


def add_title_slide(prs: Presentation, title: str, subtitle: str) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, PRIMARY)

    header_band = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.33), Inches(1.0)
    )
    header_band.fill.solid()
    header_band.fill.fore_color.rgb = PRIMARY_SOFT
    header_band.line.fill.background()

    for x in (10.8, 11.5, 12.2):
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(0.26), Inches(0.2), Inches(0.2))
        dot.fill.solid()
        dot.fill.fore_color.rgb = ACCENT
        dot.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.9), Inches(1.5), Inches(11.8), Inches(2.6))
    title_tf = title_box.text_frame
    title_tf.word_wrap = True
    title_tf.text = title
    title_paragraph = title_tf.paragraphs[0]
    title_paragraph.font.name = FONT_FAMILY
    title_paragraph.font.bold = True
    title_paragraph.font.size = Pt(44)
    title_paragraph.font.color.rgb = WHITE

    subtitle_box = slide.shapes.add_textbox(Inches(0.95), Inches(4.5), Inches(10.0), Inches(1.1))
    subtitle_tf = subtitle_box.text_frame
    subtitle_tf.text = subtitle
    subtitle_paragraph = subtitle_tf.paragraphs[0]
    subtitle_paragraph.font.name = FONT_FAMILY
    subtitle_paragraph.font.size = Pt(20)
    subtitle_paragraph.font.color.rgb = RGBColor(191, 219, 254)

    accent_bar = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(5.8), Inches(4.2), Inches(0.35)
    )
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = ACCENT
    accent_bar.line.fill.background()

    add_footer(slide, "Детализированная учебная презентация", 1)


def chunk_items(items: List[Bullet], chunk_size: int = 6) -> List[List[Bullet]]:
    if not items:
        return [[]]
    return [items[i : i + chunk_size] for i in range(0, len(items), chunk_size)]


def add_header_badge(slide, heading: str) -> None:
    if not heading.lower().startswith("блок"):
        return

    duration_match = re.search(r"\((\d+)\s*мин", heading.lower())
    if not duration_match:
        return

    badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.45), Inches(0.62), Inches(2.2), Inches(0.5)
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor(219, 234, 254)
    badge.line.color.rgb = RGBColor(147, 197, 253)
    badge.line.width = Pt(1)
    badge_tf = badge.text_frame
    badge_tf.text = f"{duration_match.group(1)} минут"
    badge_p = badge_tf.paragraphs[0]
    badge_p.font.name = FONT_FAMILY
    badge_p.font.bold = True
    badge_p.font.size = Pt(12)
    badge_p.font.color.rgb = PRIMARY_SOFT
    badge_p.alignment = PP_ALIGN.CENTER


def add_section_slides(prs: Presentation, deck_title: str, heading: str, items: List[Bullet]) -> None:
    for index, chunk in enumerate(chunk_items(items)):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        set_background(slide, SURFACE)

        top_band = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.33), Inches(1.3)
        )
        top_band.fill.solid()
        top_band.fill.fore_color.rgb = WHITE
        top_band.line.fill.background()

        accent_line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.22), Inches(13.33), Inches(0.08)
        )
        accent_line.fill.solid()
        accent_line.fill.fore_color.rgb = ACCENT
        accent_line.line.fill.background()

        title_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.33), Inches(9.5), Inches(0.85))
        title_tf = title_box.text_frame
        title_tf.word_wrap = True
        visible_heading = heading if index == 0 else f"{heading} (продолжение)"
        title_tf.text = visible_heading
        title_paragraph = title_tf.paragraphs[0]
        title_paragraph.font.name = FONT_FAMILY
        title_paragraph.font.bold = True
        title_paragraph.font.size = Pt(28)
        title_paragraph.font.color.rgb = PRIMARY

        add_header_badge(slide, heading)

        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(1.6), Inches(11.95), Inches(5.0)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = BORDER
        card.line.width = Pt(1)

        sidebar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(1.6), Inches(0.18), Inches(5.0)
        )
        sidebar.fill.solid()
        sidebar.fill.fore_color.rgb = PRIMARY_SOFT
        sidebar.line.fill.background()

        content_box = slide.shapes.add_textbox(Inches(1.05), Inches(1.92), Inches(11.2), Inches(4.55))
        content_tf = content_box.text_frame
        content_tf.word_wrap = True
        content_tf.margin_left = Pt(6)
        content_tf.margin_right = Pt(6)
        content_tf.margin_top = Pt(2)
        content_tf.margin_bottom = Pt(2)
        content_tf.clear()

        if not chunk:
            content_tf.text = "Содержимое раздела"
            paragraph = content_tf.paragraphs[0]
            paragraph.font.name = FONT_FAMILY
            paragraph.font.size = Pt(24)
            paragraph.font.color.rgb = TEXT_MAIN
        else:
            for item_index, (level, text) in enumerate(chunk):
                paragraph = content_tf.paragraphs[0] if item_index == 0 else content_tf.add_paragraph()
                paragraph.text = text
                paragraph.level = level
                paragraph.space_after = Pt(10 if level == 0 else 6)
                paragraph.font.name = FONT_FAMILY
                paragraph.font.size = Pt(22 - (level * 2))
                paragraph.font.color.rgb = TEXT_MAIN if level == 0 else TEXT_MUTED

        add_footer(slide, deck_title, len(prs.slides))


def convert(markdown_path: Path, output_path: Path) -> None:
    title, sections = parse_markdown(markdown_path)
    prs = Presentation()
    add_title_slide(prs, title, "4 блока по 40 минут")

    for heading, items in sections:
        add_section_slides(prs, title, heading, items)

    prs.save(output_path)


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    sources = [
        (
            root / "presentation-topic-1-detailed.md",
            root / "presentation-topic-1-detailed.pptx",
        ),
        (
            root / "presentation-topic-2-detailed.md",
            root / "presentation-topic-2-detailed.pptx",
        ),
    ]

    for source, output in sources:
        convert(source, output)
        print(f"created: {output}")


if __name__ == "__main__":
    main()
