# -*- coding: utf-8 -*-
"""Ribbon-like grouping of fields: feature key -> ordered [(section title, [field names])].

GUI-only concern. Any field missing from this map is shown in an automatic 'Other' section,
so adding a field to a feature never hides it (tests/test_wordtk.py checks the map is complete).
"""
SECTIONS = {
    "paragraph": [
        ("Alignment & Direction • تراز و جهت", ["alignment", "direction"]),
        ("Spacing • فاصله‌گذاری", ["line_spacing", "space_before", "space_after", "contextual_spacing"]),
        ("Indentation • تورفتگی", ["indent_start", "indent_end", "first_line_indent", "hanging_indent"]),
        ("Pagination • صفحه‌بندی", ["keep_with_next", "keep_lines", "page_break_before", "widow_control", "outline_level"]),
        ("Borders & Shading • کادر و سایه", ["shading", "border", "remove_border"]),
        ("Tabs • تب‌ها", ["tabs", "clear_tabs"]),
        ("Style & Apply • سبک و اعمال", ["style", "clear_direct", "apply_to", "style_targets"]),
        ("Change values • تغییر مقدار", ["changes"]),
    ],
    "font": [
        ("Font • قلم", ["font", "font_cs", "font_east_asia", "sync_scripts"]),
        ("Size • اندازه", ["size", "size_cs"]),
        ("Style • سبک", ["bold", "italic", "underline", "underline_color", "strike", "vertical_align"]),
        ("Color & Highlight • رنگ", ["color", "highlight", "shading"]),
        ("Effects • جلوه‌ها", ["small_caps", "all_caps", "hidden"]),
        ("Character spacing • فاصله‌ی حروف", ["char_spacing", "char_scale", "kerning", "position"]),
        ("Language & Direction • زبان و جهت", ["lang_latin", "lang_bidi", "lang_east_asia", "run_direction"]),
        ("Symbols • نمادها", ["symbols"]),
        ("Apply • اعمال", ["clear_direct", "apply_to", "style_targets"]),
        ("Change values • تغییر مقدار", ["changes"]),
    ],
    "lists": [
        ("Type • نوع", ["mode", "apply_to"]),
        ("Bullets • بولت", ["bullet_char", "bullet_font"]),
        ("Numbering • شماره‌گذاری", ["number_format", "number_text", "multilevel", "start_at", "restart"]),
        ("List indent • تورفتگی فهرست", ["indent_start", "indent_hanging", "indent_step"]),
        ("Existing list • فهرست موجود", ["use_list"]),
    ],
    "page": [
        ("Paper • کاغذ", ["paper", "width", "height", "orientation"]),
        ("Margins • حاشیه‌ها", ["margin_top", "margin_bottom", "margin_left", "margin_right",
                              "header_distance", "footer_distance"]),
        ("Printing • چاپ", ["gutter", "gutter_position", "mirror_margins", "sections"]),
        ("Columns • ستون‌ها", ["columns", "column_spacing", "column_separator"]),
        ("Layout • چیدمان", ["vertical_align", "section_direction"]),
        ("Numbering • شماره صفحه", ["page_number_format", "page_number_start", "line_numbers"]),
        ("Page borders • کادر صفحه", ["borders", "remove_borders"]),
    ],
    "breaks": [
        ("Remove breaks • حذف", ["remove_page_breaks", "remove_column_breaks", "remove_section_breaks"]),
        ("Section breaks • بخش‌ها", ["section_break_type"]),
        ("By style • بر اساس سبک", ["page_break_before_styles", "no_page_break_before_styles"]),
    ],
    "header_footer": [
        ("Header • سرصفحه", ["header", "first_header", "even_header", "clear_headers", "remove_headers"]),
        ("Footer • پاصفحه", ["footer", "first_footer", "even_footer", "clear_footers", "remove_footers"]),
        ("Options • گزینه‌ها", ["different_first_page", "different_odd_even"]),
    ],
    "background": [
        ("Page color • رنگ صفحه", ["page_color"]),
        ("Watermark • واترمارک", ["watermark_text", "watermark_color", "watermark_font", "watermark_size",
                                  "watermark_opacity", "watermark_angle", "remove_watermark"]),
    ],
    "illustrations": [
        ("Size • اندازه", ["max_width", "width", "height", "scale_percent", "fit_text_width"]),
        ("Alt text • متن جایگزین", ["alt_text", "overwrite_alt_text", "clear_alt_text"]),
        ("Layout • چیدمان", ["alignment"]),
        ("Remove • حذف", ["remove"]),
    ],
    "comments": [
        ("Comments • نظرات", ["remove_comments", "anonymize_comments"]),
        ("Revisions • ردگیری تغییرات", ["accept_changes", "reject_changes", "track_changes"]),
        ("Notes • پاورقی", ["remove_footnotes"]),
    ],
    "text": [
        ("Find & Replace • جایگزینی", ["replacements"]),
        ("Bulk replace list • فهرست جایگزینی", ["replace_list"]),
        ("Digits & Letters • ارقام و حروف", ["digits_convert", "digits", "persian_letters", "change_case"]),
        ("Spaces & Punctuation • فاصله و نشانه", ["collapse_spaces", "trim_spaces", "fix_punctuation_spacing"]),
        ("Empty lines & pages • خط و صفحه خالی", ["remove_empty_paragraphs", "remove_empty_line_breaks",
                                               "remove_empty_pages", "max_consecutive_empty"]),
        ("Cleanup • پاکسازی", ["remove_hyperlinks", "remove_hidden_text"]),
    ],
    "tables": [
        ("Select • انتخاب", ["tables"]),
        ("Borders • کادرها", ["border_target", "border_style", "border_width", "border_color", "erase_borders"]),
        ("Rows & Size • سطر و اندازه", ["row_height", "row_height_rule", "autofit", "table_width", "rows_cant_split",
                                      "header_row"]),
        ("Cells • سلول‌ها", ["cell_vertical_align", "cell_padding", "cell_spacing", "alignment", "direction"]),
        ("Style & Shading • سبک و سایه", ["table_style", "shading", "header_shading"]),
        ("Delete • حذف", ["delete_empty_rows", "delete_empty_tables", "delete_tables"]),
    ],
    "proofing": [
        ("Spelling & Grammar • املا و دستور", ["hide_spelling_errors", "hide_grammar_errors", "no_proof_text"]),
    ],
    "cover": [
        ("Cover content • محتوا", ["title", "subtitle", "author", "date"]),
        ("Look • ظاهر", ["direction", "font", "title_size", "color"]),
        ("Options • گزینه‌ها", ["remove", "replace_existing"]),
    ],
}


def sections_for(feature):
    """[(title, [Field, ...])] covering every field of ``feature`` exactly once."""
    fm = feature.field_map()
    used, out = set(), []
    for title, names in SECTIONS.get(feature.key, []):
        fields = [fm[n] for n in names if n in fm and n not in used]
        used.update(f.name for f in fields)
        if fields:
            out.append((title, fields))
    rest = [f for f in feature.fields if f.name not in used]
    if rest:
        out.append(("Other • سایر", rest))
    return out
