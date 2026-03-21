"""Document layout utilities."""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def get_usable_width(document):
    """Calculate usable width for tables based on page margins."""
    section = document.sections[0]
    page_width = section.page_width
    left_margin = section.left_margin
    right_margin = section.right_margin
    return page_width - left_margin - right_margin


def set_table_borders(table, border_spec):
    """Apply border specification to a table."""
    tbl_pr = table._tbl.tblPr
    existing = tbl_pr.find(qn("w:tblBorders"))
    if existing is not None:
        tbl_pr.remove(existing)

    tbl_borders = OxmlElement("w:tblBorders")
    for border, attrs in border_spec.items():
        border_element = OxmlElement(f"w:{border}")
        for key, value in attrs.items():
            border_element.set(qn(f"w:{key}"), str(value))
        tbl_borders.append(border_element)
    tbl_pr.append(tbl_borders)


def set_table_cell_margins(table, top=40, left=80, bottom=40, right=80):
    """Set cell margins for a table."""
    tbl_pr = table._tbl.tblPr
    existing = tbl_pr.find(qn("w:tblCellMar"))
    if existing is not None:
        tbl_pr.remove(existing)

    tbl_cell_mar = OxmlElement("w:tblCellMar")
    for side, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        side_element = OxmlElement(f"w:{side}")
        side_element.set(qn("w:w"), str(value))
        side_element.set(qn("w:type"), "dxa")
        tbl_cell_mar.append(side_element)
    tbl_pr.append(tbl_cell_mar)


def apply_table_theme_style(table, theme):
    """Apply theme styling to a table."""
    from gitdiff2docx.theme.styling import build_table_border_spec
    
    table.autofit = False

    margins = theme.get("table_cell_margins", {})
    set_table_cell_margins(
        table,
        top=int(margins.get("top", 40)),
        left=int(margins.get("left", 80)),
        bottom=int(margins.get("bottom", 40)),
        right=int(margins.get("right", 80)),
    )
    set_table_borders(table, build_table_border_spec(theme))
