from pydantic import BaseModel
from typing import List, Dict, Any, Optional

class CellFont(BaseModel):
    name: str = ""
    size: float = 0
    bold: bool = False
    italic: bool = False
    underline: str = ""
    strikethrough: bool = False
    subscript: bool = False
    superscript: bool = False
    color: str = ""
    theme_color: str = ""
    tint_and_shade: float = 0

class CellFill(BaseModel):
    color: str = ""
    pattern_color: str = ""
    pattern_type: str = ""
    theme_color: str = ""
    tint_and_shade: float = 0
    gradient: Dict = {}

class CellAlignment(BaseModel):
    horizontal: str = ""
    vertical: str = ""
    wrap_text: bool = False
    indent: int = 0
    justify_distributed: bool = False
    reading_order: int = 0
    text_orientation: int = 0

class CellBorder(BaseModel):
    style: str = ""
    color: str = ""
    weight: str = ""

class CellValidation(BaseModel):
    type: str = ""
    operator: str = ""
    formula1: str = ""
    formula2: str = ""
    ignore_blank: bool = True
    in_cell_dropdown: bool = True
    show_input_message: bool = False
    show_error_alert: bool = False
    input_title: str = ""
    input_message: str = ""
    error_title: str = ""
    error_message: str = ""
    error_style: str = ""

class CellComment(BaseModel):
    id: str = ""
    content: str = ""
    author: str = ""
    creation_date: str = ""
    resolved: bool = False

class CellHyperlink(BaseModel):
    address: str = ""
    document_reference: str = ""
    screen_tip: str = ""
    text_to_display: str = ""

class CellProtection(BaseModel):
    locked: bool = True
    hidden: bool = False

class ConditionalFormat(BaseModel):
    type: str = ""
    priority: int = 0
    stop_if_true: bool = False
    format: Dict = {}
    rule: Dict = {}

class CellMetadata(BaseModel):
    address: str
    row: int
    column: int
    value: Any
    formula: str = ""
    display_value: str = ""
    data_type: str = ""
    number_format: str = ""
    font: CellFont = CellFont()
    fill: CellFill = CellFill()
    alignment: CellAlignment = CellAlignment()
    borders: Dict[str, CellBorder] = {}
    column_width: float = 0
    row_height: float = 0
    protection: CellProtection = CellProtection()
    validation: CellValidation = CellValidation()
    comments: List[CellComment] = []
    hyperlink: CellHyperlink = CellHyperlink()
    conditional_formatting: List[ConditionalFormat] = []
    is_merged: bool = False
    hidden: bool = False
    locked: bool = True
    shrink_to_fit: bool = False
    text_rotation: int = 0

class MergedCell(BaseModel):
    address: str
    range: str

class WorksheetMetadata(BaseModel):
    worksheet_name: str
    dimensions: Dict[str, int]
    cells: List[CellMetadata]
    merged_cells: List[MergedCell]
    raw_values: List[List[Any]]