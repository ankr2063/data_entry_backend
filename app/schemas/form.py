from pydantic import BaseModel, HttpUrl
from typing import List, Dict, Any, Optional

class SharePointRequest(BaseModel):
    sharepoint_url: HttpUrl
    main_sheet_name: str = "Sheet1"
    config_sheet_name: Optional[str] = None

class FieldValidation(BaseModel):
    min: Optional[int] = None
    max: Optional[int] = None
    minLength: Optional[int] = None
    maxLength: Optional[int] = None
    pattern: Optional[str] = None

class FormField(BaseModel):
    id: str
    name: str
    label: str
    type: str
    required: bool = False
    validation: FieldValidation = FieldValidation()
    options: List[str] = []
    placeholder: str = ""
    default_value: Any = ""
    row: int
    column: int

class FormSection(BaseModel):
    id: str
    name: str
    fields: List[str]

class FormMetadata(BaseModel):
    total_rows: int
    total_columns: int
    has_sections: bool

class FormSchema(BaseModel):
    fields: List[FormField]
    sections: List[FormSection]
    metadata: FormMetadata

class FormSubmission(BaseModel):
    form_data: Dict[str, Any]
    sharepoint_url: HttpUrl
    sheet_name: str