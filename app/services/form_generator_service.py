from typing import Dict, List, Any, Optional
import re

class FormGeneratorService:
    
    @staticmethod
    def extract_form_metadata(main_sheet_data: Dict, config_sheet_data: Optional[Dict] = None) -> Dict:
        """Extract form structure and generate UI metadata"""
        
        # Parse main sheet structure
        rows = main_sheet_data.get("values", [])
        if not rows:
            return {"fields": [], "sections": []}
        
        # Extract headers and structure
        headers = rows[0] if rows else []
        form_fields = []
        sections = []
        
        # Parse configuration sheet if provided
        field_configs = {}
        if config_sheet_data:
            field_configs = FormGeneratorService._parse_config_sheet(config_sheet_data)
        
        # Generate form structure
        current_section = None
        
        for row_idx, row in enumerate(rows[1:], 1):  # Skip header row
            if not any(cell.strip() for cell in row if cell):  # Skip empty rows
                continue
                
            # Check if this is a section header (merged cells or specific pattern)
            if FormGeneratorService._is_section_header(row, headers):
                section_name = next((cell for cell in row if cell and cell.strip()), "")
                if section_name:
                    current_section = {
                        "id": f"section_{len(sections)}",
                        "name": section_name,
                        "fields": []
                    }
                    sections.append(current_section)
                continue
            
            # Process regular data rows
            for col_idx, (header, cell_value) in enumerate(zip(headers, row)):
                if not header or not header.strip():
                    continue
                    
                field_id = f"field_{row_idx}_{col_idx}"
                field_config = field_configs.get(header, {})
                
                field = {
                    "id": field_id,
                    "name": header,
                    "label": header,
                    "type": field_config.get("type", FormGeneratorService._infer_field_type(cell_value)),
                    "required": field_config.get("required", False),
                    "validation": field_config.get("validation", {}),
                    "options": field_config.get("options", []),
                    "placeholder": field_config.get("placeholder", ""),
                    "default_value": cell_value if cell_value else "",
                    "row": row_idx,
                    "column": col_idx
                }
                
                form_fields.append(field)
                
                if current_section:
                    current_section["fields"].append(field_id)
        
        return {
            "fields": form_fields,
            "sections": sections,
            "metadata": {
                "total_rows": len(rows),
                "total_columns": len(headers),
                "has_sections": len(sections) > 0
            }
        }
    
    @staticmethod
    def _parse_config_sheet(config_data: Dict) -> Dict:
        """Parse configuration sheet to extract field validation rules"""
        rows = config_data.get("values", [])
        if len(rows) < 2:
            return {}
        
        headers = rows[0]
        field_configs = {}
        
        for row in rows[1:]:
            if len(row) < 2:
                continue
                
            field_name = row[0]
            config = {}
            
            # Map configuration columns (adjust based on your config sheet structure)
            for i, header in enumerate(headers[1:], 1):
                if i < len(row) and row[i]:
                    if header.lower() == "type":
                        config["type"] = row[i]
                    elif header.lower() == "required":
                        config["required"] = row[i].lower() in ["true", "yes", "1"]
                    elif header.lower() == "validation":
                        config["validation"] = FormGeneratorService._parse_validation(row[i])
                    elif header.lower() == "options":
                        config["options"] = [opt.strip() for opt in row[i].split(",")]
                    elif header.lower() == "placeholder":
                        config["placeholder"] = row[i]
            
            field_configs[field_name] = config
        
        return field_configs
    
    @staticmethod
    def _is_section_header(row: List, headers: List) -> bool:
        """Determine if a row represents a section header"""
        # Check if only first cell has content (typical for merged section headers)
        non_empty_cells = [i for i, cell in enumerate(row) if cell and cell.strip()]
        return len(non_empty_cells) == 1 and non_empty_cells[0] == 0
    
    @staticmethod
    def _infer_field_type(value: str) -> str:
        """Infer field type from cell value"""
        if not value:
            return "text"
        
        value_str = str(value).strip()
        
        # Check for common patterns
        if re.match(r'^\d+$', value_str):
            return "number"
        elif re.match(r'^\d+\.\d+$', value_str):
            return "decimal"
        elif re.match(r'^\d{4}-\d{2}-\d{2}$', value_str):
            return "date"
        elif re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', value_str):
            return "email"
        elif value_str.lower() in ["true", "false", "yes", "no"]:
            return "boolean"
        elif len(value_str) > 100:
            return "textarea"
        else:
            return "text"
    
    @staticmethod
    def _parse_validation(validation_str: str) -> Dict:
        """Parse validation string into validation rules"""
        if not validation_str:
            return {}
        
        validation = {}
        rules = validation_str.split(";")
        
        for rule in rules:
            if ":" in rule:
                key, value = rule.split(":", 1)
                key = key.strip().lower()
                value = value.strip()
                
                if key == "min":
                    validation["min"] = int(value)
                elif key == "max":
                    validation["max"] = int(value)
                elif key == "pattern":
                    validation["pattern"] = value
                elif key == "minlength":
                    validation["minLength"] = int(value)
                elif key == "maxlength":
                    validation["maxLength"] = int(value)
        
        return validation