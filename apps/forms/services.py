import requests
from msal import ConfidentialClientApplication
from typing import Dict, List, Any
from decouple import config
from django.db import transaction
from .models import Form, FormDisplayVersion, FormEntryVersion
# import concurrent.futures  # No longer needed - was used for Graph API batch processing
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from io import BytesIO


class SharePointService:
    def __init__(self):
        self.client_app = ConfidentialClientApplication(
            config('MICROSOFT_CLIENT_ID'),
            authority=f"https://login.microsoftonline.com/{config('MICROSOFT_TENANT_ID')}",
            client_credential=config('MICROSOFT_CLIENT_SECRET'),
        )
        self.token = None
    
    def get_access_token(self) -> str:
        if not self.token:
            result = self.client_app.acquire_token_for_client(
                scopes=["https://graph.microsoft.com/.default"]
            )
            if "access_token" in result:
                self.token = result["access_token"]
            else:
                raise Exception(f"Failed to acquire token: {result.get('error_description')}")
        return self.token
    
    def create_new_form(self, sharepoint_url: str, form_name: str, created_by: str, updated_by: str) -> Dict:
        """Create new form from SharePoint URL"""
        worksheets = self.get_workbook_worksheets(sharepoint_url)
        
        display_sheet = None
        entry_sheet = None
        
        for worksheet in worksheets:
            name = worksheet['name'].lower()
            if 'display' in name:
                display_sheet = worksheet
            elif 'entry' in name:
                entry_sheet = worksheet
        
        if not display_sheet or not entry_sheet:
            raise Exception("Both 'display' and 'entry' worksheets are required")
        
        with transaction.atomic():
            form = Form.objects.create(
                form_name=form_name,
                source='sharepoint',
                url=sharepoint_url,
                created_by=created_by,
                updated_by=updated_by
            )
            
            display_metadata = self.get_display_sheet_metadata(sharepoint_url, display_sheet['name'])
            
            FormDisplayVersion.objects.create(
                form=form,
                form_display_json=display_metadata,
                form_version='1',
                approved=False,
                created_by=created_by,
                updated_by=updated_by
            )
            
            entry_data = self.get_entry_sheet_data(sharepoint_url, entry_sheet['name'])
            
            FormEntryVersion.objects.create(
                form=form,
                form_entry_json=entry_data,
                form_version='1',
                approved=False,
                created_by=created_by,
                updated_by=updated_by
            )
            
            return {
                'form_id': form.id,
                'form_name': form_name,
                'display_version': 1,
                'entry_version': 1,
                'display_sheet': display_sheet['name'],
                'entry_sheet': entry_sheet['name']
            }
    
    def update_existing_form(self, form_id: int, sharepoint_url: str, updated_by: str) -> Dict:
        """Update existing form from SharePoint URL"""
        form = Form.objects.get(id=form_id)
        worksheets = self.get_workbook_worksheets(sharepoint_url)
        
        display_sheet = None
        entry_sheet = None
        
        for worksheet in worksheets:
            name = worksheet['name'].lower()
            if 'display' in name:
                display_sheet = worksheet
            elif 'entry' in name:
                entry_sheet = worksheet
        
        if not display_sheet or not entry_sheet:
            raise Exception("Both 'display' and 'entry' worksheets are required")
        
        new_display_metadata = self.get_display_sheet_metadata(sharepoint_url, display_sheet['name'])
        new_entry_data = self.get_entry_sheet_data(sharepoint_url, entry_sheet['name'])
        
        latest_display = FormDisplayVersion.objects.filter(form=form).order_by('-form_version').first()
        latest_entry = FormEntryVersion.objects.filter(form=form).order_by('-form_version').first()
        
        versions_updated = []
        display_version = int(latest_display.form_version) if latest_display else 0
        entry_version = int(latest_entry.form_version) if latest_entry else 0
        
        with transaction.atomic():
            form.url = sharepoint_url
            form.updated_by = updated_by
            form.save()
            
            if not latest_display or latest_display.form_display_json != new_display_metadata:
                display_version += 1
                FormDisplayVersion.objects.create(
                    form=form,
                    form_display_json=new_display_metadata,
                    form_version=str(display_version),
                    approved=False,
                    created_by=updated_by,
                    updated_by=updated_by
                )
                versions_updated.append('display')
            
            if not latest_entry or latest_entry.form_entry_json != new_entry_data:
                entry_version += 1
                FormEntryVersion.objects.create(
                    form=form,
                    form_entry_json=new_entry_data,
                    form_version=str(entry_version),
                    approved=False,
                    created_by=updated_by,
                    updated_by=updated_by
                )
                versions_updated.append('entry')
            
            return {
                'form_id': form.id,
                'form_name': form.form_name,
                'display_version': display_version,
                'entry_version': entry_version,
                'versions_updated': versions_updated,
                'display_sheet': display_sheet['name'],
                'entry_sheet': entry_sheet['name']
            }
    
    def get_display_sheet_metadata(self, sharepoint_url: str, worksheet_name: str) -> Dict:
        """Get complete metadata for display sheet using openpyxl"""
        return self._get_display_metadata_from_file(sharepoint_url, worksheet_name)
    
    def get_entry_sheet_data(self, sharepoint_url: str, worksheet_name: str) -> List[Dict]:
        """Get entry sheet data and transform to JSON object list"""
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        
        site_id, file_path = self._parse_sharepoint_url(sharepoint_url)
        base_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/workbook/worksheets('{worksheet_name}')"
        
        range_response = requests.get(f"{base_url}/usedRange", headers=headers)
        if range_response.status_code != 200:
            raise Exception(f"Failed to get range: {range_response.text}")
        
        range_data = range_response.json()
        values = range_data.get("values", [])
        
        if not values or len(values) < 2:
            return []
        
        headers_row = values[0]
        data_rows = values[1:]
        
        json_objects = []
        for row in data_rows:
            row_obj = {}
            for i, header in enumerate(headers_row):
                if i < len(row):
                    row_obj[str(header)] = row[i]
                else:
                    row_obj[str(header)] = None
            json_objects.append(row_obj)
        
        return json_objects
    
    def get_workbook_worksheets(self, sharepoint_url: str) -> List[Dict]:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        
        site_id, file_path = self._parse_sharepoint_url(sharepoint_url)
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/workbook/worksheets"
        
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json().get("value", [])
        else:
            raise Exception(f"Failed to get worksheets: {response.text}")
    
    def _parse_sharepoint_url(self, sharepoint_url: str) -> tuple:
        site_id = self._get_site_id("https://persivx.sharepoint.com/sites/test")
        return site_id, "Test.xlsx"
    
    def _get_site_id(self, site_url: str) -> str:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        
        parts = site_url.replace("https://", "").split("/")
        domain = parts[0]
        site_path = "/".join(parts[1:])
        
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{domain}:/{site_path}"
        response = requests.get(graph_url, headers=headers)
        
        if response.status_code == 200:
            return response.json()["id"]
        else:
            raise Exception(f"Failed to get site ID: {response.text}")
    
    # ============================================================================
    # OBSOLETE GRAPH API METHODS - Replaced by openpyxl-based implementation
    # Can be removed in future cleanup
    # ============================================================================
    
    # def _get_batch_range_data(self, base_url: str, range_address: str, headers: Dict) -> Dict:
    #     """Get all range data in parallel batch requests"""
    #     range_url = f"{base_url}/range(address='{range_address}')"
    #     
    #     requests_to_make = [
    #         ("values", f"{range_url}"),
    #         ("formulas", f"{range_url}"),
    #         ("text", f"{range_url}"),
    #         ("format", f"{range_url}/format"),
    #         ("font", f"{range_url}/format/font"),
    #         ("fill", f"{range_url}/format/fill"),
    #         ("borders", f"{range_url}/format/borders"),
    #         ("protection", f"{range_url}/format/protection")
    #     ]
    #     
    #     results = {}
    #     with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
    #         future_to_key = {
    #             executor.submit(self._make_request, url, headers): key 
    #             for key, url in requests_to_make
    #         }
    #         
    #         for future in concurrent.futures.as_completed(future_to_key):
    #             key = future_to_key[future]
    #             try:
    #                 results[key] = future.result()
    #             except Exception as e:
    #                 results[key] = {}
    #     
    #     return results
    
    # def _make_request(self, url: str, headers: Dict) -> Dict:
    #     """Make a single HTTP request"""
    #     try:
    #         response = requests.get(url, headers=headers)
    #         return response.json() if response.status_code == 200 else {}
    #     except Exception:
    #         return {}
    
    # def _process_batch_data_to_cells(self, batch_data: Dict, row_count: int, col_count: int) -> List[Dict]:
    #     """Process batch range data into individual cell metadata"""
    #     cells_metadata = []
    #     
    #     values = batch_data.get("values", {}).get("values", [])
    #     formulas = batch_data.get("formulas", {}).get("formulas", [])
    #     text = batch_data.get("text", {}).get("text", [])
    #     format_data = batch_data.get("format", {})
    #     font_data = batch_data.get("font", {})
    #     fill_data = batch_data.get("fill", {})
    #     borders_data = batch_data.get("borders", {})
    #     protection_data = batch_data.get("protection", {})
    #     
    #     for row in range(row_count):
    #         for col in range(col_count):
    #             cell_address = self._get_cell_address(row, col)
    #             
    #             cell_value = self._safe_get_array_value(values, row, col)
    #             cell_formula = self._safe_get_array_value(formulas, row, col)
    #             cell_text = self._safe_get_array_value(text, row, col)
    #             
    #             cell_data = {
    #                 "address": cell_address,
    #                 "value": cell_value,
    #                 "formula": cell_formula,
    #                 "display_value": cell_text,
    #                 "data_type": self._infer_data_type(cell_value),
    #                 "row": row,
    #                 "column": col,
    #                 "font": self._extract_cell_font_data(font_data),
    #                 "fill": self._extract_cell_fill_data(fill_data),
    #                 "alignment": self._extract_cell_alignment_data(format_data),
    #                 "borders": self._extract_cell_border_data(borders_data),
    #                 "number_format": self._extract_cell_number_format(format_data),
    #                 "protection": self._extract_cell_protection_data(protection_data),
    #                 "column_width": format_data.get("columnWidth", 0),
    #                 "row_height": format_data.get("rowHeight", 0),
    #                 "column_hidden": format_data.get("columnHidden", False),
    #                 "row_hidden": format_data.get("rowHidden", False),
    #                 "hyperlink": {},
    #                 "comment": {},
    #                 "validation": {},
    #                 "conditional_formats": []
    #             }
    #             
    #             cells_metadata.append(cell_data)
    #     
    #     return cells_metadata
    
    # def _get_cell_address(self, row: int, col: int) -> str:
    #     col_letter = ""
    #     col_num = col + 1
    #     while col_num > 0:
    #         col_num -= 1
    #         col_letter = chr(col_num % 26 + ord('A')) + col_letter
    #         col_num //= 26
    #     return f"{col_letter}{row + 1}"
    
    # def _safe_get_array_value(self, array: List, row: int, col: int):
    #     try:
    #         if array and len(array) > row and len(array[row]) > col:
    #             return array[row][col]
    #         return ""
    #     except (IndexError, TypeError):
    #         return ""
    
    # def _extract_cell_font_data(self, font_data: Dict) -> Dict:
    #     return {
    #         "name": font_data.get("name", ""),
    #         "size": font_data.get("size", 0),
    #         "bold": font_data.get("bold", False),
    #         "italic": font_data.get("italic", False),
    #         "underline": font_data.get("underline", ""),
    #         "strikethrough": font_data.get("strikethrough", False),
    #         "subscript": font_data.get("subscript", False),
    #         "superscript": font_data.get("superscript", False),
    #         "color": self._extract_color_value(font_data.get("color", "")),
    #         "tint_and_shade": font_data.get("tintAndShade", 0)
    #     }
    
    # def _extract_cell_fill_data(self, fill_data: Dict) -> Dict:
    #     return {
    #         "color": self._extract_color_value(fill_data.get("color", "")),
    #         "pattern_color": self._extract_color_value(fill_data.get("patternColor", "")),
    #         "pattern_type": fill_data.get("patternType", ""),
    #         "tint_and_shade": fill_data.get("tintAndShade", 0),
    #         "gradient": self._extract_gradient_info(fill_data.get("gradient", {}))
    #     }
    
    # def _extract_cell_alignment_data(self, format_data: Dict) -> Dict:
    #     return {
    #         "horizontal": format_data.get("horizontalAlignment", ""),
    #         "vertical": format_data.get("verticalAlignment", ""),
    #         "wrap_text": format_data.get("wrapText", False),
    #         "indent": format_data.get("indentLevel", 0),
    #         "text_rotation": format_data.get("textRotation", 0),
    #         "justify_distributed": format_data.get("justifyDistributed", False),
    #         "reading_order": format_data.get("readingOrder", ""),
    #         "shrink_to_fit": format_data.get("shrinkToFit", False)
    #     }
    
    # def _extract_cell_border_data(self, borders_data: Dict) -> Dict:
    #     if not borders_data or "value" not in borders_data:
    #         return {}
    #     
    #     borders = {}
    #     for border in borders_data.get("value", []):
    #         side = border.get("sideIndex", "")
    #         borders[side] = {
    #             "style": border.get("style", ""),
    #             "color": self._extract_color_value(border.get("color", "")),
    #             "weight": border.get("weight", ""),
    #             "tint_and_shade": border.get("tintAndShade", 0),
    #             "line_style": border.get("lineStyle", "")
    #         }
    #     return borders
    
    # def _extract_cell_number_format(self, format_data: Dict) -> Dict:
    #     number_format = format_data.get("numberFormat", {})
    #     return {
    #         "format": number_format.get("format", ""),
    #         "format_local": number_format.get("formatLocal", ""),
    #         "category": number_format.get("category", "")
    #     }
    
    # def _extract_cell_protection_data(self, protection_data: Dict) -> Dict:
    #     return {
    #         "locked": protection_data.get("locked", False),
    #         "formula_hidden": protection_data.get("formulaHidden", False)
    #     }
    
    # def _extract_color_value(self, color_data) -> str:
    #     if isinstance(color_data, str):
    #         return color_data
    #     elif isinstance(color_data, dict):
    #         return color_data.get("index", "")
    #     return ""
    
    # def _extract_gradient_info(self, gradient_data: Dict) -> Dict:
    #     if not gradient_data:
    #         return {}
    #     
    #     return {
    #         "type": gradient_data.get("type", ""),
    #         "angle": gradient_data.get("angle", 0),
    #         "direction": gradient_data.get("direction", ""),
    #         "stops": gradient_data.get("stops", [])
    #     }
    
    # ============================================================================
    # END OF OBSOLETE METHODS
    # ============================================================================
    
    def _get_display_metadata_from_file(self, sharepoint_url: str, worksheet_name: str) -> Dict:
        """Extract complete display metadata using openpyxl"""
        try:
            token = self.get_access_token()
            headers = {"Authorization": f"Bearer {token}"}
            
            site_id, file_path = self._parse_sharepoint_url(sharepoint_url)
            download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/content"
            
            response = requests.get(download_url, headers=headers)
            if response.status_code != 200:
                raise Exception(f"Failed to download file: {response.status_code}")
            
            wb = load_workbook(BytesIO(response.content), data_only=False)
            ws = wb[worksheet_name]
            
            # Load workbook again with data_only=True to get calculated values
            wb_data = load_workbook(BytesIO(response.content), data_only=True)
            ws_data = wb_data[worksheet_name]
            
            # Determine dimensions (max 12 columns)
            max_row = ws.max_row
            max_col = min(ws.max_column, 12)
            
            # Extract cells metadata
            cells_metadata = []
            for row_idx in range(1, max_row + 1):
                for col_idx in range(1, max_col + 1):
                    cell = ws.cell(row_idx, col_idx)
                    cell_data_value = ws_data.cell(row_idx, col_idx)
                    cell_data = self._extract_openpyxl_cell_data(cell, cell_data_value, row_idx - 1, col_idx - 1, ws)
                    cells_metadata.append(cell_data)
            
            # Extract merged cells
            merged_cells = []
            for merged_range in ws.merged_cells.ranges:
                if merged_range.min_col <= max_col:
                    merged_cells.append({
                        "range": str(merged_range),
                        "start_row": merged_range.min_row - 1,
                        "start_col": merged_range.min_col - 1,
                        "row_span": merged_range.max_row - merged_range.min_row + 1,
                        "col_span": merged_range.max_col - merged_range.min_col + 1
                    })
            
            return {
                "worksheet_name": worksheet_name,
                "dimensions": {"rows": max_row, "columns": max_col},
                "cells": cells_metadata,
                "merged_cells": merged_cells
            }
            
        except Exception as e:
            raise Exception(f"Failed to extract display metadata: {e}")
    
    def _extract_openpyxl_cell_data(self, cell, cell_data_value, row: int, col: int, ws) -> Dict:
        """Extract all metadata from an openpyxl cell"""
        # Handle merged cells - they don't have full attributes
        if isinstance(cell, MergedCell):
            col_letter = chr(65 + col) if col < 26 else chr(64 + col // 26) + chr(65 + col % 26)
            return {
                "address": f"{col_letter}{row + 1}",
                "value": None,
                "formula": None,
                "display_value": "",
                "data_type": "empty",
                "row": row,
                "column": col,
                "font": {"name": "", "size": 11, "bold": False, "italic": False, "underline": "none", "strikethrough": False, "color": ""},
                "fill": {"color": "", "pattern_type": ""},
                "alignment": {"horizontal": "", "vertical": "", "wrap_text": False, "indent": 0, "text_rotation": 0},
                "borders": {"left": {"style": ""}, "right": {"style": ""}, "top": {"style": ""}, "bottom": {"style": ""}},
                "number_format": {"format": ""},
                "protection": {"locked": True},
                "column_width": ws.column_dimensions[col_letter].width if col_letter in ws.column_dimensions else 8.43,
                "row_height": ws.row_dimensions[row + 1].height if row + 1 in ws.row_dimensions else 15,
                "column_hidden": False,
                "row_hidden": False,
                "hyperlink": {"target": ""},
                "comment": {"text": ""},
                "validation": {},
                "conditional_formats": []
            }
        
        # Separate formula and value
        formula_value = cell.value if isinstance(cell.value, str) and cell.value.startswith('=') else None
        display_value = cell_data_value.value if cell_data_value else cell.value
        
        return {
            "address": cell.coordinate,
            "value": display_value,
            "formula": formula_value,
            "display_value": str(display_value) if display_value is not None else "",
            "data_type": self._infer_data_type(display_value),
            "row": row,
            "column": col,
            "font": {
                "name": cell.font.name if cell.font else "",
                "size": cell.font.size if cell.font else 11,
                "bold": cell.font.bold if cell.font else False,
                "italic": cell.font.italic if cell.font else False,
                "underline": cell.font.underline if cell.font else "none",
                "strikethrough": cell.font.strike if cell.font else False,
                "color": self._get_color_value(cell.font.color) if cell.font and cell.font.color else ""
            },
            "fill": {
                "color": self._get_color_value(cell.fill.fgColor) if cell.fill and cell.fill.patternType and cell.fill.patternType != 'none' else "",
                "pattern_type": cell.fill.patternType if cell.fill else ""
            },
            "alignment": {
                "horizontal": cell.alignment.horizontal if cell.alignment else "",
                "vertical": cell.alignment.vertical if cell.alignment else "",
                "wrap_text": cell.alignment.wrap_text if cell.alignment else False,
                "indent": cell.alignment.indent if cell.alignment else 0,
                "text_rotation": cell.alignment.text_rotation if cell.alignment else 0
            },
            "borders": {
                "left": {"style": cell.border.left.style if cell.border and cell.border.left else ""},
                "right": {"style": cell.border.right.style if cell.border and cell.border.right else ""},
                "top": {"style": cell.border.top.style if cell.border and cell.border.top else ""},
                "bottom": {"style": cell.border.bottom.style if cell.border and cell.border.bottom else ""}
            },
            "number_format": {
                "format": cell.number_format if cell.number_format else ""
            },
            "protection": {
                "locked": cell.protection.locked if cell.protection else True
            },
            "column_width": ws.column_dimensions[cell.column_letter].width if cell.column_letter in ws.column_dimensions else 8.43,
            "row_height": ws.row_dimensions[cell.row].height if cell.row in ws.row_dimensions else 15,
            "column_hidden": ws.column_dimensions[cell.column_letter].hidden if cell.column_letter in ws.column_dimensions else False,
            "row_hidden": ws.row_dimensions[cell.row].hidden if cell.row in ws.row_dimensions else False,
            "hyperlink": {"target": cell.hyperlink.target if cell.hyperlink else ""},
            "comment": {"text": cell.comment.text if cell.comment else ""},
            "validation": {},
            "conditional_formats": []
        }
    

    
    def _get_color_value(self, color_obj) -> str:
        """Extract color value from openpyxl color object"""
        if not color_obj:
            return ""
        
        # Try to get RGB value
        if hasattr(color_obj, 'rgb') and color_obj.rgb:
            rgb = str(color_obj.rgb)
            # Filter out default/empty colors (00000000, 000000, etc.)
            if rgb and rgb not in ['00000000', '000000']:
                # Remove alpha channel if present (first 2 chars)
                return rgb[-6:] if len(rgb) == 8 else rgb
        
        return ""
    
    def _infer_data_type(self, value) -> str:
        if value is None or value == "":
            return "empty"
        elif isinstance(value, bool):
            return "boolean"
        elif isinstance(value, (int, float)):
            return "number"
        elif isinstance(value, str):
            if value.startswith("="):
                return "formula"
            return "text"
        else:
            return "unknown"