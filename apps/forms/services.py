from ctypes import sizeof
from tkinter import W
import requests
from msal import ConfidentialClientApplication
from typing import Dict, List, Any
from decouple import config
from django.db import transaction
from django.db import transaction
from .models import Form, FormDisplayVersion, FormEntryVersion


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

        print('worksheets', len(worksheets));
        for worksheet in worksheets:
            print('worksheet', worksheet['name']);
        
        display_sheet = None
        entry_sheet = None
        
        for worksheet in worksheets:
            name = worksheet['name'].lower()
            if 'display' in name:
                display_sheet = worksheet
            elif 'entry' in name:
                entry_sheet = worksheet
        
        print('display_sheet', display_sheet);
        print('entry_sheet', entry_sheet);
        if not display_sheet or not entry_sheet:
            raise Exception("Both 'display' and 'entry' worksheets are required")
        
        with transaction.atomic():
            # Create form entry
            form = Form.objects.create(
                form_name=form_name,
                source='sharepoint',
                url=sharepoint_url,
                created_by=created_by,
                updated_by=updated_by
            )
            
            # Process display sheet - get complete metadata till column 12
            display_metadata = self.get_display_sheet_metadata(sharepoint_url, display_sheet['name'])

            print('display_metadata', display_metadata);
            
            FormDisplayVersion.objects.create(
                form=form,
                form_display_json=display_metadata,
                form_version='1',
                approved=False,
                created_by=created_by,
                updated_by=updated_by
            )
            
            # Process entry sheet - transform to JSON object list
            entry_data = self.get_entry_sheet_data(sharepoint_url, entry_sheet['name'])

            print('entry_data', entry_data);
            
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
        
        # Get new data
        new_display_metadata = self.get_display_sheet_metadata(sharepoint_url, display_sheet['name'])
        new_entry_data = self.get_entry_sheet_data(sharepoint_url, entry_sheet['name'])
        
        # Get latest versions
        latest_display = FormDisplayVersion.objects.filter(form=form).order_by('-form_version').first()
        latest_entry = FormEntryVersion.objects.filter(form=form).order_by('-form_version').first()
        
        versions_updated = []
        display_version = int(latest_display.form_version) if latest_display else 0
        entry_version = int(latest_entry.form_version) if latest_entry else 0
        
        with transaction.atomic():
            # Update form URL
            form.url = sharepoint_url
            form.updated_by = updated_by
            form.save()
            
            # Check if display data changed
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
            
            # Check if entry data changed
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
        """Get complete metadata for display sheet till column 12"""
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        
        site_id, file_path = self._parse_sharepoint_url(sharepoint_url)
        base_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/workbook/worksheets('{worksheet_name}')"
        
        # Get used range
        range_response = requests.get(f"{base_url}/usedRange", headers=headers)
        if range_response.status_code != 200:
            raise Exception(f"Failed to get range: {range_response.text}")
        
        range_data = range_response.json()
        row_count = range_data["rowCount"]
        
        cells_metadata = []
        for row in range(row_count):
            for col in range(12):  # Only till column 12 (L)
                cell_address = self._get_cell_address(row, col)
                cell_data = self._get_cell_complete_data(base_url, cell_address, headers)
                cell_data["row"] = row
                cell_data["column"] = col
                cells_metadata.append(cell_data)
        
        # Get merged cells information
        merged_cells = self._get_merged_cells(base_url, headers)
        
        return {
            "worksheet_name": worksheet_name,
            "dimensions": {"rows": row_count, "columns": 12},
            "cells": cells_metadata,
            "merged_cells": merged_cells
        }
    
    def get_entry_sheet_data(self, sharepoint_url: str, worksheet_name: str) -> List[Dict]:
        """Get entry sheet data and transform to JSON object list"""
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        
        site_id, file_path = self._parse_sharepoint_url(sharepoint_url)
        base_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/workbook/worksheets('{worksheet_name}')"
        
        # Get used range values
        range_response = requests.get(f"{base_url}/usedRange", headers=headers)
        if range_response.status_code != 200:
            raise Exception(f"Failed to get range: {range_response.text}")
        
        range_data = range_response.json()
        values = range_data.get("values", [])
        
        if not values or len(values) < 2:
            return []
        
        # First row as column names
        headers_row = values[0]
        data_rows = values[1:]
        
        # Transform to JSON objects
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
    
    def _get_cell_address(self, row: int, col: int) -> str:
        col_letter = ""
        col_num = col + 1
        while col_num > 0:
            col_num -= 1
            col_letter = chr(col_num % 26 + ord('A')) + col_letter
            col_num //= 26
        return f"{col_letter}{row + 1}"
    
    def _get_cell_complete_data(self, base_url: str, cell_address: str, headers: Dict) -> Dict:
        try:
            cell_url = f"{base_url}/range(address='{cell_address}')"
            response = requests.get(cell_url, headers=headers)
            
            if response.status_code != 200:
                return {"address": cell_address, "value": "", "error": response.text}
            
            cell_data = response.json()
            
            # Get basic cell data
            values = cell_data.get("values", [[""]])
            value = values[0][0] if values and len(values) > 0 and len(values[0]) > 0 else ""
            
            formulas = cell_data.get("formulas", [[""]])
            formula = formulas[0][0] if formulas and len(formulas) > 0 and len(formulas[0]) > 0 else ""
            
            text = cell_data.get("text", [[""]])
            display_value = text[0][0] if text and len(text) > 0 and len(text[0]) > 0 else ""
            
            # Get comprehensive formatting data
            format_data = self._get_cell_format_data(cell_url, headers)
            
            # Get additional cell properties
            additional_data = self._get_additional_cell_data(cell_url, headers)
            
            return {
                "address": cell_address,
                "value": value,
                "formula": formula,
                "display_value": display_value,
                "data_type": self._infer_data_type(value),
                "font": format_data.get("font", {}),
                "fill": format_data.get("fill", {}),
                "alignment": format_data.get("alignment", {}),
                "borders": format_data.get("borders", {}),
                "number_format": format_data.get("number_format", {}),
                "protection": format_data.get("protection", {}),
                "column_width": format_data.get("column_width", 0),
                "row_height": format_data.get("row_height", 0),
                "column_hidden": format_data.get("column_hidden", False),
                "row_hidden": format_data.get("row_hidden", False),
                "hyperlink": additional_data.get("hyperlink", {}),
                "comment": additional_data.get("comment", {}),
                "validation": additional_data.get("validation", {}),
                "conditional_formats": additional_data.get("conditional_formats", [])
            }
        except Exception as e:
            return {"address": cell_address, "value": "", "error": str(e)}
    
    def _get_cell_format_data(self, cell_url: str, headers: Dict) -> Dict:
        try:
            format_url = f"{cell_url}/format"
            format_response = requests.get(format_url, headers=headers)
            if format_response.status_code != 200:
                return {}
            
            format_data = format_response.json()
            
            # Get font data
            font_data = {}
            try:
                font_response = requests.get(f"{format_url}/font", headers=headers)
                if font_response.status_code == 200:
                    font_data = font_response.json()
            except:
                pass
            
            # Get fill data
            fill_data = {}
            try:
                fill_response = requests.get(f"{format_url}/fill", headers=headers)
                if fill_response.status_code == 200:
                    fill_data = fill_response.json()
            except:
                pass
            
            # Get borders data
            borders_data = {}
            try:
                borders_response = requests.get(f"{format_url}/borders", headers=headers)
                if borders_response.status_code == 200:
                    borders_data = borders_response.json()
            except:
                pass
            
            # Get protection data
            protection_data = {}
            try:
                protection_response = requests.get(f"{format_url}/protection", headers=headers)
                if protection_response.status_code == 200:
                    protection_data = protection_response.json()
            except:
                pass
            
            return {
                "number_format": {
                    "format": format_data.get("numberFormat", {}).get("format", ""),
                    "format_local": format_data.get("numberFormat", {}).get("formatLocal", ""),
                    "category": format_data.get("numberFormat", {}).get("category", "")
                },
                "font": {
                    "name": font_data.get("name", ""),
                    "size": font_data.get("size", 0),
                    "bold": font_data.get("bold", False),
                    "italic": font_data.get("italic", False),
                    "underline": font_data.get("underline", ""),
                    "strikethrough": font_data.get("strikethrough", False),
                    "subscript": font_data.get("subscript", False),
                    "superscript": font_data.get("superscript", False),
                    "color": self._extract_color_value(font_data.get("color", "")),
                    "tint_and_shade": font_data.get("tintAndShade", 0)
                },
                "fill": {
                    "color": self._extract_color_value(fill_data.get("color", "")),
                    "pattern_color": self._extract_color_value(fill_data.get("patternColor", "")),
                    "pattern_type": fill_data.get("patternType", ""),
                    "tint_and_shade": fill_data.get("tintAndShade", 0),
                    "gradient": self._extract_gradient_info(fill_data.get("gradient", {}))
                },
                "alignment": {
                    "horizontal": format_data.get("horizontalAlignment", ""),
                    "vertical": format_data.get("verticalAlignment", ""),
                    "wrap_text": format_data.get("wrapText", False),
                    "indent": format_data.get("indentLevel", 0),
                    "text_rotation": format_data.get("textRotation", 0),
                    "justify_distributed": format_data.get("justifyDistributed", False),
                    "reading_order": format_data.get("readingOrder", ""),
                    "shrink_to_fit": format_data.get("shrinkToFit", False)
                },
                "borders": self._extract_border_info(borders_data),
                "protection": {
                    "locked": protection_data.get("locked", False),
                    "formula_hidden": protection_data.get("formulaHidden", False)
                },
                "column_width": format_data.get("columnWidth", 0),
                "row_height": format_data.get("rowHeight", 0),
                "column_hidden": format_data.get("columnHidden", False),
                "row_hidden": format_data.get("rowHidden", False)
            }
        except Exception as e:
            return {}
    
    def _extract_color_value(self, color_data) -> str:
        if isinstance(color_data, str):
            return color_data
        elif isinstance(color_data, dict):
            return color_data.get("index", "")
        return ""
    
    def _extract_border_info(self, borders_data: Dict) -> Dict:
        if not borders_data or "value" not in borders_data:
            return {}
        
        borders = {}
        for border in borders_data.get("value", []):
            side = border.get("sideIndex", "")
            borders[side] = {
                "style": border.get("style", ""),
                "color": self._extract_color_value(border.get("color", "")),
                "weight": border.get("weight", ""),
                "tint_and_shade": border.get("tintAndShade", 0),
                "line_style": border.get("lineStyle", "")
            }
        return borders
    
    def _extract_gradient_info(self, gradient_data: Dict) -> Dict:
        if not gradient_data:
            return {}
        
        return {
            "type": gradient_data.get("type", ""),
            "angle": gradient_data.get("angle", 0),
            "direction": gradient_data.get("direction", ""),
            "stops": gradient_data.get("stops", [])
        }
    
    def _get_additional_cell_data(self, cell_url: str, headers: Dict) -> Dict:
        additional_data = {
            "hyperlink": {},
            "comment": {},
            "validation": {},
            "conditional_formats": []
        }
        
        # Get hyperlink data
        try:
            hyperlink_response = requests.get(f"{cell_url}/hyperlink", headers=headers)
            if hyperlink_response.status_code == 200:
                hyperlink_data = hyperlink_response.json()
                additional_data["hyperlink"] = {
                    "address": hyperlink_data.get("address", ""),
                    "document_reference": hyperlink_data.get("documentReference", ""),
                    "screen_tip": hyperlink_data.get("screenTip", ""),
                    "text_to_display": hyperlink_data.get("textToDisplay", "")
                }
        except:
            pass
        
        # Get comment data
        try:
            comment_response = requests.get(f"{cell_url}/comment", headers=headers)
            if comment_response.status_code == 200:
                comment_data = comment_response.json()
                additional_data["comment"] = {
                    "content": comment_data.get("content", ""),
                    "author": comment_data.get("author", {}).get("name", ""),
                    "creation_date": comment_data.get("creationDate", ""),
                    "replies": comment_data.get("replies", [])
                }
        except:
            pass
        
        # Get data validation
        try:
            validation_response = requests.get(f"{cell_url}/dataValidation", headers=headers)
            if validation_response.status_code == 200:
                validation_data = validation_response.json()
                additional_data["validation"] = {
                    "type": validation_data.get("type", ""),
                    "operator": validation_data.get("operator", ""),
                    "formula1": validation_data.get("formula1", ""),
                    "formula2": validation_data.get("formula2", ""),
                    "ignore_blanks": validation_data.get("ignoreBlanks", False),
                    "show_input_message": validation_data.get("showInputMessage", False),
                    "show_error_alert": validation_data.get("showErrorAlert", False),
                    "input_title": validation_data.get("inputTitle", ""),
                    "input_message": validation_data.get("inputMessage", ""),
                    "error_title": validation_data.get("errorTitle", ""),
                    "error_message": validation_data.get("errorMessage", "")
                }
        except:
            pass
        
        return additional_data
    
    def _get_merged_cells(self, base_url: str, headers: Dict) -> List[Dict]:
        try:
            merged_url = f"{base_url}/mergedCells"
            response = requests.get(merged_url, headers=headers)
            
            if response.status_code == 200:
                merged_data = response.json().get("value", [])
                return [{
                    "address": cell.get("address", ""),
                    "range": cell.get("address", "").split("!")[-1] if "!" in cell.get("address", "") else cell.get("address", "")
                } for cell in merged_data]
            return []
        except Exception as e:
            return []
    
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