import requests
from msal import ConfidentialClientApplication
from typing import Dict, List, Any
from decouple import config
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
    
    def process_sharepoint_url(self, sharepoint_url: str, created_by: str, updated_by: str) -> Dict:
        """Process SharePoint URL with display and entry worksheets"""
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
        
        # Get form name from entry sheet (remove 'entry' keyword)
        form_name = entry_sheet['name'].replace('entry', '').replace('Entry', '').strip()
        
        with transaction.atomic():
            # Create form entry
            form, created = Form.objects.get_or_create(
                form_name=form_name,
                defaults={
                    'source': 'sharepoint',
                    'url': sharepoint_url,
                    'created_by': created_by,
                    'updated_by': updated_by
                }
            )
            
            if not created:
                form.updated_by = updated_by
                form.save()
            
            # Process display sheet - get complete metadata till column 12
            display_metadata = self.get_display_sheet_metadata(sharepoint_url, display_sheet['name'])
            
            # Get next version for display
            display_version = self._get_next_version(form, 'display')
            
            FormDisplayVersion.objects.create(
                form=form,
                form_display_json=display_metadata,
                form_version=str(display_version),
                approved=False,
                created_by=created_by,
                updated_by=updated_by
            )
            
            # Process entry sheet - transform to JSON object list
            entry_data = self.get_entry_sheet_data(sharepoint_url, entry_sheet['name'])
            
            # Get next version for entry
            entry_version = self._get_next_version(form, 'entry')
            
            FormEntryVersion.objects.create(
                form=form,
                form_entry_json=entry_data,
                form_version=str(entry_version),
                approved=False,
                created_by=created_by,
                updated_by=updated_by
            )
            
            return {
                'form_id': form.id,
                'form_name': form_name,
                'display_version': display_version,
                'entry_version': entry_version,
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
        
        return {
            "worksheet_name": worksheet_name,
            "dimensions": {"rows": row_count, "columns": 12},
            "cells": cells_metadata
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
    
    def _get_next_version(self, form: Form, version_type: str) -> int:
        """Get next version number for form"""
        if version_type == 'display':
            last_version = FormDisplayVersion.objects.filter(form=form).order_by('-form_version').first()
        else:
            last_version = FormEntryVersion.objects.filter(form=form).order_by('-form_version').first()
        
        if last_version:
            try:
                return int(last_version.form_version) + 1
            except ValueError:
                return 1
        return 1
    
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
            values = cell_data.get("values", [[""]])
            value = values[0][0] if values and len(values) > 0 and len(values[0]) > 0 else ""
            
            formulas = cell_data.get("formulas", [[""]])
            formula = formulas[0][0] if formulas and len(formulas) > 0 and len(formulas[0]) > 0 else ""
            
            return {
                "address": cell_address,
                "value": value,
                "formula": formula,
                "data_type": self._infer_data_type(value)
            }
        except Exception as e:
            return {"address": cell_address, "value": "", "error": str(e)}
    
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