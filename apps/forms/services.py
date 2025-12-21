import requests
from msal import ConfidentialClientApplication
from typing import Dict, List, Any
from decouple import config


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
    
    def get_worksheet_complete_metadata(self, sharepoint_url: str, worksheet_name: str = None) -> Dict:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        
        site_id, file_path = self._parse_sharepoint_url(sharepoint_url)
        
        if not worksheet_name:
            worksheets = self.get_workbook_worksheets(sharepoint_url)
            worksheet_name = worksheets[0]["name"] if worksheets else "Sheet1"
        
        base_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/workbook/worksheets('{worksheet_name}')"
        
        range_response = requests.get(f"{base_url}/usedRange", headers=headers)
        if range_response.status_code != 200:
            raise Exception(f"Failed to get range: {range_response.text}")
        
        range_data = range_response.json()
        row_count = range_data["rowCount"]
        col_count = range_data["columnCount"]
        
        cells_metadata = []
        for row in range(row_count):
            for col in range(col_count):
                cell_address = self._get_cell_address(row, col)
                cell_data = self._get_cell_complete_data(base_url, cell_address, headers)
                cell_data["row"] = row
                cell_data["column"] = col
                cells_metadata.append(cell_data)
        
        return {
            "worksheet_name": worksheet_name,
            "dimensions": {"rows": row_count, "columns": col_count},
            "cells": cells_metadata,
            "raw_values": range_data.get("values", [])
        }
    
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
        # Simplified parsing - customize based on your SharePoint URL format
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