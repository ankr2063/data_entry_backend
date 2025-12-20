import requests
from msal import ConfidentialClientApplication
from typing import Dict, List, Any, Optional
from app.core.config import settings

class SharePointService:
    def __init__(self):
        self.client_app = ConfidentialClientApplication(
            settings.microsoft_client_id,
            authority=f"https://login.microsoftonline.com/{settings.microsoft_tenant_id}",
            client_credential=settings.microsoft_client_secret,
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
    
    def get_workbook_worksheets(self, sharepoint_url: str) -> List[Dict]:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        
        site_id, file_path = self._parse_sharepoint_url(sharepoint_url)
        
        # Try different API endpoints for file access
        endpoints = [
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/workbook/worksheets",
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_path}/workbook/worksheets",
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/documents/root:/{file_path}:/workbook/worksheets"
        ]
        
        for url in endpoints:
            print(f"DEBUG: Trying endpoint: {url}")
            response = requests.get(url, headers=headers)
            print(f"DEBUG: Response status: {response.status_code}")
            
            if response.status_code == 200:
                return response.json().get("value", [])
            else:
                print(f"DEBUG: Failed with: {response.text}")
        
        # If all endpoints fail, try searching for the file
        return self._search_and_get_worksheets(site_id, file_path, headers)
    
    def get_worksheet_complete_metadata(self, sharepoint_url: str, worksheet_name: str = None) -> Dict:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        
        site_id, file_path = self._parse_sharepoint_url(sharepoint_url)
        
        if not worksheet_name:
            worksheets = self.get_workbook_worksheets(sharepoint_url)
            worksheet_name = worksheets[0]["name"] if worksheets else "Sheet1"
        
        # Try different base URLs
        base_urls = [
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/workbook/worksheets('{worksheet_name}')",
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/documents/root:/{file_path}:/workbook/worksheets('{worksheet_name}')"
        ]
        
        # If direct access fails, search for file
        search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='{file_path}')"
        search_response = requests.get(search_url, headers=headers)
        
        if search_response.status_code == 200:
            search_results = search_response.json().get("value", [])
            for file_item in search_results:
                if file_item.get("name", "").lower() == file_path.lower():
                    file_id = file_item["id"]
                    base_urls.append(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}/workbook/worksheets('{worksheet_name}')")
                    break
        
        base_url = None
        for url in base_urls:
            test_response = requests.get(f"{url}/usedRange", headers=headers)
            if test_response.status_code == 200:
                base_url = url
                break
        
        print(f"DEBUG: {base_url}")
        if not base_url:
            raise Exception(f"Failed to access worksheet '{worksheet_name}' in file '{file_path}'")
        
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
        
        merged_cells = self._get_merged_cells(base_url, headers)
        
        return {
            "worksheet_name": worksheet_name,
            "dimensions": {"rows": row_count, "columns": col_count},
            "cells": cells_metadata,
            "merged_cells": merged_cells,
            "raw_values": range_data.get("values", [])
        }
    
    def _get_cell_address(self, row: int, col: int) -> str:
        col_letter = ""
        col_num = col + 1
        while col_num > 0:
            col_num -= 1
            col_letter = chr(col_num % 26 + ord('A')) + col_letter
            col_num //= 26
        return f"{col_letter}{row + 1}"
    
    def _parse_sharepoint_url(self, sharepoint_url: str) -> tuple:
        print(f"DEBUG: Parsing URL: {sharepoint_url}")
        site_id = self._get_site_id(f"https://persivx.sharepoint.com/sites/test")
        return site_id, "Test.xlsx"
        # try:
            # if ":x:" in sharepoint_url and "sourcedoc=" in sharepoint_url:
            #     # Handle sharing links with document ID
            #     parts = sharepoint_url.split("/")
            #     domain = parts[2]  # persivx.sharepoint.com
            #     site_name = parts[4]  # test
                
            #     # Extract filename from URL
            #     import re
            #     filename_match = re.search(r'file=([^&]+)', sharepoint_url)
            #     filename = filename_match.group(1) if filename_match else "Test.xlsx"
                
            #     print(f"DEBUG: Sharing link - domain: {domain}, site: {site_name}, filename: {filename}")
                
            #     site_id = self._get_site_id(f"https://{domain}/sites/{site_name}")
            #     return site_id, filename
            
            # elif ":x:" in sharepoint_url or ":b:" in sharepoint_url:
            #     # Handle other sharing links
            #     parts = sharepoint_url.split("/")
            #     domain = parts[2]
            #     site_name = parts[4]
                
            #     site_id = self._get_site_id(f"https://{domain}/sites/{site_name}")
            #     file_path = "Shared Documents/Test.xlsx"
            #     return site_id, file_path
            
            # else:
            #     # Handle direct URLs
            #     parts = sharepoint_url.split("/")
            #     if len(parts) < 6:
            #         raise Exception(f"Invalid direct URL format. Got {len(parts)} parts, need at least 6")
                
            #     domain = parts[2]
            #     site_name = parts[4]
            #     file_path = "/".join(parts[5:])
                
            #     site_id = self._get_site_id(f"https://{domain}/sites/{site_name}")
            #     return site_id, file_path
                
        # except Exception as e:
        #     print(f"DEBUG: Exception: {e}")
        #     raise Exception(f"Failed to parse SharePoint URL: {str(e)}")
    
    def _get_site_id(self, site_url: str) -> str:
        print(f"DEBUG: Getting site ID for: {site_url}")
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        
        parts = site_url.replace("https://", "").split("/")
        domain = parts[0]  # persivx.sharepoint.com
        site_path = "/".join(parts[1:]) if len(parts) > 1 else ""  # sites/test
        
        print(f"DEBUG: Domain: {domain}, Site path: {site_path}")
        
        # Try different Graph API endpoints
        graph_urls = [
            f"https://graph.microsoft.com/v1.0/sites/{domain}:/{site_path}",
            f"https://graph.microsoft.com/v1.0/sites/{domain}/sites/{parts[-1]}" if len(parts) > 1 else None,
            f"https://graph.microsoft.com/v1.0/sites/{domain}:/sites/{parts[-1]}" if len(parts) > 1 else None,
            f"https://graph.microsoft.com/v1.0/sites/root/sites/{parts[-1]}" if len(parts) > 1 else None
        ]
        
        for graph_url in graph_urls:
            if not graph_url:
                continue
                
            print(f"DEBUG: Trying Graph URL: {graph_url}")
            response = requests.get(graph_url, headers=headers)
            print(f"DEBUG: Response status: {response.status_code}")
            
            if response.status_code == 200:
                site_id = response.json()["id"]
                print(f"DEBUG: Got site ID: {site_id}")
                return site_id
            else:
                print(f"DEBUG: Failed with: {response.text}")
        
        # Try listing all sites to find the one we need
        print("DEBUG: Trying to list all sites")
        list_sites_url = "https://graph.microsoft.com/v1.0/sites?search=*"
        response = requests.get(list_sites_url, headers=headers)
        
        if response.status_code == 200:
            sites = response.json().get("value", [])
            print(f"DEBUG: Found {len(sites)} sites")
            
            for site in sites:
                site_name = site.get("name", "").lower()
                site_display_name = site.get("displayName", "").lower()
                web_url = site.get("webUrl", "").lower()
                
                print(f"DEBUG: Checking site: {site_name}, {site_display_name}, {web_url}")
                
                if (parts[-1].lower() in site_name or 
                    parts[-1].lower() in site_display_name or 
                    parts[-1].lower() in web_url):
                    site_id = site["id"]
                    print(f"DEBUG: Found matching site with ID: {site_id}")
                    return site_id
        
        # If all direct methods fail, try to get the root site and search from there
        print("DEBUG: Trying root site approach")
        root_site_url = f"https://graph.microsoft.com/v1.0/sites/{domain}"
        response = requests.get(root_site_url, headers=headers)
        
        if response.status_code == 200:
            root_site = response.json()
            print(f"DEBUG: Got root site: {root_site.get('id')}")
            
            # Try to get subsites
            subsites_url = f"https://graph.microsoft.com/v1.0/sites/{root_site['id']}/sites"
            subsites_response = requests.get(subsites_url, headers=headers)
            
            if subsites_response.status_code == 200:
                subsites = subsites_response.json().get("value", [])
                print(f"DEBUG: Found {len(subsites)} subsites")
                
                for subsite in subsites:
                    if parts[-1].lower() in subsite.get("name", "").lower():
                        print(f"DEBUG: Found matching subsite: {subsite['id']}")
                        return subsite["id"]
        
        raise Exception(f"Failed to get site ID for {site_url}. Check if the site exists and you have permissions.")
    
    def _get_cell_complete_data(self, base_url: str, cell_address: str, headers: Dict) -> Dict:
        try:
            cell_url = f"{base_url}/range(address='{cell_address}')"
            
            response = requests.get(cell_url, headers=headers)
            if response.status_code != 200:
                return {"address": cell_address, "value": "", "error": response.text}
            
            cell_data = response.json()
            if not isinstance(cell_data, dict):
                return {"address": cell_address, "value": "", "error": "Invalid response format"}
            
            format_data = self._get_cell_format_data(cell_url, headers)
            
            # Safely extract values with proper error handling
            values = cell_data.get("values", [[""]])
            value = values[0][0] if values and len(values) > 0 and len(values[0]) > 0 else ""
            
            formulas = cell_data.get("formulas", [[""]])
            formula = formulas[0][0] if formulas and len(formulas) > 0 and len(formulas[0]) > 0 else ""
            
            text = cell_data.get("text", [[""]])
            display_value = text[0][0] if text and len(text) > 0 and len(text[0]) > 0 else ""
            
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
                "number_format": format_data.get("number_format", ""),
                "column_width": format_data.get("column_width", 0),
                "row_height": format_data.get("row_height", 0)
            }
        except Exception as e:
            print(f"DEBUG: Error in _get_cell_complete_data for {cell_address}: {e}")
            return {"address": cell_address, "value": "", "error": str(e)}
    
    def _get_cell_format_data(self, cell_url: str, headers: Dict) -> Dict:
        try:
            format_url = f"{cell_url}/format"
            format_response = requests.get(format_url, headers=headers)
            if format_response.status_code != 200:
                return {}
            
            format_data = format_response.json()
            if not isinstance(format_data, dict):
                return {}
            
            print(f"DEBUG: Format data: {format_data}")
            # Safely get font data
            font_data = {}
            try:
                font_response = requests.get(f"{format_url}/font", headers=headers)
                if font_response.status_code == 200:
                    font_data = font_response.json()
                    if not isinstance(font_data, dict):
                        font_data = {}
            except:
                font_data = {}
            
            print(f"DEBUG: font data: {font_data}")
            # Safely get fill data
            fill_data = {}
            try:
                fill_response = requests.get(f"{format_url}/fill", headers=headers)
                if fill_response.status_code == 200:
                    fill_data = fill_response.json()
                    if not isinstance(fill_data, dict):
                        fill_data = {}
            except:
                fill_data = {}
            
            print(f"DEBUG: fill data: {fill_data}")
            
            # Safely get borders data
            borders_data = {}
            try:
                borders_response = requests.get(f"{format_url}/borders", headers=headers)
                if borders_response.status_code == 200:
                    borders_data = borders_response.json()
                    if not isinstance(borders_data, dict):
                        borders_data = {}
            except:
                borders_data = {}
            
            print(f"DEBUG: borders data: {borders_data}")
            return {
                "number_format": format_data.get("numberFormat", {}).get("format", "") if isinstance(format_data.get("numberFormat"), dict) else "",
                "font": {
                    "name": font_data.get("name", ""),
                    "size": font_data.get("size", 0),
                    "bold": font_data.get("bold", False),
                    "italic": font_data.get("italic", False),
                    "underline": font_data.get("underline", ""),
                    "color": font_data.get("color", "")
                },
                "fill": {
                    "color": fill_data.get("color", "") if isinstance(fill_data.get("color"), str) else fill_data.get("color", {}).get("index", "") if isinstance(fill_data.get("color"), dict) else "",
                    "pattern_color": fill_data.get("patternColor", ""),
                    "pattern_type": fill_data.get("patternType", "")
                },
                "alignment": {
                    "horizontal": format_data.get("horizontalAlignment", ""),
                    "vertical": format_data.get("verticalAlignment", ""),
                    "wrap_text": format_data.get("wrapText", False),
                    "indent": format_data.get("indentLevel", 0)
                },
                "borders": self._extract_border_info(borders_data),
                "column_width": format_data.get("columnWidth", 0),
                "row_height": format_data.get("rowHeight", 0)
            }
        except Exception as e:
            print(f"DEBUG: Error in _get_cell_format_data: {e}")
            return {}
    
    def _get_merged_cells(self, base_url: str, headers: Dict) -> List[Dict]:
        merged_url = f"{base_url}/mergedCells"
        response = requests.get(merged_url, headers=headers)
        
        if response.status_code == 200:
            merged_data = response.json().get("value", [])
            return [{
                "address": cell.get("address", ""),
                "range": cell.get("address", "").split("!")[-1] if "!" in cell.get("address", "") else cell.get("address", "")
            } for cell in merged_data]
        return []
    
    def _extract_border_info(self, borders_data: Dict) -> Dict:
        if not borders_data or "value" not in borders_data:
            return {}
        
        borders = {}
        for border in borders_data.get("value", []):
            side = border.get("sideIndex", "")
            color_value = border.get("color", "")
            borders[side] = {
                "style": border.get("style", ""),
                "color": color_value if isinstance(color_value, str) else color_value.get("index", "") if isinstance(color_value, dict) else "",
                "weight": border.get("weight", "")
            }
        return borders
    
    def _search_and_get_worksheets(self, site_id: str, filename: str, headers: Dict) -> List[Dict]:
        """Search for file by name and get worksheets"""
        print(f"DEBUG: Searching for file: {filename}")
        
        # Search for the file in the site
        search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='{filename}')"
        search_response = requests.get(search_url, headers=headers)
        
        if search_response.status_code == 200:
            search_results = search_response.json().get("value", [])
            print(f"DEBUG: Found {len(search_results)} files")
            
            for file_item in search_results:
                if file_item.get("name", "").lower() == filename.lower():
                    file_id = file_item["id"]
                    print(f"DEBUG: Found file with ID: {file_id}")
                    
                    # Try to get worksheets using file ID
                    worksheets_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}/workbook/worksheets"
                    worksheets_response = requests.get(worksheets_url, headers=headers)
                    
                    if worksheets_response.status_code == 200:
                        return worksheets_response.json().get("value", [])
                    else:
                        print(f"DEBUG: Failed to get worksheets: {worksheets_response.text}")
        
        raise Exception(f"File '{filename}' not found or not accessible")
    
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