from fastapi import APIRouter, HTTPException, status, Depends
from sqlalchemy.orm import Session
import requests
from app.schemas.metadata import WorksheetMetadata
from app.services.sharepoint_service import SharePointService
from app.services.metadata_service import MetadataService
from app.core.database import get_db

router = APIRouter(prefix="/public", tags=["public"])

@router.get("/cell-metadata", response_model=dict)
async def get_complete_cell_metadata_public(sharepoint_url: str, worksheet_name: str = None, db: Session = Depends(get_db)):
    """Get complete cell metadata including formatting, colors, merged cells - No authentication required"""
    try:
        sharepoint_service = SharePointService()
        metadata = sharepoint_service.get_worksheet_complete_metadata(sharepoint_url, worksheet_name)
        
        # Save metadata to database
        MetadataService.save_metadata(
            db=db,
            sharepoint_url=sharepoint_url,
            worksheet_name=metadata["worksheet_name"],
            metadata=metadata
        )
        
        return {
            "success": True,
            "message": "Cell metadata extracted and saved successfully",
            "data": metadata
        }
        
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Failed to get cell metadata: {str(e)}"
        )

@router.get("/saved-metadata")
async def get_saved_metadata(sharepoint_url: str, worksheet_name: str, db: Session = Depends(get_db)):
    """Get previously saved metadata from database"""
    try:
        saved_metadata = MetadataService.get_metadata(db, sharepoint_url, worksheet_name)
        
        if not saved_metadata:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Metadata not found for the specified worksheet"
            )
        
        return {
            "success": True,
            "data": saved_metadata.metadata_json,
            "saved_at": saved_metadata.created_at,
            "updated_at": saved_metadata.updated_at
        }
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Failed to get saved metadata: {str(e)}"
        )

@router.get("/test-token")
async def test_token():
    """Test if Microsoft Graph token works"""
    try:
        sharepoint_service = SharePointService()
        token = sharepoint_service.get_access_token()
        
        # Test with application endpoint instead of /me
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get("https://graph.microsoft.com/v1.0/sites", headers=headers)
        
        return {
            "success": True,
            "token_works": response.status_code == 200,
            "status_code": response.status_code,
            "response": response.json() if response.status_code == 200 else response.text[:500]
        }
        
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }

@router.get("/worksheets")
async def get_worksheets_public(sharepoint_url: str):
    """Get list of worksheets from SharePoint Excel file - No authentication required"""
    try:
        sharepoint_service = SharePointService()
        worksheets = sharepoint_service.get_workbook_worksheets(sharepoint_url)
        
        return {
            "success": True,
            "worksheets": [{"name": ws["name"], "id": ws["id"]} for ws in worksheets]
        }
        
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Failed to get worksheets: {str(e)}"
        )