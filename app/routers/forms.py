from fastapi import APIRouter, HTTPException, status
from app.schemas.form import SharePointRequest, FormSchema, FormSubmission
from app.services.sharepoint_service import SharePointService
from app.services.form_generator_service import FormGeneratorService

router = APIRouter(prefix="/forms", tags=["forms"])

@router.post("/generate-schema", response_model=FormSchema)
async def generate_form_schema(request: SharePointRequest):
    """Generate form schema from SharePoint Excel file"""
    try:
        sharepoint_service = SharePointService()
        
        # Get main sheet data
        main_sheet_data = sharepoint_service.get_worksheet_data(
            str(request.sharepoint_url), 
            request.main_sheet_name
        )
        
        # Get config sheet data if specified
        config_sheet_data = None
        if request.config_sheet_name:
            try:
                config_sheet_data = sharepoint_service.get_worksheet_data(
                    str(request.sharepoint_url), 
                    request.config_sheet_name
                )
            except Exception:
                # Config sheet is optional, continue without it
                pass
        
        # Generate form metadata
        form_metadata = FormGeneratorService.extract_form_metadata(
            main_sheet_data, 
            config_sheet_data
        )
        
        return FormSchema(**form_metadata)
        
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Failed to generate form schema: {str(e)}"
        )

@router.post("/submit")
async def submit_form_data(submission: FormSubmission):
    """Submit form data back to SharePoint Excel file"""
    try:
        sharepoint_service = SharePointService()
        
        # Here you would implement the logic to write data back to SharePoint
        # This involves mapping form fields back to Excel cells and updating them
        
        # For now, return success response
        return {
            "message": "Form data submitted successfully",
            "submitted_fields": len(submission.form_data)
        }
        
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Failed to submit form data: {str(e)}"
        )

@router.get("/cell-metadata")
async def get_complete_cell_metadata(sharepoint_url: str, worksheet_name: str = None):
    """Get complete cell metadata including formatting, colors, merged cells"""
    try:
        sharepoint_service = SharePointService()
        metadata = sharepoint_service.get_worksheet_complete_metadata(sharepoint_url, worksheet_name)
        
        return {
            "message": "Cell metadata extracted successfully",
            "data": metadata
        }
        
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Failed to get cell metadata: {str(e)}"
        )
async def get_worksheets(sharepoint_url: str):
    """Get list of worksheets from SharePoint Excel file"""
    try:
        sharepoint_service = SharePointService()
        worksheets = sharepoint_service.get_workbook_worksheets(sharepoint_url)
        
        return {
            "worksheets": [{"name": ws["name"], "id": ws["id"]} for ws in worksheets]
        }
        
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Failed to get worksheets: {str(e)}"
        )