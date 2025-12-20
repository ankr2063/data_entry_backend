from sqlalchemy.orm import Session
from app.models.worksheet_metadata import WorksheetMetadata
from typing import Dict, Optional

class MetadataService:
    @staticmethod
    def save_metadata(db: Session, sharepoint_url: str, worksheet_name: str, metadata: Dict) -> WorksheetMetadata:
        # Check if metadata already exists
        existing = db.query(WorksheetMetadata).filter(
            WorksheetMetadata.sharepoint_url == sharepoint_url,
            WorksheetMetadata.worksheet_name == worksheet_name
        ).first()
        
        if existing:
            existing.metadata_json = metadata
            db.commit()
            db.refresh(existing)
            return existing
        else:
            db_metadata = WorksheetMetadata(
                sharepoint_url=sharepoint_url,
                worksheet_name=worksheet_name,
                metadata_json=metadata
            )
            db.add(db_metadata)
            db.commit()
            db.refresh(db_metadata)
            return db_metadata
    
    @staticmethod
    def get_metadata(db: Session, sharepoint_url: str, worksheet_name: str) -> Optional[WorksheetMetadata]:
        return db.query(WorksheetMetadata).filter(
            WorksheetMetadata.sharepoint_url == sharepoint_url,
            WorksheetMetadata.worksheet_name == worksheet_name
        ).first()