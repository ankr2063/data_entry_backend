from sqlalchemy import Column, Integer, String, DateTime, Text, JSON
from sqlalchemy.sql import func
from app.core.database import Base

class WorksheetMetadata(Base):
    __tablename__ = "worksheet_metadata"
    
    id = Column(Integer, primary_key=True, index=True)
    sharepoint_url = Column(String(500), nullable=False)
    worksheet_name = Column(String(100), nullable=False)
    metadata_json = Column(JSON, nullable=False)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())