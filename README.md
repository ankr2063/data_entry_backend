# Data Entry Backend

A modular Python backend with FastAPI, MySQL, JWT authentication, and SharePoint Excel integration for dynamic form generation.

## Features

- **Dynamic Form Generation**: Create forms from SharePoint Excel templates
- **SharePoint Integration**: Read/write Excel files via Microsoft Graph API
- **Configurable Validation**: Define field types and validation rules
- **JWT Authentication**: Secure user authentication
- **Modular Architecture**: Clean separation of concerns

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Configure environment variables in `.env`:
- Update `DATABASE_URL` with your MySQL credentials
- Change `SECRET_KEY` to a secure random string
- Add Microsoft Graph API credentials

3. Create MySQL database:
```sql
CREATE DATABASE data_entry_db;
```

4. Run the application:
```bash
python run.py
```

## API Endpoints

### Authentication
- `POST /auth/register` - Register new user
- `POST /auth/login` - Login user

### Forms
- `POST /forms/generate-schema` - Generate form schema from SharePoint Excel
- `POST /forms/submit` - Submit form data back to SharePoint
- `GET /forms/worksheets` - Get available worksheets

### Health
- `GET /` - Health check

## SharePoint Integration

### 1. Main Sheet Structure
Your Excel sheet should follow this pattern:
- Header row with field names
- Data rows with sample/default values
- Section headers (optional) for grouping fields

### 2. Configuration Sheet (Optional)
Create a "Config" worksheet to define:
- Field types (text, number, date, select, etc.)
- Validation rules (min/max, required, patterns)
- Dropdown options
- Placeholders

### 3. Usage Example
```json
{
  "sharepoint_url": "https://company.sharepoint.com/sites/site/Shared%20Documents/form.xlsx",
  "main_sheet_name": "Sheet1",
  "config_sheet_name": "Config"
}
```

## Project Structure

```
app/
├── core/           # Configuration and utilities
├── models/         # Database models
├── schemas/        # Pydantic schemas
├── services/       # Business logic
│   ├── sharepoint_service.py    # SharePoint integration
│   ├── form_generator_service.py # Form schema generation
│   └── user_service.py          # User management
├── routers/        # API routes
└── main.py         # FastAPI application
```