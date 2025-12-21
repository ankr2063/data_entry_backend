# Django Data Entry Backend

A modular Django REST API backend with MySQL database and JWT authentication.

## Project Structure

```
apps/
├── authentication/     # User registration and login
├── organizations/      # Organization management
├── users/             # User management
└── forms/             # Forms and form data management
```

## Setup Instructions

1. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Database Configuration**
   - Copy `.env.example` to `.env`
   - Update database credentials in `.env` file
   - Create MySQL database: `CREATE DATABASE data_entry_db;`

3. **Run Migrations**
   ```bash
   python manage.py makemigrations
   python manage.py migrate
   ```

4. **Create Superuser**
   ```bash
   python manage.py createsuperuser
   ```

5. **Run Server**
   ```bash
   python manage.py runserver
   ```

## API Endpoints

### Authentication (Users)
- `POST /api/users/register/` - User registration
- `POST /api/users/login/` - User login

### Organizations (CRUD)
- `GET /api/organizations/` - List organizations
- `POST /api/organizations/` - Create organization
- `GET /api/organizations/{id}/` - Get organization
- `PUT /api/organizations/{id}/` - Update organization
- `DELETE /api/organizations/{id}/` - Delete organization

## Authentication

The API uses JWT (JSON Web Tokens) for authentication. After login/register, include the access token in the Authorization header:

```
Authorization: Bearer <access_token>
```

## Models

1. **Organizations** - Organization management with CRUD operations
2. **Users** - User management with authentication
3. **Forms** - Form definitions and related data
4. **UserFormAccess** - User permissions for forms
5. **FormEntryVersions** - Form template versions
6. **FormDisplayVersions** - Form display configurations
7. **FormData** - Form submission data
8. **FormDataHistory** - Form data change history