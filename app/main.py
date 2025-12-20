from fastapi import FastAPI
from app.core.database import engine, Base
from app.routers import auth, forms, public

Base.metadata.create_all(bind=engine)

app = FastAPI(title="Data Entry Backend", version="1.0.0")

app.include_router(auth.router)
app.include_router(forms.router)
app.include_router(public.router)

@app.get("/")
def read_root():
    return {"message": "Data Entry Backend API"}