from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv
from api import router as api_router
import uvicorn
import ssl
import os

load_dotenv()

app = FastAPI()

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)

app.include_router(api_router, prefix="/api")

@app.get("/")
async def root():
    return {"message": "Python server up and running!"}

if __name__ == "__main__":
    cert_file = "./localhost+2.pem"
    key_file = "./localhost+2-key.pem"
    
    if os.path.exists(cert_file) and os.path.exists(key_file):
        uvicorn.run(
            "main:app", 
            host="0.0.0.0", 
            port=52525, 
            reload=True,
            ssl_keyfile=key_file,
            ssl_certfile=cert_file
        )
    else:
        # http fallback
        print("SSL certificates not found. Running without HTTPS.")
        print(f"Looking for certificates: {cert_file}, {key_file}")
        uvicorn.run("main:app", host="0.0.0.0", port=52525, reload=True) 