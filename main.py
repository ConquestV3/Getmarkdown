import os
import shutil
import uuid
import zipfile
from typing import List
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from converters import convert_to_markdown

app = FastAPI(title="GetMarkdown Local Converter")

UPLOAD_DIR = os.path.abspath("./temp_uploads")
DOWNLOAD_DIR = os.path.abspath("./temp_downloads")

# Cleanup directories on start
@app.on_event("startup")
def startup_event():
    for d in [UPLOAD_DIR, DOWNLOAD_DIR]:
        if os.path.exists(d):
            shutil.rmtree(d, ignore_errors=True)
        os.makedirs(d, exist_ok=True)

# Helper to clean up source files after zip is created
def cleanup_temp_dir(path: str):
    if os.path.exists(path):
        shutil.rmtree(path, ignore_errors=True)

@app.post("/api/convert")
async def convert_file(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    extract_images: bool = Form(True),
    extract_tables: bool = Form(True)
):
    # Create unique session
    session_id = str(uuid.uuid4())
    session_upload_dir = os.path.join(UPLOAD_DIR, session_id)
    session_output_dir = os.path.join(session_upload_dir, "output")
    
    os.makedirs(session_upload_dir, exist_ok=True)
    os.makedirs(session_output_dir, exist_ok=True)
    
    # Save uploaded file
    file_path = os.path.join(session_upload_dir, file.filename)
    try:
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
    except Exception as e:
        cleanup_temp_dir(session_upload_dir)
        return JSONResponse(status_code=500, content={"error": f"Failed to save file: {str(e)}"})
        
    # Run conversion
    try:
        markdown_text = convert_to_markdown(file_path, session_output_dir, extract_images, extract_tables)
    except Exception as e:
        cleanup_temp_dir(session_upload_dir)
        return JSONResponse(status_code=500, content={"error": f"Conversion error: {str(e)}"})
        
    # Write markdown to file in output directory
    base_name, _ = os.path.splitext(file.filename)
    md_filename = f"{base_name}.md"
    md_filepath = os.path.join(session_output_dir, md_filename)
    
    try:
        with open(md_filepath, "w", encoding="utf-8") as f:
            f.write(markdown_text)
    except Exception as e:
        cleanup_temp_dir(session_upload_dir)
        return JSONResponse(status_code=500, content={"error": f"Failed to write markdown: {str(e)}"})
        
    # Check if images are extracted
    images_dir = os.path.join(session_output_dir, "images")
    has_images = os.path.exists(images_dir) and len(os.listdir(images_dir)) > 0
    zip_id = None
    
    # Create ZIP if there are images
    if has_images:
        zip_id = f"{session_id}.zip"
        zip_path = os.path.join(DOWNLOAD_DIR, zip_id)
        
        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Add markdown file
                zipf.write(md_filepath, md_filename)
                # Add images folder contents
                for root, _, files in os.walk(images_dir):
                    for img_file in files:
                        full_img_path = os.path.join(root, img_file)
                        rel_img_path = os.path.join("images", img_file)
                        zipf.write(full_img_path, rel_img_path)
        except Exception as e:
            cleanup_temp_dir(session_upload_dir)
            return JSONResponse(status_code=500, content={"error": f"Failed to create ZIP package: {str(e)}"})
            
    # Clean up uploads folder, keep download zip
    background_tasks.add_task(cleanup_temp_dir, session_upload_dir)
    
    return {
        "filename": file.filename,
        "markdown": markdown_text,
        "zip_id": zip_id,
        "has_images": has_images
    }

@app.get("/api/download/{zip_id}")
async def download_file(zip_id: str):
    zip_path = os.path.join(DOWNLOAD_DIR, zip_id)
    if not os.path.exists(zip_path):
        raise HTTPException(status_code=404, detail="File not found or link expired")
        
    # Find matching filename context if possible, or use default
    return FileResponse(
        zip_path,
        media_type="application/zip",
        filename=f"converted_markdown_assets_{zip_id[:8]}.zip"
    )

# Serve static files for frontend UI
# Mount this last so it doesn't intercept API endpoints
app.mount("/", StaticFiles(directory="static", html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
