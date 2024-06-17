import pandas as pd
import uvicorn
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from processor import process_attendance
import os
import shutil
import logging
from starlette.background import BackgroundTasks

app = FastAPI()

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def cleanup_file(file_path: str):
    if os.path.exists(file_path):
        os.remove(file_path)
        logger.info("Cleaned up file: %s", file_path)

@app.post("/process-attendance/")
async def process_attendance_file(file: UploadFile = File(...), output_format: str = 'csv', background_tasks: BackgroundTasks = BackgroundTasks()):
    logger.info("Received file: %s", file.filename)
    
    if file.content_type not in ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel"]:
        logger.error("Invalid file format: %s", file.content_type)
        raise HTTPException(status_code=400, detail="Invalid file format. Please upload an Excel file.")

    try:
        # Save the uploaded file temporarily
        temp_file_path = "temp_file.xlsx"
        with open(temp_file_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
        
        # Process the attendance file
        processed_data = process_attendance(temp_file_path)
        
        # Save the processed data to a file
        output_file = f"processed_data.{output_format}"
        if output_format == 'csv':
            processed_data.to_csv(output_file, index=False)
        elif output_format == 'xlsx':
            processed_data.to_excel(output_file, index=False)
        else:
            logger.error("Invalid output format: %s", output_format)
            raise HTTPException(status_code=400, detail="Invalid output format. Please choose 'csv' or 'xlsx'.")

        # Schedule cleanup of files
        background_tasks.add_task(cleanup_file, temp_file_path)
        background_tasks.add_task(cleanup_file, output_file)

        logger.info("Processing completed successfully")
        return FileResponse(output_file, media_type='application/octet-stream', filename=output_file)

    except Exception as e:
        logger.exception("Error processing file: %s", e)
        raise HTTPException(status_code=500, detail="Internal Server Error")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)