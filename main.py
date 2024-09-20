import pandas as pd
import json
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import Response

from format import process

app = FastAPI()

@app.post("/format-service/excel")
async def format_file(file: UploadFile = File(...)):
    content = await file.read()
    data = process(content)
    return Response(content=data,media_type="application/octet-stream")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
