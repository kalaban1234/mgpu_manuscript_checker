from fastapi import FastAPI, File, UploadFile, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from checker import check_docx

app = FastAPI()
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/check", response_class=HTMLResponse)
async def check(request: Request, file: UploadFile = File(...)):
    content = await file.read()
    report = check_docx(content)
    return templates.TemplateResponse("result.html", {"request": request, "report": report})
