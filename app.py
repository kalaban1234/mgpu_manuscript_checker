from fastapi import FastAPI, Request, UploadFile, File
from fastapi.templating import Jinja2Templates
from checker import check_docx, group_report

app = FastAPI()
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=None)
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})
@app.post("/check")
async def check(request: Request, file: UploadFile = File(...)):
    file_bytes = await file.read()
    report = check_docx(file_bytes)
    grouped = group_report(report)
    has_errors = bool(report)  # True, если есть ошибки/варнинги
    return templates.TemplateResponse(
        "result.html",
        {"request": request, "report": grouped, "has_errors": has_errors}
    )