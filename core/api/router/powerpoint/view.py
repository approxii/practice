import logging
import json
from multiprocessing import Process, Queue
from fastapi import APIRouter, Request, Form, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
import requests

router = APIRouter(prefix="/powerpoint", tags=["PowerPoint"])
templates = Jinja2Templates(directory="core/api/router/powerpoint/templates")

EXECUTOR_URL = "http://your_ip:5000"
API_KEY = "super-secret-key"
OLLAMA_URL = "http://your_ip/api/generate"
OLLAMA_MODEL = "model"
MAX_ATTEMPTS = 3 # количество попыток генерации VBA кода

# Функции генерации
def build_vba_prompt(user_prompt: str, errors: str = "") -> str:
    err_text = f"\n=== Previous Errors ===\n{errors}" if errors else ""
    return f"""
You are an expert VBA/PowerPoint macro developer. Follow this EXACT structure:

Sub ProcedureName()
    On Error GoTo ErrHandler
    
    ' Your VBA code here
    ' IMPORTANT: Use ActivePresentation for PowerPoint operations
    ' IMPORTANT: Declare variables properly
    ' IMPORTANT: Your main code goes here
    
    ppt.SaveAs "C:\\app\\res.pptm"
    Exit Sub

ErrHandler:
    On Error Resume Next
    Debug.Print "Ошибка: " & Err.Number & " - " & Err.Description
    Resume Next
End Sub

CRITICAL RULES:
1. You MUST include "ppt.SaveAs" line before "Exit Sub"
2. You MUST include "Exit Sub" before "ErrHandler:"  
3. Use proper PowerPoint objects like ActivePresentation
4. Create new slides and content based on user request

User request: {user_prompt}
{err_text}

Produce ONLY VBA code following the exact structure above. No explanations, no comments.
"""

def validate_vba_structure(vba_code):
    """Проверяет наличие обязательных элементов в VBA коде"""
    required_elements = [
        "On Error GoTo ErrHandler",
        "ppt.SaveAs",
        "Exit Sub", 
        "ErrHandler:",
        "End Sub"
    ]
    
    return all(element in vba_code for element in required_elements)



def generate_vba_with_ollama(prompt):
    payload = {"model": OLLAMA_MODEL, "prompt": prompt}
    response = requests.post(OLLAMA_URL, json=payload, stream=True, timeout=180)
    response.raise_for_status()
    
    vba_code = ""
    for line in response.iter_lines():
        if not line:
            continue
        try:
            data = json.loads(line.decode("utf-8"))
            if "response" in data:
                vba_code += data["response"]
        except json.JSONDecodeError:
            logging.warning("Не удалось распарсить строку Ollama: %s", line)
    
    return vba_code.strip()


def run_win_server(vba_code, queue):
    """Отправка VBA на Windows в отдельном процессе"""
    try:
        resp = requests.post(
            EXECUTOR_URL + "/run",
            json={"code": vba_code},
            headers={"Authorization": f"Bearer {API_KEY}"},
            timeout=300,
        )
        
        if resp.status_code != 200:
            queue.put({"errors": [f"HTTP {resp.status_code}: {resp.text}"], "slides": []})
            return
        
        queue.put(resp.json())
    except Exception as e:
        queue.put({"errors": [str(e)], "slides": []})


@router.get("/", response_class=HTMLResponse)
def show_form(request: Request):
    return templates.TemplateResponse(
        "powerpoint.html",
        {"request": request, "slides": None, "download_url": None, "errors": None},
    )


@router.post("/", response_class=HTMLResponse)
def run_vba(request: Request, user_text: str = Form(...)):
    try:
        errors = ""
        attempts = 0
        final_data = {"slides": [], "download_url": None, "errors": []}

        while attempts < MAX_ATTEMPTS:
            attempts += 1
            prompt = build_vba_prompt(user_text, errors)
            vba_code = generate_vba_with_ollama(prompt)

    
            # logging.info(
            #     "\n====== СФОРМИРОВАННЫЙ VBA-КОД (попытка #%d) ======\n%s\n%s\n",
            #     attempts,
            #     vba_code[:2000],
            #     "..." if len(vba_code) > 2000 else "",
            # )
            # print("\nИтоговый VBA-код (попытка #%d):" % attempts)
            # print(vba_code[:2000])
            # if len(vba_code) > 2000:
            #     print("... (обрезано, длина кода:", len(vba_code), ")")

        
            # if not validate_vba_structure(vba_code):
            #     logging.warning("VBA-код не соответствует шаблону — отправка во Flask пропущена.")
            #     errors = "Структура VBA не соответствует шаблону. Попробуйте снова."
            #     continue  # попробуем следующую попытку генерации


            # Только если структура валидна — отправляем
            queue = Queue()
            p = Process(target=run_win_server, args=(vba_code, queue))
            p.start()
            p.join(timeout=310)

            if p.is_alive():
                p.terminate()
                errors = "Windows процесс завис"
            else:
                final_data = queue.get()
                errors = "\n".join(final_data.get("errors", []))
                if not errors:
                    break

            logging.info("Обнаружены ошибки, попытка исправления #%d: %s", attempts, errors[:200])

        slides = final_data.get("slides", [])
        download_url = final_data.get("download_url")
        errors_list = final_data.get("errors", [])

        return templates.TemplateResponse(
            "powerpoint.html",
            {
                "request": request,
                "slides": slides,
                "download_url": download_url,
                "errors": errors_list,
            },
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
