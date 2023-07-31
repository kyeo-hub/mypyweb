from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from fastapi.requests import Request
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pywebio.platform.fastapi import webio_routes
import markdown2
from jc import jc
from bmi import main as bmi
from markdown_previewer import main as markdown_previewer
from chat_room import main as chat_room
from gomoku_game import main as gomoku_game

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def read_main(request: Request):
   with open("index.md","r",encoding='utf-8') as f:
      markdown_text= f.read()
   html = markdown2.markdown(markdown_text)   
   return templates.TemplateResponse("index.html", {"request": request,"html_content":html})
# `task_func` is PyWebIO task function
app.mount("/bmi", FastAPI(routes=webio_routes(bmi)))
app.mount("/jc", FastAPI(routes=webio_routes(jc)))
app.mount("/markdown_previewer", FastAPI(routes=webio_routes(markdown_previewer)))
app.mount("/chat_room", FastAPI(routes=webio_routes(chat_room)))
app.mount("/gomoku_game", FastAPI(routes=webio_routes(gomoku_game)))
# app.mount("/jc", FastAPI(routes=webio_routes(jc)))