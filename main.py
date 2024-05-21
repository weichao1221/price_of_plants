import base64
import json
import os
import shutil
import time
import openpyxl as xl
from openpyxl.utils import column_index_from_string
from fastapi import FastAPI, Request, WebSocket, HTTPException, Body
from typing import List
from fastapi import Form
from fastapi.responses import FileResponse
from fastapi.responses import HTMLResponse
from fastapi.responses import RedirectResponse
from fastapi.responses import Response
from fastapi.security import HTTPBasic
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
from docx.shared import Pt
from docx.oxml.ns import qn
from docx import Document
import datetime
import re
from pydantic import BaseModel
from collections import OrderedDict
import openpyxl

app = FastAPI()
templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")
app.add_middleware(SessionMiddleware, secret_key="your_secret_key")
security = HTTPBasic()
templates_week_report = Jinja2Templates(directory="templates/week_report")


def read_data(filename):
    filepath = os.path.join(os.path.dirname(__file__), 'static', 'res')
    path = os.path.join(filepath, f'{filename}.txt')
    with open(path, 'r', encoding='utf-8') as f:
        result = f.read()
    result = eval(result)
    return result


def save_data(filename, data):
    filepath = os.path.join(os.path.dirname(__file__), 'static', 'res')
    path = os.path.join(filepath, f'{filename}.txt')
    with open(path, 'w', encoding='utf-8') as f:
        f.write(str(data))
        f.close()
    return path


@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/{page}.html", response_class=HTMLResponse)
async def page(request: Request, page: str):
    return templates.TemplateResponse(f'{page}.html', {"request": request})


'''整改以后的各种价格库'''


@app.get("/library", response_class=HTMLResponse)
async def library(request: Request):
    user_name = request.session.get('username')
    return templates.TemplateResponse("price_content.html", {"request": request})


def getLibraryPlantsNameList():
    file_path = os.path.join(os.path.dirname(__file__), 'resources', '苗木价格.xlsx')
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb['常规苗木 (2)']
    name_list = []

    for row in ws.iter_rows():
        if row[0].row < 4:
            continue
        if row[column_index_from_string("C") - 1].value is None:
            continue
        name_list.append(row[column_index_from_string("B") - 1].value)
        name_list = list(OrderedDict.fromkeys(name_list))
    file_name = '秀林苗木名称列表'
    file_path = save_data(filename=file_name, data=name_list)
    return file_path


def getLibraryPlantsData():
    file_path = os.path.join(os.path.dirname(__file__), 'resources', '苗木价格.xlsx')
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb['常规苗木 (2)']
    result_list = []
    for row in ws.iter_rows():
        if row[0].row < 4:
            continue
        if not row[1].value:
            continue
        a_dict = {
            "xuhao": row[column_index_from_string("A") - 1].value,
            "mingcheng": row[column_index_from_string("B") - 1].value,
            "zhonglei": row[column_index_from_string("C") - 1].value,
            "guige_gaodu": row[column_index_from_string("D") - 1].value,
            "guige_xiongdijing": row[column_index_from_string("E") - 1].value,
            "guige_guanfu": row[column_index_from_string("F") - 1].value,
            "guige_fenzhidian": row[column_index_from_string("G") - 1].value,
            "danwei": row[column_index_from_string("H") - 1].value,
            "daochangjia_buhanshui": round(row[column_index_from_string("J") - 1].value, 2)
        }
        result_list.append(a_dict)
    file_name = '秀林苗木价格字典'
    file_path = save_data(file_name, result_list)
    return file_path


@app.get("/search_data_index", response_class=HTMLResponse)  # 默认材料价格库显示界面
async def search_data_index(request: Request):
    try:
        name_list = read_data('秀林苗木名称列表')
    except:
        # 更新苗木名称列表
        getLibraryPlantsNameList()
        # 更新苗木库
        getLibraryPlantsData()
        name_list = read_data('秀林苗木名称列表')
    # ic(name_list)
    return templates.TemplateResponse("library.html", {"request": request, "name_list": name_list})


@app.get('/refresh_library_data', response_class=HTMLResponse)
async def refresh_library_data(request: Request):
    # 更新苗木名称列表
    getLibraryPlantsNameList()
    # 更新苗木库
    getLibraryPlantsData()
    content = '''
    <script>alert('已完成更新！');window.location.href='/search_data_index'</script>
    '''
    return HTMLResponse(content)


@app.post("/search_data", response_class=HTMLResponse)
async def search_data(request: Request, name: str = Form(...)):
    result = []
    price_data = read_data('秀林苗木价格字典')
    for value in price_data:
        if name.lower() in value['mingcheng'].lower():
            result.append(value)
    name_list = read_data('秀林苗木名称列表')

    if not result:
        return "<script>alert('未查询到数据！');window.location.href='search_data_index'</script>"
    else:
        return templates.TemplateResponse("library.html", {"request": request,
                                                           "name_list": name_list, "result": result})



@app.get("/jiagequshi", response_class=HTMLResponse)
async def root(request: Request):
    from utils.get_data import get_name_list
    # 依次返回 常绿、落叶、灌木、地被类 四个列表
    changlv_name_list = get_name_list()[1]
    luoye_name_list = get_name_list()[2]
    guanmu_name_list = get_name_list()[3]
    dibeilei_name_list = get_name_list()[4]
    fig_html = "暂无数据"
    return templates.TemplateResponse("qushi.html", {"request": request,  # request参数必须有
                                                     "changlv_name_list": changlv_name_list,  # 常绿列表
                                                     "luoye_name_list": luoye_name_list,  # 落叶列表
                                                     "guanmu_name_list": guanmu_name_list,  # 灌木列表
                                                     "dibeilei_name_list": dibeilei_name_list})  # 地被类列表


# 后台手动刷新数据
@app.get("/f", response_class=HTMLResponse)  # 用于前端ajax请求
async def refresh_data(request: Request):
    from utils.get_data import get_data
    from utils.get_data import get_name_list
    from utils.get_data import draw
    name_list = get_name_list()[0]
    dir = "static/result/"
    try:
        os.makedirs(dir)
    except FileExistsError:
        pass
    for name in name_list:
        fig = draw(get_data(name)[0], get_data(name)[1], name)
        fig = fig.write_html(f"static/result/{name}.html")
    html_text = ("<script>"
                 "alert('数据刷新成功');"
                 "window.location.href='/';"
                 "</script>")
    return HTMLResponse(content=html_text, status_code=200)


@app.get('/xinxijia')
def xinxijia(request: Request):
    return templates.TemplateResponse("信息价查询.html", {"request": request})


@app.post("/xinxijia_check", response_class=HTMLResponse)
async def xinxijia_check(request: Request, sj: str = Form(default=datetime.datetime.now().strftime('%Y-%m-%d')), mc: str = Form(default="无"), gg: str = Form(default='无')):
    print(sj)
    shijian_year = sj.split('-')[0]
    shijian_month = sj.split('-')[1].replace('0', '')
    shijian = f'{shijian_year}年{shijian_month}月'
    result = []
    starttime = datetime.datetime.now()
    price_data = read_data('雄安新区信息价')
    for value in price_data:
        if mc.lower() in value['mc'].lower() and shijian == value['sj'] and gg in value['gg']:
            result.append(value)
    endtime = datetime.datetime.now()
    haoshi = (endtime - starttime).seconds
    if not result:
        result = [{
            'sj': '没',
            'mc': '找',
            'gg': '到，',
            'dw': '对',
            'p': '不',
            'p_tax': '起！',
        }]
    return templates.TemplateResponse("信息价查询.html", {"request": request, "datelist": result, "haoshi": haoshi})

@app.get("/swift_get_data")
async def swift_get_data(request: Request):
    result = []
    price_data = read_data(filename='秀林苗木价格字典')
    for value in price_data:
        mc = str(value['mingcheng'])
        zl = str(value['zhonglei'])
        gd = str(value['guige_gaodu'])
        xdj = str(value['guige_xiongdijing'])
        gg = str(value['guige_guanfu'])
        fzd = str(value['guige_fenzhidian'])
        dw = str(value['danwei'])
        jg = str(value['daochangjia_buhanshui'])
        dataList = [mc, zl, gd, xdj, gg, fzd, dw, jg]
        result.append(dataList)
    return result




if __name__ == '__main__':
    import uvicorn

    uvicorn.run(app='main:app', host="0.0.0.0", port=8000, reload=True)
