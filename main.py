from fastapi import FastAPI
from fastapi.responses import HTMLResponse, JSONResponse, PlainTextResponse
from fastapi.requests import Request
from fastapi.templating import Jinja2Templates
from fastapi import Form
import time
from fastapi.staticfiles import StaticFiles

app = FastAPI()
template = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    from utils.get_data import get_name_list
    # 依次返回 常绿、落叶、灌木、地被类 四个列表
    changlv_name_list = get_name_list()[1]
    luoye_name_list = get_name_list()[2]
    guanmu_name_list = get_name_list()[3]
    dibeilei_name_list = get_name_list()[4]
    fig_html = "暂无数据"
    return template.TemplateResponse("index.html", {"request": request,    # request参数必须有
                                                    "changlv_name_list": changlv_name_list,     # 常绿列表
                                                    "luoye_name_list": luoye_name_list,        # 落叶列表
                                                    "guanmu_name_list": guanmu_name_list,   # 灌木列表
                                                    "dibeilei_name_list": dibeilei_name_list})  # 地被类列表


# 后台手动刷新数据
@app.get("/f", response_class=PlainTextResponse)  # 用于前端ajax请求
async def refresh_data(request: Request):
    from utils.get_data import get_data
    from utils.get_data import get_name_list
    from utils.get_data import draw
    name_list = get_name_list()[0]
    for name in name_list:
        fig = draw(get_data(name)[0], get_data(name)[1], name)
        fig = fig.write_html(f"static/result/{name}.html")
    return "数据已刷新"



if __name__ == '__main__':
    import uvicorn

    uvicorn.run(app='main:app', host="127.0.0.1", port=8080, reload=True)