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
    name_list = get_name_list()[0]
    # 依次返回 常绿、落叶、灌木、地被类 四个列表
    changlv_name_list = get_name_list()[1]
    luoye_name_list = get_name_list()[2]
    guanmu_name_list = get_name_list()[3]
    dibeilei_name_list = get_name_list()[4]

    fig_html = "暂无数据"
    return template.TemplateResponse("index_2.html", {"request": request, "name_list": name_list, "fig_html": fig_html,
                                                    "changlv_name_list": changlv_name_list, "luoye_name_list": luoye_name_list,
                                                    "guanmu_name_list": guanmu_name_list, "dibeilei_name_list": dibeilei_name_list})


@app.post("/get_chart", response_class=HTMLResponse)
async def get_chart(request: Request, name: str = Form(...)):
    from utils.get_data import draw
    from utils.get_data import get_data
    start_time = time.time()
    fig = draw(get_data(name)[0], get_data(name)[1], name)
    fig_html = fig.to_html()
    fig_json = fig.to_dict()
    end_time = time.time()
    print(f"{name}已绘制完成，绘图时间：", round(end_time - start_time, 3), "s")
    return fig_json


@app.post("/get_chart_test", response_class=HTMLResponse)
async def get_chart(request: Request):
    data = await request.json()
    text_content = data.get('text', '')  # 获取前端发送的'text'字段值
    print(text_content)

    # 导入draw和get_data函数
    from utils.get_data import draw
    from utils.get_data import get_data

    start_time = time.time()
    fig = draw(get_data(start_name=text_content)[0], get_data(start_name=text_content)[1], text_content)
    fig_json = fig.to_plotly_json()
    html_str = fig.to_html(include_plotlyjs='cdn')
    end_time = time.time()

    alert_text = f'{text_content}已完成绘制，绘图时间： {round(end_time - start_time, 3)}s'  # 使用模板字符串
    print(alert_text)

    from utils.get_data import get_name_list
    name_list = get_name_list()[0]
    # 依次返回 常绿、落叶、灌木、地被类 四个列表
    changlv_name_list = get_name_list()[1]
    luoye_name_list = get_name_list()[2]
    guanmu_name_list = get_name_list()[3]
    dibeilei_name_list = get_name_list()[4]

    return template.TemplateResponse("index_2_2.html", {"request": request,
                                                      "changlv_name_list": changlv_name_list,
                                                      "luoye_name_list": luoye_name_list,
                                                      "guanmu_name_list": guanmu_name_list,
                                                      "dibeilei_name_list": dibeilei_name_list,
                                                      "html_str": html_str})