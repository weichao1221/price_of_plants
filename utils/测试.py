import plotly.express as px
import pandas as pd

# 创建示例数据帧
data = pd.DataFrame({
    'x': [1, 2, 3, 4, 5],
    'y1': [10, 11, 12, 13, 14],
    'y2': [15, 16, 17, 18, 19],
    'y3': [20, 21, 22, 23, 24]
})

# 创建图表
fig = px.line(data, x='x', y=['y1', 'y2', 'y3'], labels={'x': 'X轴', 'value': 'Y轴'})

# 默认仅显示第一种曲线
fig.data[1].visible = False
fig.data[2].visible = False

# 创建一个回调函数来处理图例名称的点击事件
# def handle_legend_click(trace, points, selector):
#     for name, curve in zip(fig.data, ['y1', 'y2', 'y3']):
#         if name.name in points.point_inds:
#             name.visible = True
#         else:
#             name.visible = False
#
# fig.for_each_trace(handle_legend_click)

# 显示图表
fig.show()
