import tensorflow as tf
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from utils.get_data import get_data
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit

def power_function(x, a, b):
    return a * np.power(x, b)

def exponential_function(x, a, b):
    return a * np.exp(b * x)


def draw(x, y):
    # 拟合数据
    params, covariance = curve_fit(exponential_function, x, y)

    # 获取拟合后的参数
    a_fit, b_fit = params

    # 获取协方差矩阵对角元素，这些元素代表参数的方差
    variance_a, variance_b = np.diag(covariance)

    # 创建用于绘图的 x 值
    x_fit = np.linspace(min(x), max(x), 100)

    # 计算拟合后的 y 值
    y_fit = exponential_function(x_fit, a_fit, b_fit)

    # 创建散点图
    scatter = go.Scatter(x=x, y=y, mode='markers', name='Data')

    # 创建拟合曲线
    fit_curve = go.Scatter(x=x_fit, y=y_fit, mode='lines',
                           name=f'Fit: y = {a_fit:.2f} * e^({b_fit:.2f} * x), Var(a) = {variance_a:.2f}, Var(b) = {variance_b:.2f}')

    # 创建图布局
    layout = go.Layout(
        title='Data and Exponential Curve Fit',
        xaxis=dict(title='x'),
        yaxis=dict(title='y')
    )

    # 绘制图形
    fig = go.Figure(data=[scatter, fit_curve], layout=layout)
    return fig  # 返回绘制的图形



if __name__ == '__main__':
    start_name = input("苗木名称：")
    x = get_data(start_name)[0]
    y = get_data(start_name)[1]
    print(x, y)
    draw(x, y)