import time

import psutil
import xlrd
import xlsxwriter
import matplotlib.pyplot as plt
from matplotlib.ticker import IndexLocator
import pylab as pl


# sheet_index: excel表中sheet索引
# col_index: 需要返回列索引
def get_col_data(fname, sheet_index, col_index):
    data = xlrd.open_workbook(fname)
    table = data.sheets()[sheet_index]

    return table.col_values(col_index)


# gap: x轴刻度间距
def draw_line_chart(x, y, title, gap, fname):
    plt.title(title)
    # 返回当前的Axes对象
    ax = plt.gca()
    # 设置边框线和水平横线
    ax.spines['left'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y')
    # 设置x/y轴刻度
    x_locator = IndexLocator(gap, 0)
    ax.xaxis.set_major_locator(x_locator)
    ax.set_xlim(0, x[-1])
    ax.set_ylim(0, 80)
    plt.plot(x, y)
    pl.xticks(rotation=90)
    # 保存
    plt.savefig(fname)
    plt.close('all')


# 计算标准差
def calc_stdev(arr, average):
    tmp = sum((i - average) ** 2 for i in arr) / len(arr)
    return tmp ** 0.5


# 获取计算机cpu使用率
def get_cpu_trend(total_time, interval):
    begin = time.time()
    x, y = [], []

    while total_time >= 0:
        time.sleep(interval)

        now = time.time() - begin
        x.append(now)
        y.append(psutil.cpu_percent())
        total_time -= interval

    draw_line_chart(x, y, 'CPU', 1, 'test.png')
    write_xlsx(y, 'test.xlsx', 'test.png')


def write_xlsx(y, xlsx_name, img_name):
    average = sum(y) / len(y)
    stdev = calc_stdev(y, average)
    ratio = len([i for i in y if i < 45]) / len(y)

    workbook = xlsxwriter.Workbook(xlsx_name)
    ws = workbook.add_worksheet('cpu趋势')
    titles = ['平均值', '标准差', '低于45的值比率']
    results = [average, stdev, ratio]
    for i in range(len(titles)):
        ws.write(0, i, titles[i])
        ws.write(1, i, results[i])

    ws.insert_image(5, 5, img_name)
    workbook.close()


x = get_col_data('cpu.xls', 0, 0)[1:]
y = get_col_data('cpu.xls', 0, 1)[1:]
draw_line_chart(x, y, 'CPU', 25, 'trend.png')
write_xlsx(y, 'result.xlsx', 'trend.png')

get_cpu_trend(20, 0.2)
