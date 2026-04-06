# coding: utf-8
# наше всё
import numpy as np
import pandas as pd

# Avoid GTK (pygobject_register_wrapper) and Tk (tkinter may be missing): use non-GUI Agg.
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
from pylab import rcParams

import plotly.express as px
import plotly.graph_objects as go


def tourn_creative_calc(file, is_plot=False):
    """
    На вход получает турнирную таблицу в формате excel
    Возвращает таблицу с мерой сложности вопросов по дистанции
    Рисуем график

    Параметры:
    file - путь к файлу с турнирной таблицей
    is_plot - флаг, рисовать ли график
    """

    # получаем файл
    source_df = pd.read_excel(file, sheet_name='Worksheet')
    
    # сложность вопросов
    res_df = source_df.copy()
    diff_df = res_df.copy()
    for i in range(1, len(diff_df.columns)-2):
        diff_df[i] = np.sum(diff_df[i]) / len(diff_df[i])

    # оценка потенциала команд
    pot_df = res_df.copy()
    for i in range(1, len(diff_df.columns)-2):
        pot_df[i] = diff_df[i] * res_df[i]

    diff_df = diff_df[0:1][diff_df.columns[3:]].T
    diff_df = diff_df.reset_index()
    diff_df.columns = ['question', 'difficulty']

    # турнирная таблица с нужными метриками
    team_stat = res_df[res_df.columns[0:3]].copy()
    team_stat['result'] = np.sum(res_df[res_df.columns[3:]], axis=1)
    team_stat['difficulty'] = np.sum(pot_df[pot_df.columns[3:]], axis=1)
    team_stat['potential'] = team_stat['difficulty'] / team_stat['result']

    # рисуем график matplotlib
    if is_plot:
        plt.figure(figsize=(8, 5))

        sns.scatterplot(
            data=team_stat, 
            x="potential", 
            y="result",
            )

        plt.title("Отбор №5:\n Распределение команд по стилю и результату")
        # add a trend line
        sns.regplot(data=team_stat, x="potential", y="result", scatter=False)

        # самая творческая
        x1 = np.min(team_stat['potential'])
        y1 = team_stat[team_stat['potential'] == x1]['result'].values[0]
        name1 = team_stat[team_stat['potential'] == x1]['Название'].values[0]
        plt.text(x1, y1, name1, fontsize=8, color='blue')

        # самая техничная
        x2 = np.max(team_stat['potential'])
        y2 = team_stat[team_stat['potential'] == x2]['result'].values[0]
        name2 = team_stat[team_stat['potential'] == x2]['Название'].values[0]
        plt.text(x2, y2, name2, fontsize=8, color='blue')

        # самая результативная
        y3 = np.max(team_stat['result'])
        x3 = team_stat[team_stat['result'] == y3]['potential'].values[0]
        name3 = team_stat[team_stat['result'] == y3]['Название'].values[0]
        plt.text(x3, y3, name3, fontsize=8, color='blue')

        # случайная команда
        x5 = team_stat['potential'].values[np.random.randint(0, len(team_stat))]
        y5 = team_stat[team_stat['potential'] == x5]['result'].values[0]
        name5 = team_stat[team_stat['potential'] == x5]['Название'].values[0]
        plt.text(x5, y5, name5, fontsize=8, color='blue')

        plt.xlabel("Творческие\техничные команды")
        plt.ylabel("Количество взятых")

        plt.savefig("creative_matplotlib.png", dpi=150, bbox_inches="tight")
        plt.close()

    return team_stat


def tourn_creative_calc_plotly(file, is_plot=False):
    """
    То же самое, только графики в plotly и экспорт в html
    """
    # получаем файл
    source_df = pd.read_excel(file, sheet_name='Worksheet')
    
    # сложность вопросов
    res_df = source_df.copy()
    diff_df = res_df.copy()
    for i in range(1, len(diff_df.columns)-2):
        diff_df[i] = np.sum(diff_df[i]) / len(diff_df[i])

    # оценка потенциала команд
    pot_df = res_df.copy()
    for i in range(1, len(diff_df.columns)-2):
        pot_df[i] = diff_df[i] * res_df[i]

    diff_df = diff_df[0:1][diff_df.columns[3:]].T
    diff_df = diff_df.reset_index()
    diff_df.columns = ['question', 'difficulty']

    # турнирная таблица с нужными метриками
    team_stat = res_df[res_df.columns[0:3]].copy()
    team_stat['result'] = np.sum(res_df[res_df.columns[3:]], axis=1)
    team_stat['difficulty'] = np.sum(pot_df[pot_df.columns[3:]], axis=1)
    team_stat['potential'] = team_stat['difficulty'] / team_stat['result']

    # рисуем график plotly
    if is_plot:
        fig = px.scatter(
            data_frame=team_stat,
            x="potential",
            y="result",
        )
        x = team_stat["potential"].to_numpy()
        y = team_stat["result"].to_numpy()
        order = np.argsort(x)
        xs, ys = x[order], y[order]
        slope, intercept = np.polyfit(xs, ys, 1)
        fig.add_trace(
            go.Scatter(
                x=xs,
                y=slope * xs + intercept,
                mode="lines",
                name="trend",
            )
        )
        fig.write_html("creative_plotly.html")

    return team_stat


if __name__ == "__main__":
    tourn_creative_calc("qval3.xlsx", is_plot=True)
    tourn_creative_calc_plotly("qval3.xlsx", is_plot=True)