import math
import sys
from itertools import accumulate

import pandas as pd
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Настройки внешнего вида, в целом можно отключать, изменять и т.д. как угодно
large = 22
med = 16
small = 12
width = 0.5
color_up = 'g'
color_down = 'g'
base_step_x = 3  # Шаг сетки по X
base_step_y = 2  # Шаг сетки по Y
params = {'axes.titlesize': large,
          'legend.fontsize': med,
          'figure.figsize': (12, 8),  # Размер изображения
          'axes.labelsize': med,
          'axes.titlesize': med,
          'xtick.labelsize': med,
          'ytick.labelsize': med,
          'figure.titlesize': large}
plt.rcParams.update(params)

# Тоже внешний вид, добавляет белые границы для столбцов
sns.set_style("white")

print(f"numpy v{np.__version__}")  # 2.1.0
print(f"pandas v{pd.__version__}")  # 2.2.2
print(f"matplotlib v{mpl.__version__}")  # 3.9.2
print(f"seaborn v{sns.__version__}")  # 0.13.2


def draw_plot(file_name, file_format='png', directory='results', labels=True, use_custom_grid=False):
    # Заагружаем данные из Excel
    data = pd.read_excel(file_name, header=None)
    # Удаляем ПОЛНОСТЬЮ пустые строки, если таковы имеются
    # Строки должны быть пустыми на протяжении ВСЕГО листа, а не только в куске этой таблицы!!!
    data = data.dropna(how='all')  # .fillna(0)

    if len(data) % 2 == 1:
        print(
            "Где-то проблема со строками, их нечётное количество!\n(Полностью пустые строки были удалены, они в учёт не идут)")
        # Если лезет ошибка, можно убрать строку ниже и код будет выполняться до тех пор, пока файл соотвествует формату
        # однако потом ТОЧНО появится другая ошибка
        return

    # Цикл обработки строк файлов. Берётся по две строки
    for i in range(len(data) // 2):
        # Переменная для удобства
        k = i * 2

        # Название сорта
        sort_name = str(data.values[k][0]).replace("/", "-")

        # Выбор междоузлий, которые сразу считаются аккумулированными,
        #  если что-то измениться и входные данные и так будут аккумулированными,
        #  комментируем строку x0 = list(accumulate(xx))
        xx = data.values[k][2:]
        xx = [x for x in xx if not np.isnan(x)]
        x0 = list(accumulate(xx))
        # Данные по длинам листов
        y1 = data.values[k + 1][2:]
        y1 = np.array([x for x in y1 if not np.isnan(x)])
        y2 = data.values[k + 1][2:]
        y2 = np.array([-x for x in y2 if not np.isnan(x)])

        # Чередуем листы
        for i in range(len(x0)):
            if i % 2 == 0:
                y1[i] = 0
            else:
                y2[i] = 0

        # Выбор данных на вывод в график
        plt.bar(x0, y1, width, facecolor=color_up)
        plt.bar(x0, y2, width, facecolor=color_down)

        # Около-оформление
        if use_custom_grid:
            step_x = base_step_x
            min_x = 0
            max_x = (math.ceil(x0[-1])) + (step_x * 2)
            plt.xticks(np.arange(min_x, max_x, step_x))
            step_y = base_step_y
            min_y = math.ceil(min(y2))
            min_y = min_y - (step_y if min_y % step_y == 0 else min_y % step_y)
            max_y = math.ceil(max(y1))
            max_y = max_y + (step_y if max_y % step_y == 0 else max_y % step_y)
            plt.yticks(np.arange(min_y, max_y, step_y))

        plt.title(sort_name, fontdict={'size': 40})
        plt.grid(True)

        # Задание подписей
        if labels:
            # Для верхних листов
            for x, y in zip(x0, y1):
                if y != 0:
                    plt.text(x, y + 0.05, f'Л: {y:.3f}\nМ: {x:.3f}', ha='center', va='bottom')

            # Для нижних листов
            for x, y in zip(x0, y2):
                if y != 0:
                    plt.text(x, y - 0.05, f'Л: {y:.3f}\nМ: {x:.3f}', ha='center', va='top')

        # Отображение если захочется посмотреть в программе а не в файле
        # Если включить, в файл пойдёт пустое изображение - учитите!
        # plt.show()

        # Создаём папку, куда будут сохранятсья изображения, если её нет
        #  Если нужно изменить название - передаём параметром
        if not os.path.exists(directory):
            os.makedirs(directory)
        # Сохранение изображения в формате png
        # Если нужно изменить формат - передаём параметром
        plt.savefig(f'{directory}/{sort_name}.{file_format}')
        plt.close()
        if True:
            print(f"{sort_name} успешно обработан.")


if __name__ == '__main__':
    # Имя файла, который загружаем
    # В файле, в именах сортов, НЕ должно содержаться символов по типу /\'" и т.д.
    # Все символы \ заменяются на - , остальные игнорируются и программа НЕ сможет выполниться корректно
    file_name = 'File.xlsx'
    # Непосредственно функция отрисовки
    draw_plot(file_name, file_format='png', directory='results', labels=False, use_custom_grid=True)

    # Пример с другими параметрами
    # draw_plot(file_name, file_format='jpg', directory='something', labels=False)
