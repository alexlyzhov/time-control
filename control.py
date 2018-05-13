from openpyxl import load_workbook
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from pandas import datetime

fname = 'data.xlsx'

wb = load_workbook(fname)
sheet_name = wb.get_sheet_names()[0]
data = pd.DataFrame(wb[sheet_name].values)

null_idx = 1
null_str = data.loc[null_idx][0]

days = np.where(data[0] == null_str)

colormap = plt.get_cmap('jet')
mycolors = [colormap(i) for i in np.linspace(0, 1, 11)]
# mycolors = [(mycolors[i][0], mycolors[i][1], mycolors[i][2], 0.6) for i in xrange(len(mycolors))]
mycolormap = matplotlib.colors.ListedColormap(mycolors)

init_date = data.iloc[[days[0][0] - 1]][0].values[0].date()

# with plt.xkcd():
for daysi, dayi in enumerate(days[0]):
# for dayi in [days[0][16]]: # debug
    dtime = data.iloc[[dayi - 1]][0].values[0]
#     if type(dtime) == str:
#         dtime = datetime(int(dtime[6:10]), int(dtime[11]), int(dtime[13]))
#     date = dtime.date()
    date = init_date + pd.Timedelta(days = daysi)

    fig, axes = plt.subplots(1, 1)
    fig.set_size_inches((12, 5))
#     axes.get_yaxis().set_visible(False)
    plt.xlim(0, 24)
    plt.ylim(0, 11)
#     plt.xlabel('hour')
#     plt.ylabel('activity')
    plt.xticks(np.arange(0, 25))
#     plt.yticks(np.arange(0, 12))
    axes.tick_params(labeltop = False, labelright = True)
    axes.yaxis.set(ticks = np.arange(0.5, 11), ticklabels = np.arange(0, 11))
    fig.suptitle(date)

    for act in range(11):
        rowi = dayi + act
        row = data.iloc[[rowi]]
        coli = 2
        while row[coli].values[0] is not None:
            if type(row[coli].values[0]) == pd.datetime:
                start = row[coli].values[0].time()
            else:
                start = row[coli].values[0]
            if type(row[coli + 1].values[0]) == pd.datetime:
                end = row[coli + 1].values[0].time()
            else:
                end = row[coli + 1].values[0]
            intensity = act + 1
    #         print(start, end, intensity)

            start_hour = start.hour + float(start.minute) / 24
            end_hour = end.hour + float(end.minute) / 24
            color = list(colormap(float(act) / 11))
#             color[3] = 0.6
            plt.broken_barh([(start_hour, end_hour - start_hour)], (act, 1),
                                   cmap = colormap,
#                                        facecolor = color)
                                   facecolor = color, edgecolor = 'white')
            coli += 2

    plt.savefig(str(date) + '.png', bbox_inches='tight')

# the new cumulative graphs
for daysi, dayi in enumerate(days[0]):
# for dayi in [days[0][16]]: # debug
    dtime = data.iloc[[dayi - 1]][0].values[0]
#     if type(dtime) == str:
#         dtime = datetime(int(dtime[6:10]), int(dtime[11]), int(dtime[13]))
#     date = dtime.date()
    date = init_date + pd.Timedelta(days = daysi)

    fig, axes = plt.subplots(1, 1)
    fig.set_size_inches((12, 5))
#     axes.get_yaxis().set_visible(False)
    plt.xlim(0, 11)
    plt.ylim(0, 15)
#     plt.xlabel('hour')
#     plt.ylabel('activity')
    plt.yticks(np.arange(0, 16))
#     plt.yticks(np.arange(0, 12))
    axes.tick_params(labeltop = False, labelright = True)
    axes.xaxis.set(ticks = np.arange(0.5, 11), ticklabels = np.arange(0, 11))
    fig.suptitle(date)

    for act in range(11):
        rowi = dayi + act
        row = data.iloc[[rowi]]
        coli = 2
        duration_total = 0
        while row[coli].values[0] is not None:
            if type(row[coli].values[0]) == pd.datetime:
                start = row[coli].values[0].time()
            else:
                start = row[coli].values[0]
            if type(row[coli + 1].values[0]) == pd.datetime:
                end = row[coli + 1].values[0].time()
            else:
                end = row[coli + 1].values[0]
            intensity = act + 1
    #         print(start, end, intensity)

            start_hour = start.hour + float(start.minute) / 24
            end_hour = end.hour + float(end.minute) / 24
            duration_cur = end_hour - start_hour
            duration_total += duration_cur
            
            coli += 2
            
        color = list(colormap(float(act) / 11))
#             color[3] = 0.6
        plt.broken_barh([(act, 1)], (0, duration_total),
                               cmap = colormap,
#                                        facecolor = color)
                               facecolor = color, edgecolor = 'white')

    plt.savefig(str(date) + '_sums.png', bbox_inches='tight')