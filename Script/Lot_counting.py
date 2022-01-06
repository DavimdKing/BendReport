import pandas as pd
from Script.SortingandFiltering import main_df

total_week_lot_list = []
total_week_lot_list_mmc = []
SO_week_lot_list = []
SO_week_lot_list_mmc = []
HTO_week_lot_list = []
HTO_week_lot_list_mmc = []
MPT_week_lot_list = []
MPT_week_lot_list_mmc = []

Three_weeks = sorted(list(main_df['Week_number'].drop_duplicates()))[-3:]

Three_weeks.reverse()
for i in range(len(Three_weeks)):
    each_week = Three_weeks[i]
    total_week_lot = main_df[main_df.Week_number == f'{each_week}']
    total_week_lot_list.append(total_week_lot)
# print(total_week_lot_list)

for x in range(len(total_week_lot_list)):
    for i in range(1,6):
        total_week_lot_mmc = (total_week_lot_list[x])[total_week_lot_list[x].CDM机号 == i]
        total_week_lot_list_mmc.append(total_week_lot_mmc)
        # print(len(total_week_lot_mmc))
# print(total_week_lot_list_mmc)

# SO_dataframe = main_df[main_df.customer == 'SO']
# for i in range(len(Three_weeks)):
#     each_week = Three_weeks[i]
#     SO_week_lot = SO_dataframe[SO_dataframe.Week_number == f'{each_week}']
#     SO_week_lot_list.append(SO_week_lot)
#
# for x in range(len(SO_week_lot_list)):
#     for i in range(1, 6):
#         SO_week_lot_mmc = (SO_week_lot_list[x])[SO_week_lot_list[x].CDM机号 == i]
#         SO_week_lot_list_mmc.append(SO_week_lot_mmc)
#         # print(len(SO_week_lot_mmc))
#
# MPT_dataframe = main_df[main_df.customer == 'MPT']
# for i in range(len(Three_weeks)):
#     each_week = Three_weeks[i]
#     MPT_week_lot = MPT_dataframe[MPT_dataframe.Week_number == f'{each_week}']
#     MPT_week_lot_list.append(MPT_week_lot)
#
# for x in range(len(MPT_week_lot_list)):
#     for i in range(1, 6):
#         MPT_week_lot_mmc = (MPT_week_lot_list[x])[MPT_week_lot_list[x].CDM机号 == i]
#         MPT_week_lot_list_mmc.append(MPT_week_lot_mmc)
#         # print(len(MPT_week_lot_mmc))

def Lot_counting(x_dataframe):
    week_lot_list = []
    week_lot_list_mmc = []
    Constant_b = []
    Constant_c = []
    y = main_df[main_df.customer == x_dataframe]
    for i in range(len(Three_weeks)):
        each_week = Three_weeks[i]
        week_lot = y[y.Week_number == f'{each_week}']
        Constant_c.append(week_lot)

    for x in range(len(Constant_c)):
        for i in range(1, 6):
            week_lot_mmc = (Constant_c[x])[Constant_c[x].CDM机号 == i]
            Constant_b.append(week_lot_mmc)
            # print(len(week_lot_mmc))
    for i in range(len(Constant_b)):
        Constant_a = len(Constant_b[i])
        week_lot_list_mmc.append(Constant_a)
    for i in range(len(Constant_c)):
        Constant_d = len(Constant_c[i])
        week_lot_list.append(Constant_d)

    return (week_lot_list, week_lot_list_mmc)
print(Lot_counting('MPT'))

