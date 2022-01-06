from Bin.All2021_Weeks import All_Weeks
from pandas import Series
from pandas import DataFrame
import pandas as pd
from collections import Counter
from Script.Function_for_sorting_data import floating_to_percentage_week, floating_to_percentage_month

Type_of_defect = 'Bend'  # bend = 'Bend'& contamination = 'Co + Adhesive Co'
File_address = '//192.168.80.20/share/Bu1-Eng/David/Bend/Bend summary2.xlsx'
# for example, '//192.168.80.20/share/Bu1-Eng/David/Bend/Bend summary2.xlsx'

All_MPT_lots = []
All_HTO_lots = []
All_SO_lots = []
All_Other_lots = []
Number_of_MPT = 0
Number_of_HTO = 0
Number_of_SO = 0
Number_of_Others = 0
lots_number = ()
CDM_Date = []
CDM_machine_no = []
MPT_perweek = []
MPT_permonth = []
SO_perweek = []
SO_permonth = []
HTO_perweek = []
HTO_permonth = []
Others_perweek = []
Others_permonth = []
Amount_of_Lotperweek = []
Amount_of_Lotpermonth = []

MPT_program = ('L-16420-08N04Z', 'L-16460-07N04Z', 'L-17290-06N04Z', 'L-17300-05N04Z', 'L-17310-06N04Z',
               'L-17370-04N04Z', 'L-17580-07N04Z', 'L-17720-02N04Z', 'L-17860-00N04Z', 'L-15750-06N04Z',
               'H-4700-04N04Z', 'L-15670-07N04Z', 'L-17460-04N04Z', 'L-15880-04 TAB.AN04Z', 'L-15880-04 TAB.BN04Z',
               'L-17480-04N04Z', 'L-17470-04N04Z', 'L-18290-02N04Z', 'L-18620-01N04Z', 'L-19120-02N04Z',
               'L-19010-01N04Z', 'L-19300-01N04Z', 'L-17460-04N04', 'L-17480-04N04', 'L-17480-05N04', 'L-18290-02N04',
               'L-18620-01N04', 'L-19010-01N04', 'L-19120-02N04', 'L-19300-01N04', 'H-4700-04N04', 'L-17370-06N04',
               'L-17370-04N04', 'L-19320-01N04', 'L-19430-01 OPT1', 'L-19320-00N04', 'L-15750-06N04', 'L-17480-05N04Z'
               )
HTO_program = {'L-18250-02N04Z', 'L-19110-03N04Z', 'L-18040-01N04Z', 'L-18600-00N04Z', 'L-19500-01N04Z',
               'L-18250-02N04', 'L-18530-02N04Z'
               }
SO_program = {'SL-13340-EN04', 'SL-13790-BN04', 'SL-13800-BN04', 'SL-13590-BN04', 'SL-13850-CN04', 'SL-13870-AN04',\
              'SL-13900-AN04'
              }

main_df = pd.read_excel(File_address, sheet_name='8', )
main_df = main_df.rename({'产品\n名称': '产品名称', '生产\n批号': '生产批号', 'CDM生产\n时间': 'CDM生产时间', 'CDM\n机号': \
    'CDM机号', '产品\n总个数': '产品总个数'}, axis=1)

lots_program = main_df['产品名称']
lots_number = main_df['生产批号']
CDM_Date = main_df['CDM生产时间']
CDM_machine_no = main_df['CDM机号']
Total_product_number = main_df['产品总个数']
Total_bend_number = main_df[Type_of_defect]
All_bend_rate = main_df['rate[%]']

for i in lots_program:
    if i in MPT_program:

        Number_of_MPT = Number_of_MPT + 1
    elif i in HTO_program:

        Number_of_HTO = Number_of_HTO + 1
    elif i in SO_program:

        Number_of_SO = Number_of_SO + 1
    else:

        Number_of_Others = Number_of_Others + 1

index_of_MPT_lots = []
index_of_HTO_lots = []
index_of_SO_lots = []
index_of_Other_lots = []
All_MPT_Date = []
All_SO_Date = []
All_HTO_Date = []
All_Other_Date = []

for i in MPT_program:
    index_of_MPT_lots.extend(lots_program[lots_program == i].index)
All_MPT_lots = lots_number[index_of_MPT_lots].values

All_MPT_Date = CDM_Date[index_of_MPT_lots].values
MPT_df = main_df.iloc[index_of_MPT_lots, :].reset_index()
main_df['customer'] = ''
main_df.iloc[index_of_MPT_lots, -1] = 'MPT'

for i in HTO_program:
    index_of_HTO_lots.extend(lots_program[lots_program == i].index)
All_HTO_lots = lots_number[index_of_HTO_lots].values
All_HTO_Date = CDM_Date[index_of_HTO_lots].values
HTO_df = main_df.iloc[index_of_HTO_lots, :].reset_index()
main_df.iloc[index_of_HTO_lots, -1] = 'HTO'

for i in SO_program:
    index_of_SO_lots.extend(lots_program[lots_program == i].index)
All_SO_lots = lots_number[index_of_SO_lots].values
All_SO_Date = CDM_Date[index_of_SO_lots].values
SO_df = main_df.iloc[index_of_SO_lots, :].reset_index()
main_df.iloc[index_of_SO_lots, -1] = 'SO'


All_Other_lots = list(
    ((set(lots_number).difference(set(All_MPT_lots))).difference((set(All_HTO_lots)))).difference((set(All_SO_lots))))

for i in All_Other_lots:
    index_of_Other_lots.extend(lots_number[lots_number == i].index)
All_Other_Date = CDM_Date[index_of_Other_lots].values
Other_df = main_df.iloc[index_of_Other_lots, :].reset_index()

for i in All_MPT_Date:
    MPT_perweek.append(All_Weeks.week_searching(pd.Timestamp(i)))
    MPT_permonth.append(All_Weeks.month_searching(i))
# print(Counter(MPT_perweek).items())

# CDM_perweek = []
# CDM_permonth = []
# for i in CDM_Date:
#     CDM_perweek.append(All_Weeks.week_searching(pd.Timestamp(i)))
#     CDM_permonth.append(All_Weeks.month_searching(i))
# print(CDM_permonth)

for i in All_SO_Date:
    SO_perweek.append(All_Weeks.week_searching(pd.Timestamp(i)))
    SO_permonth.append(All_Weeks.month_searching(i))


for i in All_HTO_Date:
    HTO_perweek.append(All_Weeks.week_searching(pd.Timestamp(i)))
    HTO_permonth.append(All_Weeks.month_searching(i))
for i in All_Other_Date:
    Others_perweek.append(All_Weeks.week_searching(pd.Timestamp(i)))
    Others_permonth.append(All_Weeks.month_searching(i))

for i in CDM_Date:
    Amount_of_Lotperweek.append(All_Weeks.week_searching(i))
    Amount_of_Lotpermonth.append(All_Weeks.month_searching(i))
# print(Counter(Amount_of_Lotperweek).items())


main_df['Week_number'] = pd.DataFrame(Amount_of_Lotperweek)
main_df['Month_number'] = pd.DataFrame(Amount_of_Lotpermonth)
Total_bend_rate_per_week = main_df.groupby('Week_number')[['产品总个数', Type_of_defect]].sum()
Total_bend_rate_per_week['bend_rate_per_week'] = (
        Total_bend_rate_per_week.iloc[:, 1] / Total_bend_rate_per_week.iloc[:, 0])
Total_bend_rate_per_week = Total_bend_rate_per_week.reset_index()
# Total_bend_rate_per_week['bend_rate_per_week'] = pd.Series(["{0:.2f}%".format(val * 100) for val in \
#     Total_bend_rate_per_week['bend_rate_per_week']], index = Total_bend_rate_per_week.index)
floating_to_percentage_week((Total_bend_rate_per_week))


Total_bend_rate_per_week_mcc = main_df.groupby(['Week_number', 'CDM机号'])[['产品总个数', Type_of_defect]].sum().unstack(
    fill_value=0).stack()
Total_bend_rate_per_week_mcc['bend_rate_per_week'] = (
        Total_bend_rate_per_week_mcc.iloc[:, 1] / Total_bend_rate_per_week_mcc.iloc[:, 0])
Total_bend_rate_per_week_mcc = Total_bend_rate_per_week_mcc.reset_index()
Total_bend_rate_per_week_mcc.iloc[:, -1] = Total_bend_rate_per_week_mcc.fillna(0).iloc[:, -1]
# Total_bend_rate_per_week_mcc['bend_rate_per_week'] = pd.Series(["{0:.2f}%".format(val * 100) for val in \
#     Total_bend_rate_per_week_mcc['bend_rate_per_week']], index = Total_bend_rate_per_week_mcc.index)
floating_to_percentage_week(Total_bend_rate_per_week_mcc)

Total_bend_rate_per_month = main_df.groupby('Month_number')[['产品总个数', Type_of_defect]].sum()
Total_bend_rate_per_month['bend_rate_per_month'] = (
        Total_bend_rate_per_month.iloc[:, 1] / Total_bend_rate_per_month.iloc[:, 0])
Total_bend_rate_per_month = Total_bend_rate_per_month.reset_index()
# Total_bend_rate_per_month['bend_rate_per_month'] = pd.Series(["{0:.2f}%".format(val * 100) for val in \
#     Total_bend_rate_per_month['bend_rate_per_month']], index = Total_bend_rate_per_month.index)
floating_to_percentage_month(Total_bend_rate_per_month)


Total_bend_rate_per_month_mcc = main_df.groupby(['Month_number', 'CDM机号'])[['产品总个数', Type_of_defect]].sum().unstack(
    fill_value=0).stack()
Total_bend_rate_per_month_mcc['bend_rate_per_month'] = (
        Total_bend_rate_per_month_mcc.iloc[:, 1] / Total_bend_rate_per_month_mcc.iloc[:, 0])
Total_bend_rate_per_month_mcc = Total_bend_rate_per_month_mcc.reset_index()
Total_bend_rate_per_month_mcc.iloc[:, -1] = Total_bend_rate_per_month_mcc.fillna(0).iloc[:, -1]
# Total_bend_rate_per_month_mcc['bend_rate_per_month'] = pd.Series(["{0:.2f}%".format(val * 100) for val in \
#     Total_bend_rate_per_month_mcc['bend_rate_per_month']], index = Total_bend_rate_per_month_mcc.index)
floating_to_percentage_month(Total_bend_rate_per_month_mcc)


MPT_df['Week_number'] = pd.DataFrame(MPT_perweek)
MPT_df['Month_number'] = pd.DataFrame(MPT_permonth)
MPT_bend_rate_per_week = MPT_df.groupby('Week_number')[['产品总个数', Type_of_defect]].sum()
MPT_bend_rate_per_week['bend_rate_per_week'] = (MPT_bend_rate_per_week.iloc[:, 1] / MPT_bend_rate_per_week.iloc[:, 0])
MPT_bend_rate_per_week = MPT_bend_rate_per_week.reset_index()
# MPT_bend_rate_per_week['bend_rate_per_week'] = pd.Series(["{0:.2f}%".format(val * 100) for val in \
#     MPT_bend_rate_per_week['bend_rate_per_week']], index = MPT_bend_rate_per_week.index)
floating_to_percentage_week(MPT_bend_rate_per_week)

MPT_bend_rate_per_week_mcc = MPT_df.groupby(['Week_number', 'CDM机号'])[['产品总个数', Type_of_defect]].sum().unstack(
    fill_value=0).stack()
MPT_bend_rate_per_week_mcc['bend_rate_per_week'] = (
        MPT_bend_rate_per_week_mcc.iloc[:, 1] / MPT_bend_rate_per_week_mcc.iloc[:, 0])
MPT_bend_rate_per_week_mcc = MPT_bend_rate_per_week_mcc.reset_index()
MPT_bend_rate_per_week_mcc.iloc[:, -1] = MPT_bend_rate_per_week_mcc.fillna(0).iloc[:, -1]
floating_to_percentage_week(MPT_bend_rate_per_week_mcc)


# MPT_bend_rate_per_week_mcc['bend_rate_per_week'] = pd.Series(["{0:.2f}%".format(val * 100) for val in \
#     MPT_bend_rate_per_week_mcc['bend_rate_per_week']], index = MPT_bend_rate_per_week_mcc.index)

MPT_bend_rate_per_month = MPT_df.groupby('Month_number')[['产品总个数', Type_of_defect]].sum()
MPT_bend_rate_per_month['bend_rate_per_month'] = (
        MPT_bend_rate_per_month.iloc[:, 1] / MPT_bend_rate_per_month.iloc[:, 0])
MPT_bend_rate_per_month = MPT_bend_rate_per_month.reset_index()
floating_to_percentage_month(MPT_bend_rate_per_month)
# MPT_bend_rate_per_month['bend_rate_per_month'] = pd.Series(["{0:.2f}%".format(val * 100) for val in \
#     MPT_bend_rate_per_month['bend_rate_per_month']], index = MPT_bend_rate_per_month.index)

MPT_bend_rate_per_month_mcc = MPT_df.groupby(['Month_number', 'CDM机号'])[['产品总个数', Type_of_defect]].sum().unstack(
    fill_value=0).stack()
MPT_bend_rate_per_month_mcc['bend_rate_per_month'] = (
        MPT_bend_rate_per_month_mcc.iloc[:, 1] / MPT_bend_rate_per_month_mcc.iloc[:, 0])
MPT_bend_rate_per_month_mcc = MPT_bend_rate_per_month_mcc.reset_index()
MPT_bend_rate_per_month_mcc.iloc[:, -1] = MPT_bend_rate_per_month_mcc.fillna(0).iloc[:, -1]
floating_to_percentage_month(MPT_bend_rate_per_month_mcc)

HTO_df['Week_number'] = pd.DataFrame(HTO_perweek)
HTO_df['Month_number'] = pd.DataFrame(HTO_permonth)
HTO_bend_rate_per_week = HTO_df.groupby('Week_number')[['产品总个数', Type_of_defect]].sum()
HTO_bend_rate_per_week['bend_rate_per_week'] = (HTO_bend_rate_per_week.iloc[:, 1] / HTO_bend_rate_per_week.iloc[:, 0])
HTO_bend_rate_per_week = HTO_bend_rate_per_week.reset_index()
floating_to_percentage_week(HTO_bend_rate_per_week)


HTO_bend_rate_per_week_mcc = HTO_df.groupby(['Week_number', 'CDM机号'])[['产品总个数', Type_of_defect]].sum().unstack(
    fill_value=0).stack()
HTO_bend_rate_per_week_mcc['bend_rate_per_week'] = (
        HTO_bend_rate_per_week_mcc.iloc[:, 1] / HTO_bend_rate_per_week_mcc.iloc[:, 0])
HTO_bend_rate_per_week_mcc = HTO_bend_rate_per_week_mcc.reset_index()
HTO_bend_rate_per_week_mcc.iloc[:, -1] = HTO_bend_rate_per_week_mcc.fillna(0).iloc[:, -1]
floating_to_percentage_week(HTO_bend_rate_per_week_mcc)

HTO_bend_rate_per_month = HTO_df.groupby('Month_number')[['产品总个数', Type_of_defect]].sum()
HTO_bend_rate_per_month['bend_rate_per_month'] = (
        HTO_bend_rate_per_month.iloc[:, 1] / HTO_bend_rate_per_month.iloc[:, 0])
HTO_bend_rate_per_month = HTO_bend_rate_per_month.reset_index()
floating_to_percentage_month(HTO_bend_rate_per_month)

HTO_bend_rate_per_month_mcc = HTO_df.groupby(['Month_number', 'CDM机号'])[['产品总个数', Type_of_defect]].sum().unstack(
    fill_value=0).stack()
HTO_bend_rate_per_month_mcc['bend_rate_per_month'] = (
        HTO_bend_rate_per_month_mcc.iloc[:, 1] / HTO_bend_rate_per_month_mcc.iloc[:, 0])
HTO_bend_rate_per_month_mcc = HTO_bend_rate_per_month_mcc.reset_index()
HTO_bend_rate_per_month_mcc.iloc[:, -1] = HTO_bend_rate_per_month_mcc.fillna(0).iloc[:, -1]
floating_to_percentage_month(HTO_bend_rate_per_month_mcc)


SO_df['Week_number'] = pd.DataFrame(SO_perweek)
SO_df['Month_number'] = pd.DataFrame(SO_permonth)
SO_bend_rate_per_week = SO_df.groupby('Week_number')[['产品总个数', Type_of_defect]].sum()
SO_bend_rate_per_week['bend_rate_per_week'] = (SO_bend_rate_per_week.iloc[:, 1] / SO_bend_rate_per_week.iloc[:, 0])
SO_bend_rate_per_week = SO_bend_rate_per_week.reset_index()
floating_to_percentage_week(SO_bend_rate_per_week)

SO_bend_rate_per_week_mcc = SO_df.groupby(['Week_number', 'CDM机号'])[['产品总个数', Type_of_defect]].sum().unstack(
    fill_value=0).stack()
SO_bend_rate_per_week_mcc['bend_rate_per_week'] = (
        SO_bend_rate_per_week_mcc.iloc[:, 1] / SO_bend_rate_per_week_mcc.iloc[:, 0])
SO_bend_rate_per_week_mcc = SO_bend_rate_per_week_mcc.reset_index()
SO_bend_rate_per_week_mcc.iloc[:, -1] = SO_bend_rate_per_week_mcc.fillna(0).iloc[:, -1]
floating_to_percentage_week(SO_bend_rate_per_week_mcc)


SO_bend_rate_per_month = SO_df.groupby('Month_number')[['产品总个数', Type_of_defect]].sum()
SO_bend_rate_per_month['bend_rate_per_month'] = (
        SO_bend_rate_per_month.iloc[:, 1] / SO_bend_rate_per_month.iloc[:, 0])
SO_bend_rate_per_month = SO_bend_rate_per_month.reset_index()
floating_to_percentage_month(SO_bend_rate_per_month)

SO_bend_rate_per_month_mcc = SO_df.groupby(['Month_number', 'CDM机号'])[['产品总个数', Type_of_defect]].sum().unstack(
    fill_value=0).stack()
SO_bend_rate_per_month_mcc['bend_rate_per_month'] = (
        SO_bend_rate_per_month_mcc.iloc[:, 1] / SO_bend_rate_per_month_mcc.iloc[:, 0])
SO_bend_rate_per_month_mcc = SO_bend_rate_per_month_mcc.reset_index()
SO_bend_rate_per_month_mcc.iloc[:, -1] = SO_bend_rate_per_month_mcc.fillna(0).iloc[:, -1]
floating_to_percentage_month(SO_bend_rate_per_month_mcc)

CDM_SO_month = {}
# try:
#     CDM_SO_month['CDM1'] = (SO_bend_rate_per_month_mcc.groupby(SO_bend_rate_per_month_mcc.CDM机号).get_group(1)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_SO_month['CDM2'] = (SO_bend_rate_per_month_mcc.groupby(SO_bend_rate_per_month_mcc.CDM机号).get_group(2)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_SO_month['CDM3'] = (SO_bend_rate_per_month_mcc.groupby(SO_bend_rate_per_month_mcc.CDM机号).get_group(3)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_SO_month['CDM4'] = (SO_bend_rate_per_month_mcc.groupby(SO_bend_rate_per_month_mcc.CDM机号).get_group(4)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_SO_month['CDM5'] = (SO_bend_rate_per_month_mcc.groupby(SO_bend_rate_per_month_mcc.CDM机号).get_group(5)).reset_index(drop=True)
# except KeyError:
#     pass

for i in range(1, 6):
    try:
        CDM_SO_month[f'CDM{i}'] = ((SO_bend_rate_per_month_mcc.groupby(SO_bend_rate_per_month_mcc.CDM机号).get_group(i)).reset_index(drop=True))
    except KeyError:
        pass



CDM_MPT_month = {}
# try:
#     CDM_MPT_month['CDM1'] = (MPT_bend_rate_per_month_mcc.groupby(MPT_bend_rate_per_month_mcc.CDM机号).get_group(1)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_MPT_month['CDM2'] = (MPT_bend_rate_per_month_mcc.groupby(MPT_bend_rate_per_month_mcc.CDM机号).get_group(2)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_MPT_month['CDM3'] = (MPT_bend_rate_per_month_mcc.groupby(MPT_bend_rate_per_month_mcc.CDM机号).get_group(3)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_MPT_month['CDM4'] = (MPT_bend_rate_per_month_mcc.groupby(MPT_bend_rate_per_month_mcc.CDM机号).get_group(4)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_MPT_month['CDM5'] = (MPT_bend_rate_per_month_mcc.groupby(MPT_bend_rate_per_month_mcc.CDM机号).get_group(5)).reset_index(drop=True)
# except KeyError:
#     pass


for i in range(1, 6):
    try:
        CDM_MPT_month[f'CDM{i}'] = ((MPT_bend_rate_per_month_mcc.groupby(MPT_bend_rate_per_month_mcc.CDM机号).get_group(i)).reset_index(drop=True))
    except KeyError:
        pass

CDM_HTO_month = {}
# try:
#     CDM_HTO_month['CDM1'] = (HTO_bend_rate_per_month_mcc.groupby(HTO_bend_rate_per_month_mcc.CDM机号).get_group(1)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_HTO_month['CDM2'] = (HTO_bend_rate_per_month_mcc.groupby(HTO_bend_rate_per_month_mcc.CDM机号).get_group(2)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_HTO_month['CDM3'] = (HTO_bend_rate_per_month_mcc.groupby(HTO_bend_rate_per_month_mcc.CDM机号).get_group(3)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_HTO_month['CDM4'] = (HTO_bend_rate_per_month_mcc.groupby(HTO_bend_rate_per_month_mcc.CDM机号).get_group(4)).reset_index(drop=True)
# except KeyError:
#     pass
# try:
#     CDM_HTO_month['CDM5'] = (HTO_bend_rate_per_month_mcc.groupby(HTO_bend_rate_per_month_mcc.CDM机号).get_group(5)).reset_index(drop=True)
# except KeyError:
#     pass

for i in range(1, 6):
    try:
        CDM_HTO_month[f'CDM{i}'] = ((HTO_bend_rate_per_month_mcc.groupby(HTO_bend_rate_per_month_mcc.CDM机号).get_group(i)).reset_index(drop=True))
    except KeyError:
        pass


CDM_SO_week = {}
for i in range(1, 6):
    try:
        CDM_SO_week[f'CDM{i}'] = ((SO_bend_rate_per_week_mcc.groupby(SO_bend_rate_per_week_mcc.CDM机号).get_group(i)).reset_index(drop=True))
    except KeyError:
        pass
CDM_MPT_week = {}
for i in range(1, 6):
    try:
        CDM_MPT_week[f'CDM{i}'] = ((MPT_bend_rate_per_week_mcc.groupby(MPT_bend_rate_per_week_mcc.CDM机号).get_group(i)).reset_index(drop=True))
    except KeyError:
        pass
CDM_HTO_week = {}
for i in range(1, 6):
    try:
        CDM_HTO_week[f'CDM{i}'] = ((HTO_bend_rate_per_week_mcc.groupby(HTO_bend_rate_per_week_mcc.CDM机号).get_group(i)).reset_index(drop=True))
    except KeyError:
        pass

CDM_Total_week = {}
for i in range(1, 6):
    try:
        CDM_Total_week[f'CDM{i}'] = ((Total_bend_rate_per_week_mcc.groupby(Total_bend_rate_per_week_mcc.CDM机号).get_group(i)).reset_index(drop=True))
    except KeyError:
        pass


CDM_numberoflot_week = (main_df.groupby(['CDM机号', 'Week_number'])[['生产批号']].count().unstack(fill_value=0).stack()).\
    reset_index()
p = (CDM_numberoflot_week.groupby(CDM_numberoflot_week.CDM机号).get_group(5)).reset_index(drop=True)
# print(p['生产批号'][2])
CDM_numberoflot_month = (main_df.groupby(['CDM机号', 'Month_number'])[['生产批号']].count().unstack(fill_value=0).stack()).\
    reset_index()
q = (CDM_numberoflot_month.groupby(CDM_numberoflot_month.CDM机号).get_group(5)).reset_index(drop=True)
print(Total_bend_rate_per_month.shape[0])

# m = CDM_Total_week['CDM1']["Week_number"].value_counts()
# print(m)

# print(Other_df['产品名称'].drop_duplicates())

# q = Total_bend_rate_per_month_mcc.sort_values((["CDM\n机号", 'Month_number']))
# print(q.iloc[:int(len(Total_bend_rate_per_month_mcc) / 5),::])
# print(q.iloc[7:(int(len(Total_bend_rate_per_month_mcc) / 5)*2),:])
# index_of_SO_lots.extend(lots_program[lots_program == i].index)
# wb = load_workbook("C:/Users/wheng/Downloads/100 vocabulary.xlsx")
# ws = wb['Sheet1']
#
# for i in range (len(All_MPT_lots)):
#     ws['B' + str(i + 2)] = All_MPT_lots[i]
#
# for i in range (len(All_HTO_lots)):
#     ws['C' + str(i + 2)] = All_HTO_lots[i]
#
# for i in range (len(All_SO_lots)):
#     ws['D' + str(i + 2)] = All_SO_lots[i]
#
# for i in range (len(All_Other_lots)):
#     ws['E' + str(i + 2)] = All_Other_lots[i]
#
# wb.save('C:/Users/wheng/Downloads/100 vocabulary.xlsx')
# print(list(main_df['Month_number'].drop_duplicates()))
# print(sorted(list(dict.fromkeys(SO_permonth))))
# def test1(customer_permonth):
#     Constant_c = []
#     Constant_d = 0
#     Constant_e = []
#     if len(list(main_df['Month_number'].drop_duplicates())) == len(list(dict.fromkeys(customer_permonth))):
#         for i in range(len(list(main_df['Month_number'].drop_duplicates()))):
#             Constant_e.append(0)
#     else:
#         for i in range(len(list(main_df['Month_number'].drop_duplicates()))):
#             if (list(main_df['Month_number'].drop_duplicates()))[i] in list(dict.fromkeys(customer_permonth)):
#                 Constant_e.append(Constant_d)
#             else:
#                 Constant_d = Constant_d + 1
#     return Constant_e

# print(sorted(list(dict.fromkeys(HTO_perweek))))
#
# print(sorted(list(main_df['Week_number'].drop_duplicates())))


# ws2['J' + str(6 - 1 * x)] = (HTO_bend_rate_per_week[HTO_bend_rate_per_week.Week_number == Three_weeks[x]].iloc[0,-1])







