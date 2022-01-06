from datetime import datetime, timedelta


class All_Weeks:
    def week_searching(dates):
        All_Week = {}
        for i in range(0, 52):
            period = [datetime(2021, 3, 29) + timedelta(days=7 * i), datetime(2021, 4, 4) + timedelta(days=7 * i)]
            i = i + 1
            All_Week['Week' + str(i)] = period

        for key, value in All_Week.items():
            if value[0] <= dates <= value[1]:
                return key

    def month_searching(dates):
        All_Month = {'01January':'01','02February':'02','03March':'03','04April':'04','05May':'05','06June':'06','07July':'07','08August':'08',
                     '09September':'09','10October':'10','11November':'11','12December':'12'}
        only_date = (str.split(str(dates))[0])
        year, month, date = only_date.split('-')
        for key, value in All_Month.items():
            if month == value:
                return key


