import pandas as pd



def floating_to_percentage_week(i):
    i['bend_rate_per_week'] = pd.Series(["{0:.2f}%".format(val * 100) for val in \
                                                                  i['bend_rate_per_week']],
                                                                 index=i.index)
    return i

def floating_to_percentage_month(i):
    i['bend_rate_per_month'] = pd.Series(["{0:.2f}%".format(val * 100) for val in \
                                                                  i['bend_rate_per_month']],
                                                                 index=i.index)
    return i

