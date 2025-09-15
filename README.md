import pandas as pd
import numpy as np
path = r"C:\Users\Alienware\Desktop\南洋\第一学期\decision making\team project\Merlion Cafe Data AC6103 25T1.xlsx"
data_org = pd.read_excel(path,sheet_name='TPY')
data_2022 = pd.read_excel(path,sheet_name='Price&cost2022')
data_2023 = pd.read_excel(path,sheet_name='Price&cost2023')
data_1 = pd.merge(data_org,data_2022,how='left',on='product')
data = pd.merge(data_1,data_2023,how='left',on='product')
data['time'] = np.where(
    data['mode'] == 'Eat In',
    np.where(data['quantity'] > 1, (data['quantity']-1)*2+3, 3),
    np.where(data['quantity'] > 1, ((data['quantity']-1)*2+3)*0.4, 3*0.4)
)
data['sales rev'] = np.where(
    data['date'].dt.year == 2022,
    data['quantity'] * data['Price1'],data['quantity'] * data['Price2']
)
data['sales cost'] = np.where(data['date'].dt.year == 2022,data['quantity'] * data['Direct Costs1'],data['quantity'] * data['Direct Costs2'])
data['year'] = data.date.dt.year
data['day'] = data.date.dt.day
def analysis(x):
    data_a = data.groupby(['year','product','mode',x])[['quantity','time','sales rev','sales cost']].sum()
    with pd.ExcelWriter(r"C:\Users\Alienware\Desktop\南洋\第一学期\decision making\team project\analysis.xlsx", mode='a') as writer:
        data_a.to_excel(writer, sheet_name='TPY-' + x)

analysis('month')
analysis('weekday')
analysis('hour')

data_month_1 = data.groupby(['year','month'])[['time']].sum()
data_month = data_month_1.groupby('month')[['time']].mean()
data_weekday_1 = data.groupby(['year','month','day','weekday'])[['time']].sum()
data_weekday = data_weekday_1.groupby(['weekday'])[['time']].mean()
data_hour_1 = data.groupby(['year','month','day','hour'])[['time']].sum()
data_hour = data_hour_1.groupby(['hour'])[['time']].mean()
data_month.insert(0,'Type','month')
data_weekday.insert(0,'Type','weekday')
data_hour.insert(0,'Type','hour')
data_time = pd.concat([data_month,data_weekday,data_hour]).round(2)
with pd.ExcelWriter(r"C:\Users\Alienware\Desktop\南洋\第一学期\decision making\team project\analysis.xlsx",mode='a') as writer:
    data_time.to_excel(writer,sheet_name='TPY-time')
<img width="73" height="856" alt="image" src="https://github.com/user-attachments/assets/66692353-b825-4965-8ec3-1786d5d611dd" />
