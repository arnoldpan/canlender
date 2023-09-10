import pandas as pd
import datetime
from dateutil import tz


from ics import Calendar, Event

# 读取Excel文件
df = pd.read_excel('schedule.xlsx')

# 提取所需列数据
title_column = df['标题']
location_column = df['位置']
start_time_column = df['开始时间']
end_time_column = df['截止时间']
reminder_column = df['30分钟前提醒内容']

# 创建日历对象
calendar = Calendar()

# 遍历每一行数据
for i in range(len(df)):
    # 提取每一行的数据
    title = title_column.iloc[i]
    location = location_column.iloc[i]
    start_time0 = start_time_column.iloc[i]
    end_time0 = end_time_column.iloc[i]
    reminder = reminder_column.iloc[i]

    # 将时间字符串转换为Timestamp对象
    start_time = pd.Timestamp(start_time0)
    end_time = pd.Timestamp(end_time0)

    # 将Timestamp对象转换为字符串
    start_time_str = start_time.strftime('%Y-%m-%d %H:%M:%S')
    end_time_str = end_time.strftime('%Y-%m-%d %H:%M:%S')
    # 添加提醒
    #reminder_time0 = datetime.datetime.strptime(start_time_str, '%Y-%m-%d %H:%M:%S') - datetime.timedelta(minutes=15)
    #reminder_time_str = reminder_time0.strftime('%Y-%m-%d %H:%M:%S')

    # 创建香港时区对象
    hk_timezone = tz.gettz('Asia/Hong_Kong')

    # 解析时间文本为datetime对象，并设置为香港时区
    parsed_start_time = datetime.datetime.strptime(start_time_str, '%Y-%m-%d %H:%M:%S').replace(tzinfo=hk_timezone)
    parsed_end_time = datetime.datetime.strptime(end_time_str, '%Y-%m-%d %H:%M:%S').replace(tzinfo=hk_timezone)
    #reminder_time = datetime.datetime.strptime(reminder_time_str, '%Y-%m-%d %H:%M:%S').replace(tzinfo=hk_timezone)
    reminder_time = parsed_start_time - datetime.timedelta(minutes=30)
    # 将香港时区时间转换为绝对时间（UTC时间）
    reminder_absolute_time = reminder_time.astimezone(tz.UTC)

    # 创建事件对象
    event = Event()
    event.name = title
    event.location = location
    event.begin = parsed_start_time
    event.end = parsed_end_time
    print(event.begin)
    event.alarm = reminder_absolute_time
    print(event.alarm)
    # 将事件添加到日历
    calendar.events.add(event)

# 生成.ics文件
with open('DMBA_Calendar.ics', 'w') as f:
    f.write(str(calendar))

print("已生成.ics文件")
