import os
import time
from datetime import datetime

import pandas as pd
import pyautogui


def click(img):
    """点击图片中心"""
    enter_mt_btn = pyautogui.locateCenterOnScreen(img)  # 锁定位置
    pyautogui.moveTo(enter_mt_btn)  # 移动到这里
    pyautogui.click()  # 单击一下
    time.sleep(1)  # 等待1s


def enter_wd(msg):
    """输入信息"""
    pyautogui.write(msg)
    time.sleep(2)


def enter_meeting(id, ps_word, way):
    # 加入会议
    if way == 'Tencent':
        # 这里需要替换自己电脑中腾讯会议应用程序的位置
        os.startfile(r"G:\SmallTools\腾讯会议\WeMeet\wemeetapp.exe")
        time.sleep(5)
        click("enter_meeting.png")
        enter_wd(id)
        click("enter_meeting2.png")
        enter_wd(ps_word)
        click("enter.png")
    if way == 'Zoom':
        os.startfile(r"C:\Users\A\AppData\Roaming\Zoom\bin\Zoom.exe")
        time.sleep(4)
        click("zoom_enter_meeting.png")
        time.sleep(2)
        enter_wd(id)
        click("zoom_enter_meeting2.png")
        time.sleep(2)
        enter_wd(ps_word)
        click("zoom_enter_meeting3.png")


df = pd.read_excel(r"timings.xlsx")


while True:
    # 获取当前时间
    now_time = datetime.now().strftime("%H:%M")
    now_day = datetime.now().strftime("%m/%d")
    now_week = str(datetime.now().strftime("%w"))
    # 获取表格中的时间
    df_time = time_list = [i.strftime("%H:%M") for i in df['Timings']]

    # 时间匹配
    matched_time_list = []  # 时间匹配的行索引
    for i, item in enumerate(df_time):  # 将时间一致的行添加到索引列表中
        if now_time == item:
            matched_time_list.append(i)

    if len(matched_time_list) != 0:  # 如果索引列表不为空，则核对日期或周是否一致
        for i in matched_time_list:
            df_day = df.loc[i, 'Date'].strftime("%m/%d") if not pd.isnull(df.loc[i, 'Date']) else -1
            df_week = int(df.loc[i, 'Week']) if not pd.isnull(df.loc[i, 'Week']) else -1
            if str(df_day) == str(now_day) or str(df_week) == str(now_week):  # 信息全部一致，开始加入会议
                print('匹配成功！')
                print('现在的日期：%s\t数据的日期：%s\t' % (now_day, df_day))
                print('现在是星期：%s\t数据要星期：%s' % (now_week, df_week))

                df_way = str(df.loc[i, 'Way']) if not pd.isnull(df.loc[i, 'Way']) else -1
                meeting_id = str(df.loc[i, 'MeetingId'])
                passcode = str(df.loc[i, 'Passcode'])
                enter_meeting(meeting_id, passcode, df_way)
                time.sleep(60)  # 等待一分钟后结束
                break


