"""
Original author: 不高兴就喝水
Improved by: Xingyu Qiu

pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159
"""
import pyautogui
import time
import xlrd
import pyperclip
import re

# 定义鼠标事件
EVENT_SINGLE_LEFT_CLICK = 1.0
EVENT_DOUBLE_LEFT_CLICK = 2.0
EVENT_SINGLE_RIGHT_CLICK = 3.0
EVENT_INPUT = 4.0
EVENT_WAIT = 5.0
EVENT_SCROLL = 6.0
EVENT_KEY = 7.0
EVENT_MOVE_TO = 8.0

IMAGE_DIR = './pictures/'


def mouse_click(clickTimes, lOrR, img, retry, repeat, x_pos=-1, y_pos=-1):
    """
    Do the left/right mouse click for specified times, retry, and repeat.
    Click on current position if img is None or empty string.
    Click on the position (xPos, yPos) if they are not the default value.
    Otherwise, click on the image position with name img.
    """
    # 转换为int
    retry = int(retry)
    repeat = int(repeat)

    # 在当前位置点击
    if not img or img == IMAGE_DIR:
        if repeat < 0:
            while True:
                print('当前位置点击')
                pyautogui.click(clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                time.sleep(0.5)
        else:
            for _ in range(repeat + 1):
                print('当前位置点击')
                pyautogui.click(clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                time.sleep(0.5)
        return

    # 在(xPos, yPos)位置点击
    if x_pos != -1 and y_pos != -1:
        if repeat < 0:
            # 无限重复
            while True:
                print(f'点击({x_pos},{y_pos})位置')
                pyautogui.click(x=x_pos, y=y_pos, clicks=clickTimes, interval=0.2, duration=0.2,
                                button=lOrR)
                time.sleep(0.1)
        else:
            # 有限重复
            for _ in range(repeat + 1):
                print(f'点击({x_pos},{y_pos})位置')
                pyautogui.click(x=x_pos, y=y_pos, clicks=clickTimes, interval=0.2, duration=0.2,
                                button=lOrR)
                time.sleep(0.1)
        return

    # helper function for click at img
    def try_click():
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            print('点击', img)
            pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2,
                            button=lOrR)
            return True
        print('未找到匹配图片', img, ',0.1秒后重试')
        time.sleep(0.1)
        return False

    # 在图片位置点击
    if retry < 0:
        # 无限重试
        if repeat < 0:
            # 无限重复
            while True:
                try_click()
        else:
            # 有限重复
            for _ in range(repeat + 1):
                while True:
                    if try_click():
                        break
    else:
        # 有限重试
        if repeat < 0:
            # 无限重复
            while True:
                for _ in range(retry + 1):
                    if try_click():
                        break
        else:
            # 有限重复
            for _ in range(repeat + 1):
                for __ in range(retry + 1):
                    if try_click():
                        break


def find_mouse_position():
    """
    Find the current mouse cursor position 100 times with 0.2 seconds time gap
    """
    for _ in range(100):
        x, y = pyautogui.position()
        print(f'鼠标坐标点为：({x},{y})')
        time.sleep(0.2)


def find_val_retry_repeat(sheet, row_i):
    """
    Get the value, retry, repeat from sheet for current row
    """
    # 默认空值，无限重试，无重复
    val = ''
    retry = -1
    repeat = 0

    val = sheet.row(row_i)[1].value
    if sheet.row(row_i)[2].ctype == 2 and sheet.row(row_i)[2].value != 0:
        retry = sheet.row(row_i)[2].value
    if sheet.row(row_i)[3].ctype == 2 and sheet.row(row_i)[3].value != 0:
        repeat = sheet.row(row_i)[3].value
    return val, retry, repeat


# 数据检查
# cmdType.value  1.0 左键单击  2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮  7.0 任意按键  8.0 鼠标移动
# ctype     空：0
#           字符串：1
#           数字：2
#           日期：3
#           布尔：4
#           error：5
def data_check(sheet):
    """
    Check validity of data in the sheet
    """
    checkCmd = True
    # 行数检查
    if sheet.nrows < 2:
        print('没数据啊哥')
        checkCmd = False
    # 每行数据检查
    for i in range(1, sheet.nrows):
        # 第1列 操作类型检查
        cmdType = sheet.row(i)[0]
        if cmdType.ctype != 2 or (
                cmdType.value != EVENT_SINGLE_LEFT_CLICK and cmdType.value != EVENT_DOUBLE_LEFT_CLICK
                and cmdType.value != EVENT_SINGLE_RIGHT_CLICK and cmdType.value != EVENT_INPUT and cmdType.value != EVENT_WAIT
                and cmdType.value != EVENT_SCROLL and cmdType.value != EVENT_KEY and cmdType.value != EVENT_MOVE_TO):
            print('第', i + 1, '行,第1列数据有毛病')
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet.row(i)[1]
        # 读图点击类型指令，内容必须为空或字符串类型
        if cmdType.value == EVENT_SINGLE_LEFT_CLICK or cmdType.value == EVENT_DOUBLE_LEFT_CLICK \
                or cmdType.value == EVENT_SINGLE_RIGHT_CLICK:
            if cmdValue.ctype != 0 and cmdValue.ctype != 1:
                print('第', i + 1, '行,第2列数据有毛病')
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == EVENT_INPUT:
            if cmdValue.ctype == 0:
                print('第', i + 1, '行,第2列数据有毛病')
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == EVENT_WAIT:
            if cmdValue.ctype != 2:
                print('第', i + 1, '行,第2列数据有毛病')
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == EVENT_SCROLL:
            if cmdValue.ctype != 2:
                print('第', i + 1, '行,第2列数据有毛病')
                checkCmd = False
        # 按键事件，内容必须为字符串类型且按键名存在
        if cmdType.value == EVENT_KEY:
            if cmdValue.ctype != 1:
                print('第', i + 1, '行,第2列数据有毛病')
                checkCmd = False
            elif cmdValue.value not in pyautogui.KEY_NAMES:
                print('The value should be in ', pyautogui.KEY_NAMES)
                checkCmd = False
        # 移动事件，内容必须为字符串类型
        if cmdType.value == EVENT_MOVE_TO:
            if cmdValue.ctype != 1:
                print('第', i + 1, '行,第2列数据有毛病')
                checkCmd = False
    return checkCmd


# 任务
def main_work(sheet):
    """
    Do the events in the sheet
    """
    for i in range(1, sheet.nrows):
        # 取本行指令的操作类型
        cmdType = sheet.row(i)[0]
        # 1代表单击左键
        if cmdType.value == EVENT_SINGLE_LEFT_CLICK:
            # 取图片名称/坐标位置,重试次数,和重复次数
            val, retry, repeat = find_val_retry_repeat(sheet, i)
            locationRegex = re.compile(r'p\(?(\d+),(\d+)\)?')
            mo = locationRegex.search(val)
            if mo is not None:
                x, y = int(mo.group(1)), int(mo.group(2))
                mouse_click(1, 'left', val, retry, repeat, x, y)
            else:
                img = IMAGE_DIR + val
                mouse_click(1, 'left', img, retry, repeat)
        # 2代表双击左键
        elif cmdType.value == EVENT_DOUBLE_LEFT_CLICK:
            # 取图片名称/坐标位置,重试次数,和重复次数
            val, retry, repeat = find_val_retry_repeat(sheet, i)
            locationRegex = re.compile(r'p\(?(\d+),(\d+)\)?')
            mo = locationRegex.search(val)
            if mo is not None:
                x, y = int(mo.group(1)), int(mo.group(2))
                mouse_click(2, 'left', val, retry, repeat, x, y)
            else:
                img = IMAGE_DIR + val
                mouse_click(2, 'left', img, retry, repeat)
        # 3代表右键
        elif cmdType.value == EVENT_SINGLE_RIGHT_CLICK:
            # 取图片名称/坐标位置,重试次数,和重复次数
            val, retry, repeat = find_val_retry_repeat(sheet, i)
            locationRegex = re.compile(r'p\(?(\d+),(\d+)\)?')
            mo = locationRegex.search(val)
            if mo is not None:
                x, y = int(mo.group(1)), int(mo.group(2))
                mouse_click(1, 'right', val, retry, repeat, x, y)
            else:
                img = IMAGE_DIR + val
                mouse_click(1, 'right', img, retry, repeat)
        # 4代表输入
        elif cmdType.value == EVENT_INPUT:
            # 取输入文本
            input_value = sheet.row(i)[1].value
            pyperclip.copy(input_value)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print('输入:', input_value)
            # 5代表等待
        elif cmdType.value == EVENT_WAIT:
            # 取等待时间
            wait_time = sheet.row(i)[1].value
            print('等待', wait_time, '秒')
            time.sleep(wait_time)
            print('等待结束')
        # 6代表滚轮
        elif cmdType.value == EVENT_SCROLL:
            # 取滚轮移动距离
            scroll = sheet.row(i)[1].value
            pyautogui.scroll(int(scroll))
            time.sleep(0.2)
            print('滚轮滑动', int(scroll), '距离')
        # 7代表任意按键
        elif cmdType.value == EVENT_KEY:
            # 取按键名称
            key_name = sheet.row(i)[1].value
            pyautogui.press(key_name)
            print('按下按键', key_name)
        # 8代表鼠标移动到指定图片
        elif cmdType.value == EVENT_MOVE_TO:
            # 取图片名称/坐标位置
            val = sheet.row(i)[1].value
            locationRegex = re.compile(r'p\(?(\d+),(\d+)\)?')
            mo = locationRegex.search(val)
            if mo is not None:
                location = (mo.group(1), mo.group(2))
                print('鼠标移动到', location)
                pyautogui.moveTo(location)
            else:
                img = IMAGE_DIR + val
                while True:
                    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
                    if location is not None:
                        print('鼠标移动到', location)
                        pyautogui.moveTo(location)
                        break
                    print('未找到匹配图片', img, ',0.1秒后重试')
                    time.sleep(0.1)


if __name__ == '__main__':
    print('\n欢迎使用RPA-简易脚本程序(・ω< )★\n')
    file = 'cmd.xls'
    # 打开文件
    wb = xlrd.open_workbook(filename=file)
    # 获取 number of sheets 和 sheet names
    text_display = ''
    idx_name_dict = {}
    num_sheets = wb.nsheets
    for i in range(num_sheets):
        curr_sheet_name = wb.sheet_by_index(i).name
        idx_name_dict[i] = curr_sheet_name
        text_display += f'{i + 1}: {curr_sheet_name}\n'
    # 选择功能：输入sheet页数，监测鼠标位置，退出
    while True:
        choice = input('\n选择功能\n'
                       f'{text_display}'
                       'm: 实时监测鼠标位置\n'
                       'q: 退出程序\n\n')
        if choice == 'q':
            quit()
        if choice == 'm':
            find_mouse_position()
        if choice.isnumeric() and 0 < int(choice) <= num_sheets:
            break
    sheet_idx = int(choice) - 1
    # 通过索引获取表格sheet页
    my_sheet = wb.sheet_by_index(sheet_idx)
    # 数据检查
    checkCmd = data_check(my_sheet)
    if checkCmd:
        # 输入循环次数或退出
        while True:
            key = input('\n选择循环次数\n'
                        '1: 做一次\n'
                        '2: 循环到死（请先确认脚本的安全性）\n'
                        'q: 退出程序\n\n')
            if key == 'q':
                quit()
            if key in {'1', '2'}:
                break

        print(f'\n开始执行脚本程序: {idx_name_dict[sheet_idx]} (￣3￣)\n')
        if key == '1':
            # 执行sheet中的指令，整体做一次
            main_work(my_sheet)
        elif key == '2':
            # 无限循环整个脚本
            while True:
                main_work(my_sheet)
                time.sleep(0.1)
                print('等待0.1秒')
    else:
        print('\n输入有误，已自动退出!\n')
