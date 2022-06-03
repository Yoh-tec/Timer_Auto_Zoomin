import openpyxl as xl
import pyautogui as pgui
from datetime import datetime
from time import sleep
import locale


def check_time(day, hour, minute, upordown):
    if upordown == 'up':
        print(f'{day.hour} >= {hour}')
        print(f'{day.minute} >= {minute}')
        if day.hour > hour:
            return True
        elif (day.hour == hour) and (day.minute >= minute):
            return True
        else:
            return False
    elif upordown == 'down':
        print(f'{day.hour} <= {hour}')
        print(f'{day.minute} <= {minute}')
        if day.hour < hour:
            return True
        elif (day.hour == hour) and (day.minute <= minute):
            return True
        else:
            return False


def date_range(day):
    now = day.time()
    print(now)
    if (check_time(now, 9, 0, "up")) and (check_time(now, 10, 40, 'down')):
        print(2)
        return 'D2'
    elif (check_time(now, 10, 50, "up")) and (check_time(now, 12, 30, 'down')):
        print(3)
        return 'D3'
    elif (check_time(now, 13, 20, "up")) and (check_time(now, 15, 00, 'down')):
        print(4)
        return 'D4'
    elif (check_time(now, 15, 10, "up")) and (check_time(now, 16, 50, 'down')):
        print(5)
        return 'D5'
    elif (check_time(now, 17, 0, "up")) and (check_time(now, 18, 40, 'down')):
        print(6)
        return 'D6'


def get_url():
    wb = xl.load_workbook('./Source/week_data2.xlsx')
    locale.setlocale(locale.LC_TIME, 'en_US.UTF-8')
    day = datetime.now()
    print(day)
    today = day.strftime('%A')
    print(today)
    sh = wb[today]
    print(sh)
    #place = sh[date_range(day)].value
    place = "https://hosei-ac-jp.zoom.us/j/87112869934?pwd=YWM3SXJIcktaY3NpVUtJM0lrREI2UT09"
    url = place.replace('=', '_')
    url = url.replace(":", "'")
    return url


def main():
    DISPLAY = 0.0
    pgui.click(x=1200, y=0, duration=0)
    url = get_url()
    print(url)
    pgui.write(url)
    sleep(0.8)
    pgui.hotkey('return')
    sleep(0.5)
    pgui.hotkey('fn', 'F')
    sleep(3)
    pgui.click(x=894, y=251, duration=0)


if __name__ == "__main__":
    main()


