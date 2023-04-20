import pyautogui
# 마우스 이벤트를 후킹하기 위해 win32 모듈을 가져옵니다.
import win32api

# 마우스 좌표값을 게속해서 출력하기 위해 와일문을 만듭니다.
while True:
    # win32 api로 마우스의 상태를 가져옵니다.
    down = win32api.GetKeyState(0x01)
    # 마우스 상태가 다운이면 와일문을 탈출합니다.
    if down == 0:
        break
    # 마우스의 현재 좌표를 출력합니다.
    print(pyautogui.position())