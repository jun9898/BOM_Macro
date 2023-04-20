import pyautogui
import keyboard
# 마우스 이벤트를 후킹하기 위해 win32 모듈을 가져옵니다.
# import win32api

# # 마우스 좌표값을 게속해서 출력하기 위해 와일문을 만듭니다.
# while True:
#     # win32 api로 마우스의 상태를 가져옵니다.
#     down = win32api.GetKeyState(0x01)
#     # 마우스 상태가 다운이면 와일문을 탈출합니다.
#     if down == 0:
#         break
#     # 마우스의 현재 좌표를 출력합니다.
#     print(pyautogui.position())
x, y = pyautogui.position()
for i in range(21):
    if keyboard.is_pressed("space"):
        break
# while True
    pyautogui.moveTo(852, y )
    pyautogui.click()
    pyautogui.moveTo(1136, 138)
    pyautogui.click()
    pyautogui.moveTo(940, 542)
    pyautogui.click()
    pyautogui.write("100")
    pyautogui.moveTo(952, 595)
    pyautogui.click()
    pyautogui.moveTo(952, 625)
    pyautogui.click()
    pyautogui.moveTo(792, 622)
    pyautogui.click()
    # pyautogui.moveTo(1047, 616)
    # pyautogui.click()
    pyautogui.press('enter')
    pyautogui.moveTo(852, y )
    pyautogui.click()
    y = y+25.4

# 라벨발행 버튼의 값은 1136, 138
