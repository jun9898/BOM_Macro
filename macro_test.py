import pyautogui
import keyboard

# 라벨발행 메크로
x, y = pyautogui.position()
for i in range(21):
    if keyboard.is_pressed('ESC'):
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
