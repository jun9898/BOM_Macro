import pandas as pd
import pyautogui
import keyboard

readData = pd.read_excel("C:/macro_test/test_1.xlsx")
print(readData)
print()

mouse_x, mouse_y = pyautogui.position() # 마우스 x, y 좌표를 전달받는다

for i in readData.index: #readData 값만큼 반복, (72, 71코드만 index 값으로 받는 방법을 찾아야함)
    if keyboard.is_pressed('ESC'):
        break
    # pyautogui.moveTo(mouse_x, mouse_y)
    pyautogui.doubleClick()
    BOM_y = (i, readData['test1'][i]) #test1 항목에는 엑셀의 BOM이 위치해있는 항목을 기입해준다. ex) 항목코드
    Materials_y = (i, readData['test2'][i]) #test2 항목에는 엑셀의 원자재 코드가 위치해있는 항목을 기입해준다. ex) 항목코드
    Amout_y = (i, readData['test3'][i]) #test3 항목에는 엑셀의 수량이 위치해있는 항목을 기입해준다. ex) 항목코드

    str_BOM_y = str(BOM_y[1])
    str_Materials_y= str(Materials_y[1])
    str_Amout_y = str(Amout_y[1])
    
    pyautogui.write("%s" % str_BOM_y)
    pyautogui.hotkey('ctrl', 'f') # ctrl+f
    pyautogui.moveTo(808, 567)
    pyautogui.doubleClick()

    pyautogui.press(['right', 'right', 'right'])
    pyautogui.press('enter')

    pyautogui.write("%s" % str_Materials_y)
    pyautogui.hotkey('ctrl', 'f') # ctrl+f
    pyautogui.moveTo(808, 567)
    pyautogui.doubleClick()

    pyautogui.press('right')
    pyautogui.write("%s" % str_Amout_y)

    if mouse_y > 970:
        mouse_y == 984
        pyautogui.moveTo(176, 98)
        pyautogui.click()
        pyautogui.press('enter')
        pyautogui.press(['left', 'left', 'left', 'left', 'down'])
        pyautogui.press('enter')
        continue
    else:
        mouse_y = mouse_y+25.4
        pyautogui.moveTo(176, 98)
        pyautogui.click()
        pyautogui.press('enter')
        pyautogui.press(['left', 'left', 'left', 'left', 'down'])
        pyautogui.press('enter')
        print("test")