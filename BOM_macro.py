from tkinter import *
import keyboard
import pandas as pd
import pyautogui
import keyboard
from tkinter import filedialog

# 파일 선택 함수.
def select_file():
    global g_add_file_path
    global g_add_file_path
    file_selected = filedialog.askopenfilename(initialdir="/", title="Select file",filetypes=(("Excel files","*.xls"),("all files","*.*")))
    g_add_file_path = file_selected # 선택한 파일 경로를 g_add_file_list에 넣는다.
    file_path_label.config(text=g_add_file_path)

    # files = file_selected
    # for i in files:
    #     g_add_file_path.insert(END, i)
        

# 경로 설정하는 함수를 만들어야 한다.

def mecroStart():
    readData = pd.read_excel(g_add_file_path)
    print(readData)
    print()

    # 생산오더 등록 메크로
    mouse_x, mouse_y = pyautogui.position() # 마우스 x, y 좌표를 전달받는다
    size_get = g_size_entry.get()
    size = int (size_get)
    

    for i in readData.index: #readData 값만큼 반복, (72, 71코드만 index 값으로 받는 방법을 찾아야함)
        if keyboard.is_pressed('ESC'):
            break
    
        BOM_y = (i, readData['시스템코드'][i]) #시스템 코드 항목에는 엑셀의 BOM이 위치해있는 항목을 기입해준다. 
        Materials_y = (i, readData['규격'][i]) #규격 항목에는 엑셀의 원자재 코드가 위치해있는 항목을 기입해준다.
        Amout_y = (i, readData['총소요량'][i]) #총소요량 항목에는 엑셀의 수량이 위치해있는 항목을 기입해준다.

        str_BOM_y = str(BOM_y[1])
        str_Materials_y= str(Materials_y[1])
        int_Amout_y = int(Amout_y[1])

        if int_Amout_y == 0:
            continue

        if "71" in (str_BOM_y) or "72" in (str_BOM_y):
            # "71" or "72" in (str_BOM_y): # BOM 안에 71과 72 문자열이 존재해야 아래 코드를 실행시킨다.
            pyautogui.doubleClick()
            str_BOM_y_replace = str_BOM_y.replace("72", "70")
            pyautogui.write(str_BOM_y_replace)
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
            pyautogui.write("%d" % (int_Amout_y*size))

            # 마우스 커서가 MES칸을 넘어가면 y값을 고정시킨다.
            if mouse_y > 970:
                mouse_y == 984
                pyautogui.moveTo(176, 98)
                pyautogui.click()
                pyautogui.press('enter')
                pyautogui.press(['left', 'left', 'left', 'left', 'down'])
                pyautogui.press('enter')

            # 아니라면 y값을 더해줘서 다음 작업을 반복한다.
            else:
                mouse_y = mouse_y+25.4
                pyautogui.moveTo(176, 98)
                pyautogui.click()
                pyautogui.press('enter')
                pyautogui.press(['left', 'left', 'left', 'left', 'down'])
                pyautogui.press('enter')


# 라벨 발행 메크로
def mecroStart2():
    x, y = pyautogui.position()
    for i in range(21):
        if keyboard.is_pressed('ESC'):
            break
        pyautogui.moveTo(x, y )
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
        pyautogui.press('enter')
        pyautogui.moveTo(852, y )
        pyautogui.click()
        y = y+25.4

# f5 키를 눌렀을때와 f6 키를 눌렀을때 해당 작업을 실행시킨다.
def on_key_press(event):
    if event.keysym == 'F5':
        mecroStart()
    elif event.keysym == 'F6':
        mecroStart2()

# 인터페이스
def main():
    root = Tk()
    root.title("파일명 변경(대흥소프트밀-전병준)")
    root.geometry("650x150")
    root.resizable(False, False)

    global file_path_label

    # file frame
    file_frame = Frame(root)
    file_frame.pack(fill="x", padx=5, pady=(5, 0))

    select_folder_button = Button(file_frame, padx=5,pady=3, width=12, text="폴더 선택", command=select_file)
    select_folder_button.pack(side="left")

    # file path frame
    file_path_Frame = LabelFrame(file_frame, text="path", relief="ridge", borderwidth=1, padx=5)
    file_path_Frame.pack(side="left", fill="both", expand=True)


    file_path_label = Label(file_path_Frame, text="", width=80)
    file_path_label.pack(side="left")

    # size entry frame
    size_entry_Frame = LabelFrame(file_frame, text="set", relief="ridge", height=50, borderwidth=1, padx=5)
    size_entry_Frame.pack(side="left", padx=(10,5))

    global g_size_entry
    g_size_entry = Entry(size_entry_Frame, width=10)
    g_size_entry.pack(side="left")

    # disable automatic resizing of file_path_Frame
    file_path_Frame.pack_propagate(False)

    # list frame
    list_frame = Frame(root)
    list_frame.pack(fill="both", padx=5, pady=(0, 5))


    # entry frame
    frame_word = LabelFrame(root, text="생산 오더 등록", relief="ridge", borderwidth=2, padx=5)
    frame_word.pack(side="left", padx=5)

    frame_word2 = LabelFrame(root, text="라벨 발행",relief="ridge", borderwidth=2, padx=5)
    frame_word2.pack(side="left", padx=5)

    frame_word3 = LabelFrame(root, text="중지",relief="ridge", borderwidth=2, padx=5)
    frame_word3.pack(side="left", padx=5)

    # entry word
    lbl_width = Label(frame_word, text=" F5 ", width=20)
    lbl_width.pack(side="left")

    lbl_width1 = Label(frame_word2, text=" F6 ", width=20)
    lbl_width1.pack(side="left")

    lbl_width2 = Label(frame_word3, text=" ESC ", width=20)
    lbl_width2.pack(side="left")

    # exit frame
    frame_end = LabelFrame(root, relief="flat", borderwidth=2, padx=5)
    frame_end.pack(side="right")

    # exit button
    break_btn = Button(frame_end, text="닫기", width=10, padx=5, command=root.quit)
    break_btn.pack(side="right")

    # Bind key press events
    root.bind("<Key>", on_key_press)

    # Start GUI event loop
    root.mainloop()


if __name__ == '__main__':
    main()

# git test