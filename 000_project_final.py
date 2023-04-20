import shutil
import datetime as dt
import os
from tkinter import *
from tkinter import filedialog, messagebox
import configparser

def load_config():
    config_path = 'config.ini'

    if not os.path.exists(config_path):
        create_config()

    config = configparser.ConfigParser()
    config.read(config_path)

    if not config.has_option('DEFAULT', 'root_directory') or not config.has_option('DEFAULT', 'old_str_1st') \
            or not config.has_option('DEFAULT', 'old_str_2nd')  or not config.has_option('DEFAULT', 'old_str_3rd') \
            or not config.has_option('DEFAULT', 'new_str_1st')  or not config.has_option('DEFAULT', 'new_str_2nd') \
            or not config.has_option('DEFAULT', 'new_str_3rd') : 
            create_config()
            config.read(config_path)
    return config

def create_config():
    config = configparser.ConfigParser()
    config['DEFAULT'] = {'root_directory': '.', 'old_str_1st': '', \
                        'old_str_2nd': '', 'old_str_3rd': '', \
                        'new_str_1st': '','new_str_2nd': '', \
                        'new_str_3rd': ''}

    with open('config.ini', 'w') as config_file:
        config.write(config_file)

# clear files
def clear_list_frame_files():
    g_add_file_list.delete(0,'end')

# select_folder
def select_folder():
    clear_list_frame_files()
    global folder_selected
    folder_selected = filedialog.askdirectory()
    files = os.listdir(folder_selected)
    for i in files:
        g_add_file_list.insert(END, i)
    # global g_add_file_list_path
    # g_add_file_list_path = forder_selected

# unselect_folder
def unselect_files():
    for index in reversed(g_add_file_list.curselection()):
        g_add_file_list.delete(index)

def rename_files():
    # folder_path = g_add_file_list_path.get()
    folder_path = os.path.join(folder_selected)
    old_str_entry_1st = g_old_str_entry_1st.get()
    old_str_entry_2nd = g_old_str_entry_2nd.get()
    old_str_entry_3rd = g_old_str_entry_3rd.get()
    new_str_entry_1st = g_new_str_entry_1st.get()
    new_str_entry_2nd = g_new_str_entry_2nd.get()
    new_str_entry_3rd = g_new_str_entry_3rd.get()
    changed_files = []
    changed_files_cnt = 0

    config = configparser.ConfigParser()
    config.read('config.ini')
    config.set('DEFAULT', 'folder_path', folder_path)
    config.set('DEFAULT', 'old_str_1st', old_str_entry_1st)
    config.set('DEFAULT', 'old_str_2nd', old_str_entry_2nd)
    config.set('DEFAULT', 'old_str_3rd', old_str_entry_3rd)
    config.set('DEFAULT', 'new_str_1st', new_str_entry_1st)
    config.set('DEFAULT', 'new_str_2nd', new_str_entry_2nd)
    config.set('DEFAULT', 'new_str_3rd', new_str_entry_3rd)
    with open('config.ini', 'w') as config_file:
        config.write(config_file)

    for file_name in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path,file_name)):
            if old_str_entry_1st in file_name:
                x = dt.datetime.now()
                old_file_path = os.path.join(folder_path, file_name)
                new_file_name = file_name.replace(old_str_entry_1st, new_str_entry_1st)
                new_file_path = os.path.join(folder_path, new_file_name)
                ori_folder = os.path.join(folder_path, x.strftime("%Y%m%d"))
                if not os.path.exists(ori_folder):
                    os.makedirs(ori_folder)
                ori_file_path = os.path.join(ori_folder, file_name)
                shutil.copyfile(old_file_path, ori_file_path)
                os.rename(old_file_path, new_file_path)
                changed_files_cnt += 1
                changed_files.append(new_file_name)
        if os.path.isfile(os.path.join(folder_path,file_name)):
            if old_str_entry_2nd in file_name:
                x = dt.datetime.now()
                old_file_path = os.path.join(folder_path, file_name)
                new_file_name = file_name.replace(old_str_entry_2nd, new_str_entry_2nd)
                new_file_path = os.path.join(folder_path, new_file_name)
                ori_folder = os.path.join(folder_path, x.strftime("%Y%m%d"))
                if not os.path.exists(ori_folder):
                    os.makedirs(ori_folder)
                ori_file_path = os.path.join(ori_folder, file_name)
                shutil.copyfile(old_file_path, ori_file_path)
                os.rename(old_file_path, new_file_path)
                changed_files_cnt += 1
                changed_files.append(new_file_name)
        if os.path.isfile(os.path.join(folder_path,file_name)):
            if old_str_entry_3rd in file_name:
                x = dt.datetime.now()
                old_file_path = os.path.join(folder_path, file_name)
                new_file_name = file_name.replace(old_str_entry_3rd, new_str_entry_3rd)
                new_file_path = os.path.join(folder_path, new_file_name)
                ori_folder = os.path.join(folder_path, x.strftime("%Y%m%d"))
                if not os.path.exists(ori_folder):
                    os.makedirs(ori_folder)
                ori_file_path = os.path.join(ori_folder, file_name)
                shutil.copyfile(old_file_path, ori_file_path)
                os.rename(old_file_path, new_file_path)
                changed_files_cnt += 1
                changed_files.append(new_file_name)
    if changed_files_cnt:
        messagebox.showinfo('Complete', 'File renaming complete!')
    else:
        messagebox.showinfo('Complete', 'None!')

# main frame
def main():

    global g_add_file_list, g_old_str_entry_1st, g_old_str_entry_2nd, g_old_str_entry_3rd, \
                            g_new_str_entry_1st, g_new_str_entry_2nd, g_new_str_entry_3rd
    

    root = Tk()
    root.title("파일명 변경(대흥소프트밀-전병준)")
    root.geometry("450x480")

    # file frame
    file_frame = Frame(root)
    file_frame.pack(fill="x", padx=5, pady=5)

    select_folder_button = Button(file_frame, padx=5,pady=5, width=12, text="폴더 선택", command=select_folder)
    select_folder_button.pack(side="left")

    unselect_files_button = Button(file_frame, padx=5,pady=5, width=12, text="파일 취소", command=unselect_files)
    unselect_files_button.pack(side="right")

    # list frame
    list_frame = Frame(root)
    list_frame.pack(fill="both", padx=5, pady=5)

    scrollbar = Scrollbar(list_frame)
    scrollbar.pack(side = "right", fill="y")

    g_add_file_list = Listbox(list_frame, selectmode="extended", height=15, yscrollcommand=scrollbar.set)
    g_add_file_list.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=g_add_file_list.yview)

    # entry frame
    frame_word = LabelFrame(root, text="옵션1", relief="ridge", borderwidth=2, padx=5)
    frame_word.pack()

    frame_word2 = LabelFrame(root, text="옵션2",relief="ridge", borderwidth=2, padx=5)
    frame_word2.pack()

    frame_word3 = LabelFrame(root, text="옵션3", relief="ridge", borderwidth=2, padx=5)
    frame_word3.pack()

    # entry word
    lbl_width = Label(frame_word, text="첫번째 변경단어 입력", width=20)
    lbl_width.pack(side="left")

    g_old_str_entry_1st = Entry(frame_word, width=10)
    g_old_str_entry_1st.pack(side="left")

    lbl_width = Label(frame_word, text="------->", width=10)
    lbl_width.pack(side="left")

    g_new_str_entry_1st = Entry(frame_word, width=10)
    g_new_str_entry_1st.pack(side="left")

    lbl_width1 = Label(frame_word2, text="두번째 변경단어 입력", width=20)
    lbl_width1.pack(side="left")

    g_old_str_entry_2nd = Entry(frame_word2, width=10)
    g_old_str_entry_2nd.pack(side="left")

    lbl_width = Label(frame_word2, text="------->", width=10)
    lbl_width.pack(side="left")

    g_new_str_entry_2nd = Entry(frame_word2, width=10)
    g_new_str_entry_2nd.pack()

    lbl_width2 = Label(frame_word3, text="세번째 변경단어 입력", width=20)
    lbl_width2.pack(side="left")

    g_old_str_entry_3rd = Entry(frame_word3, width=10)
    g_old_str_entry_3rd.pack(side="left")

    lbl_width = Label(frame_word3, text="------->", width=10)
    lbl_width.pack(side="left")

    g_new_str_entry_3rd = Entry(frame_word3, width=10)
    g_new_str_entry_3rd.pack()

    # exit frame
    frame_end = LabelFrame(root, relief="flat", borderwidth=2, padx=5)
    frame_end.pack(side="right")

    # exit button
    break_btn = Button(frame_end, text="닫기", width=10, padx=5, command=root.quit)
    break_btn.pack(side="right")

    # start button
    rename_files_btn = Button(frame_end, text="실행", width=10, padx=5, command=rename_files)
    rename_files_btn.pack(side="right")

    load_config()

    # Read config.ini for previously selected folder, old string, and new string
    config = configparser.ConfigParser()
    config.read('config.ini')
    folder_path = config.get('DEFAULT', 'root_directory') 
    old_str_entry_1st = config.get('DEFAULT', 'old_str_1st') 
    old_str_entry_2nd = config.get('DEFAULT', 'old_str_2nd') 
    old_str_entry_3rd = config.get('DEFAULT', 'old_str_3rd') 
    new_str_entry_1st = config.get('DEFAULT', 'new_str_1st') 
    new_str_entry_2nd = config.get('DEFAULT', 'new_str_2nd') 
    new_str_entry_3rd = config.get('DEFAULT', 'new_str_3rd')
    

    if folder_path:
        g_add_file_list.insert(0, folder_path)
    if old_str_entry_1st:
        g_old_str_entry_1st.insert(0,old_str_entry_1st)
    if old_str_entry_2nd:
        g_old_str_entry_2nd.insert(0,old_str_entry_2nd)
    if old_str_entry_3rd:
        g_old_str_entry_3rd.insert(0,old_str_entry_3rd)
    if new_str_entry_1st:
        g_new_str_entry_1st.insert(0,new_str_entry_1st)
    if new_str_entry_2nd:
        g_new_str_entry_2nd.insert(0,new_str_entry_2nd)
    if new_str_entry_3rd:
        g_new_str_entry_3rd.insert(0,new_str_entry_3rd)

    # Start GUI event loop
    root.mainloop()


if __name__ == '__main__':
    main()