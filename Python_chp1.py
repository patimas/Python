import tkinter as tk
from tkinter import ttk
from tkinter import font
from tkinter import Message
from tkinter import END
from tkinter import *


win = tk.Tk()

win.title("Python Test1")

# 가로x세로+x축+y축
win.geometry("500x1000+100+200")

win.resizable(False,False)


###################################################
#                Style
###################################################
style = ttk.Style()
f=font.Font(family="고딕체", size="12")
style.configure("TLabel", background="white", foreground="black", font=f)

###################################################
#
#                사용자 함수 설정
#
###################################################
def on_click():
    bn3.configure(text="클릭완성!")
    # lb3.configure(text="변경완료")

# def on_clickCombo():
# msg = tk.Message(win, text="1111")
# msg.pack()

def on_bn_clock(idx,value):
    print(value[idx])

def validate_write(t):
    win.title('Python Test1' + '  *')
    lb6['text'] = t
    return True


###################################################
# ※ Label, Button, Entry
# tk의 위젯과 ttk의 위젯 차이
# Ttk에는 18개의 위젯이 있으며, 그중 12개는 tkinter에 이미 존재합니다: 
# 존재 : (Button, Checkbutton, Entry, Frame, Label, LabelFrame, Menubutton, PanedWindow, Radiobutton, Scale, Scrollbar 및 Spinbox.) 
# 신규 : Combobox, Notebook, Progressbar, Separator, Sizegrip 및 Treeview입니다. 
# 그리고 이들은 모두 Widget의 서브 클래스입니다.
###################################################
#  tk의 Label, Button
lb1 = tk.Label(win, text="lb1 : ")
bn1 = tk.Button(win, text="bn1")

lb1.grid(row=0, column=0)
# Bn.pack()
bn1.grid(row=0, column=1)

lb2 = tk.Label(win, text="lb2 : ", bg="white", fg="black")
bn2 = tk.Button(win, text="bn2")
lb2.grid(row=0, column=2)
# Bn.pack()
bn2.grid(row=0, column=3)

# ttk의 Label, Button
lb3 = ttk.Label(win, text="lb3 : ")
bn3 = ttk.Button(win, text="bn3", command=on_click)
lb3.grid(row=1, column=0)
# Bn.pack()
bn3.grid(row=1, column=1)



lb4 = ttk.Label(win, text="lb4 : ", style="TLabel")
bn4 = ttk.Button(win, text="bn4")
lb4.grid(row=1, column=2)
# Bn.pack()
bn4.grid(row=1, column=3)


name = tk.StringVar()
en1 = ttk.Entry(win, width=20, textvariable=name)
en1.grid(row=3, column=0)


###################################################
# ※ Combobox
###################################################
comVar = tk.StringVar()
value = [1,2,3,4]
combo = ttk.Combobox(win, width=10,  textvariable=comVar, values=value, state='readonly')
combo.grid(row=3, column=0)
combo.current(0)

bn5 = ttk.Button(win, text="bn5")   #, command=on_clickCombo
bn5.grid(row=3, column=1)


###################################################
#    for문 활용
###################################################

# 1. Label
title = ('신규','추가','삭제')

# for i in title:
for i in range(len(title)):
    lb5 = ttk.Label(win, text=title[i])
    lb5.grid(row=i+5, column=0)

# 2. Entry
en = []
for i in range(len(title)):
    en.append(ttk.Entry(win, text=''))
    en[i].grid(row=i+5, column=1)

# 3. Button
bn=[]
for i in range(len(title)):
    bn.append(ttk.Button(win, text=title[i], command=lambda p=i : on_bn_clock(p,title)))
    bn[i].grid(row=i+5, column=2)

###################################################
#    validate
###################################################
v = win.register(validate_write)
en2 = ttk.Entry(win, validate="key", validatecommand=(v, '%S'))
en2.grid(row=10, column=0)

lb6 = ttk.Label(win)
en2.focus_force()
lb6.grid(row=10, column=1)


###################################################
#    Listbox
###################################################

lbox = tk.Listbox(win)
lbox.insert(1, '조우진')
lbox.insert(2, '김선애')
lbox.insert(3, '조재희')
lbox.grid(row=11, column=0)


v_list = ["조우진","김선애","조재희"]
lbox2 = tk.Listbox(win)

for ii in v_list:
    lbox2.insert(END, ii)

lbox2.grid(row=11, column=1)


###################################################
#    File Open
###################################################
f = open('Init_utf8.txt','r',encoding='utf-8') # python3부터는 ANSI기준의 파일만 읽을 수 있다. utf-8 인경우는 명시적으로 입력 해줘야 한다.
fileList = f.read()
f.close()

lbox3 = tk.Listbox(win)

data = fileList.split('\r\n')
for i in range(len(data)):
    rowdata2 = data[i].split(',')
    lbox3.insert(END, rowdata2)

lbox3.grid(row=11, column=2)


###################################################
#    LabelFrame
###################################################
lFrame = ttk.LabelFrame(win, text='종류', width=400)  #내부에 위젯이 존재할 경우, width와 height 설정을 무시하고 크기 자동 조절
lFrame.grid(row=12, column=0)

lFrameList=('선택1','선택2','선택3','선택4')

for i in range(len(lFrameList)):
    ttk.Label(lFrame, text=lFrameList[i]).grid(row=i, column=0, sticky=tk.W)


win.mainloop()