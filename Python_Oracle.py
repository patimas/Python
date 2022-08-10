from contextlib import nullcontext
import tkinter as tk
from tkinter import Image, StringVar, ttk, Menu
import cx_Oracle
import os as os

from setuptools import Command

import pyperclip

# System Tray Import
from PIL import Image, ImageTk
import pyautogui as agui
from pystray import MenuItem as item
import pystray

# 파일 저장
import pickle



# 파일명을 만들때 pyautogui.py로 만들면 에러 난다... pyautogui로 만들지 말것...


############################################################
#
#  GUI 기본값 설정
#
############################################################
win = tk.Tk()
win.title("Python 기본")

# 가로x세로+x축+y축
win.geometry("1440x800+100+100")
# win.resizable(False,False)

############################################################
#
#  설정 파일 오픈
#
############################################################

f = open('config.ini','rb')
configValueRead = pickle.load(f)
f.close()


############################################################
#
#  DB 연결
#
############################################################
dsn = cx_Oracle.makedsn("192.168.31.231", 30101, service_name="SADARI")
cx_Oracle.init_oracle_client(lib_dir=r"C:\Projects\Python\instantclient_19_9")
connection = cx_Oracle.connect(user="system", password="oracle", dsn=dsn, encoding="UTF-8")
cur = connection.cursor()


sqlStr0 = "SELECT DISTINCT OWNER,OWNER FROM DBA_TABLES WHERE  OWNER IN ('HR','BTL_REP4')"

sqlStr1 = ""
sqlStr1 +="SELECT  T1.OWNER, T1.TABLE_NAME, T2.COMMENTS      "
sqlStr1 +="FROM    DBA_TABLES          T1  ,                 "
sqlStr1 +="        DBA_TAB_COMMENTS    T2                    "
sqlStr1 +="WHERE   1=1                                       "
sqlStr1 +="AND     T1.OWNER        = :1                      "
sqlStr1 +="AND     T1.OWNER        = T2.OWNER                "
sqlStr1 +="AND     T1.TABLE_NAME   = T2.TABLE_NAME           "
# sqlStr1 +="AND     T1.TABLE_NAME   LIKE :2                   "
sqlStr1 +="ORDER BY T1.OWNER, T1.TABLE_NAME        			 "


sqlStr2 = ""
sqlStr2 += "SELECT  T1.TABLE_NAME           ,          "
sqlStr2 += "        T1.COLUMN_NAME          ,          "
sqlStr2 += "        T2.COMMENTS                        "
sqlStr2 += "FROM    DBA_TAB_COLUMNS     T1  ,          "
sqlStr2 += "        DBA_COL_COMMENTS    T2             "
sqlStr2 += "WHERE   T1.OWNER        = T2.OWNER      (+)"
sqlStr2 += "AND     T1.TABLE_NAME   = T2.TABLE_NAME (+)"
sqlStr2 += "AND     T1.COLUMN_NAME  = T2.COLUMN_NAME(+)"
sqlStr2 += "AND     T1.OWNER        = :1               "
sqlStr2 += "AND     T1.TABLE_NAME   = :2               "
sqlStr2 += "ORDER BY T1.COLUMN_ID                      "


############################################################
#
#  데이터 조회
#
############################################################
def on_tab1tv1_click(object1, object2, e):
    selectItem = object1.focus()
    # getValue = object1.item(selectItem).get('values')

    getValue = []
    getValue.append(object1.item(selectItem).get('values')[0])
    getValue.append(object1.item(selectItem).get('values')[1])

    on_get_data("sql", sqlStr2, getValue, object2)
    
    testLabel.configure(text=getValue)


def on_get_data(flag, sql, param, object):

    vdata = []

    if flag == "sql" and param == "":

        for i in cur.execute(sql):
            vdata.append(i)

        for i in range(len(vdata)):
            object.insert('', 'end', text=i, values=vdata[i], iid=str(i))

    elif flag == "sql" and param != "":
        
        # Grid 초기화
        for i in object.get_children():
            object.delete(i)

        for i in cur.execute(sql,param):
            vdata.append(i)

        for i in range(len(vdata)):
            object.insert('', 'end', text=i, values=vdata[i], iid=str(i))

    elif flag == "combo":
        cur.execute(sqlStr0)
        result = [row1[0] for row1 in cur]
        return result


def _onComboSelected(object1, object2, e):
    value = []
    value.append(object1.get())
    on_get_data("sql", sqlStr1, value, object2)
    # on_get_data("sql", sqlStr1, ['HR'], tab1tv1)

def tab1DbSave():
    configValueInput = {}
    configValueInput['user'] = db_en_User.get()
    configValueInput['pass'] = db_en_Pass.get()

    f = open('config.ini','wb')
    pickle.dump(configValueInput, f)
    f.close()

def _copy_clip_board(tree, event):
    selection = tree.selection()
    
    column = tree.identify_column(event.x)
    
    column_no = int(column.replace("#", "")) - 1  #마우스로 선택한 컬럼 index만 가져 오기
    copy_values = []
    for each in selection:
        try:
            value = tree.item(each)["values"][column_no]
            copy_values.append(str(value))
        except:
            pass
        
    copy_string = "\n".join(copy_values)
    pyperclip.copy(copy_string)



############################################################
#
#  이벤트
#
############################################################

# 화면 종료
def _exit():
    cur.close()
    connection.close()
    win.destroy()
    # exit()

# System Tray Start
def _quit_window(icon, item):
    icon.stop()
    win.destroy()

def _show_Window(icon, item):
    icon.stop()
    win.after(0,win.deiconify())

def _systemTray():
    win.withdraw()
    image=Image.open('tray.ico')
    menu = (item('Quit', _quit_window),item('Show', _show_Window))
    icon = pystray.Icon("name", image, "My System Tray Icon", menu)
    icon.run()
# System Tray End


# 창이동 이벤트
def keyPress(e):
    fw= agui.getActiveWindow()
    winstatus = 'none'

    if fw.isMaximized == True:
        winstatus = 'max'
        fw.restore()
    
    fw.moveTo(300,300)
    
    if winstatus == 'max':
        winstatus = 'none'
        fw.maximize()





############################################################
#
#  Menu 생성
#
############################################################
menu_bar = tk.Menu(win)

file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Exit", command=_exit)
menu_bar.add_cascade(label="FILE", menu=file_menu)  #화면에 그려주는 구문

win.config(menu=menu_bar)


############################################################
#
#  TAB 생성
#
############################################################
tabCtrl = ttk.Notebook(win)

tab1 = ttk.Frame(tabCtrl)
tab2 = ttk.Frame(tabCtrl)
tabCtrl.add(tab1, text="Table")
tabCtrl.add(tab2, text="Table")
tabCtrl.pack(expand=1, fill="both")


testLabel = ttk.Label(win, text="test")
testLabel.pack()


############################################################
#
#  TAB1 생성
#
############################################################
# 접속정보
tab1LFrame00 = ttk.LabelFrame(tab1, text="Connenction", width=400)
tab1LFrame00.grid(row=0, column=0, padx=8, pady=4, sticky=tk.W)
db_lb_User = ttk.Label(tab1LFrame00, text="사용자:")
db_lb_User.grid(row=0, column=0)

varDbUserStr = StringVar()
db_en_User = ttk.Entry(tab1LFrame00, text='', textvariable=varDbUserStr)
db_en_User.grid(row=0, column=1, padx=4, pady=2)
# db_en_User.insert(0, configValueRead['user'])


db_lb_Pass = ttk.Label(tab1LFrame00, text="비번:")
db_lb_Pass.grid(row=0, column=2, padx=4, pady=2)

varDbPassStr = StringVar()
db_en_Pass = ttk.Entry(tab1LFrame00, textvariable=varDbPassStr)
db_en_Pass.grid(row=0, column=3, padx=4, pady=2)
# db_en_Pass.insert(0, configValueRead['pass'])

dbBnSave = ttk.Button(tab1LFrame00, text="저장", command=tab1DbSave)
dbBnSave.grid(row=0, column=4)


tab1LFrame01 = ttk.LabelFrame(tab1, text="Table", width=400)
tab1LFrame01.grid(row=1, column=0, padx=8, pady=4, sticky=tk.W)


# 콤보박스
comVar = tk.StringVar()
value = on_get_data("combo", sqlStr0, "", "")
combo = ttk.Combobox(tab1LFrame01, width=10,  textvariable=comVar, values=value, state='readonly')
combo.grid(row=1, column=0)
combo.current(0)



# 검색바
varStr = StringVar()
tab1en1 = ttk.Entry(tab1LFrame01, textvariable=varStr)
tab1en1.grid(row=1, column=1)



# GIRD 시작
tab1LFrame02 = ttk.LabelFrame(tab1, text="Table", width=400)
tab1LFrame02.grid(row=2, column=0, padx=8, pady=4)

# 1Tree 데이터
tab1tv1=ttk.Treeview(tab1LFrame02, columns=["1","2","3"], displaycolumns=["1","2","3"], height=30)
tab1tv1.grid(row=3, column=0, columnspan=2)

tab1tv1.column("#0", width=70)
tab1tv1.heading("#0", text="NO")

tab1tv1.column("1",width=100)
tab1tv1.heading("1", text="Owner")

tab1tv1.column("2",width=200)
tab1tv1.heading("2", text="Table")

tab1tv1.column("3",width=300)
tab1tv1.heading("3", text="Comment")


combo.bind('<<ComboboxSelected>>', lambda x : _onComboSelected(combo, tab1tv1, x))
# bind가 되었을때 기본적으로 실행 한다.
combo.event_generate('<<ComboboxSelected>>')

# 2Tree 데이터
tab1tv2 = ttk.Treeview(tab1LFrame02, columns=["1","2","3"], displaycolumns=["1","2","3"], height=30)
tab1tv2.grid(row=3, column=2, padx=20, pady=0)

tab1tv2.column("#0", width=70)
tab1tv2.heading("#0", text="NO")
tab1tv2.column("1",width=150)
tab1tv2.heading("1", text="Table")
tab1tv2.column("2",width=200)
tab1tv2.heading("2", text="Column")
tab1tv2.column("3",width=300)
tab1tv2.heading("3", text="Comment")

tab1tv1.bind( '<ButtonRelease-1>', lambda x : on_tab1tv1_click(tab1tv1,tab1tv2,x))
# tab1tv2.bind( '<Control-Key-c>', lambda x : _copy_clip_board(tab1tv2,x))







############################################################
#
#  TAB2 생성
#
############################################################
# 접속정보
# tab1LFrame20 = ttk.LabelFrame(tab2, text="ImagePath", width=400)
# tab1LFrame20.grid(row=0, column=0, padx=8, pady=4, sticky=tk.W)
# ttk.Button(tab1LFrame20, text="클릭", command=ImageDoubleClick).pack()

# tab2VarStr1 = StringVar()
# tab2en1 = ttk.Entry(tab1LFrame20, text='', textvariable=tab2VarStr1)




# 종료 버튼 클릭시 System Tray로 이동
win.protocol('WM_DELETE_WINDOW', _exit)


win.mainloop()