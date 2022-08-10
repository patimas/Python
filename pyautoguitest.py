import pyautogui as pygui
from PIL import ImageDraw

import tkinter as tk
from tkinter import ttk


############################################################
#
#  GUI 기본값 설정
#
############################################################
win = tk.Tk()
win.title("Python 기본")

# 가로x세로+x축+y축
# win.geometry("1440x800+100+100")
# win.resizable(False,False)

imgPath = "./img"
serFile1 = imgPath+"/img1.png"
serFile2 = imgPath+"/img2.png"
serFile3 = imgPath+"/img3.png"

result = pygui.locateOnScreen(serFile1, confidence = 0.9)


# if result is None:
#     print("No search!!")
# else:
#     print(pygui.center(result))

result = pygui.locateCenterOnScreen(serFile1, confidence=0.9)

# 대상 에서 이미지 찾을때 (비활성 매크로 시 사용)
result = pygui.locate(serFile1, pygui.screenshot(), confidence=0.9)

#####################################################################

screenImage = pygui.screenshot()
result = pygui.locate(serFile3, screenImage, confidence=0.9)

# draw = ImageDraw.Draw(screenImage)
# draw.rectangle((result[0], result[1], result[0] + result[2], result[1] + result[3]), outline=(255, 0, 0), width=5)
# screenImage.show()

def ImageDoubleClick():

    if result is None :
        print("없음")
    else :
        print(result[0])
        print(result[1])
        pygui.doubleClick(result[0], result[1])


ttk.Button(win, text="클릭", command=ImageDoubleClick).pack()


win.mainloop()