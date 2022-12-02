#from tkinter import *
import tkinter, win32api, win32con, pywintypes
#from pymouse import PyMouse
import time
import os

#import serial
import threading
import time
#from IPython.display import clear_output
import ctypes
import inspect

#线程结束的代码
def _async_raise(tid, exctype):
    tid = ctypes.c_long(tid)
    if not inspect.isclass(exctype):
        exctype = type(exctype)
    res = ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, ctypes.py_object(exctype))
    if res == 0:
        raise ValueError("invalid thread id")
    elif res != 1:
        ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, None)
        raise SystemError("PyThreadState_SetAsyncExc failed")
        
def stop_thread(thread):
    _async_raise(thread.ident, SystemExit)
    
    
root2=tkinter.Tk()
root2.attributes("-topmost", 1)
#对应图表视图的大小
frame2 =tkinter.Frame(root2, width=50, height=50, bg='grey')
tkinter.Label(root2,text='   中文  ',bg='grey').pack()
#mouse =PyMouse()

def get_mouse():
    #root=tkinter.Tk()
    #frame=tkinter.Frame(root, width=280, height=120)
    #root2.title("hello")
    #root2.geometry("+530+560")#图表视图的位置

    root2.geometry("+780+660")
    root2.overrideredirect(True) #隐藏显示框
    frame2.bind("<Button -1>", callback)
    frame2.pack()

    root2.mainloop()
    
def callback(event):
    #win = Tk()
    #width = win.winfo_screenwidth()
    #print('width: ',width)
    #height = win.winfo_screenheight()
    #print('height: ', height)
    #tkinter.Label(root2,text='   中文  ',bg='green').pack()
    """
    x_mouse=605#800
    y_mouse=630#800
    print("点击位置：", event.x, event.y)
    root2.destroy()
    mouse.move(x_mouse,y_mouse)
    p = mouse.position()
    print('p:'+str(p))
    #mouse.click(x_mouse,y_mouse)#click是双击
    mouse.press(x_mouse, y_mouse)
    mouse.release(x_mouse, y_mouse)
    """
    """while(True):
        p = mouse.position()
        print('p:'+str(p))
        time.sleep(1)"""
    five_window()
    
    """result = os.popen("tasklist")
    str_FTMS = "FTMS.exe"
    result = result.read()
    flag = 0
    if str_FTMS in result:
        print('--yes--')
        if(flag == 0):
            #t = Thread(target = five_window).start()
            print('flag == 0--')
            flag = 1
            t = threading.Thread(target=five_window)
            t.setDaemon(True)
            t.start()
    else:
        print('--no--')
        if(flag == 1):
            #t.stop()
            print('flag == 1--')
            flag = 0
            stop_thread(t)"""

def five_window():
    root = tkinter.Tk() 

    #width = win32api.GetSystemMetrics(0)
    #height = win32api.GetSystemMetrics(1)
    root.overrideredirect(True) #隐藏显示框
    root.lift() #置顶层
    root.wm_attributes("-topmost", True) #始终置顶层
    root.wm_attributes("-disabled", True)
    root.wm_attributes("-transparentcolor", "white")#白色背景透明
    hWindow = pywintypes.HANDLE(int(root.frame(), 16))
    exStyle = win32con.WS_EX_COMPOSITED | win32con.WS_EX_LAYERED | win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOPMOST | win32con.WS_EX_TRANSPARENT
    win32api.SetWindowLong(hWindow, win32con.GWL_EXSTYLE, exStyle)

    mytext1 = "  图例  " +"\n" + "摩擦系数(跑道1)" +"\n"  +"摩擦系数(跑道2)" +"\n"+"跑道维护值        \n" + "最小允许值        \n"
    mytext1 += "水平拉力(跑道1)" + "\n" + "垂直拉力(跑道1)" + "\n" + "水平拉力(跑道2)" + "\n" + "垂直拉力(跑道2)"
    #mytext1 = "图例"
    window = tkinter.Toplevel()
    window.geometry("+1170+30")#这里400是到屏幕左边的距离，100是到屏幕上边的距离，，如果把加号换成减号，就变成距右边的距离和距下边的距离
    window.overrideredirect(True) #隐藏显示框
    window.lift() #置顶层
    window.wm_attributes("-topmost", True) #始终置顶层
    window.wm_attributes("-disabled", True)
    window.wm_attributes("-transparentcolor", "white")#白色背景透明
    #label2 = tkinter.Label(window, text=mytext1, compound = 'left',height = 10, width=14, font=('Times New Roman','12'), fg='#d5d5d5', bg='black').pack()
    label2 = tkinter.Label(window, text="图例", compound = 'left',height = 1, width=14, font=('Times New Roman','10'), fg='#d5d5d5', bg='black').pack()
    label2 = tkinter.Label(window, text="摩擦系数(跑道1)", compound = 'left',height = 1, width=14, font=('Times New Roman','10'), fg='#d5d5d5', bg='black').pack()
    label2 = tkinter.Label(window, text="摩擦系数(跑道2)", compound = 'left',height = 1, width=14, font=('Times New Roman','10'), fg='#d5d5d5', bg='black').pack()
    label2 = tkinter.Label(window, text="跑道维护值", compound = 'left',height = 1, width=14, font=('Times New Roman','10'), fg='#d5d5d5', bg='black').pack()
    label2 = tkinter.Label(window, text="最小允许值", compound = 'left',height = 1, width=14, font=('Times New Roman','10'), fg='#d5d5d5', bg='black').pack()
    label2 = tkinter.Label(window, text="水平拉力(跑道1)", compound = 'left',height = 1, width=14, font=('Times New Roman','10'), fg='#d5d5d5', bg='black').pack()
    label2 = tkinter.Label(window, text="垂直拉力(跑道1)", compound = 'left',height = 1, width=14, font=('Times New Roman','10'), fg='#d5d5d5', bg='black').pack()
    label2 = tkinter.Label(window, text="水平拉力(跑道2)", compound = 'left',height = 1, width=14, font=('Times New Roman','10'), fg='#d5d5d5', bg='black').pack()
    label2 = tkinter.Label(window, text="垂直拉力(跑道2)", compound = 'left',height = 1, width=14, font=('Times New Roman','10'), fg='#d5d5d5', bg='black').pack()
    
    #右侧的几个颜色条
    window1 = tkinter.Toplevel()
    window1.overrideredirect(True) #隐藏显示框
    window1.lift() #置顶层
    window1.wm_attributes("-topmost", True) #始终置顶层
    window1.wm_attributes("-disabled", True)
    
    window1.geometry("+1140+50")#这里400是到屏幕左边的距离，100是到屏幕上边的距离，，如果把加号换成减号，就变成距右边的距离和距下边的距离
    label2 = tkinter.Label(window1, text="", compound = 'left',height = 1, width=2, font=('Times New Roman','10'), fg='#d5d5d5', bg='#008000').pack()

    label2 = tkinter.Label(window1, text="", compound = 'left',height = 1, width=2, font=('Times New Roman','10'), fg='#d5d5d5', bg='#ff0000').pack()
    label2 = tkinter.Label(window1, text="", compound = 'left',height = 1, width=2, font=('Times New Roman','10'), fg='#d5d5d5', bg='#9966ff').pack()
    label2 = tkinter.Label(window1, text="", compound = 'left',height = 1, width=2, font=('Times New Roman','10'), fg='#d5d5d5', bg='#ffff33').pack()
    label2 = tkinter.Label(window1, text="", compound = 'left',height = 1, width=2, font=('Times New Roman','10'), fg='#d5d5d5', bg='#ffcc99').pack()
    label2 = tkinter.Label(window1, text="", compound = 'left',height = 1, width=2, font=('Times New Roman','10'), fg='#d5d5d5', bg='#ff33ff').pack()
    label2 = tkinter.Label(window1, text="", compound = 'left',height = 1, width=2, font=('Times New Roman','10'), fg='#d5d5d5', bg='#ffffff').pack()
    label2 = tkinter.Label(window1, text="", compound = 'left',height = 1, width=2, font=('Times New Roman','10'), fg='#d5d5d5', bg='#ffff66').pack()
     

    #mytext2 = "       摩擦系数测试  "
    mytext2 = "摩擦系数测试"
    window = tkinter.Toplevel()
    window.geometry("+400+85")#这里400是到屏幕左边的距离，100是到屏幕上边的距离，，如果把加号换成减号，就变成距右边的距离和距下边的距离
    window.overrideredirect(True) #隐藏显示框
    window.lift() #置顶层
    window.wm_attributes("-topmost", True) #始终置顶层
    window.wm_attributes("-disabled", True)
    window.wm_attributes("-transparentcolor", "white")#白色背景透明
    #tkinter.Label(window, text=mytext2, compound = 'left',height = 5, width=15, font=('Times New Roman','15'), fg='#ffff00', bg='black').pack()
    label2 = tkinter.Label(window, text=mytext2, compound = 'left',height = 1, width=12, font=('Times New Roman','15'), fg='#ffff00', bg='black').pack()
    
    mytext3 = "距离(米)"
    window = tkinter.Toplevel()
    window.geometry("+572+590")
    window.overrideredirect(True) #隐藏显示框
    window.lift() #置顶层
    window.wm_attributes("-topmost", True) #始终置顶层
    window.wm_attributes("-disabled", True)
    window.wm_attributes("-transparentcolor", "white")#白色背景透明
    tkinter.Label(window, text=mytext3, compound = 'left',height = 1, width=12, font=('Times New Roman','15'), fg='#ffff00', bg='black').pack()
    
    mytext4 = " 平均摩擦系数："
    window = tkinter.Toplevel()
    #window.geometry("+822+590")
    window.geometry("+816+590")
    window.overrideredirect(True) #隐藏显示框
    window.lift() #置顶层
    window.wm_attributes("-topmost", True) #始终置顶层
    window.wm_attributes("-disabled", True)
    window.wm_attributes("-transparentcolor", "white")#白色背景透明
    tkinter.Label(window, text=mytext4, compound = 'left',height = 1, width=12, font=('Times New Roman','15'), fg='#ffff00', bg='black').pack()
    
    mytext6 = "摩\n擦\n系\n数"#左侧第二个黑色竖条
    window = tkinter.Toplevel()
    window.geometry("+5+320")
    window.overrideredirect(True) #隐藏显示框
    window.lift() #置顶层
    window.wm_attributes("-topmost", True) #始终置顶层
    window.wm_attributes("-disabled", True)
    window.wm_attributes("-transparentcolor", "white")#白色背景透明
    tkinter.Label(window, text=mytext6, compound = 'left',height = 4, width=2, font=('Times New Roman','15'),fg='#008000', bg='black').pack()
    
    #mytext7 = "速\n度"#左侧第一个黑色竖条
    """mytext7 = ""#左侧第一个黑色竖条
    window = tkinter.Toplevel()
    window.geometry("+8+265")
    window.overrideredirect(True) #隐藏显示框
    window.lift() #置顶层
    window.wm_attributes("-topmost", True) #始终置顶层
    window.wm_attributes("-disabled", True)
    window.wm_attributes("-transparentcolor", "white")#白色背景透明
    tkinter.Label(window, text=mytext7, compound = 'left',height = 4, width=2, font=('Times New Roman','15'),fg='#ff33ff', bg='black').pack()
    """
    #frame=Frame(root, width=1280, height=720)
    #frame.bind("<Button -1>", callback)
    #frame.pack()

    root.mainloop() #循环
    
get_mouse()