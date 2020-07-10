# -*- coding: utf-8 -*-
"""
Created on Tue Sep 10 13:13:08 2019

@author: Qianwd
"""

from tkinter import *
import tkinter as tk
import tkinter.messagebox as messagebox
import pandas as pd
import random as rd
import xlwings as xw
import time,datetime

path="D:\\py_code\\word_gre.xlsx"
imgpath="D:\\py_code\\bdc\\picture\\background.png"

"""
窗口居中
"""
def widget_to_center(win, width, height):
    # 获取屏幕长/宽
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = sw / 2 - width / 2
    y = sh / 2 - height / 2
    win.geometry('%dx%d+%d+%d' % (width, height, x, y))

"""
建立单词库word_list
"""
class status:
    def __init__(self):
        self.RmbTime=datetime.datetime.now()
        self.position='learning'
    
    
        
class word:
    def __init__(self,English="",Chinese="",Code=0,rmbtime=0):
        self.English=English
        self.Chinese=Chinese
        self.Code=Code
        self.Status=status()
        self.Status.RmbTime=rmbtime

class RunTest:
    def __init__(self,mode,path,frame,ExcelSheet):
        self.state="English"
        self.currentIndex=0
        self.undoword=word()
        self.mode=mode
        self.frame=frame
        self.path=path
        self.word_list=[]
        self.read()
        self.start()
        self.ExcelSheet=ExcelSheet
    
    def read(self):
        wdata=pd.read_excel(io=self.path,sheet_name=0)
        wdata.index=range(1,len(wdata.index)+1)
        for index in wdata.index:
            currentWord=word(wdata.loc[index].word,wdata.loc[index].chinese,index,wdata.loc[index].rmbtime)
            self.word_list.append(currentWord) 
    
    def start(self):
        if self.mode=="FastMode":
            self.FastModeTest()
        else:
            self.SelectModeTest()
    
    def FastModeTest(self):
        TestList=[]
        text=tk.StringVar()
        if len(self.word_list)>=50:
            for count in range(1,51):
                index=rd.randint(0,len(self.word_list)-1)
                TestList.append(self.word_list.pop(index))
        elif len(self.word_list) != 0:
            for count in range(1,len(self.word_list)+1):
                index=rd.randint(0,len(self.word_list)-1)
                TestList.append(self.word_list.pop(index))
        else:
            messagebox.showwarning(title="提示",message="单词表中的单词已经学完了！")
            return 
        
        self.currentIndex=rd.randint(0,len(TestList)-1)
        text.set(TestList[self.currentIndex].English)
        self.state="English"

        def Yes():          
            self.undoword=TestList[self.currentIndex]
            text.set(TestList.pop(self.currentIndex).Chinese)
            BtnYes.place_forget()
            BtnNo.place_forget()
            BtnUndo.place(x=650,y=70,width=100,height=20)
            BtnNext.place(x=300,y=450,width=200,height=50)
            self.state="Chinese"
        
        def Next():
            if len(TestList)>0:
                self.currentIndex=rd.randint(0,len(TestList)-1)
                text.set(TestList[self.currentIndex].English)
                BtnNext.place_forget()
                BtnUndo.place_forget()
                BtnYes.place(x=100,y=450,width=200,height=50)
                BtnNo.place(x=500,y=450,width=200,height=50)
                self.state="English"
            else:
                text.set("已完成")
                BtnNext.place_forget()
                BtnUndo.place_forget()
                self.state="Finish"
                messagebox.showinfo(title="提示",message="按下回车再来一组")
        
        def No():
            text.set(TestList[self.currentIndex].Chinese)
            BtnYes.place_forget()
            BtnNo.place_forget()
            BtnNext.place(x=300,y=450,width=200,height=50)
            self.state="Chinese"
        
        def Undo():
            TestList.append(self.undoword)
            self.currentIndex=rd.randint(0,len(TestList)-1)
            text.set(TestList[self.currentIndex].English)
            BtnNext.place_forget()
            BtnUndo.place_forget()
            BtnYes.place(x=100,y=450,width=200,height=50)
            BtnNo.place(x=500,y=450,width=200,height=50)
            self.state="English"
        
        def KeyEvent(event):
            if event.keysym == '1':
                if self.state=="English":
                    Yes()
            elif event.keysym == '2':
                if self.state=="English":
                    No()
            elif event.keysym == 'F2':
                if self.state=="Chinese":
                    Undo()
            elif event.keysym == 'Right':
                if self.state=="Chinese":
                    Next()
            elif event.keysym =='Return':
                if self.state=="Finish":
                    self.FastModeTest()
            else:
                pass

                
        BtnYes=tk.Button(self.frame,text="认识",command=Yes)
        BtnNo=tk.Button(self.frame,text="不认识",command=No)
        BtnUndo=tk.Button(self.frame,text="撤销(F2)",command=Undo)
        BtnNext=tk.Button(self.frame,text="下一个",command=Next)
        BtnYes.place(x=100,y=450,width=200,height=50)
        BtnNo.place(x=500,y=450,width=200,height=50)
        Label=tk.Label(self.frame,textvariable=text,font=("Times",30))
        Label.place(x=50,y=100,width=700,height=200)
        """
        键盘控制
        """
        self.frame.bind_all("<KeyPress-1>",KeyEvent)
        self.frame.bind_all("<KeyPress-2>",KeyEvent)
        self.frame.bind_all("<KeyPress-F2>",KeyEvent)
        self.frame.bind_all("<KeyPress-Right>",KeyEvent)
        self.frame.bind_all("<KeyPress-Return>",KeyEvent)


"""
根目录下的三个按钮：分别是模式选择和退出键
"""
class MyApp:
    def __init__(self,parent):
        #photo=tk.PhotoImage(file=imgpath)
        self.root = parent
        self.root.title("要你命3000")
        self.frame = tk.Frame(parent,width=800,height=600)
        self.frame.pack()
        RootWidth=800
        RootHeight=600
        ButtonNum=3
        ButtonWidth=200
        ButtonHeight=50
        rootCb1 = tk.Button(self.frame, text = "快速模式", command = self.FastModeWindow)
        rootCb2 = tk.Button(self.frame, text = "选择模式", command = self.SelectModeWindow)
        rootCb = tk.Button(self.frame, text="退出", command=self.root.destroy)
        rootCb1.place(x=(RootWidth-ButtonWidth)/2,y=(RootHeight-ButtonHeight*ButtonNum)/(ButtonNum+3)*2,width=ButtonWidth,height=ButtonHeight)
        rootCb2.place(x=(RootWidth-ButtonWidth)/2,y=(RootHeight-ButtonHeight*ButtonNum)/(ButtonNum+3)*3+ButtonHeight,width=ButtonWidth,height=ButtonHeight)
        rootCb.place(x=(RootWidth-ButtonWidth)/2,y=(RootHeight-ButtonHeight*ButtonNum)/(ButtonNum+3)*4+ButtonHeight*2,width=ButtonWidth,height=ButtonHeight)
        """
        版本号
        """
        versLabel=tk.Label(self.frame,text="V0.0.2",fg="red")
        versLabel.place(x=(RootWidth-50)/2,y=RootHeight-15,width=50,height=10)
        """
        打开excel，进行数据读写
        """
        app=xw.App(visible=False,add_book=False)
        self.ExcelSheet=app.books.open(path)
        
    def hide(self):
        self.root.withdraw()
        
    def FastModeWindow(self):
        self.hide()
        otherFrame = tk.Toplevel()
        widget_to_center(otherFrame, 800, 600)
        otherFrame.title("快速模式")
        handler = lambda: self.onCloseOtherFrame(otherFrame)
        otherFrame.protocol("WM_DELETE_WINDOW", handler)
        run = RunTest(path=path,mode="FastMode",frame=otherFrame,ExcelSheet=self.ExcelSheet)
    
    def SelectModeWindow(self):
        self.hide()
        otherFrame = tk.Toplevel()
        widget_to_center(otherFrame, 800, 600)
        otherFrame.title("选择模式")
        handler = lambda: self.onCloseOtherFrame(otherFrame)
        otherFrame.protocol("WM_DELETE_WINDOW", handler)
    
    def onCloseOtherFrame(self, otherFrame):
        judge=messagebox.askokcancel(title="提示",message="是否确定要退出")
        if judge==True:
            otherFrame.destroy()
            self.show()
    
    def show(self):
        self.root.update()
        self.root.deiconify()
    
if __name__ == "__main__":
    word_list=[]
    RootWidth=800
    RootHeight=600
    root = tk.Tk()
    widget_to_center(root, 800, 600)
    app = MyApp(root)
    root.mainloop()
    
    
