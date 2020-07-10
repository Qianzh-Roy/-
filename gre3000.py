# -*- coding: utf-8 -*-
"""
Created on Tue Sep 10 13:13:08 2019

@author: Qianwd
"""

from tkinter import *
import tkinter as tk
import tkinter.messagebox as messagebox
import numpy as np
import pandas as pd
import random as rd
import xlwings as xw
import time,datetime
import math

path="D:\\py_code\\word_gre_0918re.xlsx"
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
        self.RmbTime=0
        self.times=0
        self.position='new'
    
    
        
class word:
    def __init__(self,English="",Chinese="",Code=0,rmbtime=0):
        self.English=English
        self.Chinese=Chinese
        self.Code=Code
        self.Status=status()
        self.Status.RmbTime=rmbtime

class RunTest:
    def __init__(self,mode,path,frame,ExcelSheet):
        self.FinishWord=int()
        self.LearnWord=int()
        self.PlanWord=int()
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
            currentWord=word(English=wdata.loc[index].word,Chinese=wdata.loc[index].chinese,Code=index,rmbtime=wdata.loc[index].rmbtime)
            self.word_list.append(currentWord) 
    
    def start(self):
        if self.mode=="FastMode":
            self.FastModeTest()
        else:
            self.SelectModeTest()
    
    def FastModeTest(self):
        num=10
        TestList=[]
        text=tk.StringVar()
        lbtext=tk.StringVar()
        fbtext=tk.StringVar()
        pbtext=tk.StringVar()
        #self.LearnWord=0
        #self.FinishWord=0
        #self.PlanWord=0
        def FormTestList(num):
            TestList=[]
            index=0
            while index<len(self.word_list):
                timedelta=datetime.datetime.now()-self.word_list[index].Status.RmbTime
                if math.isnan(timedelta.days) and len(TestList)<num:
                    TestList.append(self.word_list.pop(index))
                else:
                    index+=1
            while len(TestList)<num:    
                if len(self.word_list) != 0:
                    index=rd.randint(0,len(self.word_list)-1)
                    word2append=self.word_list.pop(index)
                    d1=datetime.datetime.now()
                    timedelta=d1-word2append.Status.RmbTime
                    if (timedelta.days>3) or math.isnan(timedelta.days) or timedelta.days==0:
                        TestList.append(word2append)
                else:
                    return TestList
            return TestList
        
        TestList=FormTestList(num)
        """
        """
        for word in TestList:
            print(word.English)
        
        if TestList != [] :
            self.PlanWord=len(TestList)
            self.currentIndex=rd.randint(0,len(TestList)-1)
            text.set(TestList[self.currentIndex].English)
            self.state="English"
        else:
            messagebox.showwarning(title="提示",message="单词表中的单词已经学完了！")
            return
        
        def Yes():
            if TestList[self.currentIndex].Status.position=="new":
                self.FinishWord+=1
                self.undoword=TestList[self.currentIndex]
                self.ExcelSheet.sheets['sheet1'].range("D"+str(TestList[self.currentIndex].Code+1)).value='1'
                self.ExcelSheet.sheets['sheet1'].range("C"+str(TestList[self.currentIndex].Code+1)).value=datetime.datetime.now().strftime("%Y/%m/%d")
                text.set(self.undoword.English+'\n'+TestList.pop(self.currentIndex).Chinese)
            else:
                TestList[self.currentIndex].Status.times-=1
                if TestList[self.currentIndex].Status.times==0:
                    self.ExcelSheet.sheets['sheet1'].range("C"+str(TestList[self.currentIndex].Code+1)).value=datetime.datetime.now().strftime("%Y/%m/%d")
                    self.undoword=TestList.pop(self.currentIndex)
                    self.undoword.Status.position='new'
                    text.set(self.undoword.English+'\n'+self.undoword.Chinese)
                    self.FinishWord+=1
                    self.LearnWord-=1
                else:
                    text.set(TestList[self.currentIndex].English+'\n'+TestList[self.currentIndex].Chinese)
            BtnYes.place_forget()
            BtnNo.place_forget()
            BtnUndo.place(x=650,y=70,width=100,height=20)
            BtnNext.place(x=300,y=450,width=200,height=50)
            lbtext.set(str(self.LearnWord))
            fbtext.set(str(self.FinishWord))
            pbtext.set(str(self.PlanWord-self.LearnWord-self.FinishWord))
            LearnBar.place(x=750-self.LearnWord/self.PlanWord*700,y=50,width=self.LearnWord/self.PlanWord*700,height=15)
            FinishBar.place(x=50,y=50,width=self.FinishWord/self.PlanWord*700,height=15)
            PlanBar.place(x=50+self.FinishWord/self.PlanWord*700,y=50,width=(self.PlanWord-self.LearnWord-self.FinishWord)/self.PlanWord*700,height=15)
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
            if TestList[self.currentIndex].Status.position=="new":
                TestList[self.currentIndex].Status.position="learning"
                self.LearnWord+=1
            
            TestList[self.currentIndex].Status.times=3
            text.set(TestList[self.currentIndex].English+'\n'+TestList[self.currentIndex].Chinese)
            BtnYes.place_forget()
            BtnNo.place_forget()
            BtnNext.place(x=300,y=450,width=200,height=50)
            lbtext.set(str(self.LearnWord))
            fbtext.set(str(self.FinishWord))
            pbtext.set(str(self.PlanWord-self.LearnWord-self.FinishWord))
            LearnBar.place(x=750-self.LearnWord/self.PlanWord*700,y=50,width=self.LearnWord/self.PlanWord*700,height=15)
            FinishBar.place(x=50,y=50,width=self.FinishWord/self.PlanWord*700,height=15)
            PlanBar.place(x=50+self.FinishWord/self.PlanWord*700,y=50,width=(self.PlanWord-self.LearnWord-self.FinishWord)/self.PlanWord*700,height=15)
            self.state="Chinese"
        
        def Undo():
            if self.undoword.Status.position=="new":
                self.undoword.Status.position="learning"
                self.FinishWord-=1
                self.LearnWord+=1
                TestList.append(self.undoword)
                self.ExcelSheet.sheets['sheet1'].range("D"+str(self.undoword.Code+1)).value=""
                self.ExcelSheet.sheets['sheet1'].range("C"+str(self.undoword.Code+1)).value=""
            self.undoword.Status.times=3
            self.currentIndex=rd.randint(0,len(TestList)-1)
            text.set(TestList[self.currentIndex].English)
            BtnNext.place_forget()
            BtnUndo.place_forget()
            BtnYes.place(x=100,y=450,width=200,height=50)
            BtnNo.place(x=500,y=450,width=200,height=50)
            lbtext.set(str(self.LearnWord))
            fbtext.set(str(self.FinishWord))
            pbtext.set(str(self.PlanWord-self.LearnWord-self.FinishWord))
            LearnBar.place(x=750-self.LearnWord/self.PlanWord*700,y=50,width=self.LearnWord/self.PlanWord*700,height=15)
            FinishBar.place(x=50,y=50,width=self.FinishWord/self.PlanWord*700,height=15)
            PlanBar.place(x=50+self.FinishWord/self.PlanWord*700,y=50,width=(self.PlanWord-self.LearnWord-self.FinishWord)/self.PlanWord*700,height=15)
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
                    self.LearnWord=self.FinishWord=0
                    self.FastModeTest()
            else:
                pass
            
        BtnYes=tk.Button(self.frame,text="认识",command=Yes)
        BtnNo=tk.Button(self.frame,text="不认识",command=No)
        BtnUndo=tk.Button(self.frame,text="撤销(F2)",command=Undo)
        BtnNext=tk.Button(self.frame,text="下一个",command=Next)
        PlanBar=tk.Label(self.frame,textvariable=pbtext,bg="gray")
        LearnBar=tk.Label(self.frame,textvariable=lbtext,bg="red")
        FinishBar=tk.Label(self.frame,textvariable=fbtext,bg="green")
        BtnYes.place(x=100,y=450,width=200,height=50)
        BtnNo.place(x=500,y=450,width=200,height=50)
        Label=tk.Label(self.frame,textvariable=text,font=("Times",25),wraplength=575)
        Label.place(x=100,y=100,width=600,height=300)
        pbtext.set(self.PlanWord)
        PlanBar.place(x=50,y=50,width=700,height=15)
        
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
        #self.frame.pack(fill="both",expand="yes") 
        RootWidth=800
        RootHeight=600
        ButtonNum=3
        ButtonWidth=200
        ButtonHeight=50
        """
        主界面背景图片及控件设置
        """
        photo=tk.PhotoImage(file=imgpath)
        panel=tk.Label(self.frame,image=photo,compound=tk.CENTER,width=800,height=600)
        rootCb1 = tk.Button(self.frame, text = "快速模式", command = self.FastModeWindow)
        rootCb2 = tk.Button(self.frame, text = "选择模式", state="disabled",command = self.SelectModeWindow)
        rootCb = tk.Button(self.frame, text="退出", command=self.exit)
        panel.pack()
        panel.image=photo
        
        rootCb1.place(x=(RootWidth-ButtonWidth)/2,y=(RootHeight-ButtonHeight*ButtonNum)/(ButtonNum+3)*2,width=ButtonWidth,height=ButtonHeight)
        rootCb2.place(x=(RootWidth-ButtonWidth)/2,y=(RootHeight-ButtonHeight*ButtonNum)/(ButtonNum+3)*3+ButtonHeight,width=ButtonWidth,height=ButtonHeight)
        rootCb.place(x=(RootWidth-ButtonWidth)/2,y=(RootHeight-ButtonHeight*ButtonNum)/(ButtonNum+3)*4+ButtonHeight*2,width=ButtonWidth,height=ButtonHeight)
        """
        版本号
        """
        versLabel=tk.Label(self.frame,text="V0.1.0",fg="red")
        versLabel.place(x=0,y=RootHeight-10,width=800,height=10)
        
        """
        打开excel，进行数据读写
        """
        self.app=xw.App(visible=False,add_book=False)
        self.ExcelSheet=self.app.books.open(path)
        handler=lambda: self.onCloseFrame(self.root)
        self.root.protocol("WM_DELETE_WINDOW", handler)
    
    def onCloseFrame(self,Frame):
        judge=messagebox.askokcancel(title="提示",message="是否确定要退出")
        if judge==True:
            Frame.destroy()
        self.app.kill()
    
    def exit(self):
        self.app.kill()
        self.root.destroy()
        
    def hide(self):
        self.root.withdraw()
        
    def FastModeWindow(self):
        self.hide()
        otherFrame = tk.Toplevel()
        widget_to_center(otherFrame, 800, 600)
        otherFrame.title("快速模式")
        """
        
        """
        photo=tk.PhotoImage(file=imgpath)
        panel=tk.Label(otherFrame,image=photo,compound=tk.CENTER,width=800,height=600)
        
        handler = lambda: self.onCloseOtherFrame(otherFrame)
        otherFrame.protocol("WM_DELETE_WINDOW", handler)
        run = RunTest(path=path,mode="FastMode",frame=otherFrame,ExcelSheet=self.ExcelSheet)
        panel.pack()
        panel.image=photo
        
    
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
        self.ExcelSheet.save()
        
    
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
    
    
