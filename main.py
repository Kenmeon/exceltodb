from tkinter import *
import tkinter.filedialog
from tkinter import scrolledtext

from utils import ImportMain
import tkinter.messagebox

root = tkinter.Tk()


def get_screen_size(window):
    return window.winfo_screenwidth(), window.winfo_screenheight()


def get_window_size(window):
    return window.winfo_reqwidth(), window.winfo_reqheight()


def center_window(root, width, height):
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    print(size)
    root.geometry(size)

def xz():
    filenames = tkinter.filedialog.askopenfilenames()

    if len(filenames) != 0:
        lb.delete(1.0, tkinter.END)
        lb.insert(1.0, '您选择的文件是：\n')
        # selectFiles = []
        for i in range(0,len(filenames)):
            lb.insert(tkinter.END, str(filenames[i]) + '\n')
            selectFiles.append(str(filenames[i]))
            print(selectFiles)

    else:
        lb.insert(tkinter.END, '您没有选择任何文件\n')

def importdb():
    if selectFiles == None or len(selectFiles) == 0:
        tkinter.messagebox.showinfo('提示', '请选择需要导入的文件')
        return
    if dbnameIn.get() == None or dbnameIn.get() == '':
        tkinter.messagebox.showinfo('提示', '请输入数据库用户名')
        return
    if pwdIn.get() == None or pwdIn.get() == '':
        tkinter.messagebox.showinfo('提示', '请输入数据库密码')
        return
    if ipIn.get() == None or ipIn.get() == '':
        tkinter.messagebox.showinfo('提示', '请输入数据库地址')
        return
    if instanceIn.get() == None or instanceIn.get() == '':
        tkinter.messagebox.showinfo('提示', '请输入数据库实例名')
        return


    # 执行导入数据库的方法
    for i in range(0, len(selectFiles)):
        impr = ImportMain()
        impr.name = dbnameIn.get()
        impr.password = pwdIn.get()
        impr.ip = ipIn.get()
        impr.instance = instanceIn.get()
        impr.filePath = selectFiles[i]
        code = impr.run()
        if code == 1:
            lb.insert(tkinter.END,selectFiles[i] + '导入成功\n')
        else:
            lb.insert(tkinter.END, selectFiles[i] + '导入失败\n')

# 数据库名

group = tkinter.Frame(root).grid(row = 0,column = 0, sticky=tkinter.W)

dbname = tkinter.Label(group,text='数据库名:',anchor='c').grid(row = 0,column = 0, sticky=tkinter.S)

dbnameIn = tkinter.Entry(group)
dbnameIn.grid(row = 0,column = 1, sticky=tkinter.W)

pwd = tkinter.Label(group,text = '密码:',anchor='c').grid(row = 0,column = 2, sticky=tkinter.S)
pwdIn = tkinter.Entry(group)
pwdIn.grid(row = 0,column = 3, sticky=tkinter.W)

#地址
group2 = tkinter.Frame(root).grid(row = 1,column = 0, sticky=tkinter.W)

ip = tkinter.Label(group2,text='IP地址:',anchor='c').grid(row = 1,column = 0, sticky=tkinter.S)

ipIn = tkinter.Entry(group2)
ipIn.grid(row = 1,column = 1, sticky=tkinter.W)

instance = tkinter.Label(group2,text = '实例名:',anchor='c').grid(row = 1,column = 2, sticky=tkinter.S)
instanceIn = tkinter.Entry(group2)
instanceIn.grid(row = 1,column = 3, sticky=tkinter.W)

root.columnconfigure(0,weight = 1)
root.columnconfigure(1,weight = 1)
root.columnconfigure(2,weight = 1)
root.columnconfigure(3,weight = 1)




btnGroup = tkinter.LabelFrame(root,bg = 'black',pady = 15).grid(row = 3,columnspan = 4)
btn = tkinter.Button(btnGroup,text="选择文件",command=xz).grid(row = 3,column = 0,columnspan = 2)
btn = tkinter.Button(btnGroup,text="导入Oracle",command=importdb).grid(row = 3,column = 2,columnspan = 2)
print(type(btn))

margine = LabelFrame(root,height = 10)
margine.grid(row = 4)

lb = scrolledtext.ScrolledText(root,width = 600,height = 50)
lb.grid(row = 5,column = 0,columnspan = 4, sticky=tkinter.W+tkinter.S+tkinter.N)
lb.insert('1.0','未选择文件\n')
# lb.configure(bg = 'green')
lb.configure(state = DISABLED)
print(type(lb))

#制定定义在这里，定义在方法只是不可以
selectFiles = []
# root的设置
root.title('Excel导入Oracle工具')
center_window(root, 600, 700)

root.mainloop()