import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog
from read_file import ReadExcel, ReadError, GradeError, ComprehensionError, ModelError, WeChatError
import os
import openpyxl


class App:
    def __init__(self, root):
        self.root = root
        # self.var = tk.StringVar()
        self.list = None
        self.sign = ""
        self.text = None
        self.data = None

        self.init_window()

    def init_window(self):
        self.root.title('自动回访')
        self.root.geometry('800x400+%d+%d' % (self.root.winfo_screenwidth()/2 - 300, self.root.winfo_screenheight()/2 - 200))
        self.root.resizable(width=False, height=False)

        # self.var.set("请导入文件")

        # 创建菜单栏
        # menubar = tk.Menu(self.root)
        # filemenu = tk.Menu(menubar, tearoff=0)
        # menubar.add_cascade(label='File', menu=filemenu)
        # filemenu.add_command(label="导入文件", command=self.import_file)
        # self.root.config(menu=menubar)

        frame1 = tk.Frame(self.root)
        frame1.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
        frame2 = tk.Frame(self.root)
        frame2.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
        frame3 = tk.Frame(frame2)
        frame3.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=1)

        # tk.Label(frame1, textvariable=self.var, font=('Arial', 12), height=2, width=2, bg="white").pack(fill=tk.X, pady=2)
        self.list = tk.Listbox(frame1)
        self.list.pack(fill=tk.BOTH, expand=1)
        self.list.bind("<ButtonRelease-1>", self.refresh_data)

        # 创建并添加滚轮
        scroll = tk.Scrollbar(self.list, command=self.list.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.list.config(yscrollcommand=scroll.set)

        self.text = tk.Text(frame2)
        self.text.pack(side=tk.TOP, fill=tk.BOTH, expand=1)

        tk.Button(frame3, text="导入文件", command=self.import_file, height=1, width=10).pack(side=tk.LEFT, expand=1)
        tk.Button(frame3, text="发送", command=self.send, height=1, width=10).pack(side=tk.LEFT, expand=1)
        tk.Button(frame3, text="自动发送", command=self.auto_send, height=1, width=10).pack(side=tk.LEFT, expand=1)

    def send(self):
        try:
            self.send_to_one(self.list.get(self.list.curselection()))
            self.list.itemconfigure(self.list.curselection(), background="red", foreground="white", selectbackground="red")
        except Exception:
            tk.messagebox.showerror(title="Error", message="请先导入Excel文件或选择需要发送信息的学生")

    def send_to_one(self, student):
        self.generate_applescript(student)
        self.run_applescript()

    def auto_send(self):
        sign = tk.messagebox.askquestion(title='Attention', message="是否开始自动发送信息？")
        try:
            if sign == "yes":
                if self.list.size() == 0:
                    raise Exception
                for index in range(self.list.size()):
                    student = self.list.get(index)
                    self.send_to_one(student=student)
                    self.list.itemconfigure(index, background="red", foreground="white", selectbackground="red")
        except Exception:
            tk.messagebox.showerror(title="Error", message="请先导入Excel文件")

    def import_file(self):
        # 选择文件，限制为xlsx格式
        filename = tkinter.filedialog.askopenfilename(filetypes=[("Excel 工作簿", "*.xlsx")], title="请选择xlsx文件")
        try:
            self.data = ReadExcel(filename)
        except ReadError as e:
            tk.messagebox.showerror(title='Error', message=e.__str__())
            return -1
        except GradeError as e:
            tk.messagebox.showerror(title='Error', message=e.__str__())
            return -1
        except ComprehensionError as e:
            tk.messagebox.showerror(title='Error', message=e.__str__())
            return -1
        except ModelError as e:
            tk.messagebox.showerror(title='Error', message=e.__str__())
            return -1
        except WeChatError as e:
            tk.messagebox.showerror(title='Error', message=e.__str__())
            return -1
        except openpyxl.utils.exceptions.InvalidFileException:
            return -1
        except Exception as e:
            tk.messagebox.showerror(title='Error', message=e.__str__())
            return -1

        # 清空列表
        self.list.delete(0, self.list.size()-1)
        for name in self.data.all_info.keys():
            self.list.insert(1, name)

    def refresh_data(self, event):
        try:
            name = self.list.get(self.list.curselection())
            # self.var.set(name)
        except Exception:
            return -1
        if self.sign != name:
            self.data.all_info[self.sign] = self.text.get("1.0", tk.END)
            self.sign = name
            self.text.delete("1.0", tk.END)
            self.text.insert(tk.END, self.data.all_info[name])

    def generate_applescript(self, name):
        script_model = """tell application "WeChat" to activate
tell application "System Events"
    tell process "WeChat"
        %s
    end tell
end tell"""
        temp = """set the clipboard to "%s"
        click menu item "查找…" of menu "编辑" of menu bar item "编辑" of menu bar 1
        key code 9 using {command down}
        delay 0.5
        key code 76
        key code 48 using {command down}
        delay 0.1
        key code 48 using {command down}
        delay 0.1
        set the clipboard to "%s"
        key code 9 using {command down}
        key code 76
        """ % (self.data.wechat[name], self.data.all_info[name])

        script_model = script_model % temp
        with open("temp.applescript", "w") as file:
            file.write(script_model)

    @staticmethod
    def run_applescript():
        os.system("osascript temp.applescript")



if __name__ == '__main__':
    root = tk.Tk()
    app = App(root=root)
    root.mainloop()