import tkinter as tk 
import tkinter.messagebox
import tkinter.filedialog
from read_file import ReadExcel
import os


class App:
    def __init__(self, root):
        self.root = root
        self.var = tk.StringVar()
        self.names = tk.StringVar()
        self.list = None
        self.sign = ""
        self.text = None
        self.data = None

        self.init_window()
        self.refresh_data()

    def init_window(self):
        self.root.title('My Window')
        self.root.geometry('1000x500')

        self.var.set("请导入文件")

        # 创建菜单栏
        menubar = tk.Menu(self.root)
        filemenu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='File', menu=filemenu)
        filemenu.add_command(label="导入文件", command=self.import_file)
        self.root.config(menu=menubar)

        frame1 = tk.Frame(self.root, width=200)
        frame1.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
        frame2 = tk.Frame(self.root)
        frame2.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        tk.Label(frame1, textvariable=self.var, font=('Arial', 12), height=2, width=2, bg="white").pack(fill=tk.X, pady=2)
        self.list = tk.Listbox(frame1)
        self.list.pack(fill=tk.BOTH, expand=1, pady=2)
        self.text = tk.Text(frame2)
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
        self.text.bind_all("<Button-1>", self.refresh_data)

        tk.Button(frame2, text="Text", command=self.button).pack(fill=tk.BOTH, expand=1)

    def button(self):
        self.generate_applescript(self.data.wechat[self.list.get(self.list.curselection())])
        self.root.clipboard_clear()
        self.root.clipboard_append(self.text.get("1.0", tk.END))
        self.run_applescript()

    def import_file(self):
        filename = tkinter.filedialog.askopenfilename()
        try:
            self.data = ReadExcel(filename)
        except Exception as e:
            tk.messagebox.showwarning(title='Error', message=e.__str__())
            return -1
        names = ""
        for name in self.data.grade.keys():
            self.list.insert(1, name)
        self.names.set(names)

    def refresh_data(self, *args):
        try:
            name = self.list.get(self.list.curselection())
            self.var.set(name)
        except Exception:
            return -1
        if self.sign != name:
            self.sign = name
            self.text.delete("1.0", tk.END)
            self.text.insert(tk.END, self.data.all_info[name])

    def generate_applescript(self, wechat):
        script_model = """tell application "WeChat" to activate
tell application "System Events"
    tell process "WeChat"
        click menu item "查找…" of menu "编辑" of menu bar item "编辑" of menu bar 1
        keystroke "%s"
        delay 0.5
        key code 76
        key code 48 using {command down}
        key code 48 using {command down}
        key code 9 using {command down}
        -- key code 76
    end tell
end tell""" % wechat
        with open("temp.applescript", "w") as file:
            file.write(script_model)

    def run_applescript(self):
        os.system("osascript temp.applescript")



if __name__ == '__main__':
    root = tk.Tk()
    app = App(root=root)
    root.mainloop()

