from read_file import ReadExcel, ReadError, GradeError, ComprehensionError, ModelError, WeChatError
import openpyxl
import os
from pathlib import Path


class CLI:
    def __init__(self):
        self.file = None
        self.path = Path("/Users/lureiny/Desktop/xueersi")
        self.data = None
        self.root = Path.home()

        self.last_file = None                          # 存放上一次读取的文件
        self.index = dict()                            # 存放索引
        self.sended = set()                            # 存放已经发送的
        self.sending = set()                           # 存放将要发送的
        self.sended_name = set()                       # 存放发送过的学生名字
        self.num_student = None

        self.file_path_get()

    def check_root(self, path):
        if path.parent.parts[:3] == self.root.parts:
            return True
        else:
            print("\033[0;30;41m{}\033[0m".format("重新定向到根目录\n"))
            self.path = self.root
            return False

    def file_input_index(self, dir_list):
        try:
            index = input("请输入文件或者文件夹对应的序号：\n")
            if index == "..":
                if self.check_root(self.path):
                    self.path = self.path.parent
                dir_list.clear()
                return True
            else:
                index = int(index)
                if index not in dir_list:
                    print("\033[0;30;41m{}\033[0m".format("序号不存在"))
                    return False
                return index
        except ValueError:
            print("\033[0;30;41m{}\033[0m".format("请输入数字"))
            print()
            return False

    def back_up_sended(self):
        for index in self.sended:
            self.sended_name.add(self.index[index])

    def file_path_get(self):
        if self.path is None or not self.path.exists():
            self.path = self.root
        elif self.path == self.root:
            pass
        else:
            self.check_root(self.path)

        output = "{0:<25}"

        dir_list = dict()
        while True:
            num = 1
            for path in self.path.iterdir():
                if path.is_dir() and path.parts[-1][0] != ".":
                    dir_list[num] = path
                    num += 1

            for path in self.path.glob("*.xlsx"):
                dir_list[num] = path
                num += 1

            print("当前所在目录：{}".format(self.path))
            temp_num = 1
            for index in dir_list:
                print(output.format("{}：{}".format(index, dir_list[index].parts[-1])), end="")
                temp_num += 1
                if temp_num > 5:
                    temp_num = 1
                    print()

            print("\n\n")

            index = False

            while index is False:
                index = self.file_input_index(dir_list)

            if index is True:
                continue

            if dir_list[index].is_dir():
                if self.check_root(dir_list[index]):
                    self.path = dir_list[index]
                dir_list.clear()
                continue
            else:
                self.file = dir_list[index]

                with open(__file__, "r") as file:
                    temp = file.read()
                temp = temp.split("\n")
                temp[9] = temp[9].replace(temp[9][temp[9].index("="):], "= Path(\"{}\")".format(self.path))
                with open(__file__, "w") as file:
                    file.write("\n".join(temp))

                self.back_up_sended()

                if self.read_file():
                    self.make_index()
                    self.num_student = len(self.data.all_info)
                    if self.file == self.last_file:
                        for index in self.index:
                            if self.index[index] in self.sended_name:
                                self.sending.remove(index)
                                self.sended.add(index)
                    else:
                        self.sended = set()
                    self.last_file = self.file
                    break
                else:
                    dir_list.clear()
                    continue

    def list_student(self):
        num = 0
        output = "{0:{1}<10}"
        print("未发送列表：")
        for index in sorted(self.sending):
            if num == 5:
                num = 0
                print()
            if index < 10:
                info = "{}：{}".format(index, self.index[index]).replace(str(index), str(index) + " ")
            else:
                info = "{}：{}".format(index, self.index[index])
            print(output.format(info, chr(12288)), end="")
            num += 1

        print("\n")
        num = 0
        print("已发送列表：")
        for index in sorted(self.sended):
            if num == 5:
                num = 0
                print()
            if index < 10:
                info = "{}：{}".format(index, self.index[index]).replace(str(index), str(index) + " ")
            else:
                info = "{}：{}".format(index, self.index[index])
            print(output.format(info, chr(12288)), end="")
            num += 1
        print("\n")

    def send(self, index):
        if index in self.index:
            try:
                student = self.index[index]
                self.send_to_one(student)
                self.sended.add(index)
                self.sending.remove(index)
                return 1
            except Exception as e:
                print(e)
                return -1
        elif index == 0:
            user_input = input("还有%d名学生没有发送回访信息，是否开始自动发送 yes/[no]：" % len(self.sending))
            if user_input == "yes":
                print("\033[0;30;41m{}\033[0m".format("开始自动发送"))
                self.auto_send()
                print("\033[0;30;41m{}\033[0m".format("自动发送结束"))

        else:
            print("\033[0;30;41m{}\033[0m".format("您输入学生不存在！"))
            return -1

    def send_to_one(self, student, auto=False):
        sign = True if self.data.wechat[student][0] and self.data.wechat[student][1] else False
        if self.data.wechat[student][0]:
            self.generate_applescript(self.data.wechat[student][0], self.data.all_info[student][0], auto=auto)
            # self.run_applescript()

        if sign and not auto:
            input("敲击回车开始给家长发送\n")

        if self.data.wechat[student][1]:
            self.generate_applescript(self.data.wechat[student][1], self.data.all_info[student][1], auto=auto)
            # self.run_applescript()

    def auto_send(self):
        try:
            if self.num_student == 0:
                print("\033[0;30;41m{}\033[0m".format("Grade表为空"))
                return -1
            temp = sorted(self.sending)
            for index in temp:
                student = self.index[index]
                self.send_to_one(student=student, auto=True)
                self.sended.add(index)
                self.sending.remove(index)
            return 1
        except Exception as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            return -1

    def read_file(self):
        try:
            self.data = ReadExcel(self.file)
            return True
        except ReadError as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            return False
        except GradeError as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            return False
        except ComprehensionError as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            return False
        except ModelError as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            return False
        except WeChatError as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            return False
        except openpyxl.utils.exceptions.InvalidFileException:
            print("\033[0;30;41m{}\033[0m".format("文件格式错误，请修改文件为xlsx格式的文件"))
            return False
        except Exception as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            return False

    def make_index(self):
        index = 1
        self.index.clear()
        for student in self.data.all_info:
            self.sending.add(index)
            self.index[index] = student
            index += 1

    @staticmethod
    def generate_applescript(wechat, info, auto=False):
        s = "  " if auto else "   --"
        script_model = """tell application "WeChat" to activate
tell application "System Events"
    tell process "WeChat"
        set the clipboard to "%s"
        click menu item "查找…" of menu "编辑" of menu bar item "编辑" of menu bar 1
        key code 9 using {command down}
        delay 0.5
        key code 76
        key code 48 using {command down}
        key code 48 using {command down}
        set the clipboard to "%s"
        key code 9 using {command down}
     %s key code 76
    end tell
end tell""" % (wechat, info, s)
        with open("temp.applescript", "w") as file:
            file.write(script_model)

    @staticmethod
    def run_applescript():
        os.system("osascript temp.applescript")


if __name__ == '__main__':
    cli = CLI()
    while True:
        cli.list_student()
        try:
            index = input("请输出学生的序号：\n")
            if index == "r":
                cli.file_path_get()
                continue
            else:
                index = int(index)
        except ValueError:
            print("\033[0;30;41m{}\033[0m".format("请输入数字"))
            print()
            continue
        print()
        code = cli.send(index=index)

