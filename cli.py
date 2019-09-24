from read_file import ReadExcel, ReadError, GradeError, ComprehensionError, ModelError, WeChatError
import openpyxl
import os


class CLI:
    def __init__(self, file):
        self.file = file
        self.data = self.read_file()
        self.index = dict()                            # 存放索引
        self.sended = set()                            # 存放已经发送的
        self.sending = set()                           # 存放将要发送的
        self.num_student = len(self.data.all_info)

        self.make_index()

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
            return self.auto_send()
        else:
            print("\033[0;30;41m{}\033[0m".format("您输入学生不存在！"))
            return -1

    def send_to_one(self, student):
        self.generate_applescript(self.data.wechat[student], self.data.all_info[student])
        # self.run_applescript()

    def auto_send(self):
        try:
            if self.num_student == 0:
                print("\033[0;30;41m{}\033[0m".format("Grade表为空"))
                return -1
            temp = sorted(self.sending)
            for index in temp:
                student = self.index[index]
                self.send_to_one(student=student)
                self.sended.add(index)
                self.sending.remove(index)
            return 1
        except Exception as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            return -1

    def read_file(self):
        try:
            return ReadExcel(self.file)
        except ReadError as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            exit(-1)
        except GradeError as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            exit(-1)
        except ComprehensionError as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            exit(-1)
        except ModelError as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            exit(-1)
        except WeChatError as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            exit(-1)
        except openpyxl.utils.exceptions.InvalidFileException:
            print("\033[0;30;41m{}\033[0m".format("文件格式错误，请修改文件为xlsx格式的文件"))
            exit(-1)
        except Exception as e:
            print("\033[0;30;41m{}\033[0m".format(e.__str__()))
            exit(-1)

    def make_index(self):
        index = 1
        for student in self.data.all_info:
            self.sending.add(index)
            self.index[index] = student
            index += 1

    @staticmethod
    def generate_applescript(wechat, info):
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
        -- key code 76
    end tell
end tell""" % (wechat, info)
        with open("temp.applescript", "w") as file:
            file.write(script_model)

    @staticmethod
    def run_applescript():
        os.system("osascript temp.applescript")


if __name__ == '__main__':
    cli = CLI("test.xlsx")
    while True:
        cli.list_student()
        try:
            index = int(input("请输出学生的序号：\n"))
        except ValueError:
            print("\033[0;30;41m{}\033[0m".format("请输入数字"))
            print()
            continue
        print()
        code = cli.send(index=index)

