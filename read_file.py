from openpyxl import load_workbook
from error import ReadError, GradeError, ComprehensionError, ModelError, WeChatError


class ReadExcel:
    def __init__(self, file):
        self.file = file
        self.wb = self.get_workbook()
        self.sheets = tuple(self.wb.get_sheet_names())
        self.grade = dict()               # 存放学生的错题和成绩
        self.comprehension = dict()       # 存放解析
        self.wechat = dict()              # 存放学生对应的wechat
        self.model = None                 # 存放模板
        self.error_nums = set()           # 记录全部学生产生的错题
        self.all_info = dict()            # 记录全部学生的回访信息

        self.grade_ws = None
        self.comprehension_ws = None
        self.wechat_ws = None
        self.model_ws = None

        # 执行Excel文件内容检查
        self.check()
        # 检查通过后生成全部的回访信息
        self.generate_all_info()

    def check(self):
        if not self.check_sheet():
            raise ReadError("Sheet错误，请检查Sheet是否包含：grade，comprehension，model，WeChat")
        else:
            self.grade_ws = self.wb.get_sheet_by_name("grade")
            self.comprehension_ws = self.wb.get_sheet_by_name("comprehension")
            self.wechat_ws = self.wb.get_sheet_by_name("WeChat")
            self.model_ws = self.wb.get_sheet_by_name("model")

        error_nums = self.check_grade()
        if error_nums:
            raise GradeError("grade表的第{}行有空值".format("，".join(error_nums)))

        error_nums = self.check_comprehension()
        if error_nums:
            raise ComprehensionError("comprehension表的第{}行有空值".format("，".join(error_nums)))

        if not self.check_model():
            raise ModelError("请检查模板中是否包含：{姓名}, {成绩}, {错题解析}")

        error_nums = self.check_wechat()
        if error_nums:
            raise WeChatError("WeChat表的第{}行有空值".format("，".join(error_nums)))

        lack_nums = self.check_grade_comprehension()
        if lack_nums:
            raise ComprehensionError("缺少以下题目的解析：\n{}".format("，".join(sorted(lack_nums))))

        lack_student = self.check_wechat_grade()
        if lack_student:
            raise WeChatError("缺少以下学生的微信联系方式：\n{}".format("，".join(lack_student)))

    def get_workbook(self):
        return load_workbook(filename=self.file)

    def check_sheet(self):
        sheets = {'grade', 'comprehension', 'model', 'WeChat'}
        for sheet in self.sheets:
            sheets.add(sheet)
        return True if len(sheets) == 4 and len(self.sheets) == 4 else False

    def check_grade(self):
        if not (self.grade_ws["A1"].value == "姓名" and self.grade_ws["B1"].value == "成绩" and self.grade_ws["C1"].value == "错题"):
            raise GradeError("grade表错误：请检查grade表第一行是否包含姓名、成绩、错题")

        error_rows = []
        row_num = 2
        for row in self.grade_ws.iter_rows(min_row=2, max_row=self.grade_ws.max_row):
            name, grade, error_nums = row
            error_nums = str(error_nums.value).split("，") if error_nums.value else []
            for error_num in error_nums:
                self.error_nums.add(error_num)
            if not (name.value and grade.value):
                error_rows.append(str(row_num))
            self.grade[name.value] = tuple((grade.value, error_nums))
            row_num += 1
        return error_rows

    def check_model(self):
        self.model = self.model_ws["A1"].value
        return True if "{姓名}" in self.model and "{成绩}" in self.model and "{错题解析}" in self.model else False

    def check_comprehension(self):
        if not (self.comprehension_ws["A1"].value == "题号" and self.comprehension_ws["B1"].value == "解析"):
            raise ComprehensionError("comprehension表错误：请检查comprehension表第一行是否包含题号，解析")

        error_rows = []
        row_num = 2
        for row in self.comprehension_ws.iter_rows(min_row=2, max_row=self.comprehension_ws.max_row):
            num, comprehension = row
            if not (num.value and comprehension.value):
                error_rows.append(str(row_num))
            self.comprehension[str(num.value)] = str(comprehension.value)
            row_num += 1
        return error_rows

    def check_wechat(self):
        if not (self.wechat_ws["A1"].value == "姓名" and self.wechat_ws["B1"].value == "微信号"):
            raise WeChatError("WeChat表错误：请检查WeChat表第一行是否包含姓名，微信号")

        error_rows = []
        row_num = 2
        for row in self.wechat_ws.iter_rows(min_row=2, max_row=self.wechat_ws.max_row):
            name, wechat_num = row
            if not (name.value and wechat_num.value):
                error_rows.append(str(row_num))
            self.wechat[name.value] = str(wechat_num.value)
            row_num += 1
        return error_rows

    def check_grade_comprehension(self):
        lack_nums = set()
        for error_num in self.error_nums:
            if error_num not in self.comprehension.keys():
                lack_nums.add(error_num)
        return lack_nums

    def check_wechat_grade(self):
        lack_student = list()
        for student in self.grade.keys():
            if student not in self.wechat.keys():
                lack_student.append(student)
        return lack_student

    def generate_one_info(self, name, grade, comprehension_nums):
        comprehension = ""
        for num in comprehension_nums:
            comprehension += ("{}:{}".format(num, self.comprehension[num] + "\n"))
        return self.model.format(姓名=name, 成绩=grade, 错题解析=comprehension)

    def generate_all_info(self):
        for name in self.grade:
            self.all_info[name] = self.generate_one_info(name=name, grade=self.grade[name][0], comprehension_nums=self.grade[name][1])


if __name__ == '__main__':
    r = ReadExcel("test.xlsx")

    print(r.all_info)
