from os.path import isfile
import operator
from xlrd import open_workbook
import csv


def detect_file_type(file):
    if isfile(file + '.csv') or file[-4:] == '.csv':
        import csv

        return 'csv'
    if (isfile(file + 'xls') or isfile(file + '.xlsx')) or (file[-4:] in ['xlsx', '.xls']):
        from xlrd import open_workbook

        return 'excel'
    else:
        raise Exception("Wrong file type!")


def detect_delimiter(csv_file):
    with open(csv_file, 'r') as my_csv_file:
        header = my_csv_file.readline()
        if header.find(";") != -1:
            return ";"
        if header.find(",") != -1:
            return ","
    return ";"


def read_file(file_names):
    with open(file_names, 'rt') as csv_file:
        spam_data = csv.reader(csv_file.read().splitlines(), delimiter=detect_delimiter(file_names))
    return spam_data


def order_file(target_file):
    target = read_file(target_file)
    target_matrix = []
    for i in target:
        target_matrix.append(i)
    return target_matrix


# Dealing with Excel workbook
def split_sheet(args):
    wb = open_workbook(args)
    mark_dict = {}
    for sheet in wb.sheets():
        mark_dict[sheet.name] = sheet
    return mark_dict


# Dealing with Excel sheets
def transfer_sheet(args):
    args_matrix = []
    for i in range(0, args.nrows):
        args_matrix.append(args.row_values(i))
    return args_matrix


def file_process(argv, b="a"):
    if detect_file_type(argv) == 'csv':
        csv_process(argv)
    if detect_file_type(argv) == 'excel':
        return excel_process(argv, b)


def csv_process(argv):
    pass


def excel_process(argv, b='a'):
    sheets = split_sheet(argv)
    mark_matrix = transfer_sheet(sheets['mark'])
    form_matrix = transfer_sheet(sheets['form'])
    comment_matrix = transfer_sheet(sheets['comment'])
    if b == 'a':
        return mark_matrix, form_matrix, comment_matrix
    elif b == 'm':
        return mark_matrix
    elif b == 'f':
        return form_matrix
    elif b == 'c':
        return comment_matrix
    else:
        raise Exception("Only m,f,c and a are valid parameters.")


def sort_grade(args):
    sorted_grade = sorted(args.items(), key=operator.itemgetter(1))
    return sorted_grade


class Profiles:
    def __init__(self, file):
        self.students_profile = file

    def get_student_info(self, student_no):
        mmark, mform, mcomment = file_process(self.students_profile)
        stud_info = {"student number:": student_no,
                     "mark": [], "comment": []}
        mark_data = []
        comment_data = []
        for i in mmark:
            if i[0] == student_no:
                mark_data = i
        for j in mcomment:
            if j[0] == student_no:
                comment_data = j
        for i, j in zip(mmark[0][1:], mark_data[1:]):
            stud_info["mark"].append((i, j))
        comment_list = list(zip(mcomment[0][1:], comment_data[1:]))
        for i, j in comment_list:
            stud_info["comment"].append((i, j))
        return stud_info

    def get_mark(self, student_no):
        mmark = file_process(self.students_profile, 'm')
        mark_data = []
        for i in mmark:
            if i[0] == student_no:
                mark_data = i
        return mark_data[1:]


    @property
    def mark_weight(self):
        mmark = file_process(self.students_profile, 'm')
        mark_weight = {}
        temp1 = [i for i in mmark[0][1:]]
        temp2 = [i for i in mmark[1][1:]]
        for i, j in list(zip(temp1, temp2)):
            mark_weight[i] = j

        return mark_weight

    @property
    def comment_weight(self):
        mcomment = file_process(self.students_profile, "c")
        comment_weight = {}
        for i, j in zip(mcomment[1][1:], mcomment[0][1:]):
            comment_weight[j] = i

        return comment_weight

    @property
    def student_list(self):
        mark_dict = split_sheet(self.students_profile)
        student_list = mark_dict['mark'].col_values(0)
        student_list = student_list[2:]
        return student_list

    def grading(self, stud_no):
        mark_grade = 0
        comment_grade = 0
        stud_info = self.get_student_info(stud_no)
        for key, value in self.mark_weight.items():
            for stud_key, stud_value in stud_info["mark"]:
                if key == stud_key:
                    mark_grade += float(value) * float(stud_value)

        filtered_comment = self.filter(stud_info["comment"])

        for key, value in self.comment_weight.items():
            for stud_key, stud_value in filtered_comment:
                if key == stud_key:
                    comment_grade += float(value) * float(stud_value)

        return mark_grade + comment_grade

    def round(self, stud_no):
        return int(round(self.grading(stud_no), 0))

    @staticmethod
    def filter(comment):
        temp = [i for i in comment]
        for idx, value in enumerate(comment):
            if isinstance(value[1], str):
                temp.remove(value)
        return temp