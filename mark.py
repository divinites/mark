#!/Users/divinites/anaconda/bin/python
from libmark import *
import os
import sys
from getopt import *

os.system('rm -rf *.docx')
try:
    opts, args = getopt(sys.argv[1:], "f:", ["file="])
except:
    print("Now this program only have one possible parameter: -f or --file")
    sys.exit(1)

for command, obj in opts:
    if command in ("-f", "--file"):
        file = obj

mark_dict = split_sheet(file)

mark_matrix = transfer_sheet(mark_dict['mark'])

weight_matrix = transfer_sheet(mark_dict['weight'])
comment_matrix = transfer_sheet(mark_dict['comment'])
statement_matrix = transfer_sheet(mark_dict['statement'])

mark_matrix.pop(0)
comment_statement = comment_matrix.pop(0)
comment_statement.pop(0)
comment_statement.pop()
std_profile = {}

for individual_mark, individual_comment in zip(mark_matrix, comment_matrix):
    individual_comment.pop(0)
    parameter = {"mark": individual_mark, "all_comment": comment_statement,
                 "comment": individual_comment,
                 "weight": weight_matrix, "statement": statement_matrix}
    std_no, final_grade = doc_process(parameter)
    std_profile[std_no] = final_grade

sorted_marks = sort_grade(std_profile)

with open ('ranking.csv', 'w+') as ranking:
    for i in sorted_marks:
        number,grade = i
        ranking.write(str(number)+','+str(grade)+'\n')



