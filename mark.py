from libmark import *
import os
os.system('rm -rf *.docx')


file = "./ugmarks.xlsx"

mark_dict = split_sheet(file)

mark_matrix = transfer_sheet(mark_dict['mark'])

weight_matrix = transfer_sheet(mark_dict['weight'])
comment_matrix = transfer_sheet(mark_dict['comment'])
statement_matrix = transfer_sheet(mark_dict['statement'])

mark_matrix.pop(0)
comment_statement = comment_matrix.pop(0)
comment_statement.pop(0)
comment_statement.pop()

for individual_mark, individual_comment in zip(mark_matrix, comment_matrix):
    individual_comment.pop(0)
    parameter = {"mark": individual_mark, "all_comment": comment_statement,
                 "comment": individual_comment,
                 "weight": weight_matrix, "statement": statement_matrix}
    doc_process(parameter)


