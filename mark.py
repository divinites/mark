from libmark import *
import os
os.system('rm -rf *.docx')

statements = "./statement.csv"
marks = "./ugmark.csv"
comment = "./ugcomment.csv"
weights = "./weights.csv"
weights_matrix = order_file(weights)


weights = [float(i) for i in weights_matrix[0]]

comment_weight = [float(j) for j in weights_matrix[1]]

mark_matrix = order_file(marks)
mark_matrix.pop(0)
comments = order_file(comment)
comment_statement = comments.pop(0)
comment_statement.pop(0)
comment_statement.pop()

for individual_mark, individual_comment in zip(mark_matrix, comments):
    individual_comment.pop(0)
    parameter = [individual_mark, comment_statement, individual_comment, comment_weight, statements, weights]
    doc_process(parameter)
