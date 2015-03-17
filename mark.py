from libmark import *
import os
os.system('rm -rf *.docx')

statements = "./statement.csv"
marks = "./mark.csv"
comment = "./comment.csv"
weights = [1,1,1,1,1,1,1,1,1,1,1,1,1,1]
comment_weight = [10,10,10,10,10,10]
mark_matrix = order_file(marks)
mark_matrix.pop(0)
comments = order_file(comment)
comment_statement = comments.pop(0)
comment_statement.pop(0)
comment_statement.pop()
for individual_mark, individual_comment in zip(mark_matrix, comments):
    individual_comment.pop(0)
    doc_process(individual_mark, comment_statement, individual_comment, comment_weight, statements, weights)
