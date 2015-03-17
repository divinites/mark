from libmark import *
statements = "./statement.csv"
marks = "./mark.csv"
weights = [10,20,30,4,5,6,7,8,9,10,11,12,13,14]
mark_matrix = order_file(marks)
mark_matrix.pop(0)
for individual_mark in mark_matrix:
    grade(individual_mark, weights)
    doc_process(individual_mark, statements,weights)
