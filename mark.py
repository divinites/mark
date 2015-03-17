from libmark import *

marks = "./mark.csv"
mark_matrix = order_file(marks)
mark_matrix.pop(0)
for individual_mark in mark_matrix:
    doc_process(individual_mark)
