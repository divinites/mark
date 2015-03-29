#!/Users/divinites/anaconda/bin/python
from template import *
import os
import sys
from getopt import *

os.system('rm -rf *.docx')
try:
    opts, args = getopt(sys.argv[1:], "vai:s:o:t:", ["all", "input=", "output=", "student=", "type="])
except:
    print("Possible parameter: -a, -v, -i, -s, --all, --input=, --output, --student=, --type=")
    sys.exit(1)

path = './'
marks = []

for o, a in opts:

    if o in ["-i", '--input']:
        marks = Result(a)
        form = FeedbackForm(a)
        score = {}
        for i in marks.student_list:
            score[i] = marks.round(marks.grading(i))

    if o in ("-o", "--output"):
        if not os.path.exists(a):
            os.mkdir(a)
            path = a
    if o in ("-t", '--type'):
        type = a
    else:
        type = 'ecs'

for o, a in opts:
    if o in ("-s", "--student"):
        std_no = a
        form.print_form(std_no,path, type)

    if o in ("--all", "a"):
        for student in marks.student_list:
            form.print_form(student, path, type)

