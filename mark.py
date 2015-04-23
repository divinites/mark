#!/Users/divinites/anaconda/bin/python
import os
from argparse import *
import lib.libmark
import template.templates
from helpstats import help_statement


parser = ArgumentParser(description=help_statement("usage"))
parser.add_argument("-i", '--input', help=help_statement('-i'), dest='input', required=True)
parser.add_argument("-o", '--output', default='./',
                    help=help_statement('-o'), dest='output', required=False)
parser.add_argument("-t", '--type', help=help_statement('-t'), dest = 'type', required=False)

exclusive_group = parser.add_mutually_exclusive_group()
exclusive_group.add_argument("-a", "--all", help=help_statement('-a'),
                             action='store_true', dest='all', required=False, default=False)
exclusive_group.add_argument("-s", "--student", type=str,
                             help=help_statement('-s'), dest='student_no', required=False)
args = parser.parse_args()
marks = lib.libmark.Profiles(args.input)
score = {}
forms = {}
current_form =None

if args.type is None:
    current_form = template.templates.DefaultForm
else:
    if args.type in template.templates.supported_forms.keys():
        current_form = template.templates.supported_forms[args.type]
    else:
        raise(ValueError)

for i in marks.student_list:
    score[i] = marks.round(marks.grading(i))
    forms[i] = current_form(args.input, i)

if not os.path.exists(args.output):
    os.mkdir(args.output)

if args.all is True:
    for i, form in forms.items():
        form.print_form(args.output)

if args.student_no is not None:
    forms[args.student_no].print_form(args.output)




