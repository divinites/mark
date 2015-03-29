def help_statement(key):
    help_dict = {"-i": "\'i\'parameter defines where the input file is.",
                 "-a": "with \'a\' parameter, will print all forms according to the student_list, "
                       "this is also the default option.",
                 "-s": "\'s\' parameter should be followed by a valid student number, print the corresponding form.",
                 "-o": " \'o\' parameter defines the location of generated forms.",
                 "usage": "For auto-marking students' essays and scripts and auto-generating feedback forms."}
    return help_dict[key]