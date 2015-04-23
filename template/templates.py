from template.base_template import *



class UGECSForm(FeedbackForm): 

    def ug_form(self, path=None):
        if path is None:
            path = './'    
        self.add_title('Economics of Corporate Strategy - ASSIGNMENT FEEDBACK FORM 2014/15')
        self.docx.write(self.title, title_font, 'new')
        self.docx.write("This form is designed to provide you with specific "
                    "feedback on your own coursework assignment. The scale "
                    "below qualitatively presents an overview of the strengths"
                    " and weaknesses of your work. Items are only ticked where"
                    " applicable. Your tutor may provide additional "
                    "comments overleaf.", intro_font, 'new')
        self.add_info()
        self.add_form()
        self.docx.document.add_page_break()
        self.add_comment_table()
        self.docx.document.save(path + self.stud_no + '.docx')  

    def print_form(self, path, form_type=None):
        if form_type is None:
            form_type == 'ecs'
            self.ug_form(path)
        else:
            print("Error Type!")


supported_forms = {"UGECSForm": UGECSForm}