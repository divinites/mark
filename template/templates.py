from template.base_template import *



class DefaultForm(FeedbackForm): 

    def ug_form(self, path=None):
        if path is None:
            path = './'    
        self.add_title('TITLE')
        self.docx.write(self.title, title_font, 'new')
        self.docx.write("Introduction....", intro_font, 'new')
        self.add_info()
        self.add_form()
        self.docx.document.add_page_break()
        self.add_comment_table()
        self.docx.document.save(path + self.stud_no + '.docx')  

    def print_form(self, path):
            self.ug_form(path)

supported_forms = {"DefaultForm": DefaultForm}