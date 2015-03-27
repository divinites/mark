import docx
from docx.shared import Pt
from docx.shared import Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH


class ECSFeedbackForm(docx.Document):
    def __init__(self, title=''):
        self.title = title

    def add_title(self, title='', font='Arial', ptsize=14):
        self.title = title
        head = self.add_paragraph()
        head_run = head.add_run(self.title)
        head_run.bold = True
        head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        head_run.font.name = font
        head_run.font.size = Pt(ptsize)

    def add_intro(self, description='', font='Arial', ptsize=14):
        intro = self.add_paragraph()
        intro_run = intro.add_run(description)
        intro_run.font.name = font
        intro_run.font.size = Pt(ptsize)
        intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def add_sentence(self, description='', font='Arial', ptsize=14):
        paragraph_list = self.paragraphs
        if not paragraph_list:
            last_paragraph = self.add_paragraph()
        else:
            last_paragraph = paragraph_list.pop()
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        sentence = last_paragraph.add_run(description)
        sentence.font.name = font
        sentence.font.size = Pt(ptsize)

    def add_info(self, number_tag='Exam Number:', mark_tag='Mark:', ptsize=14):
        info = self.add_table(1, 4)
        info_cells = info.rows[0].cells
        S0 = info_cells[0].paragraphs[0].add_run(number_tag)
        info_cells[0].width = Emu(2500000)
        S0.font.size = Pt(ptsize)
        S1 = info_cells[1].paragraphs[0]
        info_cells[1].width = Emu(3400000)
        S1.font.size = Pt(ptsize)
        S1.font.underline = True
        S2 = info_cells[2].paragraphs[0].add_run(mark_tag)
        S2.font.size = Pt(ptsize)
        info_cells[2].width = Emu(360000)
        S3 = info_cells[3].paragraphs[0]
        info_cells[3].width = Emu(360000)
        S3.font.size = Pt(ptsize)
        S0.bold = True
        S1.bold = True
        S2.bold = True
        S3.bold = True
        self.stdno = S1
        self.mark = S3

    def add_form(self, statements,font='Arial', ptsize=10, style='Light Shading'):
        neg_stats, std_mark, pos_stats = tuple(zip(*statements))
        form = self.add_table(len(statements), len(statements[0]))
        form.style = style
        pos_cells = form.columns[0].cells
        neg_cells = form.columns[2].cells
        mark_cells = form.columns[1].cells

        for idx, stats in enumerate(neg_stats):
            ptem = neg_cells[idx].paragraphs[0]
            ptem.alignment = WD_ALIGN_PARAGRAPH.LEFT
            tem = ptem.add_run(stats)
            tem.font.size = Pt(ptsize)
            tem.font.name = font
            tem.bold = False

        for idx, stats in enumerate(pos_stats):
            ptem = pos_cells[idx].paragraphs[0]
            ptem.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            tem = ptem.add_run(stats)
            tem.font.size = Pt(ptsize)
            tem.font.name = font
            tem.bold = False

        for idx, mark in enumerate(std_mark):

            ptem = mark_cells[idx].paragraphs[0]
            ptem.alignment = WD_ALIGN_PARAGRAPH.CENTER
            tem = ptem.add_run(mark)
            tem.font.size = Pt(ptsize)
            tem.font.name = font
            tem.bold = False

    def new_page(self):
        self.add_page_break()


    def add_comment_table(self, comment_catagory, font='Arial', ptsize=10, style= 'Light Shading'):

        comments = self.add_table(len(comment_catagory), 2)
        comments.style = style
        section_cells = comments.columns[0].cells
        content_cells = comments.columns[1].cells

        for i in comment_catagory:
            section_title = section_cells[0].paragraphs[0].add_run(i)
            section_title.font.size = Pt(ptsize)
            section_title.font.name = font


    for idx, i in enumerate(individual_comment):
        if i == 5:
            ps_comment = comments_cells[0].add_paragraph(style='List Bullet').add_run(comment_statement[idx])
            ps_comment.bold = False
            ps_comment.font.size =Pt(10)
        if i == 0:
            ng_comment = comments_cells[1].add_paragraph(style='List Bullet').add_run(comment_statement[idx])
            ng_comment.bold = False
            ng_comment.font.size = Pt(10)
        if i == 6:
            bs_comment = comments_cells[0].add_paragraph(style='List Bullet').add_run(comment_statement[idx]+' (Bonus Point)')
            bs_comment.bold = False
            bs_comment.font.size = Pt(10)
        else:
            pass

    ad_comment = comments_cells[3].add_paragraph(style='List Bullet').add_run(additional_comment)
    ad_comment.bold = False
    ad_comment.font.size = Pt(10)

    final_grade = grade(marks, weights)+ grade(individual_comment, comment_weight)
    while (final_grade % 10 >8):
        final_grade=final_grade+1
    finals = info_cells[3].paragraphs[0].add_run(str(final_grade))
    finals.font.size = Pt(14)
    finals.bold = True
    finals.font.underline = True

    document.save(std_no+'.docx')
    return std_no, final_grade