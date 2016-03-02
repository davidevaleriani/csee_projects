import cherrypy
from cherrypy.lib import auth_basic
import os
import zipfile
# Install antiword with apt-get install antiword
import subprocess
import shutil
# Install python-docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import xlrd

use_antiword = False

class HomePage(object):
    @cherrypy.expose
    def index(self):
        return """
            <html>
            <head>
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/css/bootstrap.min.css">
                <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
                <title>CSEE Abstract Booklet Creator</title>
                <style>
                /* Wrapper for page content to push down footer */
                html, body {
                  padding-top: 0px;
                  font-size: 15px;
                  height: 100%;
                }

                #wrap {
                  min-height: 100%;
                  height: auto !important;
                  height: 100%;
                  /* Negative indent footer by it's height */
                  margin: 0 auto -40px;
                  padding-bottom: 10px;
                }

                /* Set the fixed height of the footer here */
                #push,
                #footer {
                  height: 40px;
                  padding-top: 10px;
                }
                #footer {
                  font-size: 12px;
                  background-color: #ddd;
                  text-align: center;
                }
                </style>
            </head>
            <body>
            <div id="wrap">
                <div class="container">
                    <h2 align="center">Welcome to the Abstract Booklet Creator</h2>
                    <p align="center">This application allows you to create the booklet from a zip of abstracts.</p>
                    <br>
                    <form method="post" action="get_booklet" enctype="multipart/form-data" class="form-horizontal">
                        <div class="form-group">
                            <label for="marks" class="col-xs-2 col-xs-offset-2 control-label">Zip file</label>
                            <div class="col-xs-6">
                                <input type="file" name="marks" id="marks" />
                            </div>
                        </div>
                        <div class="form-group">
                            <div class="col-xs-offset-6 col-sm-6">
                                <button type="submit" class="btn btn-primary">Generate</button>
                            </div>
                        </div>
                    </form>
                </div>
                <div id="push"></div>
            </div>
            <div id="footer">
                <div class="container">
                    <p class="muted credit">Created by <a href="http://www.davidevaleriani.it">Davide Valeriani</a>.</p>
                </div>
            </div>
            </body>
            </html>
            """

    @cherrypy.expose
    def get_booklet(self, marks):
        # Get list of students
        students = {}
        workbook = xlrd.open_workbook("students_list.xls")
        worksheets = workbook.sheet_names()
        worksheet = workbook.sheet_by_name(worksheets[0])
        for row in range(1, worksheet.nrows):
            students[worksheet.cell_value(row, 0)] = worksheet.cell(row, 2)
        if not os.path.isdir("tmp"):
            os.mkdir("tmp")
        # Extract files
        if type(marks) != file and hasattr(marks, "file"):
            marks = marks.file
        zipf = zipfile.ZipFile(marks, 'r')
        zipf.extractall("tmp/")
        # Init the booklet document
        booklet_doc = Document("booklet_template.docx")
        booklet_doc.add_page_break()
        # For each abstract form
        for dirname, dirnames, filenames in os.walk('tmp/'):
            for f in filenames:
                # Rename file to unix-like format
                os.rename("tmp/"+f, "tmp/"+f.replace(" ", "-").lower())
                f = f.replace(" ", "-").lower()
                project_title = ""
                student_first = ""
                student_last = ""
                supervisor_first = ""
                supervisor_last = ""
                degree = ""
                abstract = ""
                if not use_antiword:
                    document = Document("tmp/"+f)
                    abstract_begin = False
                    try:
                        for p in document.paragraphs:
                            p = p.text
                            if "Project Title" in p:
                                try:
                                    project_title = p.split("Project Title:")[1].lstrip()
                                    project_title = ' '.join(project_title.split())
                                except:
                                    project_title = p
                            elif "Student First Name(s)" in p:
                                try:
                                    student_first = p.split("Student First Name(s):")[1].lstrip()
                                    student_first = ' '.join(student_first.split())
                                except:
                                    print "WARNING: check", project_title
                                    student_first = p
                            elif "Student Surname" in p:
                                try:
                                    student_last = p.split("Student Surname:")[1].lstrip()
                                    student_last = ' '.join(student_last.split())
                                except:
                                    print "WARNING: check", project_title
                                    student_last = p
                            elif "Supervisor First" in p:
                                try:
                                    supervisor_first = p.split("Supervisor First Name(s):")[1].lstrip()
                                    supervisor_first = ' '.join(supervisor_first.split())
                                except:
                                    print "WARNING: check", project_title
                                    supervisor_first = p
                            elif "Supervisor Surname" in p:
                                try:
                                    supervisor_last = p.split("Supervisor Surname:")[1].lstrip()
                                    supervisor_last = ' '.join(supervisor_last.split())
                                except:
                                    print "WARNING: check", project_title
                                    supervisor_last = p
                            elif "Degree Course" in p:
                                try:
                                    degree = p.split("Degree Course:")[1].lstrip()
                                    degree = ' '.join(degree.split())
                                except:
                                    print "WARNING: check", project_title
                                    degree = p
                            elif "Abstract." not in p and abstract_begin:
                                abstract += p
                            elif "Abstract." in p:
                                try:
                                    abstract = p.split("Abstract.")[1].lstrip()
                                except:
                                    print "WARNING: check", project_title
                                    abstract = p
                                abstract_begin = True
                    except:
                        print f

                else:
                    try:
                        # Convert to txt
                        txt_file = subprocess.check_output(["antiword", "tmp/%s" % f])
                    except:
                        continue
                    counter = 0
                    for field in txt_file.split("\n"):
                        field = field.lstrip()
                        if "Project Title" in field:
                            project_title = field[len("Project Title:"):].lstrip()
                        elif "Student" in field:
                            student = field[len("Student:"):].lstrip()
                        elif "Supervisor" in field:
                            supervisor = field[len("Supervisor:"):].lstrip()
                        elif "Abstract." in field:
                            abstract = field[len("Abstract."):].lstrip()
                            abstract += "\n".join(txt_file.split("\n")[counter+1:])
                            break
                        counter += 1
                    # get the degree from the students_list
                    reformatted_student_name = (student.split(" ")[-1]+", "+" ".join(student.split(" ")[:-1])).upper()
                    try:
                        degree = students[reformatted_student_name]
                    except:
                        degree = ""
                if not all([project_title, student_first, student_last, supervisor_first, supervisor_last, degree]):
                    print "WARNING: file %s not inserted" % f
                    continue
                print "TITLE:", project_title, f
                #print "STUDENT:", student_first, student_last
                #print "SUPERVISOR:", supervisor_first, supervisor_last
                #print "DEGREE COURSE:", degree
                #print "ABSTRACT:", abstract
                # Adding to the main document
                heading = booklet_doc.add_paragraph()
                paragraph_format = heading.paragraph_format
                run = heading.add_run("\n"+project_title.title())
                run.bold = True
                font = run.font
                font.size = Pt(13)
                paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                table = booklet_doc.add_table(3, 2)
                table.allow_autofit = False
                for i in range(3):
                    row = table.rows[i]
                    tr = row._tr
                    trPr = tr.get_or_add_trPr()
                    trHeight = OxmlElement('w:trHeight')
                    trHeight.set(qn('w:val'), "260")
                    trHeight.set(qn('w:hRule'), "exact")
                    trPr.append(trHeight)
                table.cell(0, 0).paragraphs[0].add_run('Name:').bold = True
                table.cell(0, 1).paragraphs[0].add_run(student_first.title()+" "+student_last.title()).bold = True
                table.cell(1, 0).paragraphs[0].add_run('Degree course:').bold = True
                table.cell(1, 1).paragraphs[0].add_run(degree.title()).bold = True
                table.cell(2, 0).paragraphs[0].add_run('Supervisor:').bold = True
                table.cell(2, 1).paragraphs[0].add_run(supervisor_first.title()+" "+supervisor_last.title()).bold = True
                abstract_paragraph = booklet_doc.add_paragraph()
                run = abstract_paragraph.add_run('\nAbstract.\n\n')
                run.bold = True
                font = run.font
                font.size = Pt(12)
                run = abstract_paragraph.add_run(abstract)
                run.bold = False
                font = run.font
                font.size = Pt(12)
        # Save the report
        booklet_doc.save("booklet.docx")

        # Remove the marks
        shutil.rmtree('tmp')

        # Return the file to download
        booklet_doc = open("booklet.docx", 'r')
        return cherrypy.lib.static.serve_fileobj(booklet_doc, disposition='attachment', name="booklet.docx")


if __name__ == '__main__':
    cherrypy.server.socket_host = '0.0.0.0'
    cherrypy.quickstart(HomePage())
