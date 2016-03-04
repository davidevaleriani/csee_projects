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
import csv
import numpy as np

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
                            <label for="abstracts" class="col-xs-2 col-xs-offset-2 control-label">Zip file</label>
                            <div class="col-xs-6">
                                <input type="file" name="abstracts" id="abstracts" />
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
    def get_booklet(self, abstracts):
        # Get list of students for "backup"
        students = {}
        students_by_surname = {}
        if os.path.exists("CE301LST.csv"):
            # Load list of students from CSV file
            # Full name, first names, degree, regno, surname
            with open('CE301LST.csv', 'rb') as csvfile:
                spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
                for row in spamreader:
                    # strudents[regno] = (first names, surname, degree)
                    students[row[3]] = (row[1], row[4], row[2])
                    students_by_surname[row[4].upper()] = (row[1], row[2])
        elif os.path.exists("students_list.xls"):
            # Load list of students from Excel file
            workbook = xlrd.open_workbook("students_list.xls")
            worksheets = workbook.sheet_names()
            worksheet = workbook.sheet_by_name(worksheets[0])
            for row in range(1, worksheet.nrows):
                students[worksheet.cell_value(row, 0)] = worksheet.cell(row, 2)
        # Create the directory where the abstract will be temporarily saved
        if not os.path.isdir("tmp"):
            os.mkdir("tmp")
        # Extract files from the zip file
        if type(abstracts) != file and hasattr(abstracts, "file"):
            abstracts = abstracts.file
        zipf = zipfile.ZipFile(abstracts, 'r')
        zipf.extractall("tmp/")
        # Init the booklet document
        booklet_doc = Document("booklet_template.docx")
        booklet_doc.add_page_break()
        processed_files = []
        files_not_inserted = []
        files_to_be_checked = []
        files_duplicated = []
        titles = []
        # Sort abstracts by creation time
        mtime = lambda f: os.stat(os.path.join('tmp/', f)).st_mtime
        list_of_abstracts = list(sorted(os.listdir('tmp/'), key=mtime))[::-1]
        # For each abstract form
        for f in list_of_abstracts:
            # Rename file to unix-like format (remove spaces, etc.)
            os.rename("tmp/"+f, "tmp/"+f.replace(" ", "-").lower())
            f = f.replace(" ", "-").lower()
            if f in processed_files:
                continue
            processed_files.append(f)
            project_title = ""
            student_first = ""
            student_last = ""
            supervisor_first = ""
            supervisor_last = ""
            degree = ""
            abstract = ""
            abstract_has_started = False
            if f[-4:] == "docx":
                # If the abstracts are in .docx format
                document = Document("tmp/"+f)
                # Get the degree course and the student names from the students list
                if unicode(f[:7], 'utf-8').isnumeric() and f[:7] in students.keys():
                    student_first = students[f[:7]][0].title()
                    student_last = students[f[:7]][1].title()
                    degree = students[f[:7]][2].title()
                try:
                    for p in document.paragraphs:
                        p = p.text
                        if "Title" in p and not project_title:
                            try:
                                project_title = p.split("Title:")[1].lstrip()
                                project_title = ' '.join(project_title.split())
                            except:
                                print "WARNING: check title of", project_title
                                if f in processed_files:
                                    processed_files.remove(f)
                                if f not in files_to_be_checked:
                                    files_to_be_checked.append(f)
                                project_title = p
                        elif "Student First Name" in p and not student_first:
                            try:
                                if len(p.split("Student First Name(s):")) > 1:
                                    student_first = p.split("Student First Name(s):")[1].lstrip()
                                elif len(p.split("Student First Name:")) > 1:
                                    student_first = p.split("Student First Name:")[1].lstrip()
                                else:
                                    raise
                                student_first = ' '.join(student_first.split())
                            except:
                                print "WARNING: check student name of", project_title
                                if f in processed_files:
                                    processed_files.remove(f)
                                if f not in files_to_be_checked:
                                    files_to_be_checked.append(f)
                                student_first = p
                        elif "Student Surname" in p and not student_last:
                            try:
                                student_last = p.split("Student Surname:")[1].lstrip()
                                student_last = ' '.join(student_last.split())
                            except:
                                print "WARNING: check student name of", project_title
                                if f in processed_files:
                                    processed_files.remove(f)
                                if f not in files_to_be_checked:
                                    files_to_be_checked.append(f)
                                student_last = p
                        elif "Student" in p and not student_first and not student_last:
                            try:
                                student_name = p.split(":")[1].lstrip()
                                student_first = ' '.join(student_name.split(" ")[:-1])
                                student_last = student_name.split(" ")[-1]
                            except:
                                print "WARNING: check student name of", project_title
                                if f in processed_files:
                                    processed_files.remove(f)
                                if f not in files_to_be_checked:
                                    files_to_be_checked.append(f)
                                student_last = p
                        elif "Supervisor First" in p:
                            try:
                                if len(p.split("Supervisor First Name(s):")) > 1:
                                    supervisor_first = p.split("Supervisor First Name(s):")[1].lstrip()
                                elif len(p.split("Supervisor First Name:")) > 1:
                                    supervisor_first = p.split("Supervisor First Name:")[1].lstrip()
                                else:
                                    raise
                                supervisor_first = ' '.join(supervisor_first.split())
                            except:
                                print "WARNING: check supervisor of", project_title
                                if f in processed_files:
                                    processed_files.remove(f)
                                if f not in files_to_be_checked:
                                    files_to_be_checked.append(f)
                                supervisor_first = p
                        elif "Supervisor Surname" in p:
                            try:
                                supervisor_last = p.split("Supervisor Surname:")[1].lstrip()
                                supervisor_last = ' '.join(supervisor_last.split())
                            except:
                                print "WARNING: check supervisor of", project_title
                                if f in processed_files:
                                    processed_files.remove(f)
                                if f not in files_to_be_checked:
                                    files_to_be_checked.append(f)
                                supervisor_last = p
                        elif "Supervisor" in p and not supervisor_first and not supervisor_last:
                            try:
                                supervisor = p.split(":")[1].lstrip()
                                supervisor_first = ' '.join(supervisor.split(" ")[:-1])
                                supervisor_last = supervisor.split(" ")[-1]
                            except:
                                print "WARNING: check supervisor name of", project_title
                                if f in processed_files:
                                    processed_files.remove(f)
                                if f not in files_to_be_checked:
                                    files_to_be_checked.append(f)
                                student_last = p
                        elif "Degree Course" in p and not degree:
                            try:
                                degree = p.split("Degree Course:")[1].lstrip()
                                degree = ' '.join(degree.split())
                            except:
                                print "WARNING: check degree of", project_title
                                if f in processed_files:
                                    processed_files.remove(f)
                                if f not in files_to_be_checked:
                                    files_to_be_checked.append(f)
                                degree = p
                        elif "Abstract." not in p and abstract_has_started:
                            abstract += p
                        elif "Abstract." in p:
                            try:
                                abstract = p.split("Abstract.")[1].lstrip()
                                abstract_has_started = True
                            except:
                                print "WARNING: check abstract of", project_title
                                if f in processed_files:
                                    processed_files.remove(f)
                                if f not in files_to_be_checked:
                                    files_to_be_checked.append(f)
                                abstract = p
                except:
                    print "ERROR with file", f, "- SKIP"
                    continue
            elif f[-3:] == "doc":
                # If the abstracts are in .doc format
                try:
                    # Convert the .doc to .txt
                    txt_file = subprocess.check_output(["antiword", "tmp/%s" % f])
                except:
                    print "WARNING: unable to convert", f, "- SKIP"
                    continue
                counter = 0
                for field in txt_file.split("\n"):
                    field = field.lstrip()
                    if "Title" in field:
                        project_title = field.split("Title:")[1].lstrip()
                    elif "Student".upper() in field.upper():
                        student = field.split(":")[1].lstrip()
                        student_first = ' '.join(student.split(" ")[:-1]) if len(student.split(" ")[:-1]) > 1 else student.split(" ")[:-1]
                        student_last = student.split(" ")[-1]
                    elif "Supervisor".upper() in field.upper():
                        supervisor = field[len("Supervisor:"):].lstrip()
                        supervisor_first = ' '.join(supervisor.split(":")[:-1]) if len(supervisor.split(":")[:-1]) > 1 else supervisor.split(":")[:-1]
                        supervisor_last = supervisor.split(":")[-1]
                    elif "Abstract".upper() in field.upper():
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
            if not all([project_title, student_first, student_last, supervisor_first, supervisor_last]):
                print "WARNING: file %s not inserted" % f
                if f in processed_files:
                    processed_files.remove(f)
                if f not in files_not_inserted:
                    files_not_inserted.append(f)
                print "TITLE:", project_title
                print "STUDENT:", student_first, " > ", student_last
                print "SUPERVISOR:", supervisor_first, " > ",supervisor_last
                print "DEGREE COURSE:", degree
                print "ABSTRACT:", abstract
                print
                continue
            if not degree:
                if student_last.upper() in students_by_surname.keys():
                    degree = students_by_surname[student_last.upper()][1]
                else:
                    print "WARNING: file %s doesn't have a degree course" % f
                    if f not in files_to_be_checked:
                        files_to_be_checked.append(f)
                    print "TITLE:", project_title
                    print "STUDENT:", student_first, student_last
                    print "SUPERVISOR:", supervisor_first, supervisor_last
                    print "DEGREE COURSE:", degree
                    print "ABSTRACT:", abstract
                    print
            if not abstract:
                print "WARNING: file %s doesn't have an abstract" % f
                if f not in files_to_be_checked:
                    files_to_be_checked.append(f)
                print "TITLE:", project_title
                print "STUDENT:", student_first, student_last
                print "SUPERVISOR:", supervisor_first, supervisor_last
                print "DEGREE COURSE:", degree
                print "ABSTRACT:", abstract
                print
            if project_title.title() in titles:
                #print "WARNING: file %s is a duplicate" % f
                if f in processed_files:
                    processed_files.remove(f)
                if f not in files_duplicated:
                    files_duplicated.append(f)
                #print "TITLE:", project_title
                #print "STUDENT:", student_first, student_last
                #print "SUPERVISOR:", supervisor_first, supervisor_last
                #print "DEGREE COURSE:", degree
                #print "ABSTRACT:", abstract
                #print
                continue
            titles.append(project_title.title())
            # Adding to the main document
            heading = booklet_doc.add_paragraph()
            paragraph_format = heading.paragraph_format
            run = heading.add_run("\n"+project_title.title())
            run.bold = True
            font = run.font
            font.size = Pt(13)
            paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            table = booklet_doc.add_table(3, 3)
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
            run = abstract_paragraph.add_run('\nAbstract.\n')
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

        return """
        <html>
        <body>
        <h2>Files added: %d</h2>
        <h2>Files to be checked: %d</h2>
        <h3><ul>%s</ul><h3>
        <h2>Files not inserted: %d</h2>
        <h3><ul>%s</ul><h3>
        <h2>Files duplicated: %d</h2>
        <h3><ul>%s</ul><h3>
        <h2>Total: %d</h2>
        <h1 align="center"><a href="download_booklet">Download booklet</a></h1>
        </body>
        </html>
        """ % (len(processed_files),
               len(files_to_be_checked),
               ''.join(["<li>"+filename+"</li>" for filename in files_to_be_checked]),
               len(files_not_inserted),
               ''.join(["<li>"+filename+"</li>" for filename in files_not_inserted]),
               len(files_duplicated),
               ''.join(["<li>"+filename+"</li>" for filename in files_duplicated]),
               len(processed_files+files_not_inserted+files_duplicated))

    @cherrypy.expose
    def download_booklet(self):
        # Return the file to download
        booklet_doc = open("booklet.docx", 'r')
        return cherrypy.lib.static.serve_fileobj(booklet_doc, disposition='attachment', name="booklet.docx")

if __name__ == '__main__':
    cherrypy.server.socket_host = '0.0.0.0'
    cherrypy.quickstart(HomePage())
