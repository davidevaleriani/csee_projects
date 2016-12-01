import cherrypy
import os
import numpy as np
import zipfile
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Border, Side


class HomePage(object):
    def index(self):
        return """
            <html>
            <head>
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/css/bootstrap.min.css">
                <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
                <title>CSEE CE201 Grades Automator</title>
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
                    <h2 align="center">CSEE CE201 Grades Automator</h2>
                    <p>Ensure that in the folder of the server there are the following files:
                    <ul>
                    <li><em>students_list.csv</em> containing the list of students of the class
                    <li><em>template.xlsx</em> containing the template of the marking form
                    </ul>
                    <form method="get" action="generate" class="form-horizontal">
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
    index.exposed = True

    @cherrypy.expose
    def generate(self):
        # Settings
        template_name = "template.xlsx"
        students_list = "students_list.csv"
        marks_dir = "marks/"
        # Load the template
        mark_form_doc = load_workbook(filename=template_name)
        first_sheet = mark_form_doc.get_sheet_names()[0]
        mark_form = mark_form_doc.get_sheet_by_name(first_sheet)
        border = Border(left=Side(border_style='thin', color='00000000'),
                        right=Side(border_style='thin', color='00000000'),
                        top=Side(border_style='thin', color='00000000'),
                        bottom=Side(border_style='thin', color='00000000'),
                        vertical=Side(border_style='thin', color='00000000'),
                        horizontal=Side(border_style='thin', color='00000000'))

        # Load the students list
        students = np.loadtxt(students_list, delimiter=",", dtype={'names': ('group', 'sup1', 'sup2',
                                                                             'stud1_regno', 'stud1_name',
                                                                             'stud2_regno', 'stud2_name',
                                                                             'stud3_regno', 'stud3_name',
                                                                             'stud4_regno', 'stud4_name',
                                                                             'stud5_regno', 'stud5_name',
                                                                             'stud6_regno', 'stud6_name'),
                                                                   'formats': ('S10', 'S30', 'S30',
                                                                               'S7', 'S30',
                                                                               'S7', 'S30',
                                                                               'S7', 'S30',
                                                                               'S7', 'S30',
                                                                               'S7', 'S30',
                                                                               'S7', 'S30',
                                                                               )}, skiprows=1)
        # Skip the header row
        for student in students:
            # Populating
            # Group name
            mark_form['C2'] = student[0]
            # Supervisor
            mark_form['C3'] = student[1]
            # Second assessor
            mark_form['C4'] = student[2]
            # Group composition
            for member in range(3, len(student), 2):
                regno, name = student[member], student[member+1]
                row_index = 37+(member-3)/2
                mark_form['B'+str(row_index)] = regno
                mark_form['C'+str(row_index)] = name
            # Saving files in the specific folders
            if not os.path.isdir(marks_dir):
                os.makedirs(marks_dir)
            mark_form_doc.save(marks_dir+"/comarking_"+"_".join(student[0].split(" "))+'.xlsx')

        # Zip the marks
        zipf = zipfile.ZipFile('marks.zip', 'w')
        for root, dirs, files in os.walk(marks_dir):
            for f in files:
                zipf.write(os.path.join(root, f))
        zipf.close()
        # Remove the marks directory
        shutil.rmtree(marks_dir)
        # Return the file to download
        marks_file = open("marks.zip", 'r')
        return cherrypy.lib.static.serve_fileobj(marks_file, disposition='attachment', name="marks.zip")


if __name__ == '__main__':
    cherrypy.server.socket_host = '0.0.0.0'
    cherrypy.quickstart(HomePage())
