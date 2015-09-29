import cherrypy
from cherrypy.lib import auth_basic
import os
import csv
import zipfile
import tempfile
import shutil
from openpyxl import load_workbook
import xlrd  # Support to old xls files

USERS = {'csee': 'csee'}
marks_dir = "marks/"

class HomePage(object):
    def index(self):
		return """
			<html>
			<head>
				<meta name="viewport" content="width=device-width, initial-scale=1.0">
			    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/css/bootstrap.min.css">
			    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
			    <title>CSEE Grades Automator</title>
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
					<h2 align="center">Welcome to CSEE Grades Automator</h2>
					<p align="center">This application allows you to generate the mark forms for final year projects.</p>
					<br>
					<form method="get" action="generate" enctype="multipart/form-data" class="form-horizontal">
					    <div class="form-group">
							<label for="template_sup" class="col-xs-2 col-xs-offset-2 control-label">Template Supervisor</label>
							<div class="col-xs-6">
					    		<input type="file" name="template_sup" id="template_sup" />
					    	</div>
		          	    </div>
		          	    <div class="form-group">
							<label for="template_sec" class="col-xs-2 col-xs-offset-2 control-label">Template Second Assessor</label>
							<div class="col-xs-6">
					    		<input type="file" name="template_sec" id="template_sec" />
					    	</div>
		          	    </div>
					    <div class="form-group">
							<label for="students_list" class="col-xs-2 col-xs-offset-2 control-label">Students' list</label>
							<div class="col-xs-6">
								<input type="file" name="students_list" id="students_list" />
							</div>
					    </div>
					    <div class="form-group">
							<label for="name_column" class="col-xs-2 col-xs-offset-2 control-label">Student name column</label>
							<div class="col-xs-6">
								<input type="number" class="form-control" name="name_column" min="1" value="2" id="name_column" />
							</div>
					    </div>
					    <div class="form-group">
							<label for="regno_column" class="col-xs-2 col-xs-offset-2 control-label">Student registration number column</label>
							<div class="col-xs-6">
								<input type="number" class="form-control" name="regno_column" min="1" value="1" id="regno_column" />
							</div>
					    </div>
					    <div class="form-group">
							<label for="sup_column" class="col-xs-2 col-xs-offset-2 control-label">Student supervisor column</label>
							<div class="col-xs-6">
								<input type="number" class="form-control" name="sup_column" min="1" value="7" id="sup_column" />
							</div>
					    </div>
					    <div class="form-group">
							<label for="sec_column" class="col-xs-2 col-xs-offset-2 control-label">Student second assessor column</label>
							<div class="col-xs-6">
								<input type="number" class="form-control" name="sec_column" min="1" value="8" id="sec_column" />
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
    index.exposed = True

    @cherrypy.expose
    def generate(self, template_sup="template_supervisor.xlsx", template_sec="template_second_assessor.xlsx", students_list="list_of_students.xls", name_column=2, regno_column=1, sup_column=7, sec_column=8):
    	# Load the template
    	if template_sup.split(".")[-1].upper() == "XLSX" and template_sec.split(".")[-1].upper() == "XLSX":
    		mark_form_sup_doc = load_workbook(filename = template_sup)
    		first_sheet = mark_form_sup_doc.get_sheet_names()[0]
    		mark_form_sup = mark_form_sup_doc.get_sheet_by_name(first_sheet)
    		mark_form_sec_doc = load_workbook(filename = template_sec)
    		first_sheet = mark_form_sec_doc.get_sheet_names()[0]
    		mark_form_sec = mark_form_sec_doc.get_sheet_by_name(first_sheet)
    	else:
    		return """
    		<html>
    		<head>
				<meta name="viewport" content="width=device-width, initial-scale=1.0">
			    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/css/bootstrap.min.css">
			    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
			    <title>CSEE Grades Automator</title>
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
    				<h2 align="center" style="color:red">You have to select a xlsx file for the template.</h2>
    				<div align="center">
    					<a href="index" class="btn btn-default">Back</a>
    				</div>
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
    		raise Exception("Only xlsx files are supported so far for templates")

    	# Load the students list
    	if students_list.split(".")[-1].upper() == "XLS":
    		workbook = xlrd.open_workbook(students_list)
    		worksheets = workbook.sheet_names()
    		worksheet = workbook.sheet_by_name(worksheets[0])
    		# Skip the header row
    		for row in range(1, worksheet.nrows):
    			student = worksheet.row(row)
    			student_names = worksheet.cell_value(row, int(name_column)-1)
    			try:
    				student_surname = student_names.split(",")[0].strip(" ")
    				student_name = student_names.split(",")[1].strip(" ")
    			except:
    				# Handle the cases where there is no comma in the name cell
    				student_surname = student_names.split(" ")[0].strip(" ")
    				student_name = student_names.split(" ")[1].strip(" ")
    			student_regno = worksheet.cell_value(row, int(regno_column)-1)
    			student_sup = worksheet.cell_value(row, int(sup_column)-1)
    			student_sec = worksheet.cell_value(row, int(sec_column)-1)
    			# Populating
    			# Name
    			mark_form_sup['C3'] = student_name+" "+student_surname
    			mark_form_sec['C3'] = student_name+" "+student_surname
    			# Registration number
    			mark_form_sup['C4'] = student_regno
    			mark_form_sec['C4'] = student_regno
    			# Supervisor
    			mark_form_sup['C5'] = mark_form_sec['C5'] = student_sup
    			# Second assessor
    			mark_form_sup['C6'] = mark_form_sec['C6'] = student_sec
    			# Saving files in the specific folders
    			if not os.path.isdir(marks_dir):
        			os.makedirs(marks_dir)
        		if not os.path.isdir(marks_dir+student_sup+"/"):
        			os.makedirs(marks_dir+student_sup+"/")
    			mark_form_sup_doc.save(marks_dir+student_sup+"/"+student_surname+'_sup.xlsx')
    			if not os.path.isdir(marks_dir+student_sec+"/"):
        			os.makedirs(marks_dir+student_sec+"/")
    			mark_form_sec_doc.save(marks_dir+student_sec+"/"+student_surname+'_sec.xlsx')
    	elif students_list.split(".")[-1].upper() == "XLSX":
    		workbook = load_workbook(students_list, use_iterators=True)
    		first_sheet = workbook.get_sheet_names()[0]
    		worksheet = workbook.get_sheet_by_name(first_sheet)
    		row_iterator = worksheet.iter_rows()
    		# Skip the first row
    		next(row_iterator)
    		for row in row_iterator:
    			student_names = row[int(name_column)-1].value
    			student_regno = row[int(regno_column)-1].value
    			student_sup = row[int(sup_column)-1].value
    			student_sec = row[int(sec_column)-1].value
    			try:
    				tmp = int(student_regno)
    			except:
    				break
    			try:
    				student_surname = student_names.split(",")[0].strip(" ")
    				student_name = student_names.split(",")[1].strip(" ")
    			except:
    				# Handle the cases where there is no comma in the name cell
    				student_surname = student_names.split(" ")[0].strip(" ")
    				student_name = student_names.split(" ")[1].strip(" ")
    			# Populating
    			# Name
    			mark_form_sup['C3'] = student_name+" "+student_surname
    			mark_form_sec['C3'] = student_name+" "+student_surname
    			# Registration number
    			mark_form_sup['C4'] = student_regno
    			mark_form_sec['C4'] = student_regno
    			# Supervisor
    			mark_form_sup['C5'] = mark_form_sec['C5'] = student_sup
    			# Second assessor
    			mark_form_sup['C6'] = mark_form_sec['C6'] = student_sec
    			# Saving files in the specific folders
    			if not os.path.isdir(marks_dir):
        			os.makedirs(marks_dir)
        		if not os.path.isdir(marks_dir+student_sup+"/"):
        			os.makedirs(marks_dir+student_sup+"/")
    			mark_form_sup_doc.save(marks_dir+student_sup+"/"+student_surname+'_sup.xlsx')
    			if not os.path.isdir(marks_dir+student_sec+"/"):
        			os.makedirs(marks_dir+student_sec+"/")
    			mark_form_sec_doc.save(marks_dir+student_sec+"/"+student_surname+'_sec.xlsx')
    	else:
    		raise Exception("Unable to serve this type of file")
    	
    	print "Generate zip"
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

    @cherrypy.expose
    def get_marks(self):
    	marks_file = open("marks.zip", 'r')
    	return cherrypy.lib.static.serve_fileobj(marks_file, disposition='attachment', name="marks.zip")

if __name__ == '__main__':
    cherrypy.server.socket_host = '0.0.0.0'
    cherrypy.quickstart(HomePage())
