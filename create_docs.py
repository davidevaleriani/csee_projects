#######################################################################
# CE101 Marking
#
#   This script reads a list of students from a CSV file and a template
#   document from a docx file and generates several documents, one
#   per student, with the format regNo_Team_Report_Feedback.docx and
#   including the name of the student.
#   The CSV file should be formatted with
#       registrationNumber; surname, first name(s)
#   It has been used in the marking process of CE101 module @
#   University of Essex
#
#   Command line:
#   python create_docs.py [template.docx] [list_of_students.csv]
#
#
# Author: Davide Valeriani
#         Brain-Computer Interfaces Lab
#         University of Essex
#
#######################################################################

import sys
from docx import *
import csv
import zipfile
import tempfile
import os
import shutil

# CONFIGURATION

# Type of document (rs = Reflective Statement; tr = Team Report)
type_of_document = "rs"
# Main working directory
main_folder = ""
# Model of the docx file to fill
template_filename = "CE101 Reflective Statement Feedback Form 2015.docx"
# List of students of CE101 (registration number; surname, first name(s); mark1; mark2; ...)
csv_filename = "ce101_2016.csv"
# Directory where the filled feedback documents will be saved
feedback_directory = type_of_document+"/"
# Columns settings (in the CSV file)
regno_col = 0
fullname_col = 1
mark1_col = 2
group_name_col = 3
num_options_available_marks = [2, 4, 2, 3, 3, 3, 2, 2]

# MAIN

# Check if the template is accessible
if not os.path.isfile(main_folder+csv_filename):
    print("Error: unable to open the file", main_folder+csv_filename)
    exit(1)
csv_file = open(main_folder+csv_filename, 'rU')
files_created = 0
for row in csv.reader(csv_file, delimiter=",", quotechar='"'):
    # Skip the header row
    try:
        reg_number = int(row[regno_col])
    except:
        continue
    # Skip empty rows
    if row[regno_col] == "":
        continue
    # Extract information about the student
    #surname = row[1].split(" ")[0]
    #first_name = row[1].split(",")[1]
    full_name = row[fullname_col]
    if type_of_document == "tr":
        group_name = row[-1].split("Group ")[1]
        output_filename = str(reg_number)+"_Group_"+group_name+"_Team_Report_Feedback.docx"
    elif type_of_document == "rs":
        output_filename = str(reg_number)+"_Reflective_Statement.docx"
    else:
        print "Type of document not supported"
        exit(1)
    print("Processing", reg_number, full_name)
    # Create the directory to store feedback
    if not os.path.isdir(main_folder+feedback_directory):
        os.makedirs(main_folder+feedback_directory)
    # Check if the template is accessible
    if not os.path.isfile(main_folder+template_filename):
        print("Error: unable to open the file", main_folder+template_filename)
        exit(1)
    else:
        # Open template docx
        zip = zipfile.ZipFile(main_folder+template_filename)
        # Read the xml document, that is basically the "proper" document file
        word_xml = zip.read('word/document.xml')
        # Get the XML tree
        tree = etree.fromstring(word_xml)
        word_schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        # Add the name of the student near the Name label
        if type_of_document == "rs":
            current_mark_index = 0
            skip_checkbox = 0
            if row[mark1_col+current_mark_index] == "":
                current_mark = num_options_available_marks[current_mark_index]-1
                enable_marking = False
            else:
                current_mark = int(row[mark1_col+current_mark_index])
                enable_marking = True
        for node in tree.iter(tag=etree.Element):
            if node.tag == '{%s}%s' % (word_schema, 't'):
                if "Name:" in node.text:
                    node.text = "Name: "+full_name
                if len(node.text) == 1:
                    node.text = ""
            if type_of_document == "rs":
                # Fill reflective statement with the marks
                if node.tag == '{%s}%s' % (word_schema, 'checkBox'):
                    for c in node.iterchildren():
                        if c.tag == '{%s}%s' % (word_schema, 'default'):
                            if skip_checkbox > 0:
                                skip_checkbox -= 1
                            elif current_mark == 0:
                                if enable_marking:
                                    c.set(c.keys()[0], "1")
                                    skip_checkbox = num_options_available_marks[current_mark_index] - int(row[mark1_col+current_mark_index]) - 1
                                current_mark_index += 1
                                if (mark1_col + current_mark_index) >= len(row):
                                    continue
                                if row[mark1_col + current_mark_index] == "":
                                    current_mark = num_options_available_marks[current_mark_index] - 1
                                    enable_marking = False
                                else:
                                    current_mark = int(row[mark1_col + current_mark_index])
                                    enable_marking = True
                            else:
                                current_mark -= 1

        # Make temporary directory
        tmp_dir = tempfile.mkdtemp()
        # Extract all the files contained in the docx
        zip.extractall(tmp_dir)

        # Overwrite the xml file with new data
        with open(os.path.join(tmp_dir, 'word/document.xml'), 'wb') as f:
            xmlstr = etree.tostring(tree, pretty_print=True)
            f.write(xmlstr)

        # Get a list of all the files in the original docx zipfile
        filenames = zip.namelist()
        # Now, create the new zip file and add all the filex into the archive
        zip_copy_filename = output_filename
        with zipfile.ZipFile(main_folder+feedback_directory+zip_copy_filename, "w") as docx:
            for filename in filenames:
                docx.write(os.path.join(tmp_dir, filename), filename)

        # Clean up the temp dir
        shutil.rmtree(tmp_dir)
        files_created += 1

print("Process finished:", files_created, "file created")
