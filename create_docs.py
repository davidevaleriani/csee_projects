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

template_filename = "template_rs.docx"     # Model of the docx file to fill
csv_filename = "ce101_2016.csv"        # List of students of CE101 (registration number; surname, first name(s); .....)
feedback_directory = "feedback/"        # Directory where the filled feedback documents will be saved
marks_filename = "marks.csv"            # registration number;mark1;mark2...

if len(sys.argv) > 3:
    print("Using: python %s [docx template] [CSV list of students]" % (sys.argv[0]))
    exit(1)
elif len(sys.argv) == 3:
    template_filename = sys.argv[1]
    csv_filename = sys.argv[2]
elif len(sys.argv) == 2:
    template_filename = sys.argv[1]
else:
    print("INFO: Using default %s and %s files" % (template_filename, csv_filename))
# MAIN

# Check if the template is accessible
if not os.path.isfile(csv_filename):
    print("Error: unable to open the file", csv_filename)
    exit(1)
csv_file = open(csv_filename, 'rU')
counter = 0
for row in csv.reader(csv_file, delimiter=",", quotechar='"'):
    # Skip the header row
    try:
        reg_number = int(row[0])
    except:
        continue
    # Skip empty rows
    if row[0] == "":
        continue
    # Extract information about the student
    #surname = row[1].split(" ")[0]
    #first_name = row[1].split(",")[1]
    full_name = row[3]
    group_name = row[-1].split("Group ")[1]
    print("Processing", reg_number, full_name)
    #output_filename = str(reg_number)+"_Group_"+group_name+"_Team_Report_Feedback.docx"
    output_filename = str(reg_number)+"_Reflective_Statement.docx"
    # Create the directory to store feedback
    if not os.path.isdir(feedback_directory):
        os.makedirs(feedback_directory)
    # Check if the template is accessible
    if not os.path.isfile(template_filename):
        print("Error: unable to open the file", template_filename)
        exit(1)
    else:
        # Open template docx
        zip = zipfile.ZipFile(template_filename)
        # Read the xml document, that is basically the "proper" document file
        word_xml = zip.read('word/document.xml')
        # Get the XML tree
        tree = etree.fromstring(word_xml)
        word_schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        # Add the name of the student near the Name label
        for node in tree.iter(tag=etree.Element):
            if node.tag == '{%s}%s' % (word_schema, 't'):
                if "Name:" in node.text:
                    node.text = "Name: "+full_name
                if len(node.text) == 1:
                    node.text = ""
            if node.tag == '{%s}%s' % (word_schema, 'checkBox'):
                for c in node.iterchildren():
                    if c.tag == '{%s}%s' % (word_schema, 'default'):
                        #c.set(c.keys()[0], "1")
                        pass


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
        with zipfile.ZipFile(feedback_directory+zip_copy_filename, "w") as docx:
            for filename in filenames:
                docx.write(os.path.join(tmp_dir, filename), filename)

        # Clean up the temp dir
        shutil.rmtree(tmp_dir)
    counter += 1

print("Process finished:", counter, "file created")
