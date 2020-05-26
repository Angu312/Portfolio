# This script takes user inputs to create directories (if not already existing) and moves files into their respective folders based on file extensions.

import os
import shutil
import pickle

if os.path.isfile('C:/Users/Angu312/Desktop/Current Build.txt'):
    os.chdir('C:/Users/Angu312/Desktop/')
    #open/read the file and assign each list item in a separate variable
    fileObject = open('Current Build.txt','rb')
    #load the object from the file into var Inputs
    Inputs = pickle.load(fileObject)
    Year = Inputs[0]
    Project = Inputs[1]
    Trial = Inputs[2]

if not os.path.isfile('C:/Users/Angu312/Desktop/Current Build.txt'):
    os.chdir('C:/Users/Andrew/Desktop/')
    Year = str(input("What year does this project take place in? "))
    if os.path.exists(Year):
        os.chdir(Year)
    else:
        os.mkdir(Year)
        os.chdir(Year)

    Project = str(input("What is the name of this project? "))
    if os.path.exists(Project):
        os.chdir(Project)
    else:
        os.mkdir(Project)
        os.chdir(Project)

    Trial = "Trial " + str(input("What trial number is this? "))
    if not os.path.exists(Trial):
        folders = ['Build', 'JobReports', 'DataReports', 'Backup']
        for folder in folders:
            os.makedirs(os.path.join(Trial, folder))

    Inputs = [Year, Project, Trial]
    os.chdir('C:/Users/Angu312/Desktop/')
    file_Name = "Current Build.txt"
    # open the file for writing
    fileObject = open(file_Name,'wb')
    # this writes to the object Inputs to the file named 'Current Build.txt'
    pickle.dump(Inputs,fileObject)
    # here we close the fileObject
    fileObject.close()

# Lines 8-9 sets both the source and destination directories in variables
DESKTOP = 'C:/Users/Angu312/Desktop'
MUSIC = 'C:/Users/Angu312/Music'
VIDEOS = 'C:/Users/Angu312/Videos'
DESTINATION1 = os.path.join('C:/Users/Angu312/Desktop/', Year, Project, Trial, 'Build')
# For example, the above function creates a path such as so 'C:/Users/Angu312/Desktop/(Year)/(Project)/(Trial)/Build'
DESTINATION2 = os.path.join('C:/Users/Angu312/Desktop/', Year, Project, Trial, 'JobReports')
DESTINATION3 = os.path.join('C:/Users/Angu312/Desktop/', Year, Project, Trial, 'DataReports')
DESTINATION4 = os.path.join('C:/Users/Angu312/Desktop/', Year, Project, Trial, 'Backup')

# Lines 14-24 will check the source directory for desired file extension. It will then copy files with the extension to the destination directory.
for files in os.listdir(DESKTOP):
    if files.lower().endswith('.xlsx'):
        shutil.copy(os.path.join(DESKTOP, files), DESTINATION1)
for files in os.listdir(MUSIC):
    if files.lower().endswith('.pptx'):
        shutil.copy(os.path.join(MUSIC, files), DESTINATION2)
for files in os.listdir(VIDEOS):
    if files.lower().endswith('.docx'):
        shutil.copy(os.path.join(VIDEOS, files), DESTINATION3)
