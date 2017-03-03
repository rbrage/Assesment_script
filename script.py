from tempfile import NamedTemporaryFile
from shutil import copyfile
from openpyxl import load_workbook
import zipfile, csv, re, fnmatch, os, time, datetime, random

# Change the names in here to the ones you have available.
staff = []
assesment_for_each_staff = []
id_staff = []
num_assesmentfolder = 0
log_fil = open("script_Log.log", "w")


def log(msg):
    ts = time.time()
    st = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')
    log_fil.write(st + ':\t' + msg + '\n')
    log_fil.close()

def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file))

def make_zip():
    print('creating archive')
    for i in range(len(staff)):
        zf = zipfile.ZipFile(staff[i] +'.zip', 'w', zipfile.ZIP_DEFLATED)
        for j in range(len(id_staff[i])):
            try:
                zipdir(id_staff[i][j]+'/', zf)
            finally:
                print()
        log('Created: '+staff[i] + '.zip')
        zf.close()


def print_dir(end_sufix):
    i = 0
    file_list = []
    for dirname, dirnames, filenames in os.walk('.'):
        # print path to all subdirectories first.1
        for docxname in filenames:
            if docxname.endswith(end_sufix):
                print(i, ': ' + docxname)
                file_list.append(docxname)
                i += 1
    return file_list

def distribute_number_of_exam():
    number_of_staff = len(staff)
    number_for_each_staff = (num_assesmentfolder / number_of_staff)
    remaining_assesments = (num_assesmentfolder % number_of_staff)

    for i in range(1, number_of_staff + 1):
        print(i , ' - ', staff[i-1])
        extra = 1 if i <= remaining_assesments else 0
        assesment_for_each_staff.append(int(number_for_each_staff + extra))

    log('Distribution: '+str(assesment_for_each_staff))
    log('Sum of Distribution: '+str(sum(assesment_for_each_staff)))

def select_staff():
    input_name = input('Type in names on how is going to assess. Use comma between if there is more than one.\n:')
    global staff
    staff = input_name.split(',')
    random.shuffle(staff)
    log('#{} - Staff: {}'.format(len(staff), str(staff)))
    global num_assesmentfolder
    num_assesmentfolder = count_assesment_folders()
    global id_staff
    id_staff = [[] for i in range(len(staff))]
    distribute_number_of_exam()


def count_assesment_folders():
    number_of_assesments = 0
    for dirname, dirnames, filenames in os.walk('.'):
        for subdirname in dirnames:
            if fnmatch.fnmatch(subdirname, '*_assign*'):# or fnmatch.fnmatch(subdirname, 'D*'):
                number_of_assesments += 1
    return number_of_assesments


def create_feedbackfiles():
    log('Create_feedbackfile')
    file_list = print_dir(('.docx', '.doc'))
    log(str(file_list))
    feedback_file_path = file_list[int(input('Type in the number of the feedback file: '))]
    path = os.getcwd()
    folder_name = os.path.dirname(path)
    log("Path: " + path)
    log("Folder name: " + folder_name)
    file = open("StudentID.txt", "w")
    log('Created StudentID')
    i = 0
    x = 0
    for dirname, dirnames, filenames in os.walk('.'):
        # print path to all subdirectories first.
        for subdirname in dirnames:
            if fnmatch.fnmatch(subdirname, '*_assign*'):# or fnmatch.fnmatch(subdirname, 'D*'):
                log('subdirname: ' + subdirname)
                studentID = re.findall(r'\d+', subdirname)
                log('len(studentID):' + str(len(studentID)))
                log('assessment_for_each_staff {} '.format(assesment_for_each_staff))
                log('x: {} -- i: {}'.format(x, i))
                log('Staff: {}'.format(staff))
                if x >= assesment_for_each_staff[i]:
                    i += 1
                    x = 0
                staff_name = staff[i]
                x += 1
                if len(studentID) == 1:
                    log('studentID:' + studentID[0])
                    file.write(studentID[0] + '\t' + staff_name + '\n')
                id_staff[i].append(subdirname)
                copyfile(path + '/' + feedback_file_path, path + '/' + subdirname + '/' + subdirname + '.docx')
                log('Made new file: ' + path + '/' + subdirname + '/' + subdirname + '.docx')
    log('Folders to zip: '+ str(id_staff))
    file.close()

    make_zip()

    print(
        '\nProcess done!\nYou will find a document in each folder and a StudentID.txt with all student identifikation numbers')
    log(
        '\nProcess done!\nYou will find a document in each folder and a StudentID.txt with all student identifikation numbers')


def merge_csv_sheet():
    file_list = print_dir(('.csv'))
    print(file_list)
    dist_csv_path = file_list[int(input('Type in the number of the Distribution sheet: '))]
    print(dist_csv_path)



    # with open(dist_csv_path) as csv_file:
    #     reader = csv.reader(csv_file, delimiter=',')
    #     print(reader)
    #     for row in reader:
    #         print(row)
    # grade_sheet_path = = file_list[int(input('Type in the number of the Grade sheet from moodle: '))]


def delete_student_exam():
    for dirname, dirnames, filenames in os.walk('.'):
        # print path to all subdirectories first.
        for docxname in filenames:
            print(dirname)
            print(': ' + docxname)
            if not fnmatch.fnmatch(docxname, 'FS*'):
                if not fnmatch.fnmatch(docxname, '*py'):
                    print('File: ' + docxname)
                    os.remove(dirname + '/' + docxname)

def read_xlsx_file():
    file_list = print_dir(('.xlsx'))
    print(file_list)
    dist_xlsx_path = file_list[int(input('Type in the number of the Distribution sheet: '))]
    print(dist_xlsx_path)
    wb = load_workbook(filename=dist_xlsx_path, read_only=False)
    print(wb.sheetnames)
    ws = wb['Sheet1']
    for a_cell in ws['A']:
        print(a_cell.value)

def read_csv_file():
    file_list = print_dir(('.csv'))
    print(file_list)
    dist_csv_path = file_list[int(input('Type in the number of the Distribution sheet: '))]
    print(dist_csv_path)




    # with open(dist_csv_path) as csv_file:
    #     reader = csv.reader(csv_file, delimiter=',')
    #     for row in reader:
    #         print(row)

prog_to_run = input('What program/operation do you want to run? Type in the number:\n'
                    '\t1: Create feedback file in each folder, and collect the student ID in a list.\n')
                    # '\t2: Merge grades into feedback file with merge dist.list and Moodle grade sheet.\n'
                    # '\t3: Read csv file.\n')
                    #'\t3: Keep only the feedback file and remove the students exam in the folder.\n:')
prog_to_run = int(prog_to_run)


if prog_to_run == 1:
    select_staff()
    create_feedbackfiles()
elif prog_to_run == 2:
    merge_csv_sheet()
elif prog_to_run == 3:
    read_csv_file()
elif prog_to_run == 4:
    delete_student_exam()
else:
    print('Type in one of the number to choose select a script!')

'''
#writer = csv.writer(open('Karakterer-IIS2016-Course Project IIS - Response Upload-26344.csv', 'w', delimiter=',', quotechar='"'))


print('Dist_file done\n')

with open('Karakterer-IIS2016-Course Project IIS - Response Upload-26344.csv') as csv_file:
    spam_reader = csv.reader(csv_file, delimiter=',', quotechar='"')
    for row1 in spam_reader:
        print(row1)





from shutil import copyfile
import os
import fnmatch
import re

log_fil = open("Log.txt", "w")
def log(msg):
    log_fil.write(msg+ '\n')

path = os.getcwd()
folder_name = os.path.dirname(path)
log("Path: " + path)
log("Folder name: " + folder_name)

print('path: '+ path+'\n')
print('-----------------------')
log('-----------------------')
file = open("StudentID.txt", "w")
log('Created StudentID')
docXname = ''
for dirname, dirnames, filenames in os.walk('.'):
    # print path to all subdirectories first.
    for docxname in filenames:
        log('docxname: '+docxname)
        if fnmatch.fnmatch(docxname, 'FS*'):
            docXname = docxname

if docXname == '':
    print('Error! No FS file in this folder!')
    log('Error! No FS file in this folder!')
    quit()

for dirname, dirnames, filenames in os.walk('.'):
    # print path to all subdirectories first.
    for subdirname in dirnames:
        if 'P' or 'D' in subdirname:
            log('subdirname: ' + subdirname)
            studentID = re.findall(r'\d+', subdirname)
            log('len(studentID):'+ str(len(studentID)))
            if len(studentID) == 1:
                log('studentID:'+studentID[0])
                file.write(studentID[0]+'\n')
            print('Made new file: ' + path + '/' + docXname)
            log('Made new file: ' + path + '/' + docXname)
            copyfile(path + '/' + docXname, path+'/'+subdirname+'/' + subdirname+'.docx')
        else:
            print('ERROR! - Could not make file!')
            log('ERROR! - Could not make file!')

file.close()

print('\nProcess done!\nYou will find a document in each folder and a StudentID.txt with all student identifikation numbers')
log('\nProcess done!\nYou will find a document in each folder and a StudentID.txt with all student identifikation numbers')


        '''
