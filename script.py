from shutil import copyfile
from openpyxl import load_workbook, Workbook
import zipfile, csv, re, fnmatch, os, time, datetime, random, statistics

# Change the names in here to the ones you have available.
staff = []
assesment_for_each_staff = []
id_staff = []
num_assesmentfolder = 0
log_fil = open("script_Log.log", "w")
distribution_grade = dict()
grade_value = []

def log(msg):
    ts = time.time()
    st = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')
    log_fil.write(st + ':\t' + msg + '\n')

def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file))

def make_zip():
    for i in range(len(staff)):
        zf = zipfile.ZipFile(staff[i] +'.zip', 'w', zipfile.ZIP_DEFLATED)
        for j in range(len(id_staff[i])):
            try:
                zipdir(id_staff[i][j]+'/', zf)
            finally:
                pass
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
    input_name = input('Type in names of who is going to assess. Use comma between if there is more than one:')
    global staff
    staff = [x.strip() for x in input_name.split(',')]
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


def create_sheet_header_info(wb):
    ws = wb.create_sheet("Distribution", 0)
    ws.cell(row=1, column=1, value='ID')
    ws.cell(row=1, column=2, value='Teacher')
    ws.cell(row=1, column=3, value='Start')
    ws.cell(row=1, column=4, value='Finished')
    ws.cell(row=1, column=5, value='Points')
    ws.cell(row=1, column=6, value='Grade')
    ws.cell(row=1, column=7, value='Comments')

    for x in range(0,len(staff)):
        ws.cell(row=x+3, column=10, value=staff[x])
        ws.cell(row=x+3, column=11, value=assesment_for_each_staff[x])

    return ws


def create_feedbackfiles():
    log('Create_feedbackfile')
    file_list = print_dir(('.docx', '.doc'))
    log(str(file_list))
    feedback_file_path = file_list[int(input('Type in the number of the feedback file: '))]
    path = os.getcwd()
    folder_name = os.path.dirname(path)
    log("Path: " + path)
    log("Folder name: " + folder_name)
    wb = Workbook()
    ws = create_sheet_header_info(wb)
    log('Created Headers to Distribution.xlsx')
    i = 0
    x = 0
    z = 1
    for dirname, dirnames, filenames in os.walk('.'):
        # print path to all subdirectories first.
        for subdirname in dirnames:
            if fnmatch.fnmatch(subdirname, '*_assign*'):
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
                    z += 1
                    log('studentID:' + studentID[0])
                    ws.cell(row=z, column=1, value=studentID[0])
                    ws.cell(row=z, column=2, value=staff_name)
                    wb.save('Distribution.xlsx')
                id_staff[i].append(subdirname)
                copyfile(path + '/' + feedback_file_path, path + '/' + subdirname + '/' + subdirname + '.docx')
                log('Made new file: ' + path + '/' + subdirname + '/' + subdirname + '.docx')

    log('Saved Distribution.xlsx')
    log('Folders to zip: '+ str(id_staff))

    make_zip()

    print(
        '\nProcess done!\nYou will find a document in each folder and a Distribution.xlsx with all student identifikation numbers')
    log(
        'Process done! You will find a document in each folder and a Distribution.xlsx with all student identifikation numbers')


def merge_csv_sheet():
    read_xlsx_file()
    log('Done read_xlsx_file')
    read_csv_file()
    log('Done read_csv_file')
    print(
        '\nProcess done!\nYou will find a NEW csv file ready to upload to moodle')
    log(
        'Process done! You will find a NEW csv file ready to upload to moodle')

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

def calculate_stats():
    try:
        stat_file = open('CP_statistics.txt', 'w')
        stat_file.write('Mean:\t {} - Arithmetic mean (“average”) of data.\n'.format(statistics.mean(grade_value)))
        stat_file.write('Median:\t {} - Median (middle value) of data.\n'.format(statistics.median(grade_value)))
        stat_file.write('Median_low:\t {} - Low median of data.\n'.format(statistics.median_low(grade_value)))
        stat_file.write('Median_high:\t {} - High median of data.\n'.format(statistics.median_low(grade_value)))
        stat_file.write('Median_grouped:\t {} - Median, or 50th percentile, of grouped data.\n'.format(statistics.median_grouped(grade_value)))
        stat_file.write('Population standard deviation:\t {} - Population standard deviation of data.\n'.format(statistics.pstdev(grade_value)))
        stat_file.write('Population variance:\t {} - Population variance of data.\n'.format(statistics.pvariance(grade_value)))
        stat_file.write('Standard deviation:\t {} - Sample standard deviation of data.\n'.format(statistics.stdev(grade_value)))
        stat_file.write('Variance:\t {} - Sample variance of data.\n'.format(statistics.variance(grade_value)))
        stat_file.close()
        log('Done calculated stats')
    except TypeError as msg:
        log(msg)

def read_xlsx_file():
    log('read_xlsx_file')
    file_list = print_dir(('.xlsx'))
    dist_xlsx_path = file_list[int(input('Type in the number of the Distribution sheet:'))]
    wb = load_workbook(filename=dist_xlsx_path, read_only=False)
    ws = wb['Distribution']
    for row in ws.rows:
        if row[0].value == 'ID':
            continue
        global distribution_grade
        distribution_grade[row[0].value] = []
        distribution_grade[row[0].value].append([row[4].value, row[5].value])
        #print(distribution_grade)
        global grade_value
        grade_value.append(row[4].value)
   # calculate_stats()

def read_csv_file():
    log('read_csv_file')
    file_list = print_dir(('.csv'))
    dist_csv_path = file_list[int(input('Type in the number of the Grade sheet from Moodle sheet:'))]

    with open(dist_csv_path, "r",newline='', encoding='utf-8') as csv_file:
        reader = csv.DictReader(csv_file,delimiter=',')

        with open('NEW-Greading-upload.csv', 'w', encoding='utf-8') as csvfile:
            fieldnames = ['\ufeffIdentifier', 'Status', 'Grade', 'Maximum Grade', 'Grade can be changed',
                          'Last modified (submission)',
                          'Last modified (grade)', 'Feedback comments']
            writer = csv.DictWriter(csvfile,delimiter=',', fieldnames=fieldnames)
            writer.writeheader()
            for row in reader:
                p_id = re.findall(r'\d+', row['\ufeffIdentifier'])[0]
                print(row)
                writer.writerow({'\ufeffIdentifier': row['\ufeffIdentifier'],
                                 'Status': row['Status'],
                                 'Grade': distribution_grade[p_id][0][0],
                                 'Maximum Grade': row['Maximum Grade'],
                                 'Grade can be changed': row['Grade can be changed'],
                                 'Last modified (submission)': row['Last modified (submission)'],
                                 'Last modified (grade)': row['Last modified (grade)'],
                                 'Feedback comments': 'Grade = {}'.format(distribution_grade[p_id][0][1])})
prog_to_run = -1
while prog_to_run != 0:
    prog_to_run = int(input('What program/operation do you want to run? Type in the number, 0 to quit:\n'
                        '\t1: Create feedback file in each folder, and collect the student ID in a list.\n'
                        '\t2: Merge grades into feedback file with merge dist.list and Moodle grade sheet.\n'))
                        #'\t3: Keep only the feedback file and remove the students exam in the folder.\n:')

    if prog_to_run == 1:
        select_staff()
        create_feedbackfiles()
    elif prog_to_run == 2:
        merge_csv_sheet()
    elif prog_to_run == 0:
        exit(0)
    else:
        print('Type in one of the number to choose select a script!')

log_fil.close()