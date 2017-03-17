from shutil import move
import os, fnmatch, sys
path = os.path.dirname(os.path.realpath(sys.argv[0]))

def makedir(path_competed):
    files = os.listdir(path=path_competed)
    for docxname in files:
        os.mkdir(os.path.join(path_competed,docxname.strip('.docx')))
        move(os.path.join(path_competed,docxname), os.path.join(path_competed,docxname.strip('.docx')))


def print_dir(end_sufix):
    print("Path: " + path)
    i = 0
    file_list = []
    files = os.listdir(path=path)
    for docxname in files:
        if docxname.endswith(end_sufix):
            print(i, ': ' + docxname)
            file_list.append(os.path.join(path,docxname))
            i += 1
    return file_list

