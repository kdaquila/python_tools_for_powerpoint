import os
import glob
import shutil
from datetime import datetime


def get_curr_date_str():
    return datetime.now().strftime('%m.%d.%y')


# TODO revaluate if we need this
def copy_files(dest_folder, src_folder, patterns=()):
    for pattern in patterns:
        files = glob.glob(os.path.join(src_folder, pattern))
        for f in files:
            shutil.copyfile(src=f, dst=os.path.join(dest_folder, os.path.split(f)[1]))


# TODO revaluate if we need this
def make_dir_if_needed(directory_path):
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)


# TODO revaluate if we need this
def delete_previous_pptx(directory):
    # Delete all previous saved images
    files = glob.glob(os.path.join(directory, '*.pptx'))
    for f in files:
        os.remove(f)


# TODO revaluate if we need this
def clear_directory(directory, exception_list=()):
    files = os.listdir()
    for f in files:
        if f not in exception_list:
            os.remove(f)


# TODO revaluate if we need this
def modify_path(input_path, suffix="_edit"):
    """
    This function modifies an file path by inserting the given suffix between the base file name
    and the file extension. For example, C:/dog/cat.txt becomes C:/dog/cat_edit.txt, for a
    suffix '_edit'

    :param input_path: the path string to modify
    :param suffix: the string to insert into the path
    :return: the modified path string
    """
    head, tail = os.path.split(input_path)
    basename, extension = os.path.splitext(tail)
    return os.path.join(head, basename + suffix + extension)





