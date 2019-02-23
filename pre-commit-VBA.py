'''
pre-commit-VBA.py

Last updated 23/02/2019

Script extracts VBA modules, forms and class modules into a new 'FILENAME.VBA' subdirectory within the repository root folder (by default)

With the standard .gitignore entries, only Excel files located within the root directory of the repository will be processed & added to a commit
'''

import os
import shutil

from oletools.olevba3 import VBA_Parser

# list excel extensions that will be processed and have VBA modules extracted
excel_file_extensions = ('xlsb', 'xls', 'xlsx', 'xlsm', 'xla', 'xlt', 'xlam', 'xltm')

# Set this to 'True' if you would like to retain the file headers in CLASS (.cls) and FORM (.frm) modules
keep_header = False

# process Excel files in this directory (not recursive to subdirectories)
directory = '.'

# String to append to end of subdirectory
VBA_suffix = '.VBA'

# String to append to end of subdirectory based on filename
Rev_tag = ' - Rev '


# function to extract the VBA modules from an Excel archive
def extract_VBA_files(workbook_name):

    # Find index of standard '- Rev X.xxx' tag in workbook filename if it exists
    index_rev = workbook_name.find(Rev_tag)

    # If revision tag existed in the filename, remove it based on index position, else remove only the file extension
    if index_rev != -1:
        filename = workbook_name[0:index_rev]
    else:
        filename = os.path.splitext(workbook_name)[0]

    # Setup directory name based on filename
    vba_path = filename + VBA_suffix

    # Make new directory (to replace existing where it previously existed and was removed)
    os.mkdir(vba_path)

    # print to console for information/debugging purposes
    max_len = max(len('Directory name -- ' + vba_path), len('Workbook name  -- ' + workbook_name))
    print()
    print('-' * max_len)
    print('Workbook name  -- ' + workbook_name)
    print('Directory name -- ' + vba_path)
    print('-' * max_len)

    # setup for removal of VBA modules
    vba_parser = VBA_Parser(workbook_name)
    vba_modules = vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []

    # extract VBA modules
    for _, _, module_name, content in vba_modules:
        decoded_content = content.decode('latin-1')
        lines = []
        if '\r\n' in decoded_content:
            lines = decoded_content.split('\r\n')
        else:
            lines = decoded_content.split('\n')
        if lines:
            # check if module type is a CLASS (.cls) or FORM (.frm)
            if module_name[-4:] in ['.cls', '.frm']:
                # CLASS (.cls) & FORM (.frm) modules have more non-code lines at the beginning,
                # these headers can be left or removed based on the keep_header variable
                content = lines if keep_header else lines[8:]
            else:
                content = lines if keep_header else lines[1:]
            if content and content[-1] == '':
                content.pop(len(content) - 1)
                # check for empty modules (not processed if empty, modules won't be empty if headers are kept even if there is no content)
                non_empty_lines_of_code = len([c for c in content if c])

                if non_empty_lines_of_code > 0:

                    with open(os.path.join(vba_path, module_name), 'w', encoding='utf-8') as module_file:
                        module_file.write('\n'.join(content))
                        # print to console for information/debugging purposes
                        print('Extracted -- ' + module_name + ' from ' + workbook_name)


# remove all previous '.VBA' directories including contents
for directories in os.listdir(directory):
    if directories.endswith(VBA_suffix):
        shutil.rmtree(directories)

# loop through files in given directory and process those that are Excel files (excluding temporary Excel files ~$*)
for filenames in os.listdir(directory):
    if filenames.endswith(excel_file_extensions):
        # skip temporary excel files ~$*.*, otherwise process Excel file and extract VBA modules
        if not filenames.startswith('~$'):
            extract_VBA_files(filenames)

# print trailing line to separate output
print()

# loop through directories & remove '.VBA' directory if nothing was extracted and it is therefore empty
for directories in os.listdir(directory):
    if directories.endswith(VBA_suffix):
        if not os.listdir(directories):
            os.rmdir(directories)
