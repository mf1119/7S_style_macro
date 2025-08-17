import sys
import os
import shutil
from datetime import datetime
import re

# print (sys.argv)

try:
    file_dir = sys.argv[1]
except IndexError:
    print("You didn't give me a file to work on!")
    exit(1)

if not os.path.isfile(file_dir):
    print(f"File {file_dir} does not exist! Quitting...")
    exit(1)

filename, file_extension = os.path.splitext(file_dir)

now = datetime.now()
datetime_string = now.strftime("%Y%m%d_%H%M%S")

# print(datetime_string)
print(f"Original: {file_dir}\n Backup: {filename}_{datetime_string}{file_extension}")

# Backup original
shutil.copyfile(file_dir, f"{filename}_{datetime_string}{file_extension}")

PAGE_PATTERN = "^Page ([0-9]+)"
PANEL_PATTERN = "^[.]([0-9]+)"
ELEMENT_PATTERN = "^>([A-Za-z])"

output_file = []
page_num = 0
panel_num = 0
element_num = 0
shorthand_dict = {}

with open(file_dir) as f:
    for line in f:
        if re.search(PAGE_PATTERN, line):
            searched_page_num = re.match(PAGE_PATTERN, line)
            page_num = searched_page_num.groups(0)[0]
            element_num = 0
            panel_num = 0

        if re.search(PANEL_PATTERN, line):
            searched_panel_num = re.match(PANEL_PATTERN, line)
            panel_num = searched_panel_num.groups(0)[0]
            line = page_num + line
        
        if re.search(ELEMENT_PATTERN, line):
            element_num+=1
            line = re.sub(ELEMENT_PATTERN, f"{page_num}.{panel_num}.{element_num} \\1", line)
        
        output_file.append(line)


for line in output_file:
    print(line)

with open(file_dir, "w") as f:
    f.writelines(output_file)
