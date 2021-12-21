
# import OS module
import os

# Get the list of all files and directories
path = "C:\\Users\\wisitl\\Downloads"
dir_list = os.listdir(path)

print("Files and directories in '", path, "' :")

# prints all files
for i in dir_list:
    print(i)