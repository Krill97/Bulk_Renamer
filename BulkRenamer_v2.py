# Python3 script to rename files in a specified folder using an input CSV
#
# Author: Simon Di Florio
# Email: simon.diflorio@apa.com.au
#
# Date: 5/03/2021

import os
import csv
import time
import matplotlib
import pandas


def main():
    """
    Main Script Function
    """

    # "C:\Users\simdif\OneDrive - APA Group\Downloads\Book1.csv"
    # C:\\Users\\simdif\\OneDrive - APA Group\\Krill49\\Text_Files

    # print("PYTHONPATH:", os.environ.get('PYTHONPATH'))
    # print("PATH:", os.environ.get('PATH'))
    # print("\n")

    print("This script takes an input CSV containing document numbers and an input directory,\n" +
          "then renames the files in the directory to the specified document numbers in the CSV\n\n" +
          "NOTE: This script WILL NOT create new files or change file extensions of existing files,\n" +
          "so ensure that the target directory already contains the \n" +
          "correct number and types of files prior to running\n")

    # Prompt user for CSV, and try reading it
    print("1. Input full path to CSV containing document numbers\n" +
          "   (Shift + Right click CSV file in file explorer, select 'Copy as Path' and paste below)")
    csv_list = readInputCSV()

    # Prompt user for target directory and get list of directory contents
    print("2. Input path to target directory, or leave blank to target the script's current directory")
    directory_list, input_directory = openInputDir()

    t0 = time.perf_counter()  # start timer

    # Get file extensions from CSV and directory, and add full path to every item
    csv_list_ext = [os.path.splitext(document) for document in csv_list]
    csv_list = ["{0}\\{1}".format(input_directory, document)
                for document in csv_list]

    directory_list_ext = [os.path.splitext(
        document) for document in directory_list]
    directory_list = ["{0}\\{1}".format(
        input_directory, document) for document in directory_list]

    # iterate through document and directory list, renaming files in target directory
    try:
        renameDocuments(csv_list, directory_list,
                        csv_list_ext, directory_list_ext)
    except Exception as e:
        print(e)

    t1 = time.perf_counter()
    print("\nScript took " + "{:.6f}".format(t1 -
                                             t0) + " seconds to run")  # stop timer

    # exit program prompt
    input("Press enter to exit ")


def readInputCSV():
    """
    Try reading an input CreateDoc CSV file and return the list of contained document numbers
    If invalid, try again
    """
    return_doc_list = []
    input_CSV = input("CSV Path: ")
    print()

    while True:
        try:
            input_CSV = input_CSV.replace('\"', "")

            with open(input_CSV, "r") as csvFile:
                csv_reader = csv.reader(csvFile)

                for line in csv_reader:
                    if line and line[0] not in return_doc_list:
                        stripped_line = str.strip(line[0])
                        if ' ' in stripped_line:
                            print(stripped_line +
                                  " contains whitespace... skipping")
                        else:
                            return_doc_list.append(stripped_line)

            print("Successfully read CSV\n")
            return return_doc_list
        except Exception as e:
            print(e)
            input_CSV = input("\nInvalid CSV, try again: ")
            print()


def openInputDir():
    """
    Try opening a directory and listing its contents
    If invalid, try again
    """
    input_dir = input("Directory Path: ")
    print()
    while True:
        try:
            if input_dir == "":
                dir_list = os.listdir()
                input_dir = os.path.dirname(os.path.abspath(__file__))
            else:
                input_dir = input_dir.replace('\"', "")
                dir_list = os.listdir(input_dir)

            print("Successfully found directory\n")
            return dir_list, input_dir
        except:
            input_dir = input("Invalid directory path, try again: ")
            print()


def renameDocuments(number_list, dir_list, number_ext, dir_ext):
    """
    Loop through document and directory list
    and rename documents accordingly
    """
    files_renamed = 0
    dir_renamed_index = [False] * len(dir_list)

    for i, document in enumerate(number_list):
        for j, file in enumerate(dir_list):
            # avoid making copies of existing document numbers
            if (document in dir_list):
                print(str(i) + ". " + os.path.basename(document) +
                      " already exists in directory... skipping")
                break
            else:
                # avoid renaming previously written doc numbers and only rename matching extensions
                if dir_list[j] in number_list or number_ext[i][1] != dir_ext[j][1] or dir_renamed_index[j]:
                    if j == len(dir_list) - 1:
                        print("Couldn't find suitable file to rename to " +
                              os.path.basename(document))
                    else:
                        continue
                # rename current file to specified document number
                else:
                    os.rename(file, document)
                    files_renamed += 1
                    dir_renamed_index[j] = True
                    print(str(i) + ". " + os.path.basename(file) +
                          " -> " + os.path.basename(document))
                    break

    # final status message
    if files_renamed == 1:
        print("\nRenamed " + str(files_renamed) +
              " file out of " + str(len(number_list)))
    elif files_renamed > 1:
        print("\nRenamed " + str(files_renamed) +
              " files out of " + str(len(number_list)))
    else:
        print("Finished running, no files renamed")


# driver code
if __name__ == "__main__":
    main()
