# Python3 script to rename files in a specified folder using an input CSV
#
# Author: Simon Di Florio
# Email: simon.diflorio@apa.com.au
#
# Date: 5/03/2021

import os
import sys
import csv
import time
import xlsxwriter
import pandas


def main(argv):
    """
    Main Script Function
    """

    # "C:\Users\simdif\OneDrive - APA Group\Downloads\Book1.csv"
    # C:\Users\simdif\OneDrive - APA Group\Krill49\Text_Files

    print("This script takes an input CSV containing document numbers and an input  directory,\n" +
          "then renames the files in the directory to the specified document numbers in the CSV\n\n" +
          "NOTE: This script WILL NOT create new files or change file extensions of existing files,\n" +
          "so ensure that the target directory already contains the \n" +
          "correct number and types of files prior to running\n")

    # Prompt user for CSV, and try reading it
    print("1. Input full path to CSV containing document numbers\n" +
          "   (Shift + Right click CSV file in file explorer, select 'Copy as Path' and paste below)")

    # user_input = input("CSV Path: ")
    print()
    csv_list = readInputCSV(
        "C:\\Users\\simdif\\OneDrive - APA Group\\Downloads\\Book1.csv")
    while isinstance(csv_list, Exception):
        user_input = input("\nInvalid CSV Path, try again: ")
        print()
        csv_list = readInputCSV(user_input)

    # Prompt user for target directory and get list of directory contents
    print("2. Input path to target directory, or leave blank to target the script's current directory")

    user_input = input("Directory Path: ")
    print()
    directory_list, input_directory = openInputDir(user_input)
    while isinstance(directory_list, Exception):
        user_input = input("\nInvalid directory path, try again: ")
        print()
        directory_list, input_directory = openInputDir(user_input)

    t0 = time.perf_counter()  # start timer

    # Get file extensions from CSV and directory, and add full path to every item
    csv_list = list(zip(csv_list, [os.path.splitext(document)[1]
                                   for document in csv_list]))
    csv_list = [("{0}\\{1}".format(input_directory, document[0]), document[1])
                for document in csv_list]

    directory_list = list(zip(directory_list, [os.path.splitext(document)[1]
                                               for document in directory_list]))
    directory_list = [("{0}\\{1}".format(input_directory, document[0]), document[1])
                      for document in directory_list]

    # iterate through document and directory list, renaming files in target directory
    try:
        renameDocuments(csv_list, directory_list)
    except Exception as e:
        print(e)

    t1 = time.perf_counter()  # stop timer
    print("\nScript took " +
          "{:.6f}".format(t1 - t0) +
          " seconds to run")

    # exit program prompt
    input("Press enter to exit ")


def readInputCSV(input_CSV: str):
    # TODO Implement PANDAS variant
    """
    Try reading an input CreateDoc CSV file and return the list of contained document numbers
    If invalid, try again
    """
    return_doc_list = []
    input_CSV = input_CSV.replace('\"', "")
    try:
        with open(input_CSV, "r") as csvFile:
            csv_reader = csv.reader(csvFile)
            print("Valid CSV, reading...")

            for line in csv_reader:
                if line and line[0] not in return_doc_list:
                    stripped_line = str.strip(line[0])
                    if ' ' in stripped_line:
                        print(stripped_line +
                              " contains whitespace... skipping")
                    else:
                        return_doc_list.append(stripped_line)
    except Exception as exc:
        print(exc)
        return Exception

    print("Finished reading CSV\n")
    return return_doc_list


def openInputDir(input_dir: str):
    """
    Try opening a directory and listing its contents
    If invalid, try again
    """
    try:
        if input_dir == "":
            dir_list = os.listdir()
            input_dir = os.path.dirname(os.path.abspath(__file__))
        else:
            input_dir = input_dir.replace('\"', "")
            dir_list = os.listdir(input_dir)

        print("Successfully found directory\n")
        return dir_list, input_dir
    except Exception as exc:
        print(exc)
        return exc, input_dir


def buildDocument(doc_number: tuple):
    """
    Parse the input document extension, build the appropriate file, and close it
    Supports: .xlsx, and .txt files
    """
    if doc_number[1] == ".xlsx":
        excel_doc = xlsxwriter.Workbook(doc_number[0])
        excel_doc.close()
        return True

    elif doc_number[1] == ".txt":
        txt_doc = open(doc_number[0], "w")
        txt_doc.close()
        return True

    else:
        return False


def renameDocuments(number_list: list, dir_list: list):
    # TODO create docx, xlsx and txt files if they don't already exist
    """
    Loop through document and directory list and rename documents accordingly
    """
    rename_counter = 0
    file_renamed = [False] * len(dir_list)
    max_dir_index = len(dir_list) - 1

    for i, doc_number in enumerate(number_list):
        for j, initial_file in enumerate(dir_list):
            # avoid making copies of existing document numbers
            if (doc_number in dir_list):
                print(str(i) + ".\t" + os.path.basename(doc_number[0]) +
                      " already exists in directory... skipping")
                break
            else:
                # avoid renaming previously written doc numbers and only rename matching extensions
                if (initial_file[0] in [doc_number[0] for doc_number in number_list]
                        or doc_number[1] != initial_file[1]
                        or file_renamed[j]):
                    if j == max_dir_index:
                        if buildDocument(doc_number):
                            print(str(i) + ".\t" + "Built new file: " +
                                  os.path.basename(doc_number[0]))
                            rename_counter += 1
                        else:
                            print(str(i) + ".\t" + "Couldn't find suitable file to rename to " +
                                  os.path.basename(doc_number[0]))
                    else:
                        continue
                # rename current file to specified document number
                else:
                    os.rename(initial_file[0], doc_number[0])
                    rename_counter += 1
                    file_renamed[j] = True
                    print(str(i) + ".\t" + os.path.basename(initial_file[0]) +
                          " -> " + os.path.basename(doc_number[0]))
                    break

    # final status message
    if rename_counter == 1:
        print("\nRenamed " + str(rename_counter) +
              " file out of " + str(len(number_list)))
    elif rename_counter > 1:
        print("\nRenamed " + str(rename_counter) +
              " files out of " + str(len(number_list)))
    else:
        print("Finished processing, no files renamed")


# driver code
if __name__ == "__main__":
    main(sys.argv)
