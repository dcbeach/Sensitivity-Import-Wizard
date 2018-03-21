import csv
import os
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from collections import Counter

newFileName = 'SensitivityExcelData'


# Determine list of all saved data fileNames in folder
def getFileNames(inputRoot):
    fileNames = []
    for root, dirs, files in os.walk(inputRoot):  # this was 'HVLD Data'
        for fileName in files:
            fileNames.append(fileName)
    return (fileNames)


# Imports HVLD Data and returns in 3 seperate lists
def importHVLD(fileLocation):
    results = []
    sensitivity = []
    date = []
    tempint = 1
    with open(fileLocation, 'r') as inputfile:
        for line in inputfile:
            tempint += 1
            if tempint == 19:
                sensitivity.append(line.strip().split())
                del sensitivity[0][0]
                del sensitivity[0][1]
                sensitivity = int(sensitivity[0][0])
            elif tempint == 48:
                date.append(line.strip().split())
                del date[0][0]
                date = str(date[0][0])
            elif tempint > 50:
                results.append(line.strip().split())

    return results, sensitivity, date


# Restructures HVLD Import Data into Correct Format
def restructureHVLD(fullResults, sensitivity, date):
    results2 = []
    for result in fullResults:
        result.append(sensitivity)
        result.insert(0, date)
        result.insert(3, result[7])
        del result[8]

        # Convert . and x to pass and fail
        if str(result[7]) == '.':
            result[7] = 'Pass'
        else:
            result[7] = 'Fail'
        results2.append(result)
    return results2


# Appends data to CSV file
def writeDataCSV(results, newFileName):
    with open(newFileName + '.csv', 'a') as csvfile:
        writer = csv.writer(csvfile, lineterminator='\n')
        writer.writerows(results)


# Asks the user to chose a file. Every file within the chosen files folder will be imported
def DetermineFolderPath():
    # code to promt user to chose file
    root = tk.Tk()
    root.withdraw()
    root.fileName = filedialog.askopenfilename()

    # split file string by '/'
    wordlist = root.fileName.split('/')
    # delete the portion with the filename so we can open the whole directory
    del wordlist[len(wordlist) - 1]
    # rejoin the array with '\'
    fileFolderPath = '\\'.join(wordlist)

    return fileFolderPath


# Runs the whole script
def ExecuteSensitivityAnalysis():
    result_total = []
    folder = DetermineFolderPath()
    fileNames = getFileNames(folder)

    for x in fileNames:
        results, sensitivity, date = importHVLD(folder + '\\' + x)
        restructureHVLD(results, sensitivity, date)
        writeDataCSV(results, folder + '\\' + newFileName)
        result_total.append(results)

    draw_graph(result_total, fileNames)


def draw_graph(result_total, file_names):
    simple_results = []
    sensitivities = []
    sample_count_per_file = []
    possible_missing_files = []
    for sensitivity_group in result_total:
        sample_count_per_file.append(len(sensitivity_group))
        for sample_run in sensitivity_group:
            if sensitivities:
                if sample_run[8] in sensitivities:
                    simple_results[sensitivities.index(sample_run[8])].append(float(sample_run[4]))
                else:
                    sensitivities.append(int(sample_run[8]))
                    if len(simple_results) < sensitivities.index(sample_run[8]) + 1:
                        simple_results.append([])
                        simple_results[sensitivities.index(sample_run[8])].append(float(sample_run[4]))
                    else:
                        simple_results[sensitivities.index(sample_run[8])].append(float(sample_run[4]))
            else:
                sensitivities.append(int(sample_run[8]))
                if len(simple_results) < sensitivities.index(sample_run[8]) + 1:
                    simple_results.append([])
                    simple_results[sensitivities.index(sample_run[8])].append(float(sample_run[4]))
                else:
                    simple_results[sensitivities.index(sample_run[8])].append(float(sample_run[4]))

    # Missing Samples
    flagged_files = []
    intended_sample_count = Counter(sample_count_per_file).most_common(1)
    intended_sample_count = intended_sample_count
    x = 0
    for count in sample_count_per_file:
        x += 1
        if int(count) != int(intended_sample_count[0][0]):
            flagged_files.append(x)

    for files in flagged_files:
        print("Index of flagged file: " + str(files))
        print("PLEASE NOTE: " + file_names[int(files) - 1] + " has abnormal test count")



    # Draw Graph
    fig = plt.figure()
    fig.suptitle('Sensitivity Determination', fontsize=14, fontweight='bold')
    ax = fig.add_subplot(111)
    ax.boxplot(simple_results)
    ax.set_xlabel('Sensitivity (%)')
    ax.set_ylabel('Voltage Readings (VDC)')
    plt.xticks(list(range(1, len(sensitivities) + 1)), sensitivities)
    plt.show()

def runProgram():
    main = tk.Tk()
    main.title("Sensitivity Data Importer")
    main.geometry("300x150+300+300")

    instruction_label = tk.Message(main, text="Make sure all the sensitivity data is saved in a single " +
                                              "folder. Hit the Import Data button below. Open the folder and double click on any test run. Your excel file will be saved inside that folder.",
                                   width=250, anchor="center", justify="center")
    instruction_label.pack()

    go_button = ttk.Button(main, text="Import Data", command=ExecuteSensitivityAnalysis, width=25)
    go_button.pack()

    main.mainloop()


runProgram()
