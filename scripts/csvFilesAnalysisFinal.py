import fnmatch
import os
import sys
import csv
import codecs
from collections import defaultdict

folderPath = "large_image"

pausedStatusFiles = os.listdir(folderPath + "/paused")
finishedStatusFiles = os.listdir(folderPath + "/finished")
queuedStatusFiles = os.listdir(folderPath + "/queued")
uploadingStatusFiles = os.listdir(folderPath + "/uploading")
failedStatusFiles = os.listdir(folderPath + "/failed")

pausedStatusFiles = sorted(pausedStatusFiles)
finishedStatusFiles = sorted(finishedStatusFiles)
queuedStatusFiles = sorted(queuedStatusFiles)
uploadingStatusFiles = sorted(uploadingStatusFiles)
failedStatusFiles = sorted(failedStatusFiles)


def checkerAndWriter(statusFolder, statusFiles):
    finalStringToPrint = ""
    for g in range(0, len(statusFiles)):
        csv_file = folderPath + statusFolder + statusFiles[g]
        columns = defaultdict(list)  # each value in each column is appended to a list

        with open(csv_file) as f:
            reader = csv.DictReader(f)  # read rows into a dictionary format
            for row in reader:  # read a row as {column1: value1, column2: value2,...}
                for (k, v) in row.items():  # go over each column name and value
                    columns[k].append(v)  # append the value into the appropriate list
                    # based on column name k
        with open(csv_file, "r") as f:
            reader = csv.reader(f, delimiter=",")
            columName = next(reader)

        emptyColums = []

        finalStringToPrint += "\n" + statusFiles[g] + " : "

        for i in range(1, len(columName)):
            if columns[columName[i]] != '':
                z = ((columns[columName[i]].count('') / len(columns[columName[i]])) * 100) if len(columns[columName[i]]) != 0 else 0
                finalStringToPrint += columName[i] + "-" + str(z) + ", "

        finalStringToPrint += str(len(columns[columName[i]]))
    return finalStringToPrint


# print("the empty columns are: ",emptyColums)
# print("We have ", len(emptyColums), " empty colum(s) out of ", len(columName) -1, " so the percentage is ", format(100 * (len(emptyColums))/float(len(columName) -1), '.2f'), "%")

def printToFile():
    resutlFile = open(folderPath + "_" + "result.txt", "w")
    resutlFile.write(folderPath + "\n")
    resutlFile.write("Paused Status")
    resutlFile.write(checkerAndWriter("/paused/", pausedStatusFiles) + "\n")
    resutlFile.write("$$$\n")
    resutlFile.write("Finished Status")
    resutlFile.write(checkerAndWriter("/finished/", finishedStatusFiles) + "\n")
    resutlFile.write("$$$ \n")
    resutlFile.write("Queued Status")
    resutlFile.write(checkerAndWriter("/queued/", queuedStatusFiles) + "\n")
    resutlFile.write("$$$ \n")
    resutlFile.write("Uploading Status")
    resutlFile.write(checkerAndWriter("/uploading/", uploadingStatusFiles) + "\n")
    resutlFile.write("$$$\n")
    resutlFile.write("Failed Status")
    resutlFile.write(checkerAndWriter("/failed/", failedStatusFiles) + "\n")
    resutlFile.close()


# output to csv file
printToFile()

# z = ((columns[columName[i]].count('') / len(columns[columName[i]])) * 100 ) if len(columns[columName[i]]) != 0 else 0
# str(columns[columName[i]].count('') / len(columns[columName[i]]) * 100 )
