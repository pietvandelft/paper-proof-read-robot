# coding=utf-8

import os
from paper import Paper
import docx
import csv
import time
from pathlib import Path

inputPath = Path("static/papers")
CSVPath = Path("static/output")


def main():
    """ Goes through every file in the input dir and tries to score the paper.
        Scores are written to a CSV-file: output_%timestamp%.csv
    """
    
    CSVFilename = "scores_" + time.strftime("%Y%m%d%H%M%S") + ".csv"  
    
    with open(CSVPath / CSVFilename, 'w', newline='') as CSVFile:
        # write delimiter (separator) to first line of file, because otherwise # Excel defaults to ';' on some locales.
        CSVFile.write("sep=,\n")

        writer = csv.DictWriter(CSVFile, Paper.getCSVHeaders())
        writer.writeheader()
        
        try:
            fileNames = os.listdir( inputPath )
        except:
            print("Could not open directory: " + str(inputPath))
            return
        
        for filename in fileNames:
            print("scoring file: " + filename)
            
            singleCSVRow = {}

            # We identify the paper by filename
            singleCSVRow[Paper.FILENAME_HEADER] = filename
            
            # add the scores
            singleCSVRow.update(getPaperScores(filename))

            writer.writerow(singleCSVRow)

    
def getPaperScores(filename):
    filePath = inputPath / filename
    try:
        paper = Paper(filePath) # attempt to parse the docx-file
    except:
        print("There was an error opening file: " + filename)
        return
    
    return paper.getScores() # get a dict of the scores of this paper. 
 
if __name__ == '__main__':
    main()

