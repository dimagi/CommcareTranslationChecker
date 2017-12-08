import sys
import re
import argparse
import openpyxl

def parseArguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", help="Location of Translation file to check", type=str)
    return parser.parse_args()

def convertCellToOutputValueList(cell):
    '''
    Convert an Excel cell to a list of <output value...> tags contained within the cell

    Input:
    cell (openpyxl.cell.cell.Cell): Cell whose contents are to be parsed

    Output:
    List of unicode objects, each representing an instance of <output value...> in cell
    '''
    #### First pass: find an instance of "<output value" and pull whole string until next ">"
    #### Second pass: for each "<" after output value, ignore the next ">"
    openTag = "output value=\""
    closeTag ="\"/>"
    outputList = []
    currentIndex = 0
    try:
        while cell.value[currentIndex:].find(openTag) != -1:
            currentIndex += cell.value[currentIndex:].find(openTag) + len(openTag)
            outputList.append(cell.value[currentIndex:cell.value[currentIndex:].find(closeTag) + currentIndex])
    except TypeError, e:
        return []

    return outputList

def checkRowForMismatch(row, columnDict, baseColumnName = None):
    '''
    Check all of the given columns in a row provided for any mismatch in the columns' OutputValueList 

    Input:
    row(list): list of openyxl.cell.cell.Cell objects representing a single row in an Excel sheet 
    columnDict(dict): dictionary mapping column index to column name, representing every column to be checked against the baseColumn 
    baseColumnName(str [opt]): CURRENTLY UNIMPLEMENTED. Name of the column to be considered "correct." Defaults to first column in columnDict.

    Output:
    Tuple consisting of a single-element dictionary mapping the baseColumn's index to its outputValueList, and a dictionary mapping the column indexes of mismatched cells to their OutputValueList.
    '''
    mismatchDict = {}
    baseColumnDict=  {}

    realOutputValueList = None

    for colIdx in columnDict.keys():
        try:
            curOutputValueList = convertCellToOutputValueList(row[colIdx])
            curCellValue = row[colIdx].value
            if realOutputValueList is None:
                realOutputValueList = curOutputValueList
                realCellValue = curCellValue
                baseColumnDict = {colIdx : realOutputValueList}
            elif realOutputValueList != curOutputValueList:
                mismatchDict[colIdx] = curOutputValueList
                # print "Warning at row %s of sheet %s: output value mismatch" % (idx,ws.title)
                # print "%s: %s \n%s: %s" % (curOutputValueList, curCellValue.encode('utf8'), realOutputValueList, realCellValue.encode('utf8'))
        except AttributeError, e:
            pass

    return (baseColumnDict, mismatchDict)


def main(argv):
    args = parseArguments()
    try:
        wb = openpyxl.load_workbook(args.file)
        print "Loaded"
    except openpyxl.exceptions.InvalidFileException, e:
        print "Invalid File!"
        exit(-1)

    ## Iterate through WorkSheets
    testCount = 0
    for ws in wb:

        ## Dictionary mapping column index to column name
        defaultColumnDict = {}

        ## Find all columns of format "default_[CODE]"
        for idx, cell in enumerate(ws.rows[0]):
            if cell.value[:8] == "default_":
                defaultColumnDict[idx] = cell.value

        for idx, row in enumerate(ws.rows[1:]):
            rowCheckResults = checkRowForMismatch(row, defaultColumnDict)
            if len(rowCheckResults[1]) > 0:
                baseColumnName = defaultColumnDict[rowCheckResults[0].keys()[0]]
                baseColumnOutputValueList = rowCheckResults[0][rowCheckResults[0].keys()[0]]
                mismatchColumnNames = ",".join(defaultColumnDict[i] for i in rowCheckResults[1].keys())
                print "WARNING %s row %s: the output values in %s do not match %s" % (ws.title,idx, mismatchColumnNames, baseColumnName)
                # print "%s: %s \n%s: %s" % (curOutputValueList, curCellValue.encode('utf8'), realOutputValueList, realCellValue.encode('utf8'))


if __name__ == "__main__":
    main(sys.argv[1:])