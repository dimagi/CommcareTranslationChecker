import sys
import re
import argparse
import openpyxl

def parseArguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", help="Location of Translation file to check", type=str)
    parser.add_argument("--columnList", help="[Opt] Comma-separated list of column names to check. By default, all columns that start with 'default_' will be checked.", type=str, default=None)
    parser.add_argument("--baseColumn", help="[Opt] Name of column that others are to be compared against. Warnings are flagged for all columns that do not match the baseColumn. Defaults to leftmost column in columnList.", type=str, default=None)
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

def checkRowForMismatch(row, columnDict, baseColumnIdx = None):
    '''
    Check all of the given columns in a row provided for any mismatch in the columns' OutputValueList 

    Input:
    row(list): list of openyxl.cell.cell.Cell objects representing a single row in an Excel sheet 
    columnDict(dict): dictionary mapping column index to column name, representing every column to be checked against the baseColumn 
    baseColumnIdx(int [opt]): Index of the column to be considered "correct." Defaults to lowest-indexed column in columnDict.

    Output:
    Tuple consisting of a single-element dictionary mapping the baseColumn's index to its outputValueList, and a dictionary mapping the column indexes of mismatched cells to their OutputValueList.
    '''
    mismatchDict = {}
    baseColumnDict=  {}

    baseOutputValueList = None

    ## Build baseColumnDict
    if baseColumnIdx is None:
        baseColumnIdx = sorted(columnDict.keys())[0]
    baseOutputValueList = convertCellToOutputValueList(row[baseColumnIdx])
    baseColumnDict = {baseColumnIdx : baseOutputValueList}

    for colIdx in columnDict.keys():
        try:
            curOutputValueList = convertCellToOutputValueList(row[colIdx])
            if colIdx != baseColumnIdx and baseOutputValueList != curOutputValueList:
                mismatchDict[colIdx] = curOutputValueList
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
            if args.columnList:
                if cell.value in args.columnList:
                    defaultColumnDict[idx] = cell.value
            elif cell.value[:8] == "default_":
                defaultColumnDict[idx] = cell.value

        for idx, row in enumerate(ws.rows[1:]):
            baseColumnIdx = None
            if args.baseColumn:
                for colIdx in defaultColumnDict.keys():
                    if defaultColumnDict[colIdx] == args.baseColumn:
                        baseColumnIdx = colIdx 

            rowCheckResults = checkRowForMismatch(row, defaultColumnDict, baseColumnIdx)
            if len(rowCheckResults[1]) > 0:
                baseColumnName = defaultColumnDict[rowCheckResults[0].keys()[0]]
                baseColumnOutputValueList = rowCheckResults[0][rowCheckResults[0].keys()[0]]
                mismatchColumnNames = ",".join(defaultColumnDict[i] for i in rowCheckResults[1].keys())
                print "WARNING %s row %s: the output values in %s do not match %s" % (ws.title,idx, mismatchColumnNames, baseColumnName)
                # print "%s: %s \n%s: %s" % (curOutputValueList, curCellValue.encode('utf8'), baseOutputValueList, baseCellValue.encode('utf8'))


if __name__ == "__main__":
    main(sys.argv[1:])