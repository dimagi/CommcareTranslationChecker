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
        print ws.title

        ## Dictionary mapping column index to column name
        defaultColumnDict = {}

        ## Find all columns of format "default_[CODE]"
        for idx, cell in enumerate(ws.rows[0]):
            if cell.value[:8] == "default_":
                defaultColumnDict[idx] = cell.value

        for idx, row in enumerate(ws.rows[1:]):
            ## Right now, just compare counts of output_value string
            realOutputValueList = None

            for colIdx in defaultColumnDict.keys():
                testCount += 1
                try:
                    curOutputValueList = convertCellToOutputValueList(row[colIdx])
                    curCellValue = row[colIdx].value
                    # if testCount%100 == 0:
                    #     print curOutputValueList, realOutputValueList
                    #     print curCellValue.encode('utf8'), realCellValue.encode('utf8')
                    if realOutputValueList is None:
                        realOutputValueList = curOutputValueList
                        realCellValue = curCellValue
                    elif realOutputValueList != curOutputValueList:
                        print "Warning at row %s of sheet %s: output value mismatch" % (idx,ws.title)
                        print "%s: %s \n%s: %s" % (curOutputValueList, curCellValue.encode('utf8'), realOutputValueList, realCellValue.encode('utf8'))
                except AttributeError, e:
                    pass


if __name__ == "__main__":
    main(sys.argv[1:])