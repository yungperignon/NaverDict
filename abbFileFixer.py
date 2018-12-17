# -*- encoding: utf-8 -*-

import sys
import os
import pandas as pd

fileAddress = 'C:/Users/USER/Documents/All Work/Rando'
fileName = 'input.xlsx'
# id_col = 0
entryORG_col = 0
entry_col = 1
defORG_col = 2
def_col = 3
catID_col = 4
catName_col = 5
pCatID_col = 6
pCatName_col = 7
rank_col = 8
voteNumber_col = 9
ave_col = 10
isBot_col = 11
cat1_col = 12
def mergerows(start, dfs):
    oneRowList = [(dfs.index.values[start - 1], dfs.iat[start - 1, ave_col], dfs.iat[start - 1, cat1_col])]
    firstDef = dfs.iat[start - 1, def_col]
    try:
        firstDef = str(firstDef)
    except:
        firstDef = dfs.iat[start - 1, def_col]
    firstDef = firstDef.replace("-", "").replace(" ", "").replace(",","").lower()
    resumeIndex = start
    ##get id and cat1_col average
    while resumeIndex < len(dfs.index):
        currentDef = dfs.iat[resumeIndex, def_col]
        try:
            currentDef = str(currentDef)
        except:
            currentDef = dfs.iat[resumeIndex, def_col]
        if not pd.isnull(dfs.iat[resumeIndex, def_col]) and currentDef.replace("-", "").replace(" ", "").replace(",","").lower() != firstDef:
            break
        oneRowList.append((dfs.index.values[resumeIndex], dfs.iat[resumeIndex, ave_col], dfs.iat[resumeIndex, cat1_col]))
        resumeIndex += 1
    sorted_OneRowList = sorted(oneRowList, key=lambda tup: tup[1])
    while (len(dfs.columns) - 13 + 1) < len(oneRowList):
        dfs.insert(cat1_col + len(dfs.columns) - 13 + 1, unicode("카테고리", "utf-8") + str(len(dfs.columns) - 13 + 2), "")
    dropList = []
    for index, rowTuple in enumerate(sorted_OneRowList):
        dfs.iat[start - 1, cat1_col + len(oneRowList) - (index + 1)] = str(rowTuple[0]) + "|" + rowTuple[2]
        if index < len(oneRowList) - 1:
            dropList.append(rowTuple[0])
    dfs.drop(dropList, inplace=True)
def inspect():
    cwd = os.getcwd()
    os.chdir(fileAddress)
    dataSheet = unicode("정리","utf-8")
    dfs = pd.read_excel(fileName, sheet_name=dataSheet, index_col = 0, usecols="A:N", encoding='utf-8')
    rowIndex = 0
    while rowIndex < len(dfs.index):
        print dfs.index.values[rowIndex]
        if not pd.isnull(dfs.iat[rowIndex, entry_col]):
            rowIndex += 1
        else:
            currDef = dfs.iat[rowIndex, def_col]
            try:
                currDef = str(currDef)
            except:
                currDef = dfs.iat[rowIndex, def_col]
            if not pd.isnull(dfs.iat[rowIndex, def_col]):
                if rowIndex > 0:
                    oldDef = dfs.iat[rowIndex - 1, def_col]
                    try:
                        oldDef = str(oldDef)
                    except:
                        oldDef = dfs.iat[rowIndex - 1, def_col]
                    if oldDef.replace(" ", "").replace(",", "").replace("-", "").lower() == currDef.replace(" ", "").replace(",", "").replace("-","").lower():
                        mergerows(rowIndex, dfs)
                    else:
                        rowIndex += 1
                else:
                    rowIndex += 1
            else:
                mergerows(rowIndex, dfs)
            #means beginning started
    writer = pd.ExcelWriter("edited_" + fileName, engine='xlsxwriter')
    # Write your DataFrame to a file
    dfs.to_excel(writer, 'Sheet1', index=True)
    # Save the result
    writer.save()
if __name__ == '__main__':
    inspect()
