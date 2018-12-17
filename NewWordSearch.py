# -*- encoding: utf-8 -*-

import urllib2
from bs4 import BeautifulSoup
import sys
import os
import pandas as pd

fileAddress = 'C:/Users/USER/Documents/All Work/New Word Search'
fileName = ''
def search(searchWord):
    foundBoolean = False
    resultToReturn = None
    ###Word Prep####
    wordIsAcronym = searchWord.isupper()
    keyWord = searchWord
    if not wordIsAcronym:
        keyWord = keyWord.lower()
    ##Getting ready for webscraping
    wordToSearchURL = searchWord.replace(" ", "%20")
    quote_page = "".join(['https://endic.naver.com/search.nhn?sLn=en&searchOption=entry_idiom&query=', wordToSearchURL])
    try:
        print quote_page
    except:
        print "Couldn't print quote page."
    searchPageCount = 1
    current_page_num = 1
    ##Start Webscraping
    while(current_page_num <= searchPageCount):
        if current_page_num > 1:
            quote_page = "".join([quote_page, '&theme=&pageNo=', str(current_page_num)])
        try:
            page = urllib2.urlopen(quote_page)
            page_decoded = page.read().decode('utf-8', 'ignore')
            soup = BeautifulSoup(page_decoded, 'html.parser')
            listToLookAt = soup.find_all("dl", class_ = "list_e2 mar_left")
            if len(listToLookAt) == 0:#No Search Results
                return None, None
            if current_page_num == 1:
                pagesDiv = soup.find_all("div", class_ = "sp_paging")
                pagesString = pagesDiv[0].stripped_strings
                searchPageCount = 0
                for item in pagesString:
                    searchPageCount += 1
                if searchPageCount > 5:
                    searchPageCount = 5
            dTerms = listToLookAt[0].find_all("span", class_="fnt_e30")
            for index, item in enumerate(dTerms):
                linkForEntry = item.a['href']
                #get current term of search result
                termStringBuilder = []
                for contentItem in item.a.contents:
                    if contentItem.name != 'sup' and contentItem.string is not None:
                        termStringBuilder.append(contentItem.string)
                entireTerm = "".join(termStringBuilder)
                lEntireTerm = entireTerm.lower()
                if lEntireTerm == keyWord:
                    resultToReturn = findMatchingDef(linkForEntry)
                    if resultToReturn is not None and len(resultToReturn.replace(" ","")) > 0:
                        return True, resultToReturn
        except:
            break
        current_page_num += 1
    return foundBoolean, resultToReturn
def findMatchingDef(link):
    #returns an int that signifies how the def matches
    #               -1: ERROR
    #               0: does not match in whatever way, just returning first def
    #               1: if it contains all parts of def separately
    #               2: if it contains def
    #               3: if it is exact match to def
    #returns the actual definition found
    defToReturn = None
    try:
        quote_page = "".join(['https://endic.naver.com', link])
        page = urllib2.urlopen(quote_page)
        page_decoded = page.read().decode('utf-8', 'ignore')
        soup = BeautifulSoup(page_decoded, 'html.parser')
    except:
        print("Error with entry page at: " + quote_page)
        return None
    listsToLookAt = soup.find_all("dl", class_ = "list_a3")
    ##prepping def
    for list in listsToLookAt:
        for currDef in list.find_all("span", class_="fnt_k06"):
            defStringBuilder = []
            for contentItem in currDef.contents:
                if contentItem.string is not None:
                    defStringBuilder.append(contentItem.string)
                else:
                    defStringBuilder.append("")
            finalCurrDef = "".join(defStringBuilder)
            noSFinalCurrDef = finalCurrDef.replace(" ","")
            if len(noSFinalCurrDef) > 0:
                return finalCurrDef
    return defToReturn

def inspectFile():
    ##open fileName
    cwd = os.getcwd()
    os.chdir(fileAddress)
    wordsFile = pd.ExcelFile(fileName)
    # Specify a writer
    writer = pd.ExcelWriter("edited_" + fileName, engine='xlsxwriter')
    for sName in wordsFile.sheet_names:
        dfs = wordsFile.parse(sheet_name=sName, usecols="A:E", encoding='utf-8')#pd.read_excel(fileName, sheet_name=sheetNum, usecols="A:C", encoding='utf-8')
        print sName
        dfs.insert(3, "Found", "")
        dfs.insert(4, "First Result", "")
        wordCol = dfs.columns[0]
        foundCol = dfs.columns[3]
        resultCol = dfs.columns[4]
        for index, row in dfs.iterrows():
            foundBoolean, result = search(row[wordCol])
            if foundBoolean is None:
                dfs.loc[index, foundCol] = "No search results"
            else:
                dfs.loc[index, foundCol] = foundBoolean
                if foundBoolean:
                    dfs.loc[index, resultCol] = result
        # Write your DataFrame to a file
        dfs.to_excel(writer, sName, index=False)
        # Save the result
    writer.save()
    ##make data construct

if __name__ == '__main__':
    sheetNumber = 1
    if len(sys.argv) < 2:
        print("Make sure to input name of file to inspect, number of sheets to inspect (first whatever), and if necessary, its location (default is C:/Users/USER/Documents/All Work/New Word Search)")
        sys.exit(0)
    fileName = sys.argv[1]
    if len(sys.argv) > 2:
        fileAddress = sys.argv[2]
    inspectFile()
