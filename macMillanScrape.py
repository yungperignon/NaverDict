# -*- encoding: utf-8 -*-
import urllib2
from bs4 import BeautifulSoup
import sys
import os
import xlsxwriter

fileAddress = 'C:/Users/USER/Documents/All Work/New Word Search'
fileName = ''
def search(searchWord):
    foundBoolean = False
    resultToReturn = True
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
                return False, None
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

def scrapeMcMillan(startPage, endPage):
    ##open fileName
    cwd = os.getcwd()
    os.chdir(fileAddress)
    workbook = xlsxwriter.Workbook(fileName)
    worksheet = workbook.add_worksheet()

    row = 1
    word_Col = 0
    page_Col = 1
    def_Col = 2
    found_Col = 3
    result_Col = 4

    worksheet.write(0, word_Col, "Word")
    worksheet.write(0, page_Col, "Page")
    worksheet.write(0, def_Col, "Word Definition")
    worksheet.write(0, found_Col, "Found on Naver")
    worksheet.write(0, result_Col, "Result")
    currPage = startPage

    while currPage <= endPage:
        quote_page = "".join(['https://www.macmillandictionary.com/open-dictionary/index-chronological-order_page-', str(currPage), '.htm'])
        page = urllib2.urlopen(quote_page)
        page_decoded = page.read().decode('utf-8', 'ignore')
        soup = BeautifulSoup(page_decoded, 'html.parser')
        listToLookAt = soup.find(id="odatozindex")
        listEntries = listToLookAt.find_all("li")

        for index, entry in enumerate(listEntries):
            linkForEntry = entry.a['href']
            #get current term of search result
            termStringBuilder = []
            for contentItem in entry.a.contents:
                if contentItem.string is not None:
                    termStringBuilder.append(contentItem.string)
            entireTerm = "".join(termStringBuilder)
            worksheet.write(row, word_Col, entireTerm)
            worksheet.write(row, page_Col, currPage)
            mcMillanDef = mcMillanEntryDef(linkForEntry)
            if mcMillanDef is not None and len(mcMillanDef) > 0:
                worksheet.write(row, def_Col, mcMillanDef)
            foundBoolean, result = search(entireTerm)
            worksheet.write(row, found_Col, foundBoolean)
            if foundBoolean:
                worksheet.write(row, result_Col, result)
            row += 1
        currPage += 1
    workbook.close()
def mcMillanEntryDef(link):
    page = urllib2.urlopen(link)
    page_decoded = page.read().decode('utf-8', 'ignore')
    soup = BeautifulSoup(page_decoded, 'html.parser')
    definitionSpan = soup.find_all("span", class_="DEFINITION")[0]
    defStringBuilder = []
    for defPart in definitionSpan.strings:
        defStringBuilder.append(defPart)
    entireDef = "".join(defStringBuilder)
    entireDef = entireDef.strip()
    return entireDef


if __name__ == '__main__':
    sheetNumber = 1
    if len(sys.argv) < 4:
        print("Make sure to input name for new excel file, number for start page, and number for end page. Also write address if you want.")
        sys.exit(0)
    fileName = sys.argv[1]
    scrapeMcMillan(int(sys.argv[2]), int(sys.argv[3]))
    if len(sys.argv) > 4:
        fileAddress = sys.argv[4]
