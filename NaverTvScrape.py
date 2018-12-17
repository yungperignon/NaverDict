# -*- encoding: utf-8 -*-

import urllib2
from bs4 import BeautifulSoup
import sys
import os
import pandas as pd
import requests
import time
#just gets list of channel names and links for corresponding channels into a list
fileAddress = 'C:/Users/USER/Documents/All Work/NaverTVScrape'


def getVideoInfo(link, dfs, itemNum):
    try:
        page = requests.get(link)
        contents = page.content
        soup = BeautifulSoup(contents, 'lxml')
        divToLookAt = soup.find(id="clipInfoArea")
        titleDatePlayDiv = divToLookAt.find("div", class_="watch_title")
        titleDatePlayItems = titleDatePlayDiv.select("._clipTitle,.play,.date")
        dfs.loc[itemNum, 'Latest Video'] = titleDatePlayItems[0].string.strip()
        dfs.loc[itemNum, 'Latest Video Views'] = int(titleDatePlayItems[1].contents[2].replace(",", ""))
        dfs.loc[itemNum, 'Latest Video Date'] = titleDatePlayItems[2].contents[1]
        hashTagBuilder = []
        hashTags = divToLookAt.select(".hash_box a")
        for index, item in enumerate(hashTags):
            hashTagBuilder.append(item.string)
            if index < len(hashTags) - 1:
                hashTagBuilder.append(", ")
        dfs.loc[itemNum, 'Latest Video Tags'] = "".join(hashTagBuilder)
    except Exception as e:
        print(link + ": " + str(e))
def getLatestVideo(link, dfs, itemNum):
    try:
        page = requests.get(link + "/clips")
        contents = page.content
        soup = BeautifulSoup(contents, 'lxml')
        divToLookAt = soup.select_one("body.ch_home > div#wrap > div#container > div#content > div#cds_flick > div.flick-container > div.flick-panel > div.ch_content > div.ch_clip > div.cate_wrap > div.wrp_cds > div._infiniteCardArea")
        dfs.loc[itemNum, 'Latest Video URL'] = "https://tv.naver.com" + divToLookAt.div.div.a['href']
        getVideoInfo(dfs.loc[itemNum, 'Latest Video URL'], dfs, itemNum)
    except Exception as e:
        print(link + ": " + str(e))
def getChannelInfo(link, dfs, itemNum):
    try:
        page = requests.get(link)
        contents = page.content
        soup = BeautifulSoup(contents, 'lxml')
        divToLookAt = soup.find("div", class_="ch_content")
        asideDiv = divToLookAt.find("div", class_="ch_aside")
        chDateDiv = asideDiv.find("div", class_="ch_date")
        chBDay = (chDateDiv.select('span[class~=date]'))[0].string
        dfs.loc[itemNum, 'Channel Creation Date'] = chBDay
        chInfoList = asideDiv.find("ul", class_="view").find_all("li")
        dfs.loc[itemNum, 'Subscribers'] = int(chInfoList[0].span.string.replace(",", ""))
        dfs.loc[itemNum, 'Videos'] = int(chInfoList[4].span.string.replace(",", ""))
        dfs.loc[itemNum, 'Total Views'] = int(chInfoList[1].span.string.replace(",", ""))
        dfs.loc[itemNum, 'Total Likes'] = int(chInfoList[2].span.string.replace(",", ""))
        dfs.loc[itemNum, 'Total Comments'] = int(chInfoList[3].span.string.replace(",", ""))
        dfs.loc[itemNum, 'Playlists'] = int(chInfoList[5].span.string.replace(",", ""))
        print(link)
        if (dfs.loc[itemNum, 'Videos'] > 0):
            getLatestVideo(link, dfs, itemNum)
        #getVideoInfo(dfs.loc[itemNum, 'Latest Video URL'], dfs, itemNum)
    except Exception as e:
        print(link + ": " + str(e))

def searchForChannels(channelName, limit):
    ##store dataframe
    cwd = os.getcwd()
    os.chdir(fileAddress)
    dfs = pd.DataFrame(columns = ['Channel Name', 'Category', 'URL', 'Channel Creation Date','Subscribers', 'Videos', 'Total Views', 'Total Likes', 'Total Comments', 'Playlists', 'Latest Video', 'Latest Video Date', 'Latest Video Tags', 'Latest Video Views', 'Latest Video URL'])
    #wordToSearchURL = channelName.replace(" ", "%20")
    currPage = 1
    pageLim = limit
    itemNum = 0
    while currPage <= pageLim:
        if currPage == 1:
            quote_page = "".join(["https://tv.naver.com/search/channel?query=", channelName, "&isTag=false"])#"https://tv.naver.com/search/channel?query=영어&isTag=false"#.join([https://tv.naver.com/search/channel?query=영어&isTag=false'])
        else:
            quote_page = "".join(["https://tv.naver.com/search/channel?query=", channelName, "&sort=rel&page=",str(currPage),"&isTag=false"])
        try:
            page = requests.get(quote_page)
            contents = page.content
            soup = BeautifulSoup(contents, 'lxml')
            divToLookAt = soup.find("div", class_="my_ch")
            channelList = divToLookAt.find("ul", class_="ch_list")
            if currPage == 1:
                pagingDiv = divToLookAt.find("div", class_="paging")
                lastPage = pagingDiv.find("a", class_="_click next next_end")
                if lastPage is None:
                    lastPageNumber = 1
                else:
                    lastPageNumber = int(lastPage['data-page'])
                if pageLim > lastPageNumber:
                    pageLim = lastPageNumber
        except Exception as e:
            print(e)
        listElements = channelList.select('a[class~=info_a]')
        for item in listElements:
            dfs.loc[itemNum, 'Channel Name'] = item['title']
            dfs.loc[itemNum, 'URL'] = "https://tv.naver.com" + item['href']
            getChannelInfo(dfs.loc[itemNum, 'URL'], dfs, itemNum)
            itemNum += 1
        currPage += 1
    writer = pd.ExcelWriter(channelName.decode('utf-8').encode('euc-kr') + " search results.xlsx", engine='xlsxwriter')
    dfs.to_excel(writer, 'Sheet1', index=False)
    writer.save()

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Make sure to input file to write into, search and if necessary, limit of pages to search")
        sys.exit(0)
    searchWord = sys.argv[1].decode('euc-kr').encode('utf-8')
    limitPages = float('inf')
    if len(sys.argv) >= 3:
        limitPages = int(sys.argv[2])
    if len(sys.argv) >= 4:
        fileAddress = sys.argv[3]
    searchForChannels(searchWord, limitPages)
