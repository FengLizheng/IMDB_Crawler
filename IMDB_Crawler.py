# -*- coding: utf-8 -*-
"""
Created on Fri Dec  6 14:09:52 2019

@author: FENG Lizheng The universitu of Hong Kong

"""


import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import xlrd




key = ""

df={'Movie':[],'ID':[],'Author':[],'Date':[],'Rating':[],'Up Vote':[],'Total Vote':[],'Subject':[],'Review':[]}
excel_source=xlrd.open_workbook('title.xlsx')
sheet_source = excel_source.sheets()[0]
nrows = sheet_source.nrows
for i in range(0, nrows):
    
 try:
    flag=True
    MovieName=sheet_source.cell_value(i,1)
    ID=sheet_source.cell_value(i,4)
    print(i,MovieName)
    if ID!='no_result':
            base_url = "https://www.imdb.com/"
            url = 'https://www.imdb.com/title/'+ID+'/reviews?ref_=tt_ql_3'
            MAX = 5000
            label = 1
           
            res = requests.get(url)
            res.encoding = 'utf-8'
            soup = BeautifulSoup(res.text, "lxml")


            for item in soup.select(".lister-list"):
                Subject = item.select(".title")[0].text
                author = item.select(".display-name-link")[0].text
                date = item.select(".review-date")[0].text
                votetext = item.select(".text-muted")[0].text
                upvote = re.findall(r"\d+",votetext)[0]
                totalvote = re.findall(r"\d+", votetext)[1]
                rating = item.select("span.rating-other-user-rating > span")
                if len(rating) == 2:
                    rating = rating[0].text
                else:
                    rating = ""
                review = item.select(".text")[0].text
                df['Movie'].append(MovieName)
                df['Review'].append(review)  
                df['ID'].append(ID)
                df['Author'].append(author)
                df['Date'].append(date)
                df['Rating'].append(rating)
                df['Up Vote'].append(upvote)
                df['Total Vote'].append(totalvote)
                df['Subject'].append(Subject)                             
                label = label + 1
            

            load_more = soup.select(".load-more-data")
            flag = True
            if len(load_more):
                    ajaxurl = load_more[0]['data-ajaxurl']
                    base_url = base_url + ajaxurl + "?ref_=undefined&paginationKey="
                    try:
                        key = load_more[0]['data-key']
                    except:
                        flag = False
            else:
                    flag = False

            while flag:
                    url = base_url + key
                    
                    res = requests.get(url)
                    res.encoding = 'utf-8'
                    soup = BeautifulSoup(res.text, "lxml")
                    for item in soup.select(".lister-item-content"):
                        Subject = item.select(".title")[0].text
                        author = item.select(".display-name-link")[0].text
                        date = item.select(".review-date")[0].text
                        votetext = item.select(".text-muted")[0].text
                        vote = re.findall(r"\d+", votetext)[0]
                        totalvote = re.findall(r"\d+", votetext)[1]
                        rating = item.select("span.rating-other-user-rating > span")
                        if len(rating) == 2:
                            rating = rating[0].text
                        else:
                            rating = ""
                        review = item.select(".text")[0].text
                        df['Movie'].append(MovieName)
                        df['Review'].append(review) 
                        df['ID'].append(ID)
                        df['Author'].append(author)
                        df['Date'].append(date)
                        df['Rating'].append(rating)
                        df['Up Vote'].append(upvote)
                        df['Total Vote'].append(totalvote)
                        df['Subject'].append(Subject)
                                      
                        label = label + 1          
                    if label >= MAX:
                        break
                    load_more = soup.select(".load-more-data")
                    if len(load_more):
                        key = load_more[0]['data-key']
                    else:
                        flag = False

 except:
        print('have no review')                      
print('Under Saving Process') 
                
destinationFileName='ReviewIMDB.csv'                
pd.DataFrame(df).to_csv(destinationFileName, index=False)

print('Already Saved')    

