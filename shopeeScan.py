import tkinter as tk
import os
from tkinter.constants import LEFT
from typing import List
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.keys import Keys
import xlsxwriter
import sys

if getattr(sys, 'frozen', False): 
    chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
    br = webdriver.Chrome(chromedriver_path)
else:
    br = webdriver.Chrome()
br.set_window_position(0, 0)
br.set_window_size(1020, 1180)

br.get('https://vietdoo.ml/')
ratings = []
solds = []
count = 1
links = []

def getLinks(keyword, page):
    
    for i in range(page):
        link = "https://shopee.vn/search?keyword=" + keyword.replace(" ", "%20") + "&page=" + str(i)
        
        br.get(link)
        sleep(2)
        br.execute_script("window.scrollTo(100, 1000)") 
        sleep(0.5)
        br.execute_script("window.scrollTo(100, 2000)") 
        sleep(0.5)
        br.execute_script("window.scrollTo(100, 3000)")

        items = []
        while len(items) != 50:
            items = br.find_elements_by_xpath("//div[@class='row shopee-search-item-result__items']/div/a")

        for item in items :
            links.append(item.get_attribute("href"))
        #print("Page", count ,"Scan - Done")
    return links

def run (links):
   
    for link in links:
        br.get(link)
        rating = ''
        while (br.find_elements_by_xpath("//div[@class='container']/div[2]/div[3]/div[1]/div[2]") == []):
            sleep(0.05)
        a = br.find_elements_by_xpath("//div[@class='container']/div[2]/div[3]/div[1]/div[2]/div")
        if len(a) == 3:
            rating = br.find_element_by_xpath("//div[@class='container']/div[2]/div[3]/div[1]/div[2]/div[2]/div[1]").get_attribute("innerText")
            sold = br.find_element_by_xpath("//div[@class='container']/div[2]/div[3]/div[1]/div[2]/div[3]/div[1]").get_attribute("innerText")
            ks = '000'
            kr = '000'
            if sold.find(',') > -1:
                ks = '00'
            if sold.find(',') > -1:
                kr = '00'
            sold = int(sold.replace(",", "").replace("k", ks))
            rating = int(rating.replace(",", "").replace("k", kr))

            solds.append(sold)
            ratings.append(rating)
            good = round(rating/sold * 100)
            buy = ''
            if good >= 40 and sold > 100 :
                buy = 'good'
            if good < 40 :
                buy = 'normal'
            if good < 10 :
                buy = 'bad'

           # print("Rating:", rating, "sold:", sold, "---", good, '%', buy)
    else :
        solds.append(0)
        ratings.append(0)


def excelExport (f) :
    
    f.write('A' + '1', 'Sold')
    f.write('B' + '1', 'Rating')
    f.write('C' + '1', 'Link')

    for i in range(len(ratings)):
        f.write('A' + str(i + 2), solds[i])
        f.write('B' + str(i + 2), ratings[i])
        f.write('C' + str(i + 2), links[i])

links = []

def show_entry_fields():
    keyword = str(e1.get())
    pages = int(e2.get())
    links = getLinks(keyword, pages)
    file_exl = xlsxwriter.Workbook(str(len(links)) + '_' + keyword + ".xlsx")
    f = file_exl.add_worksheet(keyword)
    run(links)
    excelExport (f)
    file_exl.close()
   # print(e1.get() + e2.get())

master = tk.Tk()
tk.Label(master,  
         text="vietdoo").grid(row=3,sticky= "W", column=0)
tk.Label(master,  
         text="SHOPEE SCAN ").grid(sticky= "W", row=0, column=1)
tk.Label(master,
         text="Product name").grid(row=1,sticky= "W", column=0)
tk.Label(master,
         text="Pages").grid(row=2,sticky= "W", column=0)
e1 = tk.Entry(master)
e2 = tk.Entry(master)
e2.insert(2, "4")
e1.grid(row=1, column=1)
e2.grid(row=2, column=1)

tk.Button(master, 
          text='     Scan Now     ', command=show_entry_fields).grid(row=3, 
                                                       column=1 , 
                                                       sticky= "W", 
                                                       pady=1)
tk.mainloop()


br.close()
br.quit()