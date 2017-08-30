from commentHelper import getArticle
from commentHelper import getCode
import openpyxl

'''Script to scrape Gawker/Jezebel articles completely, uncomment items to switch from Gawker/Jezebel'''

articleCodes = getCode()
#articleCodes = getCode2() #uncomment this for Jezebel article scrape instead

#index for mass scraping
startIndex = 1
endIndex = 500

#change this if required scrape to start at specific row
excelRow = 1

#change worksheets/workbook here
wb = openpyxl.load_workbook('GawkerArticleScrape.xlsx')
sheet = wb.get_sheet_by_name("Gawker")

article = input("Press enter to mass scrape or enter link for single scrape: ")
if article == "":
    for i in range(startIndex, endIndex):
    #Index to keep track of the comments (used to change link and get new comments)

        currentCode = articleCodes[i]
        webURL = "http://gawker.com/{}".format(currentCode)
        #webURL = "http://jezebel.com/{}".format(currentCode) #uncomment for jezebel link
        currentArticle = getArticle(webURL)
        sheet.cell(row = excelRow, column = 1).value = currentArticle
        excelRow+= 1
        wb.save("GawkerArticleScrape.xlsx")
#remember to change save workbook to match the load workbook
wb.save("GawkerArticleScrape.xlsx")
