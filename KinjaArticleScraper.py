import openpyxl
from commentHelper import *

'''Script to scrape Kinja articles completely'''

excelRow = 2
debugcounter = 0
wb = openpyxl.load_workbook('KinjaArticles.xlsx')
sheet = wb.get_sheet_by_name("Sheet1")
debugCounter = 1
approved = True

print("This program only works on Kinja websites, links intended for scraping should include a 10 digit article code.")
print("As is currently implemented, the following links in 'KinjaLinks.txt' will be ignored:")

validLinks = getLinks()

sheet.cell(row = 1, column = 1).value = "Article Link"
sheet.cell(row = 1, column = 2).value = "Num Comments"
sheet.cell(row = 1, column = 3).value = "Num Likes"
sheet.cell(row = 1, column = 4).value = "Word Count"
sheet.cell(row = 1, column = 5).value = "Character Count"
sheet.cell(row = 1, column = 6).value = "Article Text"

for articleLink in validLinks:

    currentSource = findSource(articleLink)
    currentCode = findCode(articleLink)
    webURL = currentSource + currentCode
    try:
        currentArticle = getArticle(webURL)
    except:
        continue

    articleCharCount = countCharacters(currentArticle)
    sheet.cell(row = excelRow, column = 1).hyperlink = webURL
    sheet.cell(row = excelRow, column = 2).value = findRepliesFromLink(webURL)
    sheet.cell(row = excelRow, column = 3).value = findLikes(webURL)
    sheet.cell(row = excelRow, column = 4).value = countWords(currentArticle)
    sheet.cell(row = excelRow, column = 5).value = articleCharCount
    if articleCharCount > 32767:
        sheet.cell(row = excelRow, column = 6).value = currentArticle[:32767]
        excelRow += 1
        sheet.cell(row = excelRow, column = 6).value = currentArticle[32767:]
    else:
        sheet.cell(row = excelRow, column = 6).value = currentArticle

    excelRow += 1
    print("Article {} done".format(debugCounter))
    debugCounter += 1
    wb.save("KinjaArticles.xlsx")

wb.save("KinjaArticles.xlsx")
