from bs4 import BeautifulSoup
import urllib.request
import json
import openpyxl
from commentHelper import *

'''Script to retrieve comment data from Kinja articles'''

# row to start parsing at (change if needed)
excelRow = 2
wb = openpyxl.load_workbook('KinjaData.xlsx')
sheet = wb.get_sheet_by_name("Sheet1")
debugCounter = 1
approved = True

print("This program only works on Kinja websites, links intended for scraping should include a 10 digit article code.")
print("As is currently implemented, the following links in 'KinjaLinks.txt' will be ignored:")

validLinks = []
with open("KinjaLinks.txt", "r") as text_file:
    for line in text_file:
        try:
            findCode(line.strip())
            validLinks.append(line.strip())
        except:
            if line.strip() != "":
                print(line.strip())
print("")
print("There are " + str(len(validLinks)) + " valid articles: \n" + '\n'.join(validLinks))
print("")

for articleLink in validLinks:

    startIndex = 0
    numberOfComments = 0
    approvedChildComments = 0

    currentSource = findSource(articleLink)
    currentCode = findCode(articleLink)
    webURL = currentSource + currentCode
    print("")
    print("link: " + articleLink)
    print("source: " + currentSource)
    print("code: " + currentCode)

    try:
        web = urllib.request.urlopen(webURL)
    except:
        print("Error, cannot open URL: " + webURL)
        break

    soup = BeautifulSoup(web.read(), "html.parser")

    # Find the specific HTML element that holds the number of total replies
    try:
        totalNumComments = findReplies(soup)
    except:
        print("Error, cannot find headline (maybe headline does not exist?)")
        headline = "(no headline)"


    headlineRow = excelRow
    # sheet.cell(row=headlineRow, column=1).value = (headline)
    # stats to keep track of
    avgMainWord = 0
    avgMainChar = 0
    avgChildWord = 0
    avgChildChar = 0
    imageCount = 0

    # Keep looping until we get all comments (calling different JSON links)
    while startIndex < totalNumComments:

        if approved:
            jsonURL = currentSource + "api/comments/views/replies/{0}?dap=true&startIndex={1}&maxReturned" \
                  "=100&maxChildren=100&approvedOnly=true&cache=true".format(currentCode, startIndex)
        else:
            jsonURL = currentSource + "api/comments/views/replies/{0}?dap=true&startIndex={1}&maxReturned" \
                  "=100&maxChildren=100&approvedOnly=false&cache=true".format(currentCode, startIndex)

        page = urllib.request.urlopen(jsonURL).read()
        pageString = page.decode('utf-8')
        decoded = json.loads(pageString)
        dataSet = decoded["data"]["items"]

        counter = 0
        while counter < len(dataSet) and len(dataSet) != 0:

            mainComment = dataSet[counter]["reply"]["deprecatedFullPlainText"]
            try:
                imageCount += len(dataSet[counter]["reply"]["images"])
            except:
                pass

            if mainComment != "":
                mainCommentWordCount = countWords(mainComment)
                mainCommentCharacterCount = countCharacters(mainComment)
                avgMainWord += mainCommentWordCount
                avgMainChar += mainCommentCharacterCount

            numberOfComments+=1

            childSet = dataSet[counter]["children"]["items"]

            childCounter = 0
            while childCounter < len(childSet):

                childComment = childSet[childCounter]["deprecatedFullPlainText"]
                try:
                    imageCount += len(childSet[childCounter]["reply"]["images"])
                except:
                    pass

                if childComment != "":
                    childWordLen = countWords(childComment)
                    childCharLen = countCharacters(childComment)
                    avgChildWord += childWordLen
                    avgChildChar += childCharLen
                approvedChildComments+= 1
                childCounter+=1
            counter += 1
        startIndex+=100

    sheet.cell(row = excelRow, column = 1).hyperlink = webURL
    sheet.cell(row = excelRow, column = 2).value = totalNumComments
    sheet.cell(row = excelRow, column = 3).value = ((numberOfComments + approvedChildComments))
    sheet.cell(row = excelRow, column = 4).value = ((numberOfComments))
    sheet.cell(row = excelRow, column = 5).value = ((approvedChildComments))
    try:
        sheet.cell(row = excelRow, column = 6).value = ((avgMainWord/numberOfComments))
        sheet.cell(row = excelRow, column = 7).value = ((avgMainChar/numberOfComments))
    except:
        sheet.cell(row = excelRow, column = 6).value = numberOfComments
        sheet.cell(row = excelRow, column = 7).value = numberOfComments
    try:
        sheet.cell(row = excelRow, column = 8).value = ((avgChildWord/approvedChildComments))
        sheet.cell(row = excelRow, column = 9).value = ((avgChildChar/approvedChildComments))
    except:
        sheet.cell(row = excelRow, column = 8).value = ((approvedChildComments))
        sheet.cell(row = excelRow, column = 9).value = ((approvedChildComments))

    sheet.cell(row = excelRow, column = 10).value = imageCount

    # if approved:
    # else:
    #     sheet.cell(row = excelRow, column = 3).value = (numberOfComments + approvedChildComments)
    #     sheet.cell(row = excelRow, column = 6).value = ((numberOfComments))
    #     sheet.cell(row = excelRow, column = 8).value = ((approvedChildComments))
    #     sheet.cell(row = excelRow, column = 13).value = ((avgMainWord/numberOfComments))
    #     sheet.cell(row = excelRow, column = 15).value = ((avgMainChar/numberOfComments))
    #     if approvedChildComments != 0:
    #         sheet.cell(row = excelRow, column = 14).value = ((avgChildWord/approvedChildComments))
    #         sheet.cell(row = excelRow, column = 16).value = ((avgChildChar/approvedChildComments))
    #     else:
    #         sheet.cell(row = excelRow, column = 14).value = ((approvedChildComments))
    #         sheet.cell(row = excelRow, column = 16).value = ((approvedChildComments))

    print("Article {} done".format(debugCounter))
    debugCounter+=1
    excelRow += 1
    wb.save('KinjaData.xlsx')

wb.save('KinjaData.xlsx')
