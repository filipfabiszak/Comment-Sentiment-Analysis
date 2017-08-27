from bs4 import BeautifulSoup
import urllib.request
import json
import openpyxl

from commentHelper import countWords
from commentHelper import countCharacters
from commentHelper import getCode2
from commentHelper import findCode
from commentHelper import findReplies

'''Program to scrape comment data from Kinja articles'''

# row to start parsing at (change if needed)
excelRow = 2
# sheets and workbook to use, should be included in same directory
wb = openpyxl.load_workbook('data.xlsx')
sheet = wb.get_sheet_by_name("Sheet1")


# Change index here to look for specific articles
articleStartIndex = 0
articleEndIndex = 1

userInput = input("Press Enter for mass scraping OR paste link/article code for specific link scraping: ")
if userInput == "":
    getSpecific = False

    articleStartIndex = int(input("Please choose the start index for the articles you want to scrap: ")) - 1
    articleEndIndex = articleStartIndex + int(input("Please choose how many articles you would like to scrap: "))

else:
    getSpecific = True

approved = False
currentIndex = articleStartIndex + 1



# Comment holds comments with HTML styling, plain holds text only, childPlain for child comments
plain = []
childPlain = []
debugCounter = 1
articleCodes = getCode2() # this is for jezebel


for i in range(articleStartIndex, articleEndIndex):
# Index to keep track of the comments (used to change link and get new comments)
    startIndex = 0
    numberOfComments = 0
    approvedChildComments = 0

    currentCode = articleCodes[i]
    print(currentCode)
    webURL = "http://jezebel.com/{}".format(currentCode)
    if getSpecific:
        webURL = userInput
        currentCode = findCode(webURL)


    # Open request to webpage
    try:
        web = urllib.request.urlopen(webURL)
    except:
        print("Error, cannot open URL")
        break

    soup = BeautifulSoup(web.read(), "html.parser")


    # Find the specific HTML element that holds the number of total replies
    try:
        r = findReplies(soup)
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

    dataSetIsEmpty = False;

    # Keep looping until we get all comments (calling different JSON links)
    while dataSetIsEmpty != True:



        # Link can be changed to included non approved comments as well
        if approved:
            jsonURL = "http://jezebel.com/api/comments/views/replies/{0}?dap=true&startIndex={1}&maxReturned" \
                  "=100&maxChildren=100&approvedOnly=true&cache=true".format(articleCodes[i], startIndex)

        else:
            jsonURL = "http://jezebel.com/api/comments/views/replies/{0}?dap=true&startIndex={1}&maxReturned" \
                  "=100&maxChildren=100&approvedOnly=true&cache=true".format(articleCodes[i], startIndex)


        page = urllib.request.urlopen(jsonURL).read()
        pageString = page.decode('utf-8')

        # Turns JSON file into dictionary
        decoded = json.loads(pageString)
        dataSet = decoded["data"]["items"]

        if len(dataSet) == 0:
            dataSetIsEmpty = True

        counter = 0
        while counter < len(dataSet) and len(dataSet) != 0:
            print(counter)
            # Going through the content and taking what we need
            htmlLines = BeautifulSoup(dataSet[counter]["reply"]["display"], "html.parser")
            mainComment = htmlLines.findAll('p')
            fullComments = ""
            for comment in mainComment:
                text = comment.getText()
                fullComments += " " + text

            # making sure the comment is not empty
            if mainComment != "":
                mainWordLen = countWords(fullComments)
                mainCharLen = countCharacters(fullComments)
                avgMainWord += mainWordLen
                avgMainChar += mainCharLen

            numberOfComments+=1

            childSet = dataSet[counter]["children"]["items"]

            childCounter = 0
            while childCounter < len(childSet):
                htmlLine = BeautifulSoup(childSet[childCounter]["display"], "html.parser")
                childComment = htmlLine.findAll('p')
                fullComment = ""
                for comment in childComment:
                    text = comment.getText()
                    fullComment += " " + text

                childCounter+=1
                if fullComment != "":
                    childWordLen = countWords(fullComment)
                    childCharLen = countCharacters(fullComment)
                    avgChildWord += childWordLen
                    avgChildChar += childCharLen
                approvedChildComments+= 1
            counter += 1
        startIndex+=100

    # Output to file after collection
    if approved:
        sheet.cell(row = excelRow, column = 3).value = (numberOfComments + approvedChildComments)
        sheet.cell(row = excelRow, column = 6).value = ((numberOfComments))
        sheet.cell(row = excelRow, column = 8).value = ((approvedChildComments))
        sheet.cell(row = excelRow, column = 13).value = ((avgMainWord/numberOfComments))
        sheet.cell(row = excelRow, column = 15).value = ((avgMainChar/numberOfComments))
        if approvedChildComments != 0:
            sheet.cell(row = excelRow, column = 14).value = ((avgChildWord/approvedChildComments))
            sheet.cell(row = excelRow, column = 16).value = ((avgChildChar/approvedChildComments))
        else:
            sheet.cell(row = excelRow, column = 14).value = ((approvedChildComments))
            sheet.cell(row = excelRow, column = 16).value = ((approvedChildComments))
    else:
        sheet.cell(row = excelRow, column = 1).value = currentIndex
        sheet.cell(row = excelRow, column = 2).value = r
        sheet.cell(row = excelRow, column = 3).value = ((numberOfComments + approvedChildComments))
        sheet.cell(row = excelRow, column = 4).value = ((numberOfComments))
        sheet.cell(row = excelRow, column = 5).value = ((approvedChildComments))
        sheet.cell(row = excelRow, column = 6).value = ((avgMainWord/numberOfComments))
        sheet.cell(row = excelRow, column = 7).value = ((avgMainChar/numberOfComments))
        if approvedChildComments != 0:
            sheet.cell(row = excelRow, column = 8).value = ((avgChildWord/approvedChildComments))
            sheet.cell(row = excelRow, column = 9).value = ((avgChildChar/approvedChildComments))
        else:
            sheet.cell(row = excelRow, column = 8).value = ((approvedChildComments))
            sheet.cell(row = excelRow, column = 9).value = ((approvedChildComments))

    currentIndex += 1
    childPlain.clear()

    print("{} Article done".format(debugCounter))
    debugCounter+=1
    excelRow += 1
    wb.save('data.xlsx')

wb.save('data.xlsx')
