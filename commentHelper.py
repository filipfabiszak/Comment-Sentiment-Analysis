from bs4 import BeautifulSoup
import openpyxl
import urllib.request
import re


# helper function to count words in a line
def countWords(line):
    '''for word in wordList:
        if word.isspace() == False:
            words += 1'''
    punct = re.compile(r'([^A-Za-z0-9 ])')
    wordList = punct.sub("", line).split()

    return len(wordList)

# helper function to counter character in a line
def countCharacters(line):
    return len(line) - line.count(' ')


# Helper function to find article code in a link
def findCode(link):
    code = ""
    regex = re.compile('[-+]?\d+[\.]?\d*')

    matchCode = re.search("\d{6,12}", link)
    return (matchCode.group())


# Helper function to remove unneeded html tags
def remove(soup, tagname):
    for tag in soup.findAll(tagname):
        contents = tag.contents
        parent = tag.parent
        tag.extract()
        for tag in contents:
            parent.append(tag)


# helper function to get article, filtering out most unneeded garbage
def getArticle(link):
    try:
        webLink = urllib.request.urlopen(link)
    except:
        print("Error, cannot open URL")

    soup = BeautifulSoup(webLink.read(), "html.parser")
    for tag in soup.find_all('small'):
        tag.replaceWith('')
    for tag in soup.find_all('script'):
        tag.replaceWith('')

    for tag in soup.find_all('aside'):
        tag.replaceWith('')
    """for tag in soup.find_all('em'):
        tag.replaceWith('')"""
    article = soup.findAll("div", class_="post-content")
    fullComments = ""
    for comment in article:
        text = comment.getText()
        fullComments += " " + text.strip()
    print(fullComments)
    return fullComments


# Helper function to get list of code from text file (GAWKER)
def getCode():
    links = []
    articleCodes = []
    with open("linksg.txt", "r") as text_file:
        for line in text_file:
            links.append(line.strip())

    for link in links:
        articleCodes.append(findCode(link))
    return articleCodes


# Helper function to get list of code from text file (JEZEBEL)
def getCode2():
    links = []
    articleCodes = []
    with open("linksj.txt", "r") as text_file:
        for line in text_file:
            links.append(line.strip())

    for link in links:
        articleCodes.append(findCode(link))
    return articleCodes


# Helper function to get headline
def findHeadline(soup):
    headlineHTML = soup.find("h1", {"class": "headline"})
    headline = headlineHTML.getText()
    return headline


# Helper function to get reply count
def findReplies(soup):
    replies = soup.find("span",{"id": "js_reply-count"})
    r = int(replies.getText())
    return r
