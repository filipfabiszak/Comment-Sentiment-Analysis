from bs4 import BeautifulSoup
import openpyxl
import urllib.request
import re

# helper function to count words in a line
def countWords(line):
    punct = re.compile(r'([^A-Za-z0-9 ])')
    wordList = punct.sub("", line).split()
    return len(wordList)

# helper function to counter character in a line
def countCharacters(line):
    return len(line)

# Helper function to find article source from a link
def findSource(link):
    matchCode = re.search(".*[/]", link)
    return (matchCode.group())

# Helper function to find article code from a link
def findCode(link):
    matchCode = re.search("\d{10}$", link.strip().strip("/"))
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
        print("Error, cannot open URL: " + webURL)

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
    # print(fullComments)
    return fullComments

def getLinks():
    validLinks = []
    with open("KinjaLinks.txt", "r") as text_file:
        for line in text_file:
            try:
                findCode(line)
                validLinks.append(line.strip())
            except:
                if line.strip() != "":
                    print(line.strip())
    print("")
    print("There are " + str(len(validLinks)) + " valid articles")
    print("")
    return validLinks


# Helper function to get headline
def findHeadline(soup):
    headlineHTML = soup.find("h1", class_="headline")
    headline = headlineHTML.getText()
    return headline


# Helper function to get reply count (defined at top of article)
def findReplies(soup):
    replies = soup.find("section", class_="js_discussion-region")
    return int(replies['data-reply-count-total'])

def findLikes(webURL):
    web = urllib.request.urlopen(webURL)
    soup = BeautifulSoup(web.read(), "html.parser")
    likes = soup.find("a", class_="js_like")
    liketext=likes.find("span", class_="text")
    try:
        r = int(liketext.getText())
        return r
    except:
        return 0


def findRepliesFromLink(webURL):
    web = urllib.request.urlopen(webURL)
    soup = BeautifulSoup(web.read(), "html.parser")
    replies = soup.find("section", class_="js_discussion-region")
    return int(replies['data-reply-count-total'])
