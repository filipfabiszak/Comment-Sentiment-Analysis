# Comment-Sentiment-Analysis
Instructions:
1. Download this repository
2. Unzip/extract the folder
3. Open Pycharm
4. Open Project
   * Navigate to "Comment-Sentiment-Analysis-master" and open it as a project
5. Ensure that a python interpreter is selected by navigating to File -> Settings -> Project -> Project Interpreter
   * if not, select python3.x (as long as it's 3.something it should be fine)
6. On the same "Project Interpreter" view, there should be a list of "Packages" available to the project
7. To run these scripts, we must install 2 packages using the "+" on the right hand side
   * bs4
   * openpyxl
8. Search for each of these packages, and click "Install Package"
9. Upon installing these pacakges successfully, there shouldn't be any errors/underlines at the top of the file when opening any of the "Scraper.py" files saying that the packages are undefined
10. All of the scripts should now be functional
11. To begin scraping, paste any links your webcrawler generated into the "KinjaLinks.txt" file, ensuring each link is on its own line. Invalid links are allowed, as they will be ignored when scraping.
12. Right click one of KinjaArticleScraper.py, KinjaCommentScraper.py, or KinjaDataScraper.py and "Run" it.
   * Note that: each of these scripts overwrite their corresponding Kinja.xlsx Excel files. (Any previous data will be lost)
13. Once the scrips prompts that it is finished, open the associated Excel file to observe your data.
