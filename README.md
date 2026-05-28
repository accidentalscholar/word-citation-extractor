# Word Citation Extractor
This simple script attempts to extract well-structured citations from Word DOCX files in a folder and provides the output as tables in an Excel files. It also attempts to do basic sanity tests on the extracted citations.

## The Problem
Verifying individual citations / references in a document manually is quite onerous. While it may not be a [wicked problem](https://en.wikipedia.org/wiki/Wicked_problem), automating this in a robust way is quite challenging. One issue is that there are numerous referencing styles in use, and any _generalised_ tool that attempts to boil the ocean without targeted knowledge of which referencing style has been used in the presented documents, is likely to suffer on accuracy. Tools targeting specific referencing styles have the potential to be more accurate.

A bigger challenge, however, is that real-life documents are "messy" - there are unforeseen aspects of these documents, and the referencing styles are not always strictly or consistently applied. Hence, the problem of caliberation - matching too tighly would miss some of the citations in some documents, while relaxing the patterns too much would make the tool idenitfy as citations some text that is clearly not.

This script uses algorithmic pattern matching (commonly known as regular expressions or _regex_) to identify citations in the provided documents. It has not achieved the _goldilocks_ level, and I welcome feedback on what it gets wrong, especially if I can identify patterns, which will allow me to improve its logic.

Primarily intended for _author-date_ type referencing, this script is also intended to be servicable for footnote / endnote type of referencing.

## Usage

### Preparation
Before using this script, ensure all your Word DOCX files are stored in a single folder.

### Running the script
When you run the Python script, it will pop up a file explorer/finder window asking you to select the folder in which your Word DOCX files are stored.

Once you have done that, it will extract the citations from all Word DOCX files in the chosen folder and save the output to an Excel file in the same folder.

### Output
The script will save the output as an Excel file in the same folder as your Word DOCX files.

The Excel workbook will have a sheet per Word DOCX file and a summary sheet providing an overview for the folder.

Each row in the summary table represents a Word DOCX file, and the columns are as following:
1. Filename
2. Inferred referencing style
3. Total unique references
4. % In-text matched to ref	- this indicates the percent of the in-text citations that have been matched to full bibliographical references.
5. % Ref matched to in-text	- this indicates the percent of the full bibliographical references that have been matched to in-text citations.
6. % Real sources (API)	- this is the percent of the sources determined to be real.
7. % Refs with URL - this is the percent of the sources for which a link has been provided.
8. % Direct Links - Of the sources for which a link has been provided, this is the percent of the links that are deep links rather than just a base/ home page.
9. % HTTP 200 - Of the sources for which a link has been provided, this is the percent of the links that don't return a system error.
10. % 200 and NOT 404 - Of the sources for which a link has been provided, this is the percent of the links that don't return a soft 404 error, e.g., a "404 page not found" page. 
11. Alphabetical References? - Many referencing styles require the bibliographic references to be listed in alphabetical order. This column captures whether the bibliographical references in this document are in such an order.


Each row in the respective document tables represents an identified citation, and the columns are as following:
1. Filename
2. In text citation
3. Estimated page in-text - this provides the page number on which the in-text citation was found.
4. Matched full reference
5. Estimated page reference - this provides the page number on which the full bibliographical reference was found.
6. Inferred style
7. Source type - the script attempts to categorise the source. This is done using very simple heuristics.
8. Source URL - if the full reference includes a link, this is captured.
9. Date accessed - if the full reference includes a link, and states the access date, this is captured.
10. Direct link? - The script attempts to determine whether the link provided is a deep link, as opposed to just a base/ home page.
11. 200 OK? - The script attempts to determine whether the link provided actually takes to a working page.
12. 404 Page? - If the link provided takes to a technically working page, the script attempts to determine whether it is a "404 Page Not Found" page.
13. Real Source (API)? - The script attempts to determine whether the reference is real. For this, it checks the source against crossref.org


These individual citations have some formatting:
1. If there is not a direct link to the source, then the background is light orange.
2. If the link provided is not working, then the background is light orange.
3. If the reference is assessed to be not real, then the background is light orange.

Each of the sheets for individual files also include a chart for the determined types of the references found.

### Note
The script uses some standard Python libraries. If you don't have them installed on your system, then in the first run, the script will try to install these dependencies. 

## Caveat
Tested on Windows 11 Education 64-bit

Not tested on Apple iOS or Linux.

## Never run a Python script before?
It's straightforward, but you may need to install Python on your machine first.

### Install Python
Anaconda is one of the most popular distributions of Python. Download and install from https://www.anaconda.com/download

Installation is straightforward, but if you need help, check out https://www.anaconda.com/docs/getting-started/anaconda/install/overview

### Start Spyder
Anaconda comes with Spyder IDE. Start Spyder.

Once Spyder is ready, open the file 'word-citation-extractor.py' that has the script.

All that's left is for you to hit 'Run', i.e. the green 'Play' button.