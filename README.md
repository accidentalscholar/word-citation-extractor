# Word Citation Extractor
This simple script attempts to extract well-structured citations from _Word DOCX_ files in a folder and provides the output as tables in an _Excel_ file. It also attempts to do basic sanity tests on the extracted citations.

## The Problem
Verifying individual citations / references in a document manually is quite onerous. While it may not be a [wicked problem](https://en.wikipedia.org/wiki/Wicked_problem), automating this in a robust way is quite challenging. One issue is that there are numerous referencing styles in use, and any _generalised_ tool that attempts to boil the ocean without targeted knowledge of which referencing style has been used in the presented documents, is likely to suffer on accuracy. Tools targeting specific referencing styles have the potential to be more accurate.

A bigger challenge, however, is that real-life documents are "messy" - there are unforeseen aspects of these documents, and the referencing styles are not always strictly or consistently applied. Hence, the problem of caliberation - matching too tighly would miss some of the citations in some documents, while relaxing the patterns too much would make the tool idenitfy as citations some text that is clearly not.

This script uses algorithmic pattern matching (commonly known as regular expressions or _regex_) to identify citations in the provided documents. It has not achieved the _goldilocks_ level, and I welcome feedback on what it gets wrong, especially if I can identify patterns, which will allow me to improve its logic.

Primarily intended for _author-date_ type referencing, this script is also intended to be servicable for footnote type of referencing.

## Usage

### Preparation
Before using this script, ensure all your _Word DOCX_ files are stored in a single folder.

### Running the script
When you run the _Python_ script, it will pop up a file explorer/finder window asking you to select the folder in which your _Word DOCX_ files are stored.

Once you have done that, it will extract the citations from all _Word DOCX_ files in the chosen folder and save the output to an _Excel_ file in the same folder.

### Output
The script will save the output as an _Excel_ file in the same folder as your _Word DOCX_ files.

The Excel workbook will have a sheet per _Word DOCX_ file and a summary sheet providing an overview for the folder.

Each row in the summary table represents a _Word DOCX_ file, and the columns are as following:
1. _Filename_
2. _Inferred referencing style_
3. _Total unique references_
4. _% In-text matched to ref_ - this indicates the percent of the in-text citations that have been matched to full bibliographical references.
5. _% Ref matched to in-text_ - this indicates the percent of the full bibliographical references that have been matched to in-text citations.
6. _% Real sources (API)_ - this is the percent of the sources determined to be real.
7. _% Refs with URL_ - this is the percent of the sources for which a link has been provided.
8. _% Direct Links_ - Of the sources for which a link has been provided, this is the percent of the links that are deep links rather than just a base/ home page.
9. _% HTTP 200_ - Of the sources for which a link has been provided, this is the percent of the links that don't return a system error.
10. _% 200 and NOT 404_ - Of the sources for which a link has been provided, this is the percent of the links that don't return a soft 404 error, e.g., a "404 page not found" page. 
11. _Alphabetical References?_ - Many referencing styles require the bibliographic references to be listed in alphabetical order. This column captures whether the bibliographical references in this document are in such an order.


Each row in the respective document tables represents an identified citation, and the columns are as following:
1. _Filename_
2. _In text citation_
3. _Estimated page in-text_ - this provides the page number on which the in-text citation was found.
4. _Matched full reference_
5. _Estimated page reference_ - this provides the page number on which the full bibliographical reference was found.
6. _Inferred style_
7. _Source type_ - the script attempts to categorise the source. This is done using very simple heuristics.
8. _Source URL_ - if the full reference includes a link, this is captured.
9. _Date accessed_ - if the full reference includes a link, and states the access date, this is captured.
10. _Direct link?_ - The script attempts to determine whether the link provided is a deep link, as opposed to just a base/ home page.
11. _200 OK?_ - The script attempts to determine whether the link provided actually takes to a working page.
12. _404 Page?_ - If the link provided takes to a technically working page, the script attempts to determine whether it is a "404 Page Not Found" page.
13. _Real Source (API)?_ - The script attempts to determine whether the reference is real. For this, it checks the source against crossref.org


These individual citations have some formatting:
1. If there is not a direct link to the source, then the background is _light orange_.
2. If the link provided is not working, then the background is _light orange_.
3. If the reference is assessed to be not real, then the background is _light orange_.

Each of the sheets for individual files also include a chart for the determined types of the references found.

### Note
The script uses some standard _Python_ libraries. If you don't have them installed on your system, then in the first run, the script will try to install these dependencies. 

## Caveat
Tested on _Windows 11 Education 64-bit_.

Not tested on _Apple iOS_ or _Linux_.

## Never run a Python script before?
It's straightforward, but you may need to install _Python_ on your machine first.

### Install Python
_Anaconda_ is one of the most popular distributions of _Python_. Download and install from https://www.anaconda.com/download

Installation is simple, but if you need help, check out https://www.anaconda.com/docs/getting-started/anaconda/install/overview

### Start Spyder
_Anaconda_ comes with _Spyder IDE_. Start _Spyder_.

Once _Spyder_ is ready, open the file 'word-citation-extractor.py' that has the script.

All that's left is for you to hit 'Run', i.e. the green 'Play' button.