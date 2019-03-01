# Archival Data Organization for Howard Gotlieb Archival Research Center (HGARC) 2019 #

## The Project: ##
Converting a PDF scan of a Finding Aid into an Excel Spreadsheet, with individual items sorted under their respective categories.
## Versions: ##
__Previous versions can be found in our Wiki Page__

### Current Version: Mk. 2 ###

#### Primary Revisions from Mk. 1 ####
* Image editing step deseriable
  * Refined preservation of indentation / space counting
* "Clustering" overhaul due to multi-leveled lines
* Consideration of potential switch / usage to Python from Java
* Improve OCR or create common error patcher

#### Project Plan ####
__I. File Conversion__
Break large PDF files of Archival Data (Inventories, Descriptions, Listings) to smaller series of images to allow the OCR string to fit

__II. Clean File__
Reduce the background noise in images for accurate scanning

__III. Scan__
Scan the image and extract the text and create a text file

__IV. Clean Text__
Split the string down to individual lines, then separate on spacing from left

__V. Extract Prompt__
Construct a depth structure based off of prompt
* Considered creating a program to tag the beginning of the type space to keep the white space

__VI. Data Categorization__
Categorize the data

__VII. Reassemble Data__
Create a CSV or Excel file with the processed data








## Team Members:
Jennifer (Jaehei) Kim, Richard Xiao

## Supervisor:
Claudia Friedel
