# Archival Data Organization for Howard Gotlieb Archival Research Center (HGARC) 2019 #

## The Project: ##
Converting a PDF scan of a Finding Aid into an Excel Spreadsheet, with individual items sorted under their respective categories.

## Current Version In Development: Mk. 3 ##
__Previous versions can be found in our Wiki Page__

### Primary Revisions from Mk. 2 ###
* Accurate white space preservation method

### Project Plan ###
#### I. Data Selection ####
Prompts the user to select pages they want to read in.

#### II. File Conversion ####
Read in PDF to images

#### III. Enhance Image ####
Reduce the background noise in images for accurate scanning

#### III - IV. Horizontal and Vertical Scan ####
Scan the image and extract the text and create a text file, both horizontally and vertically

* Scan vertically, to determine how many lines are in each column
* Scan horizontally, to determine the actual words
* Run through each vertical String searching for its horizontal counterpart, which pinpoints how far away each horizontal string begins from the left
* Must override pre-set language settings (i.e. English reads in text from left to right)

#### V. Transpose ####
Transpose the vertically scanned result to horizontal

#### VI. Quality Check ####
Match the transposed vertical text file with horizontally scanned result to quality check

#### VII. Structure Selection & Data Categorization ####
Prompts the user for structure (formatting and organization of data)

#### VIII. Reassemble Data ####
Create a CSV or Excel file with the processed data

## Team Members:
Jennifer (Jaehei) Kim, Richard Xiao

## Supervisor:
Claudia Friedel
