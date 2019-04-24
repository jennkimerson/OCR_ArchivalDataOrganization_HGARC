# Archival Data Organization for Howard Gotlieb Archival Research Center (HGARC) 2019 #

## The Project: ##
Converting a PDF scan of a Finding Aid into an Excel Spreadsheet, with individual items sorted under their respective categories.

## In Development: Mk. 3 ##
* Scan vertically, to determine how many lines are in each column
* Scan horizontally, to determine the actual words
* Run through each vertical String searching for its horizontal counterpart, which pinpoints how far away each horizontal string begins from the left
* Must override pre-set language settings (i.e. English reads in text from left to right)

## Current Version: Mk. 2 ##
__Previous versions can be found in our Wiki Page__

### Primary Revisions from Mk. 1 ###
* Image editing step deseriable
  * Refined preservation of indentation / space counting
* "Clustering" overhaul due to multi-leveled lines
* Consideration of potential switch / usage to Python from Java
* Improve OCR or create common error patcher

### Project Plan ###
#### I. File Conversion ####
Break large PDF files of Archival Data (Inventories, Descriptions, Listings) to smaller series of images to allow the OCR string to fit

#### II. Clean File ####
Reduce the background noise in images for accurate scanning

#### III. Scan ####
Scan the image and extract the text and create a text file

#### IV. Clean Text ####
Split the string down to individual lines, then separate on spacing from left

#### V. Extract Prompt ####
Construct a depth structure based off of prompt
* Considered creating a program to tag the beginning of the type space to keep the white space

#### VI. Data Categorization ####
Categorize the data

#### VII. Reassemble Data ####
Create a CSV or Excel file with the processed data








## Team Members:
Jennifer (Jaehei) Kim, Richard Xiao

## Supervisor:
Claudia Friedel
