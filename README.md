# Exhibit A Generator

## An application to streamline the creation of Exhibit A documents for the Sustainable Development department for the City of Chicago.

### [Download Here](https://github.com/chukwumaokere/exhibitagenerator/releases/download/3.1.1/ExhibitAGenerator-Setup3.1.1.exe)  

### How to use:
1. Download and Run the setup file   
2. Press "Okay!" to start creating documents   
3. Select a spreadsheet that has your data in it   
4. Enter the name of sheet with your data (usually at the bottom of the excel sheet)   
5. Select your range of rows (inclusive. If you enter "2-5" you will get rows 2, 3, 4, 5) (NOT INCLUDING THE HEADER. Rows usually start at 2)   
6. Or select click "All Rows" if you want to select them all   
7. Press "Create Documents"   

### How to test:
1. Run the program from an install and run the shortcut or use the dev method   
2. On the UI, select "Generate Exhibit A Word Documents from Excel File"   
3. Select the included "Dummy_List.xlsx" as your spreadsheet   
4. Enter "2018 Applicants" (without quotes) as the Sheet Name   
5. Select the "Exhibit-A-Python-1.docx" template as your DOCX Template   
6. Enter any rows like "2-5" (without quotes) or select "All Rows"   
7. Press "Create Documents"   

### Stable version coming soon...

#### For dev, go into the folder exhibitagenerator/exhibitagenerator/ and in here run "npm install" and if needed "npm install electron --save" then in that directory "exhibitagenerator/exhibitagenerator/" run "npm start"
