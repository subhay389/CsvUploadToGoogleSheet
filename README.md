This script uploads any csv data into Google Sheets

# PRE REQUISITES
You need to change some variables. 

1. credFileName: Change to your Google service account json file 
* You can get from here: https://developers.google.com/workspace/guides/create-credentials#service-account
2. outputFilePath: (OPTIONAL) Any data that is uploaded to Google sheets will also be written here 
3. inputFilePath: CSV file you are going to upload.
4. numOfRowToSkip: If you want to skip any rows from the CSV file 
5. requiredColumns: The columns from csv file you want to extract
