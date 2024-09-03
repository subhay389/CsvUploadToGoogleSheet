import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def read_csv_from_row(inputFilePath, numOfRowToSkip):
    # Step 1: Read the Excel file, skipping the first 3 rows
    df = pd.read_csv(inputFilePath, header=numOfRowToSkip)
    return df

def filter_row(inputDf, requiredColumns):    
    validColumns = [col for col in requiredColumns if col in inputDf.columns]
    if not validColumns:
        print("No valid columns found in the DataFrame.")
        return
    
    # Filter the DataFrame to include only the valid columns
    dfFiltered = inputDf[validColumns]

    return dfFiltered   

def transform_column_type(df):
    # Convert barcode to string since barcode number is too large value. Eg 10000000000 is too large but "10000000000" is OK
    df['Barcode'] = df['Barcode'].astype(str)
    return df

def write_to_local_file(dfToWrite, outputFilePath):
    # Write the filtered DataFrame to a new sheet in the Excel file
    with pd.ExcelWriter(outputFilePath, engine='openpyxl') as writer:
        # Write the filtered DataFrame to a new sheet named 'FilteredData'
        dfToWrite.to_excel(writer, sheet_name='FilteredData', index=False)

def write_to_google_sheet(credFileName ,spreadSheetId, df, startRow, startCol, chunkSize):
    # Write the filtered DataFrame to Google Sheets
    # Authenticate and create a client
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(credFileName, scope)
    client = gspread.authorize(creds)

    # Open the spreadsheet
    sheet = client.open_by_key(spreadSheetId).sheet1

    # Prepare the data for writing
    dataToWrite = df.values.tolist()

    # Batch write data
    for i in range(0, len(dataToWrite), chunkSize):
        chunk = dataToWrite[i:i + chunkSize]
        numRows = len(chunk)
        numCols = len(df.columns)
        
        # Define the range for batch update
        cellRange = sheet.range(startRow + i, startCol, startRow + i + numRows - 1, startCol + numCols - 1)
        
        # Flatten the chunk for cell assignment
        cellValues = [item for sublist in chunk for item in sublist]
        for cell, value in zip(cellRange, cellValues):
            cell.value = value
        
        # Update cells in the sheet
        sheet.update_cells(cellRange)


# CONSTANT VARIABLES
credFileName = "../credentials/credentials.json" # DO NOT CHANGE
chunkSize = 1000 # DO NOT CHANGE
outputFilePath = "../outputFiles/pos_filtered.xlsx" # NAME OF OUTPUT FILE WHICH WILL CONTAIN FILTERED OUTPUT DATA. THIS IS OPTIONAL

# INPUT VARIABLES (NEED TO CHANGE)
inputFilePath = "../pos_input.csv" # POS FILE DOWNLOADED FROM SOURCE
numOfRowToSkip = 3 # THE UNWANTED ROWS IN POS FILE (inputFilePath)
requiredColumns = ["Barcode", "Barcode", "Item Name", "categoryID", "Category 1", "Sale Price", "Stock"] # THE ROWS NEEDED TO BE EXTRACTED FROM POS FILE IN ORDER 

# GOOGLE SHEET VARIABLES (NEED TO CHANGE).
startRow = 3 # THE ROW NUMBER IN GOOGLE SHEET THAT WE WANT TO START WRITING
startCol = 9 # THE COLUMN NUMBER IN GOOGLE SHEET THAT WE WANT TO START WRITING. Eg 9 is column I. 
# IF YOU WANT TO WRITE DATA TO GOOGLE SHEETS STARTING FROM TOP LEFT, SET startRow=0;startCol=0;
googleSheetFileId = "" # GOOGLE SHEET ID 

## MAIN FUNCTION STARTS HERE ## 
inputDf = read_csv_from_row(inputFilePath, numOfRowToSkip) # read POS file
filteredDf = filter_row(inputDf, requiredColumns) # extract only required columns 
transformedDf = transform_column_type(filteredDf) # make transformation necessary to be compatible with python
write_to_local_file(transformedDf, outputFilePath) # This is optional 
write_to_google_sheet(credFileName, googleSheetFileId, transformedDf, startRow, startCol, chunkSize) # write to Google sheets 



