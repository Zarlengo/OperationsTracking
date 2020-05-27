# [CC_Load_AE.xlsm](https://bitbucket.org/FlowCorp_CC/contract_cutting/src/master/CC_Load_AE.xlsm)

---
Hidden sheet

# LoadAEData

## Global variables

 * `MacroFileName` Macro file name, to know what excel file to close at the end of the script
 * `dataSource` SQL server location
 * `cn` SQL connection variable
 * `BladeColl` Dictionary variable for storing data

 ## Load_AE_SN()
 
		Sub to develop the data to send to the database

 * Creates initial variables for the script
 * Prompts the user to provide the Serial Control Date
 * InvoiceMessage:
	* Prompts the user to provide the Invoice Number
	* Checks for the slug column
 * SlugColMessage:
	* Creates the slug format string
	* Checks the worksheet for the slug
	* If not found, prompts the user for the column number or letter
 * SlugNumeric:
	* Converts the letter to a number
	* Stores the slug column data
 * BladeMessage:
	* Creates the blade format string
	* Checks the worksheet for the blade
	* If not found, prompts the user for the column number or letter
 * BladeNumeric:
	* Converts the letter to a number
	* Stores the blade column data
 * SlugMessage:
	* Not used - Checks for LPT5 vs LPT7
	* Sends the data to the user to verify everything is correct
	* Cycles through each row in the worksheet
		* Adds all of the row data to a collection
		* Sends the collection to the LoadAEData2Access function
		* Changes row color based on duplicate blades
		* Updates the status bar with the progress
 * EndSub:
	* Resets the calculation mode
	* Clears the status bar
	* Closes the macro workbook
	
 ## LoadAEData2Access(VarIn)
 
		Function to load the data into the database.
		
 * Splits the blade collection object into each variable
 * Queries blade and slug entries in the database
	* If the blade is found, marks the result as a duplicate
	* If the slug is found, codes the blade PN as -2, else as a -1
 * Creates the SQL string to add the data to the database
 * Sends the SQL string to the database
 
 ## AccessRecordCheck(VarIn)
 
		Function check for existing entries and to assign the dash number for the blade
 
 ## CollLoop(ColLCnt)
 
		Function to check all the blades were assigned the dash number correctly
 
 ## OpenAccess(VarIn)
 
		Sub to open the SQL connection
		
 ## CloseAccess(VarIn)
 
		Sub to close the SQL connection
 
---

# Modules

## Global variables

 * `folderLocation` Macro location
 * `myDB` ACCESS database location (not used)
 * `Op10FileName` OP10 file location (not used)
 * `Op20FileName` OP20 file location (not used)
 * `MacroFileName` Marking macro file location (not used)

 ## IsWorkBookOpen(VarIn)
 
		Function to check if the workbook is already open

 ## ModifyString(VarIn)
 
		Function to create the part marker string with the blade serial numbers

 ## Excel2Access(VarIn)
 
		Sub to create the access connection, sends the VarIn string to the SQL database, and closes the connection.

 ## ConvertToLetter(iCol)
 
		Function to convert a number to the associated column
