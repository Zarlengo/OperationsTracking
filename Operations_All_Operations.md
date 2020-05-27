# [All Operations.vbs](https://bitbucket.org/FlowCorp_CC/contract_cutting/src/master/All%20Operations.vbs)

---

# Settings

 * adminMode: true to show borders on the window
 * debugMode: debugMode = true to ignore error bypassing, use to check for connection issues
 * meMode: true to allow the ME option as a scanner.  Uses a TCP connection directly to the computer for testing purposes
 
# Connections

 * \[00_Script\]
 * \[00_Machine_IP\]
 
# Common Code

 * `dataSource`: location of the SQL server on the network
 * `initialCatalog`: database name within the SQL server
 * Initializes variables
 * Adds database connection constants
 * Closes any open mshta.exe processes, all other open scripts
 * Checks for any arguments being passed into the script, stores as `sArg`
 * Calls function `Load_Access` to check for database connections
 * Creates the ie window from function `HTABox`
 * Calls function `checkAccess` to update the window based on the connection results
 
# HTA Loop

Stop conditions: "cancel", "access", "done", "allOps"  
Checks if the window has been closed, ends script if it has
 
 * "cancel": Closes the window, activated by pressing the red X  
	* Closes the window
	* Ends the script
 * "access": Retry the database connection, activated by pressing the database button
	* Resets the loop variable
	* Changes the color of the database button
	* Updates the database connection text
	* Calls function `Load_Access` to check for database connections
	* Calls function `checkAccess` to update the window based on the connection results
 * "done": Loads the selected script, activated by the button/dropdowns in the window
 * "allOps": Reloads the All Operations script, activated by pressing the blue +
 
 ---
 
# Functions

## Function HTABox(`sBgColor`, `h`, `w`, `l`, `t`)

Function to create the HTA window

 * `sBgColor` background color of the window
 * `h` window height
 * `w` window width
 * `l` location of the window from the left
 * `t` location of the window from the top
 
Sets a random integer to assign to this instance  
Defines the window parameters  
Creates the new ie instance  
Loads the function `LoadHTML` containing all of the HTML information  
Adjusts the window size based on the number of scripts loaded  
Exits the function if the window if there is an error loading the instance

## Function checkAccess()

Checks if the connection was successful and modifies the window accordingly

## Function GetNewConnection()

Function to make the connection to the database  

 * Creates the connection object
 * Loads the connection options
 * If the connection is successful passes the object out of the function
 * If the connection is unsuccessful then passes false out of the function

## Function Load_Access()

Function to check for the database connection and then load initial items to the window

 * Obtains the connection to the database
 * Gets the number of production scripts
 * Gets the number of admin scripts
 * Gets all of the scripts from the database
	* Stores the script columns and loads them to the production or admin array
	* Determines if the script is going to be a button or a dropdown (if there are scanners used in the script)
		* Button: Creates button HTML
		* Dropdown:
			* Loads the machine options from the database for that scanner option
			* Creates the dropdown options for each scanner
			* If meMode is true then adds the Manufacturing Engineering computers to the list, allows using a TCPIP program to feed data into the script without using a scanner
	* Stores all of the script variables to `ArgDict`

## Function LoadOPScript(OPNum, ScriptArg)

Function to load the HTA script or the PS1 script

 * HTA executes through a different protocol than wscript.exe
 * PS1 will not run depending on security options, this will prompt the user with steps and then open the folder location

## Function ServerClose()

Function to run through several clean up commands when closing the window

 * Sets up to ignore any errors
 * Closes the ie window
 * Ends the script

## Function LoadHTML(sBgColor)

Function to create all of the HTML objects for the ie window  
  
HTA String

 * Sets up the window parameters
	* border:			Sets or retrieves the type of window border for the HTML Application (HTA).
	* contextMenu:		Sets or retrieves whether the context menu is displayed when the right mouse button is clicked.
	* maximizeButton:	Sets or retrieves a Boolean value that indicates whether a Maximize button is displayed in the title bar of the HTML Application (HTA) window.
	* minimizeButton:	Sets or retrieves a Boolean value that indicates whether a Minimize button is displayed in the title bar of the HTML Application (HTA) window.
	* sysMenu:			Sets or retrieves a Boolean value that indicates whether a system menu is displayed in the HTML Application (HTA).
 
CSS

 * Sets up all of the CSS parameters for the HTML
 * adminMode is true will show borders for the elements in the window
 
JavaScript

 * operationFunction(`OpNum`) 
	* Activated by the button press
	* Ends the HTA loop
	* Loads the operation number to a saved variable
 * argumentFunction(`OpNum`)
	* Activated by the dropdown selection
	* Ends the HTA loop 
	* Loads the operation number and scanner ID to saved variables
 
Body

 * Database connection button and string
 * Operation header
	* Cycles through `ArgDict` for each operation script and loads to the window
 * Admin header
	* Cycles through `ArgDict` for each admin script and loads to the window
	* Creates a window height variable based on the total number of scripts loaded
 * Blue + Load all operations script
 * Red X close the script window
	* `done` stored variable for the HTA loop
	* `argText` stored variable for the scanner ID
	* `OPText` stored variable for the operation number