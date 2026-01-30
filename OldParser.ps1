# Script for getting information from an LSM Log.

# Adjust security setting to allow windows to run Scripts.
#	Launch Powershell as an Administrator
#	type:  Set-ExecutionPolicy RemoteSigned
#	Select "Yes To All"
# .-)

#Clear the screen of junk, prior to starting and have user close Excel.
	Clear-Host

# Present Program Header, Company ownership and author
	Write-Host ""
	Write-Host "      ----------------------------------------------------------------" -ForegroundColor Green
	Write-Host -NoNewLine "     |" -ForegroundColor Green
	Write-Host -NoNewLine "     Samsung Electronics America, Customer Service Division     " -ForegroundColor White
	Write-Host "|" -ForegroundColor Green
	Write-Host -NoNewLine "     |"  -ForegroundColor Green
	Write-Host -NoNewLine "              Business to Business, Visual Display              " -ForegroundColor White
	Write-Host "|" -ForegroundColor Green
	Write-Host "     |                                                                |"-ForegroundColor Green
	Write-Host -NoNewLine "     |"  -ForegroundColor Green
	Write-Host -NoNewLine "                LED Signage Manager Log Analyzer                " -ForegroundColor Cyan
	Write-Host "|" -ForegroundColor Green
	Write-Host -NoNewLine "     |"  -ForegroundColor Green
	Write-Host -NoNewLine "                         Version 1.6.3                          " -ForegroundColor Cyan
	Write-Host "|" -ForegroundColor Green
	Write-Host "     |                                                                |" -ForegroundColor Green
	Write-Host -NoNewLine "     |"  -ForegroundColor Green
	Write-Host -NoNewLine "                          Written by                            " -ForegroundColor DarkGray
	Write-Host "|" -ForegroundColor Green
	Write-Host -NoNewLine "     |"  -ForegroundColor Green
	Write-Host -NoNewLine "                         Chuck Tinker                           " -ForegroundColor DarkGray
	Write-Host "|" -ForegroundColor Green
#	Write-Host -NoNewLine "     |"  -ForegroundColor Green
#	Write-Host -NoNewLine "                     Field Service Engineer                     " -ForegroundColor DarkGray
#	Write-Host "|" -ForegroundColor Green
	Write-Host -NoNewLine "     |"  -ForegroundColor Green
	Write-Host -NoNewLine "                  ChuckTinker@SEA.Samsung.com                   " -ForegroundColor DarkGray
	Write-Host "|" -ForegroundColor Green
	Write-Host "     |                                                                |" -ForegroundColor Green
	Write-Host "      ----------------------------------------------------------------" -ForegroundColor Green
	sleep -s 1
	Write-Host ""

# Present Warnings and such
	Write-Host "      ----------------------------------------------------------------" -ForegroundColor White
	Write-Host -NoNewLine "     |"  -ForegroundColor White
	Write-Host -NoNewLine " The LSM Log Analyzer is compatible with SBB-SNOW S-Box's Only. " -ForegroundColor Green
	Write-Host "|" -ForegroundColor White
	Write-Host -NoNewLine "     |"  -ForegroundColor White
	Write-Host -NoNewLine " Results are inconsistent with The Wall Luxury (SBB-MBox).      " -ForegroundColor Yellow
	Write-Host "|" -ForegroundColor White
	Write-Host "     |                                                                |" -ForegroundColor White
	Write-Host -NoNewLine "     |"  -ForegroundColor White
	Write-Host -NoNewLine " Ensure the LSMLog filename or folder does not contain a ( . ). " -ForegroundColor Yellow
	Write-Host "|" -ForegroundColor White
	Write-Host -NoNewLine "     |"  -ForegroundColor White
	Write-Host -NoNewLine " The file will not be saved to the folder you selected.         " -ForegroundColor Yellow
	Write-Host "|" -ForegroundColor White
	Write-Host "     |                                                                |" -ForegroundColor White
	Write-Host -NoNewLine "     |"  -ForegroundColor White
	Write-Host -NoNewLine " SAVE AND CLOSE ALL EXCEL WORKSHEETS BEFORE CONTINUING          " -ForegroundColor Red
	Write-Host "|" -ForegroundColor White
	Write-Host "     |                                                                |" -ForegroundColor White
	Write-Host -NoNewLine "     |"  -ForegroundColor White
	Write-Host -NoNewLine " Press Enter When Ready to Continue.                            " -ForegroundColor Cyan
	Write-Host "|" -ForegroundColor White
	Write-Host "      ----------------------------------------------------------------" -ForegroundColor White
	Read-Host 

#Set Some Variables, Define Others
if($Halt -ne "Halt"){ #Skip this Section Logic, Set by changing $Halt from "" to "Halt"
	$LLAver	= "LLAv163"	# Variable used in createing the new filename output of the Script File
	$Halt = ""			# Variable used to exit the program immeiately, AND to group code together
	$xlCellTypeLastCell = 11	# Defines a Cell Locator Variable I know nothing about, But need
	$FileBrowser = ""	# Temporary Variable used while setting up File Open Dialog
	$LSMLogFile = ""	# The Filename (And Directory) of the LSM log we're analyzing
	$Excel = ""			# Indicates the Excel Program itself
	$ExcelDoc = ""		# The Document opened in Excel
	$Sheet = ""			# The Sheet in the Document opened in Excel
	$CellFound = ""		# Holds XLS Cell Location of data I'm looking for
	$CellBegin = ""		# Used to remember where I started, so I don't read it again when I loop in the xls
	$CellCurrent = ""	# Current XLS Cell I'm working with
	$CellRow = ""		# Current XLS Cell Row number I'm working on
	$LSMLog = ""		# Array used to store Device Responses from the Log File
	$SBBList = ""		# Array of SBox IP's and ID's used to search for device specific Data
	$CabList = ""		# Array of Cabinet IP's and ID's used to search for device specific Data
	$DevLog = ""		# LSM Log filtered to one Device by IP and ID
	$DevIP = ""			# IP Address of the current Device I'm Gathering Data for
	$DevID = ""			# Device ID used by MDC to Identify the Device
	$DevIDx = ""		# Device ID used by MDC in Hex (From the MDC Response)
	$DevLogEntry = ""	# One Log Entry with IP, ID and MDC Response (pulled from any group of Logs)
	$MDCResp = ""		# MDC Response from the Device, in Hex String with Headers and CRC
	$DataStart = ""		# Array Entry number, of the first byte of Data in the MDCResp
	$DataLen = ""		# Number of Bytes contained in the Data Block
	$DataEnd = ""		# Array Entry number, of the Last byte of Data in the MDCResp
	$DataX = ""			# Array of MDC Data pulled from MDC Response in HEX
	$DataA = ""			# Array of MDC data pulled from MDC Response in ASCII
	$DataD = ""			# Array of MDC data pulled from MDC Response in Decimal
	$enc = ""			# Variable used to hold Data Encoding parameters
	$line = ""			# Holds each Line of an Array's Data, while it is processed in a ForEach loop
	$byteX = ""			# Stores a full HEX Byte, i.e. 0x2E.  Used in ASCII conversion loop
	$Count = ""			# Counter used in Loops based on fixed numbers, Reset Frequently
	$FieldCount = ""	# Number of Fields I need to pull from one MDC Response
	$FieldCountLoc = ""	# Location containing the number of Data Fields getting pulled back
	$FunIn = ""			# Used to hand a value to a Function
	$FunOut = ""		# Output Value of a Function
	$SBBEntry = 0		# Set Back Box Counter used in Loops
	$Row1 = 1			# Default Excel Spreadsheet Cell Row for Sheet 1 (Layout)
	$Row2 = 1			# Default Excel Spreadsheet Cell Row for Sheet 2 (SBoxes)
	$Row3 = 1			# Default Excel Spreadsheet Cell Row for Sheet 3 (Cabinets)
	$Column1 = 1		# Default Excel Spreadsheet Cell Column for Sheet 1 (Layout)
	$Column2 = 1		# Default Excel Spreadsheet Cell Column for Sheet 2 (SBoxes)
	$Column3 = 1		# Default Excel Spreadsheet Cell Column for Sheet 3 (Cabinets)
	$SBBType = ""		# Stores the type of SBB, which determines Settings types and availability
	$CabType = ""		# Stores the type of LED Cabinet, which determines Settings Types and Availability
	$1stVerCabMain = "None"	# Used to store the First LED Cabinet Main FW version
	$1stVerCabFPGA = "None" # Used to store the First LED Cabinet FPGA version
}#End of Skip Section Logic

# Setup Functions / Subroutines for repeated use
if($Halt -ne "Halt"){ #Skip this Section Logic, Set by changing $Halt from "" to "Halt"
# Build a Function to extract MDC Data from MDCResponse
	Function Get-MDCData{
	# Set the incoming Varible to $Function INput
		Param($FunIn)
	# Find Full String containing Search Text
		$DevLogEntry = $DevLog -match "$FunIn" | Select-Object -last 1
	# Determine if the Device Log Entry is populated
		if($DevLogEntry.count -ne 0){  #If Data is present do...
		# Grab only the MDC REsponse, remove IP and Device from LogLine
			$MDCResp = $DevLogEntry.split(",")[2]
		# Create an Array from the String
			$MDCResp = $MDCResp.split(" ")
		# Find Data Length from MDC Response (4th charater in MDC Response)
			$DataLen = $MDCResp[3]
		# Convert Hex to Decimal for math purposes
			$DataLen = [System.Convert]::ToInt16($DataLen,16)
		# Set the End point of the MDC Response
			$DataEnd = $DataLen+3
		# Capture the Data as it is (Hex), based on the Data Length
			$Data = $MDCResp[6..$DataEnd]
			} #End of If Device Log Entry is not empty
		else{ #If Device Log Entry is Blank, return "No Data Found"
			$Data = "NULL"
		}
	# Return the End Result back to the code that asked for it.
		Return $Data
	} # End of Get MDC Data Function

# Build a Function to Return MDC Data from a REsponse, converted to ASCII
	Function Get-MDCASCII{
	# Set the incoming Varible to $Function INput
		Param($FunIn)
	# Find Full String containing Search Text
		$DevLogEntry = $DevLog -match "$FunIn" | Select-Object -last 1
	# Determine if the Log Entry is Empty
		if($DevLogEntry.count -ne 0){
		# Grab only the MDC REsponse, remove IP and Device from LogLine
			$MDCResp = $DevLogEntry.split(",")[2]
		# Create an Array from the String
			$MDCResp = $MDCResp.split(" ")
		# Find Data Length from MDC Response (5th charater in MDC Response)
			$DataLen = $MDCResp[3]
		# Convert Hex to Decimal for math purposes
			$DataLen = [System.Convert]::ToInt16($DataLen,16)
		# Set the End point of the MDC Response
			$DataEnd = $DataLen+3
		# Capture the Data as it is (Hex), based on the Data Length
			$Data = $MDCResp[6..$DataEnd]
		# Convert the Response Hex Array ASCII Text
			$enc = [System.Text.Encoding]::ASCII #Set Encoding to ASCII
		# Setup an Array for ASCII version of the MDCResp Data
			$FunOut = New-Object System.Collections.ArrayList
		# Process each Array Entry and convert it from HEX to ASCII
			ForEach($line in $Data){ #For every line in DataX Array... do...
			# Determine if the value being process is 00
				if($line -eq "00"){
				# If the value is 00, do nothing
				}
			# Else, If the line is not NULL (00), convert it
				else{ 
				# Create a full Hex Byte from the 2 digit Byte i.e. 0xE4
					$byteX = "0x$line"
				# Convert to Ascii and add to DataA Array
					$FunOut.add($enc.GetString($byteX)) >null.txt
				} #End of If it's not NULL (00)
			} # End of For Each instance.
		# Convert the result into a Text String
			$FunOut = "$FunOut"
		# Remove all Spaces from the Text String
			$FunOut = $FunOut | foreach {$_ -replace " ",""}
		} #End of If Device Log Entry is not empty
	# If the Device Log Entry is Blank
		else{ 
		# Set the value to "unknown"
			$FunOut = "NULL"
		} # End of If the DevLogEntry is Blank
	# Return the End Result back to the code that asked for it.
		Return $FunOut
	} # End of Function MDC-ASCII

# Build a Function to convert Hex to ASCII
	Function Conv-x2a{
		Param($FunIn)
		if($FunIn -ne "NULL"){
			$enc = [System.Text.Encoding]::ASCII #Set Encoding to ASCII
		#Setup an Array for ASCII version of the MDCResp Data
			$FunOut = New-Object System.Collections.ArrayList
		#Process each Array Entry and convert it from HEX to ASCII
			ForEach($line in $FunIn){ #For every line in DataX Array... do...
				if($line -eq "00"){ # If a Line contains a NULL (00), Skip it
				}
				else{ #If the line is not NULL (00), convert it
			# Create a full Hex Byte from the 2 digit Byte i.e. 0xE4
				$byteX = "0x$line"
			# Convert to Ascii and add to DataA Array
				$FunOut.add($enc.GetString($byteX)) >null.txt
				} #End of If it's not NULL (00)
			} # End of For Each instance.
		#Convert the Array of ASCII's into a String
			$FunOut = "$FunOut"
		# Mark Double spaces as an _
			$FunOut = $FunOut | foreach {$_ -replace "  ","!"}
		#Remove the Single Spaces from the String
			$FunOut = $FunOut | foreach {$_ -replace " ",""}
		#Change _ back to a Space
			$FunOut = $FunOut | foreach {$_ -replace "!"," "}
		} # End of If FunIn does NOT equal NULL
		else{
			$FunOut = $FunIn
		}
		Return $FunOut
	}

#Build a Function to convert Hex to Decimal
	Function Conv-x2d{
		Param($FunIn)
		if($FunIn -ne "NULL"){
		#Setup an Array for Conversion of the MDCResp Data
			$FunOut = New-Object System.Collections.ArrayList
		#Process each Array Entry and convert it from HEX to ASCII
			ForEach($line in $FunIn){ #For every line in DataX Array... do...
			# Create a full Hex Byte from the 2 digit Byte i.e. 0xE4
				$byteD = [System.Convert]::ToInt16($line,16)
			# Convert to Ascii and add to DataA Array
				$FunOut.add($byteD) >null.txt
			} # End of For Each instance.
		} # End of If FunIn does NOT equal NULL
		else{
			$FunOut = $FunIn
		}
		Return $FunOut
	} # End of Hex to Dec Function

# Build a Function to convert number Cell positions into Alpha cell positions
	Function ExcelA1{ 
		Param([parameter(Mandatory=$true)] 
		[int]$number) 
		$a1Value = $null 
		While ($number -gt 0) { 
			$multiplier = [int][system.math]::Floor(($number / 26)) 
			$charNumber = $number - ($multiplier * 26) 
			If ($charNumber -eq 0) { $multiplier-- ; $charNumber = 26 } #End of If
			$a1Value = [char]($charNumber + 64) + $a1Value 
			$number = $multiplier 
		} # End of While
		Return $a1Value 
	}

}#End of Skip Section Logic

#Build some Refernece Tables and Arrays to store data
if($Halt -ne "Halt"){ #Skip this Section Logic, Set by changing $Halt from "" to "Halt"
# SBB Model Type (MBOX or SBOX Only)
	$tSBBType = New-Object System.Collections.ArrayList
	$tSBBType.add("TB-MSM,SBB-MBOX") >null.txt
	$tSBBType.add("TB-HMS,SBB-SNOW") >null.txt
	$tSBBType.add("TB-KTM,SBB-SNOW") >null.txt
	$tSBBType.add("TB-GSS,SBB-ISO8E") >null.txt
	$tSBBType.add("NULL,No Data Found") >null.txt

# LED Device Type:  0xD0_81
	$tDevType = New-Object System.Collections.ArrayList
	$tDevType.add("00,Reserved") >null.txt
	$tDevType.add("01,SBox") >null.txt
	$tDevType.add("02,Cabinet-IS/IFH/IFH-D") >null.txt
	$tDevType.add("03,Cabinet-IFJ/2in1") >null.txt
	$tDevType.add("NULL,No Data Found") >null.txt

# SBB Video Wall Mode:  0x84
	$tSBBVWMode = New-Object System.Collections.ArrayList
	$tSBBVWMode.add("00,Off") >null.txt
	$tSBBVWMode.add("01,On") >null.txt
	$tSBBVWMode.add("NULL,No Data Found") >null.txt

# Seam Correction Enabled/Disable (D0 98)
	$tSeamC = New-Object System.Collections.ArrayList
	$tSeamC.add("00,Off") >null.txt
	$tSeamC.add("01,On") >null.txt
	$tSeamC.add("NULL,No Data Found") >null.txt

# Cabinet RGB Color Correction Enabled/Disable (IWR and IWJ only)
	$tCabCC = New-Object System.Collections.ArrayList
	$tCabCC.add("00,Off") >null.txt
	$tCabCC.add("01,On") >null.txt
	$tCabCC.add("NULL,No Data Found") >null.txt

# Module RGB Color Correction Enabled/Disable (D0 99)
	$tModCC = New-Object System.Collections.ArrayList
	$tModCC.add("00,Off") >null.txt
	$tModCC.add("01,On") >null.txt
	$tModCC.add("NULL,No Data Found") >null.txt

# Pixel RGB Color Correction Enabled/Disable (D0 95)
	$tPixCC = New-Object System.Collections.ArrayList
	$tPixCC.add("00,Off") >null.txt
	$tPixCC.add("01,On") >null.txt
	$tPixCC.add("NULL,No Data Found") >null.txt

# Create the LSM Log Array to hold all Device Responses, and Add a Header
	$LSMLog = New-Object System.Collections.ArrayList
	#$LSMLog.add("IPAddress, ID, Response") >null.txt

# Create the SBox Device List Array and add a Header
	$SBBList = New-Object System.Collections.ArrayList
	#$SBBList.add("IPAddress,ID") >null.txt

# Create the Cabinet List Array (IP Address, ID)
	$CabList = New-Object System.Collections.ArrayList
	#$CabList.add("IP Address, Cabinet ID") >null.txt

# Create an SBB-GroupIP Table (SBB IP, Group IP)
	$GroupIPTable = New-Object System.Collections.ArrayList
	#$GroupIPTable.add{"SBB IP, Group IP"} >null.txt

}#End of Skip Section Logic

# Open the LSM Log Excel file and gather Data from it
if($Halt -ne "Halt"){ #Skip this Section Logic, Set by changing $Halt from "" to "Halt"
# Load the Windows Forms Library
	Add-Type -AssemblyName System.Windows.Forms
# Find the LSM Log file we want to Analyze
	Write-Host " Please select the LSM Log File to Analyze..."
	$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
		InitialDirectory = [Environment]::GetFolderPath('MyComputer') 
		Filter = 'SpreadSheet (*.xlsx)|*.xlsx'
	}
	$null = $FileBrowser.ShowDialog()
	$LSMLogFile = $FileBrowser.FileName
	Write-Host "File Selected:  $LSMLogFile"
	Write-Host "Opening Excel File... Please wait."
# Create a SaveAs File Name with -LLAv163 at the end
	$LSMLogOutput = $LSMLogFile.split(".")[0]
	$LSMLogOutput = "$LSMLogOutput-$LLAver.xlsx"
# Start MSExcel program
	$Excel = New-Object -ComObject Excel.Application
# Load LSM Log Spreadsheet into MSExcel
	$ExcelDoc = $excel.Workbooks.Open($LSMLogFile)
# Save LSM Log File as a new File
	$ExcelDoc.SaveAs($LSMLogOutput)
# Make Excel Visible (Debug Line)
	#$Excel.Visible = $True
# Make Sheet 1 the active sheet in the LSM Log we're working on
	$Sheet = $ExcelDoc.Worksheets.Item(1)
# find the first Device Response in the Log File and set variable "CellFound"
	$CellFound = $Sheet.Cells.Find("AA FF ")
# Set the search Beginning point so we know when we cycle back to beginning
	$CellBegin = $CellFound.address(0,0,1,1)
# Set the Next Location to Blank, so it doesn't match Begin
	$CellNext = ""
# Import Each Line into the LSMLog, and populate Device Lists
	Write-Host "Importing Data Now... Please Wait..."
	While($CellNext -ne $CellBegin){
	# Set the current location, which will be incremented later
		$CellCurrent = $CellFound.address(0,0,1,1)
	# set the Row Value for use in gathering information
		$CellRow = $CellFound.row
	# Get the entire DevIP Cell (with properites)
		$DevIP = $Sheet.Cells.Item($CellRow,4)
	# Get the DevIP Txt Only, from the CellDevIP
		$DevIP = $DevIP.text
	# Get the entire MDCResp Cell (With properties) from XLS
		$MDCResp = $Sheet.Cells.Item($CellRow,5)
	# Get the MDCResp Text Only from the CellMDCResp Variable
		$MDCResp = $MDCResp.text
	# Copy the MDC Response into an Array called DEV ID to be filtered
		$DevID = $MDCResp.split(" ")
	# Take the Third byte from the MDC Response Array (ID)
		$DevID = $DevID[2]
	# Make the 2-digit hex into a Qualified Hex code 0x05
		$DevID = "0x$DevID"
	#convert that Hex Code to Decimal value
		$DevID = [Byte[]] $DevID
	# Add the Decimal ID number to the LSMLog Array
		$LSMLog.add("$DevIP,$DevID,$MDCResp") >null.txt
	# Add Device to either the SBox or Cabinet Device List
		if($DevID -eq "1"){
		# Add the device to the Sbox List
			$SBBList.add("$DevIP,$DevID") >null.txt
		# Put a Dev Type on screen so I know it's not locked up.
			Write-Host -NoNewLine "SBB:  "
		}
		else{
		# Add the device to the Cabinet List (Guessing it's a cabinet)
			$CabList.add("$DevIP,$DevID") >null.txt
		# Put the Dev type on-screen so I know it's not locked up
			Write-Host -NoNewLine "Cab:  "
		}
	# Add the Log Entry information on-screen so I see it
		Write-host "$DevIP, $DevID, $MDCResp" # MDC command removed
	#Find the next instance and create "next"
		$CellFound = $Sheet.Cells.FindNext($CellFound)
		$CellNext = $CellFound.address(0,0,1,1)
	} #End of While loop to gather data from the Excel Log.
# Gather count of Excel Lines Processed
	$LSMLogCountTotal = $LSMLog.count
# Filter and remove duplicate entries from the Data
	$LSMLog = $LSMLog | Select -unique
	$SBBlist = $SBBList | Select -unique
	$CabList = $CabList | Select -unique
# Generate and Present Data Summary
	$SBBCount = $SBBlist.count
	$CabCount = $Cablist.count
	$LSMLogCount = $LSMLog.count
	Write-Host "-----"
	Write-Host "LSM Log Lines processed:  $LSMLogCountTotal"
	Write-Host "Unique Log count:  $LSMLogCount"
	Write-Host "SBoxes found:  $SBBCount"
	Write-Host "Cabinets found:  $CabCount"
}#End of Skip Section Logic

# Setup Excel Sheet and set Pointers for Data Gathering
if($Halt -ne "Halt"){ #Skip this Section Logic, Set by changing $Halt from "" to "Halt"
# Create a New Sheets in Excel to add Data to (4)
	Write-Host "Adding Worksheets to LSM Log File to store data."
	$Excel.Visible = $True
	sleep -m 500
	$Sheet = $excel.Worksheets.Add()
	$Sheet.Name = "CabInfo"
	$Sheet = $excel.Worksheets.Add()
	$Sheet.Name = "SBBInfo"
	$Sheet = $excel.Worksheets.Add()
	$Sheet.Name = "CabLayouts"
}#End of Skip Section Logic

# Gather data for each SBB in the SBB List
if($Halt -ne "Halt"){ #Skip this Section Logic, Set by changing $Halt from "" to "Halt"
# Setup up SBB Page Header and Settings
	Write-Host "Adding Header and Settings for SBB Data Sheet"
# Move to the SBB Info Worksheet (2)
	$Sheet = $ExcelDoc.Worksheets.Item(2)
	$Sheet.Activate()
# Add the Headers to the Sheet
	$Column2 = 1
	$sheet.Cells.Item(1,1) = "Setting"
	$sheet.Cells.Item(2,1) = "SBox IP"
	$sheet.Cells.Item(3,1) = "Model"
	$sheet.Cells.Item(4,1) = "Serial Number"
	$sheet.Cells.Item(5,1) = "Main Version"
	$sheet.Cells.Item(6,1) = "Additional Versions"
	$sheet.Cells.Item(7,1) = "SBox Name"
	$sheet.Cells.Item(8,1) = "Group 1 IP"
	$sheet.Cells.Item(9,1) = "Group 2 IP"
	$sheet.Cells.Item(10,1) = "Group 3 IP"
	$sheet.Cells.Item(11,1) = "Group 4 IP"
	$sheet.Cells.Item(12,1) = "Video Wall Mode"
	$sheet.Cells.Item(13,1) = "Cabinet Layout"
	$sheet.Cells.Item(14,1) = "SBox Output Resolution"
	$sheet.Cells.Item(15,1) = "Cabinet Resolution"
# Bold the Setting Header
	$sheet.Cells.Item(1,1).Font.Bold=$True

# Start Gathering indvidual settings from each SBB
Write-Host "Getting Data for SBox's"
$SBBEntry = 1
	ForEach($SBB in $SBBList){ 
	# DEBUGGING:  If list has mulitple SBB, run $SBB = $SBBList[0] for the first line
	# DEBUGGING:  If the list is only one SBB, run $SBB = $SBBList
	# Create a Verbose SBB Name for the Column Header
		$SBBName = "SBB $SBBEntry"

	#Setup Row and column for Data
		$Column2 = $SBBEntry + 1

	# Add Header to SBB Spreadsheet
		$sheet.Cells.Item(1,$Column2) = "$SBBName"
		$sheet.Cells.Item(1,$Column2).Font.Bold=$True

	# Get the IP Address of this SBB and add it to the Spreadsheet
		$DevIP = $SBB.split(",")[0]
		$Value = $DevIP
		$Sheet.Cells.Item(2,$Column2) = $Value
		Write-Host -NoNewLine "$Value, "

	# Get and Set the SBB ID number
		$DevID = $SBB.split(",")[1]
		Write-Host "SBox ID:  $DevID"

	# Gather all logs for this device
		Write-Host "Getting all Log Entries for SBB $DevIP"
		$DevLog = $LSMLog -match "$DevIP,$DevID,"
		
	# Get the Model Name (0x8A) from $DevLog
		$ModelName = Get-MDCAscii " 41 8A "
		# Test to see if the Model Number comes back valid, if not, use use FW to get model
		if($ModelName -eq 0){  #If the value is blank, get the Model from the FW Version
		# Get the FW Version of the SBB  (41 0E)
			$Value = Get-MDCASCII " 41 0E "
		# Pull the first 6 Digits for the Model number "TB-HMS"
			$Value = $Value[0..5]
		# Turn this into one string of text
			$Value = "$Value"
		# Remove the Spaces from the String
			$Value = $Value | foreach {$_ -replace " ",""}
		# Lookup the FW Name in the SBB Type Table (Above)
			$SBBType = $tSBBType -match $Value
		# Pull the Verbose SBB type from the Table Entry
			$SBBType = $SBBType.split(",")[1]
		# Enter the Model Name into the Spreadsheet
			$Sheet.Cells.Item(3,$Column2) = $SBBType	
		}
		else{
			$Sheet.Cells.Item(3,$Column2) = $ModelName
			$SBBType = $ModelName[0..7]
			$SBBType = "$SBBType"
			$SBBType = $SBBType | foreach {$_ -replace " ",""}
		}

	# Get the Serial Number 0x0B
		$Value = Get-MDCASCII " 41 0B "
		$Sheet.Cells.Item(4,$Column2) = $Value

	# Get the Main Firmware Version (0x0E)
		$Value = Get-MDCASCII " 41 0E "
		$Sheet.Cells.Item(5,$Column2) = $Value
		
	# Get the additional Firmware Versions of the SBB
		# Determine what kind of SBB it is...
		if($ModelName -match "AU"){
		# Set the Search String for JAU/HAU
			$SearchTxt = " 41 D2 32 "  #Dave and Busters Example
		#set Version Count Location in MDCResp
			$FieldCountLoc = 8
		#Set Initial starting point for First Version
			$DataStart = 11
		} #End of If it's a SBB-SNOWJAU
		elseif($ModelName -match "3U"){ #If it match's H3U or J3U
		#Set the Search String for J3U and H3U
			$SearchTxt = " 41 1B A4 "
		#set Version Count Location in MDCResp
			$FieldCountLoc = 7
		#Set Initial starting point for First Version
			$DataStart = 10
		}#End of if it's an SBB-SNOWJ8U
		else {
			Write-Host "Unknown SBB type, Skipping Section"
			$SkipSection = "True"
		} #Setup for other SBB Types
		if($SkipSection -ne "True"){
			#Capture the Line of data that matchs
			$DevLogEntry = $DevLog -match "$SearchTXT" | Select-Object -last 1
			#Grab the third part of the log line (separated by comma's)
			$MDCResp = $DevLogEntry.split(",")[2]
			# Create an Array from the Hex Data String
			$MDCResp = $MDCResp.split(" ")
			# Determine how many fields I'm pulling
			$FieldCount = $MDCResp[$FieldCountLoc]
			#converti it to Decimal for Counting
			$FieldCount = [System.Convert]::ToInt16($FieldCount,16)
			#Set Counter to 0 so I know how many I've counted
			$Count = 0
			# Gather the first Version number in the data
			$Count ++
			#Set the Field Length (Digit Prior to Field Start)
			$DataLen = $MDCResp[$DataStart-1]
			#convert to Decimal for Calculations
			$DataLen = [System.Convert]::ToInt16($DataLen,16)
			#Set the End of the Data Field postion
			$DataEnd = $DataStart + $DataLen - 1
			#Capture Field Data to an Array
			$DataX = $MDCResp[$DataStart..$DataEnd]
			#Start an Array to capture the Field Data
			$enc = [System.Text.Encoding]::ASCII #Set Encoding to ASCII
			$DataA = New-Object System.Collections.ArrayList #Create new Array
			ForEach($line in $DataX){ #Start processing Loop and set exit parameter
				if($line -eq "00"){}
				else{
					$byteX = "0x$line"
					$DataA.add($enc.GetString($byteX)) >null.txt
				} #End of If it's not NULL 0x00
			} #End of For Each entry in the Array
			#Make String out of the Array
			$DataA = "$DataA"
			#Remove Spaces from the Response
			$Value = $DataA | foreach {$_ -replace " ",""}
			#Enter the Data into the Version List Variable
			$Ver2s = $Value
			$DataStart = $DataEnd + 3
			#Gather each Additional Version and increment Current Field, Displaying on-screen (for now)
			While($Count -ne $FieldCount){
				#now that we started the gathering, inc the Current field to what we're gathering
				$Count ++
				#Set the Field Length (Digit Prior to Field Start)
				$DataLen = $MDCResp[$DataStart-1]
				#convert to Decimal for Calculations
				$DataLen = [System.Convert]::ToInt16($DataLen,16)
				#Set the End of the Data Field postion
				$DataEnd = $DataStart + $DataLen - 1
				#Capture Field Data to an Array
				$DataX = $MDCResp[$DataStart..$DataEnd]
				#Start an Array to capture the Field Data
				$enc = [System.Text.Encoding]::ASCII #Set Encoding to ASCII
				$DataA = New-Object System.Collections.ArrayList #Create new Array
				ForEach($line in $DataX){ #Start processing Loop and set exit parameter
					if($line -eq "00"){}
					else{
						$byteX = "0x$line"
						$DataA.add($enc.GetString($byteX)) >null.txt
					} #End of If it's not NULL 0x00
				} #End of For Each entry in the Array
				#Make String out of the Array
				$DataA = "$DataA"
				#Remove Spaces from the Response
				$Value = $DataA | foreach {$_ -replace " ",""}
				#Add the New Value to the Version List Ver2s
				$Ver2s = "$Ver2s `r`n$Value"
				#Setup Data Start Point for the next data Field
				$DataStart = $DataEnd + 3
			} # End of While Loop to get data for each Version
			#Enter Data into Spreadsheet
			$Value = $Ver2s
			$Sheet.Cells.Item(6,$Column2) = $Value
		} #End of If Halt is NOT True

	# Get the SBox Friendly Name ( 41 67 )
		$Value = Get-MDCData " 41 67 "
###### Check to see if the value is Null, or if it contains information
		if($Value -match "NULL") {
			$Value = "No Data Found"
		} # End of if value was not found
		else{ # If Value has Hex Data, Convert it and add it.
			$Value = Conv-x2a $Value
		} # End of converion to text
		# Write the results to the Sheet
		$Sheet.Cells.Item(7,$Column2) = $Value

	# Get the Group IP Addresses from the SBox (Decimal)
		#Pull data for the Group IP Addresses	
		$Lookup = Get-MDCData " 41 1B 84 "
		# convert the HEX to Decimal
		$Lookup = Conv-x2d $Lookup
		# Set the starting row for the 4 IP Address's
		$Row2 = 8  #Temporary Place Holder
		# Set an IP Loop Counter (1-4)
		$count = 0
		# Set IP Start poing
		$IPStart = 1
		# Gather each IP and add to Spreadsheet
		While($count -ne 4){
			# Increment Group# to first group
			$count++
			# Set the Group Setting Name for the spreadsheet
			# Create a IP location stop point
			$IPEnd = $IPStart + 3
			# Get the Array of IP's for this group
			$GroupIP = $Lookup[$IPStart..$IPEnd]	
			# Convert it to text
			$GroupIP = "$GroupIP"
			# Change the Spaces to Dots for IP formatting
			$Value = $GroupIP | foreach {$_ -replace " ","."}
			# Add the value to the Spreadsheet
			$Sheet.Cells.Item($Row2,$Column2) = $Value
			$Row2++
			# Add the value to the SBB-GroupIP Table, if the value is not 0.0.0.0
			if($Value -ne "0.0.0.0"){
				$GroupIPTable.add("$SBBEntry,$count,$DevIP,$Value") >null.txt
			}
			else{} # Do nothing
			# Setup the next Startpoing for IP's.
			$IPStart = $IPStart + 4
		} # End of Gather Group IP Loop
		
	# Get the Video Wall mode from the SBB
		$Lookup = Get-MDCData " 41 84 "
		$Value = $tSBBVWMode -match $Lookup
		$Value = $Value.split(",")[1]
		$Sheet.Cells.Item(12,$Column2) = $Value
		# Change the Cell based on ON/Off
		if($Value -match "ON"){  #Set Cell to Light Green if ON
			$Sheet.Cells(12,$Column2).Interior.ColorIndex = 35
		}
#		elseif($Value -match "OFF"){ # Set Cell to Light Gray if OFF
#			$Sheet.Cells(12,$Column2).Interior.ColorIndex = 15
#		 } 
		else{} # If not either, do nothing

	# Increment the SBB Entry by one
		$SBBEntry++
	} # End of Multi SBB Gather Loop
# Auto Fit the Cells so they can be seen
	# Select all Cells that have data in them
	$usedRange = $Sheet.UsedRange	
	# Set the Column Width's to something really Wide
	$usedRange.EntireColumn.ColumnWidth = 48
	# Auto Fit the Row Heighth
	$usedRange.EntireRow.AutoFit() | Out-Null
	# Auto Fit the Column Width
	$usedRange.EntireColumn.AutoFit() | Out-Null
	# Set the Text to be Top Aligned
	$usedRange.VerticalAlignment = -4160
}# End of Halt Section Logic

#Gather Cabinet Information
if($Halt -ne "Halt"){ #Skip this Section Logic, Set by changing $Halt from "" to "Halt"
# Move to the Layout Sheet (1)
	$Sheet = $ExcelDoc.Worksheets.Item(3)
	$Sheet.Activate()
# Build a Header for the Sheet
	$sheet.Cells.Item(1,1) = "SBox IP"
	$sheet.Cells.Item(1,2) = "Group IP"
	$sheet.Cells.Item(1,3) = "Cab ID"
	$sheet.Cells.Item(1,4) = "Serial#"
	$Sheet.Cells.Item(1,5) = "Model#"
	$sheet.Cells.Item(1,6) = "Cabinet FW"
	$sheet.Cells.Item(1,7) = "Cabinet FPGA"
	$sheet.Cells.Item(1,8) = "Temp(c)"
	$sheet.Cells.Item(1,9) = "BackLt"
	$sheet.Cells.Item(1,10) = "Cab RGB CC"
	$sheet.Cells.Item(1,11) = "Mod RGB CC"
	$sheet.Cells.Item(1,12) = "Pix RGB CC"
	$sheet.Cells.Item(1,13) = "Seam Cor"
	
# Set Starting Row for Cabinet Data
	$Row3 = 2
	
# Create a Cabinet Location Array for ALL Cabinets to be analyzed Later
	$CabLocs = New-Object System.Collections.ArrayList

# Gather Data for each entry in the GroupIPTable
	ForEach($Group in $GroupIPTable){
	# DEGUBBING $Group = $GroupIPTable[0]
	# Extract SBox number from $Group
		$SBoxID = $Group.split(",")[0]
		# Remove the spaces from the Split command
		$SBoxID = $SBoxID | foreach {$_ -replace " ",""}
	# Define the SBox IP Address by extracting it from $Group
		$SBoxIP = $Group.split(",")[2]
		# Remove the spaces from the Split command
		$SBoxIP = $SBoxIP | foreach {$_ -replace " ",""}
	# Extract Group Number from $Group
		$GroupID = $Group.split(",")[1]
		# Remove the spaces from the Split command
		$GroupID = $GroupID | foreach {$_ -replace " ",""}
	# Define the Group IP by extracting it from $Group
		$GroupIP = $Group.split(",")[3]
		# Remove the spaces from the Split command
		$GroupIP = $GroupIP | foreach {$_ -replace " ",""}
	#Build a Cabinet list for this group (Include a "," to annote IP address End .12, not .121,)
		$GroupCabList = $CabList -Match "$GroupIP,"
	# Start getting data for each Cabinet entry
		ForEach($Cabinet in $GroupCabList){
		# DEBUGGING:  $Cabinet = $GroupCabList[0]
		# Define the Cabinet ID by extracting it from $Cabinet
			$CabID = $Cabinet.split(",")[1]
			# Remove the spaces from the Split command
			$CabID = $CabID | foreach {$_ -replace " ",""}
		
		# Build a log for this device only (indluce a , to annotate end of Cabinet number 2, not 21 or 22 or 23)
			$DevLog = $LSMLog -match "$Cabinet,"
		
		# SBox IP Address
			$Value = $SBoxIP
			$sheet.cells.item($Row3,1) = $Value
			Write-Host -NoNewLine "$Value, "
	
		# Group IP Address
			$Value = $Cabinet.split(",")[0]
			$sheet.cells.item($Row3,2) = $Value
			Write-Host -NoNewLine "$Value, "
			
		# Cabinet ID
			$Value = $Cabinet.split(",")[1]
			$Sheet.Cells.Item($Row3,3) = $Value
			Write-Host -NoNewLine "$Value, "
			
		# Cabinet SN ( 41 0B )
			$Value = Get-MDCASCII " 41 0B "
			$Sheet.Cells.Item($Row3,4) = $Value
			Write-Host -NoNewLine "$Value, "
					
		# Cabinet Model Number ( 41 8A )
			$Value = Get-MDCASCII " 41 8A "
			$Sheet.Cells.Item($Row3,5) = $Value
			Write-Host -NoNewLine "$Value, "
			
		# Capture the Version information from the DevLog
			$DevLogEntry = $DevLog -match " 41 1B A4 " | Select-Object -last 1
			#Grab the third part of the log line (separated by comma's)
			$MDCResp = $DevLogEntry.split(",")[2]
			# Create an Array from the Hex Data String
			$MDCResp = $MDCResp.split(" ")
		# Capture the Main FW version, starting in position 10
			$DataStart = 10
			#Set the Field Length (Digit Prior to Field Start)
			$DataLen = $MDCResp[$DataStart-1]
			#convert to Decimal for Calculations
			$DataLen = [System.Convert]::ToInt16($DataLen,16)
			#Set the End of the Data Field postion
			$DataEnd = $DataStart + $DataLen - 1
			#Capture Field Data to an Array
			$DataX = $MDCResp[$DataStart..$DataEnd]
			#Start an Array to capture the Field Data
			$enc = [System.Text.Encoding]::ASCII #Set Encoding to ASCII
			$DataA = New-Object System.Collections.ArrayList #Create new Array
			ForEach($line in $DataX){ #Start processing Loop and set exit parameter
				if($line -eq "00"){}
				else{
					$byteX = "0x$line"
					$DataA.add($enc.GetString($byteX)) >null.txt
				} #End of If it's not NULL 0x00
			} #End of For Each entry in the Array
			#Make String out of the Array
			$DataA = "$DataA"
			#Remove Spaces from the Response
			$Value = $DataA | foreach {$_ -replace " ",""}
			#Enter Data into Spreadsheet
			$Sheet.Cells.Item($Row3,6) = $Value
			#Test to see if this is the first version recorded, if it is, create the Reference Version
			if($1stVerCabMain -eq "None"){
				$1stVerCabMain = $Value
			} # End of If it's the first version
			else{ # If it is not the first Cabinet
				if($Value -ne $1stVerCabMain){ #If it does not match, Change color to Tan/Gold
					$Sheet.Cells($Row3,6).Interior.ColorIndex = 40
				} # End of if version does NOT match
				else {} # If does match, do nothing
			} # End of IF version is not the first
			Write-Host -NoNewLine "$Value, "
		#Capture the FPGA Version (Next Version after Main)
			#Setup Data Start Point for the next data Field
			$DataStart = $DataEnd + 3
			#Set the Field Length (Digit Prior to Field Start)
			$DataLen = $MDCResp[$DataStart-1]
			#convert to Decimal for Calculations
			$DataLen = [System.Convert]::ToInt16($DataLen,16)
			#Set the End of the Data Field postion
			$DataEnd = $DataStart + $DataLen - 1
			#Capture Field Data to an Array
			$DataX = $MDCResp[$DataStart..$DataEnd]
			#Start an Array to capture the Field Data
			$enc = [System.Text.Encoding]::ASCII #Set Encoding to ASCII
			$DataA = New-Object System.Collections.ArrayList #Create new Array
			ForEach($line in $DataX){ #Start processing Loop and set exit parameter
				if($line -eq "00"){}
				else{
					$byteX = "0x$line"
					$DataA.add($enc.GetString($byteX)) >null.txt
				} #End of If it's not NULL 0x00
			} #End of For Each entry in the Array
			#Make String out of the Array
			$DataA = "$DataA"
			#Remove Spaces from the Response
			$Value = $DataA | foreach {$_ -replace " ",""}
			#Enter Data into Spreadsheet
			$Sheet.Cells.Item($Row3,7) = $Value
			#Test to see if this is the first version recorded, if it is, create the Reference Version
			if($1stVerCabFPGA -eq "None"){
				$1stVerCabFPGA = $Value
			} # End of If it's the first version
			else{ # If it is not the first Cabinet
				if($Value -ne $1stVerCabFPGA){ #If it does not match, Change color to Tan/Gold
					$Sheet.Cells($Row3,7).Interior.ColorIndex = 40
				} # End of if version does NOT match
				else {} # If does match, do nothing
			} # End of IF version is not the first
			Write-Host -NoNewLine "$Value, "

		# Pull the Video Wall Data ( 41 8C A0 )
			$Lookup = Get-MDCData " 41 8C A0 "
			#Create a Text String from the response
			$Lookup = "$Lookup"
			# Remove spaces so I can grab 3 digits at a time (02800002D0)
			$Lookup = $Lookup | foreach {$_ -replace " ",""}
		# Cabinet Location X-Axis (Bytes 3 & 4)
			$CabLocXGroup = $Lookup[(6..9)]
			# Convert into a String (one line)
			$CabLocXGroup = "$CabLocXGroup"
			# Remove Spaces (000)
			$CabLocXGroup = $CabLocXGroup | foreach {$_ -replace " ",""}
			# Convert from Hex to Decimal
			$CabLocXGroup = [System.Convert]::ToInt16($CabLocXGroup,16)
		# Cabinet Location Y-Axis (Bytes 5 & 6)
			$CabLocYGroup = $Lookup[(10..13)]
			# Convert into a String (one line)
			$CabLocYGroup = "$CabLocYGroup"
			# Remove Spaces (000)
			$CabLocYGroup = $CabLocYGroup | foreach {$_ -replace " ",""}
			# Convert from Hex to Decimal
			$CabLocYGroup = [System.Convert]::ToInt16($CabLocYGroup,16)
		# Cabinet Width ( Bytes 7 & 8)
			$CabWidth = $Lookup[(14..17)]
			# Convert into a String (one line)
			$CabWidth = "$CabWidth"
			# Remove Spaces (000)
			$CabWidth = $CabWidth | foreach {$_ -replace " ",""}
			# Convert from Hex to Decimal
			$CabWidth = [System.Convert]::ToInt16($CabWidth,16)
		# Cabinet Height (Bytes 9 & 10)
			$CabHeight = $Lookup[(18..21)]
			# Convert into a String (one line)
			$CabHeight = "$CabHeight"
			# Remove Spaces (000)
			$CabHeight = $CabHeight | foreach {$_ -replace " ",""}
			# Convert from Hex to Decimal
			$CabHeight = [System.Convert]::ToInt16($CabHeight,16)
		# Calculate SBB Starting Pixel Location for the Layout
			#Apply Offsets for the SBB locations
			if($GroupID -eq 2){	# Apply Group 2 Offset
				$CabLocXSbb = $CabLocXGroup	
				$CabLocYSbb = $CabLocYGroup + 1080	# Apply Downard Offset
			} # End of Group 2 Offset
			elseif($GroupID -eq 3){ # Apply Group 3 Offset
				$CabLocXSbb = $CabLocXGroup + 1920	# Apply Offset to the Right
				$CabLocYSbb = $CabLocYGroup	
			} # End of Group 3 Offset
			elseif($groupID -eq 4){ # Apply Group 4 Offets
				$CabLocXSbb = $CabLocXGroup + 1920 	# Apply Offset to the Right
				$CabLocYSbb = $CabLocYGroup + 1080	# Apply Downward Offset
			} # End of Group 4 Offsets
			else{ # Group 1 or other... No Offest Needed
				$CabLocXSbb = $CabLocXGroup	
				$CabLocYSbb = $CabLocYGroup	
			} # End of No Offset Needed
		# Capture Cabinet information into Tables for parsing later
			$CabLocs.add("$SBoxIP,$GroupIP,$SBoxID,$GroupID,$CabID,$CabWidth,$CabHeight,$CabLocXGroup,$CabLocYGroup,$CabLocXSbb,$CabLocYSbb") >null.txt

		# Get the Temperature ( 41 D0 84 ) Byte 4
			$Lookup = Get-MDCData " 41 D0 84 "
			# Check to see if data is Null, and simply make the cell Gray
			if($Lookup -match "Null"){
				# Set Cell to Light Gray if no data found
				$Value = " "
				$Sheet.Cells($Row3,8).Interior.ColorIndex = 15
			} # End of Null Value Routine
			else{ # If the Value is Not Null, process the information
				# Grab the Value from the Return
				$Value = $Lookup[3]
				# Convert the Hex into a number
				$Value = [System.Convert]::ToInt16($Value,16)
				# Enter the value in the Spreadsheet
				$Sheet.Cells.Item($Row3,8) = $Value	
				# Check to see if the Value is 66 or Higher
				if($Value -gt 59){
					# Set the Background Color to Yellow
					$Sheet.Cells($Row3,8).Interior.ColorIndex = 6
				}
				else{} #If not too high, do nothing
			}
			Write-Host -NoNewLine "$Value, "
		
		# Cabinet Backlight ( 41 D0 94 ) 
			$Lookup = Get-MDCData " 41 D0 94 "
			If($Lookup -match "NULL"){ # If Get Data returns a NULL
				$Value = " "
				$Sheet.Cells($Row3,9).Interior.ColorIndex = 15
			} # End of NULL Returned Option
			else{ # If Lookup is Not Null
				$Value = $Lookup[1]
				$Value = [System.Convert]::ToInt16($Value,16)
				$Sheet.Cells.Item($Row3,9) = $Value
			} # End of if Loopup is not NULL
			Write-Host -NoNewLine "$Value, "
			
		# Cabinet RGB ON/Off ( 41 D0 9E ) 
			$Lookup = Get-MDCData " 41 D0 9E "
			$Value = $Lookup[1]
			$Value = $tCabcc -match $Value
			# Check to see if data is Null, and simply make the cell Gray
			if($Value -match "Null"){
				# Set Cell to Light Gray if no data found
				$Value = " "
				$Sheet.Cells($Row3,10).Interior.ColorIndex = 15
			} # End of Null Value Routine
			else{ # If the Value is Not Null, process the information
				# Split the return and grab the Verbose Setting
				$Value = $Value.split(",")[1]
				# Enter the Value into the Spreadsheet
				$Sheet.Cells.Item($Row3,10) = $Value	
				# Check to see if the value is ON
				if($Value -match "ON"){
					#change the Background Color to Light Green if it's on
					$Sheet.Cells($Row3,10).Interior.ColorIndex = 35
				}
				else{} #If it's another setting, don't color code it
			} # End of routine if it's NOT null.
			Write-Host -NoNewLine "$Value, "

		# Module RGB ON/Off ( 41 D0 99 ) 
			$Lookup = Get-MDCData " 41 D0 99 "
			$Value = $Lookup[1]
			$Value = $tModcc -match $Value
			# Check to see if data is Null, and simply make the cell Gray
			if($Value -match "Null"){
				# Set Cell to Light Gray if no data found
				$Value = " "
				$Sheet.Cells($Row3,11).Interior.ColorIndex = 15
			} # End of Null Value Routine
			else{ # If the Value is Not Null, process the information
				# Split the return and grab the Verbose Setting
				$Value = $Value.split(",")[1]
				# Enter the Value into the Spreadsheet
				$Sheet.Cells.Item($Row3,11) = $Value	
				# Check to see if the value is ON
				if($Value -match "ON"){
					#change the Background Color to Light Green if it's on
					$Sheet.Cells($Row3,11).Interior.ColorIndex = 35
				}
				else{} #If it's another setting, don't color code it
			} # End of routine if it's NOT null.
			Write-Host -NoNewLine "$Value, "

		# Pixel RGB CC On/Off ( 41 D0 95 ) 
			$Lookup = Get-MDCData " 41 D0 95 "
			$Value = $Lookup[1]
			$Value = $tPixcc -match $Value
			# Check to see if data is Null, and simply make the cell Gray
			if($Value -match "Null"){
				# Set Cell to Light Gray if no data found
				$Value = " "
				$Sheet.Cells($Row3,12).Interior.ColorIndex = 15
			} # End of Null Value Routine
			else{ # If the Value is Not Null, process the information
				# Split the return and grab the Verbose Setting
				$Value = $Value.split(",")[1]
				# Enter the Value into the Spreadsheet
				$Sheet.Cells.Item($Row3,12) = $Value	
				# Check to see if the value is ON
				if($Value -match "ON"){
					#change the Background Color to Light Green if it's on
					$Sheet.Cells($Row3,12).Interior.ColorIndex = 35
				}
				else{} #If it's another setting, don't color code it
			} # End of routine if it's NOT null.
			Write-Host -NoNewLine "$Value, "
		
		# Seam correction on/off ( 41 D0 98 ) 
			$Lookup = Get-MDCData " 41 D0 98 "
			$Value = $Lookup[1]
			$Value = $tSeamC -match $Value
			# Check to see if data is Null, and simply make the cell Gray
			if($Value -match "Null"){
				# Set Cell to Light Gray if no data found
				$Value = " "
				$Sheet.Cells($Row3,13).Interior.ColorIndex = 15
			} # End of Null Value Routine
			else{ # If the Value is Not Null, process the information
				# Split the return and grab the Verbose Setting
				$Value = $Value.split(",")[1]
				# Enter the Value into the Spreadsheet
				$Sheet.Cells.Item($Row3,13) = $Value	
				# Check to see if the value is ON
				if($Value -match "ON"){
					#change the Background Color to Light Green if it's on
					$Sheet.Cells($Row3,13).Interior.ColorIndex = 35
				}
				else{} #If it's another setting, don't color code it
			} # End of routine if it's NOT null.
			Write-Host -NoNewLine "$Value, "

		# End the Data Line On-Screen
			Write-Host ""
		# Set Column Width to Auto
			$usedRange = $Sheet.UsedRange	
			$usedRange.EntireColumn.AutoFit() | Out-Null
		# Prepare for the next Cabinet Entry
			$Row3++
		} #End of this Cabinet
	} #End of This Group's
}#End of Skip Section Logic

# Analyze Cabinet Location Table Information and present Data accordingly
if($Halt -ne "Halt"){ #Skip this Section Logic, Set by changing $Halt from "" to "Halt"
# Pull Data for each SBB one at a time and process
	Write-Host "Analyzing Cabinet Location Data..."
	ForEach($SBB in $SbbList){
	# DEBUGGING $SBB = $SbbList[0]
	# Move to the Layout Sheet (1)
		$Sheet = $ExcelDoc.Worksheets.Item(1)
		$Sheet.Activate()
	# Create an Array to store all the different X Locations for later calculations
		$XLocs = New-Object System.Collections.ArrayList
	# Create an Array to store all the different Y Locations For later calculations
		$YLocs = New-Object System.Collections.ArrayList
	# Setup Hi and Low Value holders for X&Y to be overwritten as values are found
		$XLocLow = 999999
		$XLocHigh = 0
		$YLocLow = 999999
		$YLocHigh = 0
	# Extract the SBB IP Address from SBBList
		$SbbIP = $SBB.split(",")[0]
	# Pull all Cabinets for this SBox and add to an SBB table
		$CabLocsSbb = $CabLocs -match "$SBBIP"
	# Create an SBB name to use in the Spreadsheet
		$Line = $CabLocsSBB[0]  			# Pull a Cab Location line and put it into a variable
		$SBBID = $line.split(",")[2]  		# grab the 3rd entry (sbb ID)
		$SBBName = "Sbox $SbbID`: $SbbIP"  	# Make a Verbose Name
	# Create an entry in excel to indicate this Layout's Name
		$Sheet.Cells.Item($Row1,1) = "$SbbName"
		$Sheet.Cells.Item($Row1,1).Font.Bold=$True
	# Increment row counter to Prepare for adding Cabinets below the header
		$Row1++
	# Create a Reference cell to use as an offset for placing cabinets in cells 
		$StartRow = $Row1
		$StartCol = $Column1
	# Start Processing Cabinets information one Cabinet at a time
		ForEach($Line in $CabLocsSbb){ 
		#DEBUGGING:  $Line = $CabLocsSBB[0]
		# Get the Group ID Name (1-4)
			$CabGrp = $Line.split(",")[3]
		# Get the Cabinet ID, X and Y Starting Pixel Location
			$CabID = $Line.split(",")[4]
		# Get the Cabinet X-Axis Resolution
			$CabResX = $Line.Split(",")[5]
		# Get the Cabinet Y-Axis Resolution
			$CabResY = $Line.Split(",")[6]
		# Get the X-Axis Pixel Location (SBB Location, not the Group Location)
			$XLocPix = $Line.Split(",")[9]
		# Determine Which Column this is (Starts with 0)
			$CabCol = $XLocPix / $CabResX
		# Add this value to the X Location Array
			$XLocs.add($XLocPix) >null.txt  #Originally it was a String with ""'s
		# See if the value is Highest or Lowest and if yes, use it
			If($XLocPix -ne 0){ #Skip if the value is 0
				if($XLocPix -gt $XLocHigh){
					$XlocHigh = $XLocPix
				} else{} # Do NOthing}
				if($XLocPix -lt $XLocLow){
					$XLocLow = $XLocPix
				} else{} # Do NOthing}
			}Else{} #Do Nothing
		# Get the Y-Axis Pixel Location (SBB Location, not the Group Location)
			$YLocPix = $Line.Split(",")[10]
		# Determine Which Row this is (Starts with 0)
			$CabRow = $YLocPix / $CabResY
		# Add this value to the Y Locations Arrary
			$YLocs.add($YLocPix) >null.txt
		# See if the value is Highest or Lowest and if yes, use it
			If($YLocPix -ne 0){ #Skip if the value is 0
				if($YLocPix -gt $YLocHigh){
					$YlocHigh = $YLocPix
				} else{} # Do NOthing}
				if($YLocPix -lt $YLocLow){
					$YLocLow = $YLocPix
				} else{} # Do NOthing}
			}Else{} #Do Nothing
		# Create a Pix Location Value for documentation
			$CabPixLoc = "$XLocPix x $YLocPix"
		# Set the Excel Cell Row
			$CellRow = $StartRow + $CabRow
		# Set the Excel Cell Column
			$CellCol = $StartCol + $CabCol
		# Place the ID number in the Excel Sheet
			$sheet.Cells.Item($CellRow,$CellCol) = "Group-$CabGrp`r`nCabID-$CabID`r`n$CabPixLoc"
		# Shade the Cell based on the Group Number
			if($CabGrp -eq 1){ #If the Entry is Group 1, Set Color to Green
				$Sheet.Cells($CellRow,$CellCol).Interior.ColorIndex = 35
			} #End of Cabinet Group 1
			elseif($CabGrp -eq 2){ #If the Entry is Group 2, Set Color to Violet
				$Sheet.Cells($CellRow,$CellCol).Interior.ColorIndex = 24
			} #End of Cabinet Group 2
			elseif($CabGrp -eq 3){ #If the Entry is Group 3, Set Color to Blue
				$Sheet.Cells($CellRow,$CellCol).Interior.ColorIndex = 37
			} #End of Cabinet Group 2
			elseif($CabGrp -eq 4){ #If the Entry is Group 4, Set Color to tan/Gold
				$Sheet.Cells($CellRow,$CellCol).Interior.ColorIndex = 40
			} #End of Cabinet Group 2
			else{} # If nothing Match's, do nothing
		#Debugging Lines
			Write-Host "Sbb-$SbbID, Group-$CabGrp, Cabinet-$CabID, Location-$CabPixLoc, Row-$CabRow, Column-$CabCol"
		} # End of Cabinet
	# Clean up Counters and stuff
		# Remove Duplicate X Columns List
		$XLocs = $XLocs | Select -unique
		# Sort XLoc's by Values
		$XLocs = $XLocs | Sort-Object 
		# Get rid of duplicate Rows
		$YLocs = $YLocs | Select -unique
		# Sort Y Locs
		$Ylocs = $Ylocs | Sort-Object
		# Determine how many Columns there are (X)
		$Columns = $XLocs.count
		# Determine how many Rows there are (Y)
		$Rows = $YLocs.count

	# Format the Text for the Cabinet Cells 
		#Create the A1 Version of the Start Column
		$StartColA1 = ExcelA1 $StartCol
		# Create the Starting Cell in A1 Format
		$StartCellA1 = "$StartColA1$StartRow"
		#Calculate End Row and Column
		$EndCol = $StartCol + $Columns -1
		$EndRow = $StartRow + $Rows -1
		# Create the A1 Version of the Ending Column
		$EndColA1 = ExcelA1 $EndCol
		# Create the Ending Cell in A1 Format
		$EndCellA1 = "$EndColA1$EndRow"
		#Highlight the range and format it.
		$GroupRange = "$StartCellA1`:$EndCellA1"
		$Selection = $Sheet.range("$GroupRange")
		$Selection.Select() >null.txt
		$Selection.Columns().ColumnWidth = 12
		$Selection.Columns().HorizontalAlignment = -4108

	# Move to the SBB Spreadsheet
		$Sheet = $ExcelDoc.Worksheets.Item(2)
		$Sheet.Activate()
	# Move to the proper Column for this SBB
		$SBBIDN = [System.Convert]::ToInt16($SBBID,10) # Create a Decimal SBBID
		$Column2 = $SbbIDN + 1
	
	# Document the Cabinet Layout (10x5, 6x3, etc)
		$Setting = "Cabinet Layout"
		$Value = "$Columns x $Rows"
		$Sheet.Cells.Item(13,$Column2) = $Value

	# SBox Video wall Resolution
		$Setting = "SBox Output Resolution"
		$CabResXD = [System.Convert]::ToInt16($CabResX,10) # Create a Decimal Value
		$CabResYD = [System.Convert]::ToInt16($CabResY,10) # Create a Decimal Value
		$VWResX = $CabResXD * $Columns
		$VWResY = $CabResYD * $Rows
		$Value = "$VWResX x $VWResY"
		$Sheet.Cells.Item(14,$Column2) = $Value
		
	# Get the Cabinet Resolutions Values X adn Y
		$Setting = "Cabinet Resolution"
		$Value = "$CabResX x $CabResY"
		$Sheet.Cells.Item(15,$Column2) = $Value
	
	#Prepare Spreadsheet for next SBox
		$row1 = $row1 + $Rows + 1
	}# End of pulling data for each SBB
# Set the last row used on sheet two to the last row used by cab analysis
	$Row2 = $Row
}#End of Skip Section Logic

# Save, Close and End the Script.
if($Halt -ne "Halt"){ #Skip this Section Logic, Set by changing $Halt from "" to "Halt"
# Move to the Layout Sheet (1) and make visible
	$Sheet = $ExcelDoc.Worksheets.Item(1)
	$Sheet.Activate()
# Save the spreadsheet with this current name (LLAv163)
	$ExcelDoc.Save()
#Close the spreadsheet, kill the process, then re-open the spreadsheet
	Write-Host "Clearing out the excel Program... Standby"
	$excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) >null.txt
	kill -processname excel
	sleep -s 1
	Write-Host "Launching a clean copy of Excel"
# Start MSExcel program
	$Excel = New-Object -ComObject Excel.Application
# Load LSM Log Spreadsheet into MSExcel
	$ExcelDoc = $excel.Workbooks.Open($LSMLogOutput)
# Make Excel Visible
	$Excel.Visible = $True
# Make Sheet 1 the active sheet in the LSM Log we're working on
	$Sheet = $ExcelDoc.Worksheets.Item(1)

# Thank the User and prompt to close the Powershell Window
	Write-Host "      ------------------------------------------" -ForeGroundColor Blue
	Write-Host -NoNewLine "     |" -ForegroundColor Blue
	Write-Host -NoNewLine " Log File Analysis Complete               " -ForegroundColor Green
	Write-Host "|" -ForegroundColor Blue
	Write-Host "     |                                          |" -ForeGroundColor Blue
	Write-Host -NoNewLine "     |" -ForegroundColor Blue
	Write-Host -NoNewLine " Thank you for using the LSM Log Analyzer " -ForegroundColor Cyan
	Write-Host "|" -ForegroundColor Blue
	Write-Host -NoNewLine "     |"  -ForeGroundColor Blue
	Write-Host -NoNewLine "     -Tinker .-)>                         " -ForeGroundColor Cyan
	Write-Host "|" -ForeGroundColor Blue
	Write-Host "     |                                          |" -ForeGroundColor Blue
	Write-Host -NoNewLine "     |" -ForegroundColor Blue
	Write-Host -NoNewLine " Press Enter to close PowerShell.         " -ForegroundColor Green
	Write-Host "|" -ForegroundColor Blue
	Write-Host "      ------------------------------------------" -ForeGroundcolor Blue
	Read-Host
}#End of Skip Section Logic
