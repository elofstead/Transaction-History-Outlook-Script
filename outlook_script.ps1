<# This script uses COM objects to access transaction notification emails from my bank and 
   compiles them into an excel spreadsheet. This is my first time using powershell and my primary
   objective is to learn some of the basics of powershell scripting while also helping me track my
   monthly expenditures. 
   
   Note that this script requires Excel and the Microsoft office version of Outlook installed. #>


<# Create the outlook COM object and login to the default profile.
   Note that you will need to configure a default outlook profile in the mail settings in control panel#>
Add-Type -assembly "Microsoft.Office.Interop.Outlook" | Out-Null
$outlook = new-object -comobject outlook.application
$namespace = $outlook.GetNameSpace("MAPI")
$namespace.Logon($null, $null, $false, $true) 


# Create the excel COM object and configure column headers
$excel = New-Object -ComObject excel.application
$excel.visible = $true
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)
$worksheet.Cells.Item(1,1) = 'Date'
$worksheet.Cells.Item(1,2) = 'Amount'
$worksheet.Cells.Item(1,3) = 'Monthly Total'


# Access the folder the transaction notifications are stored in and retrieve the messages
$transactionFolder = $namespace.Folders(1).Folders(12)
$messages = $transactionFolder.items


# Iterate through the messages, extract the amount spent and the date recieved, and add to the spreadsheet
$i = 2
foreach($message in $messages){ 
	$amount = $message.Body | Select-String -Pattern "\$\d*\.\d*" | %{ $_.Matches.Value }
	$date = $message.ReceivedTime | Get-Date -Format "yyyy-MM-dd"
	$worksheet.Cells.Item($i,1) = $date
	$worksheet.Cells.Item($i,2) = $amount
	$i++
}


# Logoff and close the Outlook COM object
$namespace.Logoff()
$outlook.Quit()


# Sort by date and format the spreadsheet
$emptyVar = [System.Type]::Missing
$sortCol = $worksheet.Range("A2:A"+$i)
$worksheet.UsedRange.Sort($sortCol,2,$emptyVar,$emptyVar,$emptyVar,$emptyVar,$emptyVar,1) | Out-Null
$worksheet.Range("A:C").EntireColumn.AutoFit() | Out-Null


# Loop through the sorted spreadsheet to find the ranges that coorespond to each month and add a formula to get the monthly totals
$startRow = 2
$currentRow = 3
$currentMonth = $worksheet.Cells.Item(2,1).Text | Select-String -Pattern "\d*" | %{ $_.Matches.Value }
while($true){

	# checks if the current date cell is empty, and if so we are at the end of the spreadsheet, so we add the final formula and break the loop
	if ([string]::IsNullOrWhiteSpace($worksheet.Cells.Item($currentRow,1).text)) { 
		$worksheet.Cells($startRow,3).Formula = "=SUM(B" + $startRow + ":B" + ($currentRow - 1) + ")"
		break;
	}

	# when this condition is true we are at the end of the range for a particular month, so we add the formula for that month and update the 
	# currentMonth variable to be the next month
	if ($currentMonth -ne ($worksheet.Cells.Item($currentRow,1).Text | Select-String -Pattern "\d*" | %{ $_.Matches.Value } ) ) {
		$worksheet.Cells($startRow,3).Formula = "=SUM(B" + $startRow + ":B" + ($currentRow - 1) + ")"
		$currentMonth = $worksheet.Cells.Item($currentRow,1).Text | Select-String -Pattern "\d*" | %{ $_.Matches.Value }
		$startRow = $currentRow
	}

	#increment counter
	$currentRow++
}

# Save excel file and quit excel COM object
$workbook.SaveAs("C:\Users\USERNAME\Desktop\Transactions.xlsx") #update this line to reflect your actual file structure/desired location

$excel.Quit()

