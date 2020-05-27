<#	PowerShell file that runs on a Server 2012 machine (sstgkendba10.shapetechnologies.com).
	This is set up as a repeating task occuring every 5 minutes. Looks through all the CMM files and uploads new ones into the SQL database.
	*********** VERSION HISTORY **************
	v1.0 - 9/16/2018	Initial release for production
	v1.1 - 10/20/2018	Added excel loading for final and MRB
#>
Clear-Host

$dataSource = 		"PRODSQLAPP01\PRODSQLAPP01"
$SQLDBName = 		"CMM_Repository"
$DataSet = 			New-Object System.Data.DataSet
$SourceLocation =	"G:\Flow\Operations\Seattle\Quality\Contract Cutting\"
<# $FolderLocations =	"$($script:SourceLocation)CMM Results\LPT5_WJ1 RESULTS\060053-1\",
					"$($script:SourceLocation)CMM Results\LPT5_WJ1 RESULTS\060053-2\" #>
$FolderLocations =	"$($script:SourceLocation)CMM Results\LPT5_WJ1 RESULTS\MRB Inventory Scrap Re-run\Blades that Passed_Text Files\",
					"$($script:SourceLocation)CMM Results\LPT5_WJ1 RESULTS\MRB Inventory Scrap Re-run\Blades that Failed_Text Files\",
					"$($script:SourceLocation)Salvino\2019-6-18 Rerun Blades\",
					"$($script:SourceLocation)Salvino\2019-6-20 Rerun Blades\" 
$SNString = 		"|"
$finalDate = 		[datetime]"1/1/2000"
$ArchiveDate = 		[datetime]"05/13/2019"
$SNString = 		"|"
$FinalSNString = 	"|"
$sqlInsertString = 	"INSERT INTO [50_Final] ([Blade S/N], [Blade Inspected Date], [Accepted Y/N], [Final Insp Inspector Last Name], [Comments]) VALUES"
$sqlMRBString = 	"INSERT INTO [40_MRB] ([Serial Number], [Location]) VALUES"
$minTol =			40.795, 		40.795, 	155.1, 	155.1, 	168, 	168,  	155.1, 	155.1, 	26.7, 	26.7, 	16.55,  16.55,  32,  	32,  	-0.7,  	 	-0.7,  		-0.7,  		-0.7
$maxTol =			41.805, 		41.805, 	156.5, 	156.5,  169.4,  169.4,  156.5,  156.5,  28.1,  	28.1,  	17.95,  17.95,  99.99,  99.99,   0.7,   	 0.7,   	 0.7,  		 0.7

Function Access_SQL($SqlQuery) {
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = "Server = $dataSource; Database = $SQLDBName; Integrated Security = SSPI;"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	$SqlCmd.CommandText = $SqlQuery
	$SqlCmd.Connection = $SqlConnection
	try {
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCmd
		$SqlAdapter.Fill($DataSet)
		return $true
	}
	catch {
		return $false
	}
} #End Function  Access_SQL
Function Upload_SQL($SqlArray) {
	$SQLDBName = "CMM_Repository"
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = "Server = $dataSource; Database = $SQLDBName; Integrated Security = SSPI;"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	forEach ($SQLString in $SQLArray) {
		$SqlCmd.CommandText = $SQLString
		$SqlCmd.Connection = $SqlConnection
		try {
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$SqlAdapter.Fill($DataSet)
		}
		catch {
			return $false
		}
	}
	return $true
} #End Function  Access_SQL
Function SQL_Connect {
	#"Select [File Name] From [40_CMM_LPT5];"
	$SqlQuery = "Select [File Name] FROM [40_CMM_LPT5] WHERE [Date] >= '$($ArchiveDate)';"
	$AccessLoaded = Access_SQL($SqlQuery)
	ForEach ($row in $DataSet.Tables[0]) {$script:SNString += "$($row[0])|"}
	
	$script:DataSet = New-Object System.Data.DataSet
	$SqlQuery = "Select Top 1 [Blade Inspected Date] From [50_Final] Order By [Blade Inspected Date] Desc;"
	$AccessLoaded = Access_SQL($SqlQuery)
	ForEach ($row in $DataSet.Tables[0]) {$script:finalDate = "$($row[0].AddDays(-6).ToString('MM/dd/yyy'))"}
	write-host $finalDate
	$script:DataSet = New-Object System.Data.DataSet
	$SqlQuery = "Select [Blade S/N] From [50_Final] Where [Blade Inspected Date] >= '$($finalDate)';"
	$AccessLoaded = Access_SQL($SqlQuery)
	ForEach ($row in $DataSet.Tables[0]) {
		$script:FinalSNString += "$($row[0])|"
	}
		
	If ($AccessLoaded) {
		return $true
	} Else {
		return $false
	}
} # End Function SQL_Connect
Function Start_Search{
	$SQLArray = @()
	$newSQL = $false
	If (SQL_Connect) {
		$filetype = "*.txt"
		Foreach ($filepath in $script:FolderLocations) {
			cmd.exe /c dir "$($filepath)$($filetype)" /a:-d /b | foreach {
				if ($script:SNString.IndexOf("|$($_)|") -lt 0) {
					$SQLString = DisplayTextFile("$($filepath)$($_)", $_)
					if ($SQLString -ne $false) {
						If (Upload_SQL($SQLString)) {
							write-host "File uploaded: $($filepath)$($_)"
						} Else {
							write-host ""
							write-host "Failure found in: $($filepath)$($_)"
							write-host $SQLString
							write-host ""
						}
					}
				}
			}
		}
		#if ($newSQL -eq $true) {Upload_SQL($SQLArray)}
	}
} # End Function Start_Search
Function DisplayTextFile($varIN) {
	$strFilePath = $varIN[0]
	$myfile = $varIN[1]
	Set-Variable ForReading -option Constant -value 1
	
	Get-Content $strFilePath | ForEach-Object {
		$fileString += "$($_)$([char]10)"
	}
	$fileString = $fileString.ToUpper()
	
	$NewLineSearch = "DIM"
	$NewLine = $fileString.IndexOf($NewLineSearch)
	If ($NewLine -lt 0) {return $false}
	
	$headString = $fileString.SubString(0, $NewLine - 1)
	$TrimStr = $fileString.SubString($NewLine, $fileString.length - $NewLine - 1).trim()
	
	$DisplayTextFile = "NO SQL"
    $StringQueryMid = ") VALUES ("
    $StringQuerySuf = "); "

	$DateSearch = "DATE="
	$TimeSearch = "TIME="
	$RevSearch = "REV NUMBER :"
	$SNSearch = "<SERIALNUMBER="
	$PartSearch = "<PARTNUMBER="
	$OperSearch = "<OPERATOR="
	
	
	$DateString = $headString.SubString($headString.IndexOf($DateSearch) + $DateSearch.Length, $headString.IndexOf($TimeSearch) - $headString.IndexOf($DateSearch) - $DateSearch.Length).Trim().Trim(">")
	$TimeString = $headString.SubString($headString.IndexOf($TimeSearch) + $TimeSearch.Length, $headString.IndexOf("PART NAME") -   $headString.IndexOf($TimeSearch) - $TimeSearch.Length).Trim().Trim(">")
	$RevString = $headString.SubString($headString.IndexOf($RevSearch) + $RevSearch.Length, $headString.IndexOf("SER NUMBER") -   $headString.IndexOf($RevSearch) - $RevSearch.Length).Trim().Trim(">")
	$SerialNumber = $headString.SubString($headString.IndexOf($SNSearch) + $SNSearch.Length, $headString.IndexOf("<RUNNUMBER") -   $headString.IndexOf($SNSearch) - $SNSearch.Length).Trim().Trim(">")
	$PartNumber = $headString.SubString($headString.IndexOf($PartSearch) + $PartSearch.Length, $headString.IndexOf($OperSearch) -   $headString.IndexOf($PartSearch) - $PartSearch.Length).Trim().Trim(">")
	$Operator = $headString.SubString($headString.IndexOf($OperSearch) + $OperSearch.Length, $headString.IndexOf("<MACHINE") - $headString.IndexOf($OperSearch) - $OperSearch.Length).Trim().Trim(">")
	
	$StringQueryPre = "INSERT INTO [40_CMM_LPT5] ([Serial Number], [File Name], [Date], [Revision], [Part Number], [Operator]"
	$StringQueryMid += "'$($SerialNumber)', '$($myfile)', '$($DateString) $($TimeString)', '$($RevString)', '$($PartNumber)', '$($Operator)'"

	While ($TrimStr -match $NewLineSearch) {
		$NewDimLine = $TrimStr.IndexOf($NewLineSearch)
		$TrimStr = $TrimStr.SubString($NewDimLine, $TrimStr.length - $NewDimLine).trim()
			
		$nextLineStart = GetNewLine($TrimStr)
		$LineString = $TrimStr.SubString(0, $nextLineStart).trim()
		$DimString = TrimDim($LineString)
		$MinMax = $false
		Do {
			$TrimStr = $TrimStr.SubString($nextLineStart, $TrimStr.length - $nextLineStart).trim()
			$nextLineStart = GetNewLine($TrimStr)
			If ($nextLineStart -eq 0) {$nextLineStart = $TrimStr.length}
			If (($TrimStr -match "^AX") -and ($TrimStr.SubString(0, $nextLineStart).IndexOf("MAX") -ge 0)) {
				$MinMax = $true
			}
		} Until (($TrimStr -match "^M") -or ($TrimStr -match "^T"))
		If ($TrimStr -match "^T") {$DimString = $DimString -replace "DIM", "T"}
		If ($MinMax) {
			$DimValue = GetDimValue($TrimStr.SubString(1, $nextLineStart).Trim())
			$MinDimValue = GetMinDimValue($TrimStr.SubString(1, $nextLineStart).Trim())
			$StringQueryPre += ", [$($DimString) Max], [$($DimString) Min]"
			$StringQueryMid += ", $($DimValue), $($MinDimValue)"
		} Else {
			If ($nextLineStart -eq $TrimStr.Length) {$nextLineStart -= 1}
			$DimValue = GetDimValue($TrimStr.SubString(1, $nextLineStart).Trim())
			$StringQueryPre += ", [$($DimString)]"
			$StringQueryMid += ", $($DimValue)"
		}
		$TrimStr = $TrimStr.SubString($nextLineStart, $TrimStr.length - $nextLineStart).trim()
	}	
	return "$($StringQueryPre)$($StringQueryMid)$($StringQuerySuf)"
} #End Function DisplayTextFile
Function GetNewLine($TextString) {
    $GetNewLine = $TextString.length
    #010     new line
    #011     vertical tab
    #012     new page
    #013     carriage return
    For ($n = 10; $n -le 13; $n++) {
        If ($TextString -match "$([char]$n)") {
            If ($GetNewLine -gt $TextString.IndexOf("$([char]$n)")) {
				$GetNewLine = $TextString.IndexOf("$([char]$n)")
			}
        }
    }
	return $GetNewLine
} #End Function GetNewLine
Function TrimDim($TXTString) {
                          # (			# )				# =			# >
    $SearchArray = "$([char]40)", "$([char]41)", "$([char]61)", "$([char]62)"
    ForEach ($SearchElement In $SearchArray) {
        While ($TXTString.IndexOf($SearchElement) -ge 0) {
            $TXTString = $TXTString.SubString(0, $TXTString.IndexOf($SearchElement))
        }
    }
    $TrimDim = $TXTString.Replace("$([char]46)", "$([char]95)")			#Replace . with _
    $TrimDim = $TrimDim.Replace("$([char]60)", "")			# Replace < with null
	return $TrimDim
} #End Function TrimDim
Function GetDimValue($TXTString) {
    $TXTString = $TXTString.SubString(0, $TXTString.IndexOf("$([char]32)"))
    return [float]::Parse($TXTString)
} #End Function GetDimValue
Function GetMinDimValue($TXTString) {
	# [char]32 = Space
    $TXTString = $TXTString.SubString($TXTString.IndexOf("$([char]32)"), $TXTString.Length - $TXTString.IndexOf("$([char]32)")).trim()
	$TXTString = $TXTString.SubString(0, $TXTString.IndexOf("$([char]32)"))
	return [float]::Parse($TXTString)
} #End Function GetMinDimValue
Function Import-Xls { 
 
<# 
.SYNOPSIS 
Import an Excel file. 
 
.DESCRIPTION 
Import an excel file. Since Excel files can have multiple worksheets, you can specify the worksheet you want to import. You can specify it by number (1, 2, 3) or by name (Sheet1, Sheet2, Sheet3). Imports Worksheet 1 by default. 
 
.PARAMETER Path 
Specifies the path to the Excel file to import. You can also pipe a path to Import-Xls. 
 
.PARAMETER Worksheet 
Specifies the worksheet to import in the Excel file. You can specify it by name or by number. The default is 1. 
Note: Charts don't count as worksheets, so they don't affect the Worksheet numbers. 
 
.INPUTS 
System.String 
 
.OUTPUTS 
Object 
 
.EXAMPLE 
".\employees.xlsx" | Import-Xls -Worksheet 1 
Import Worksheet 1 from employees.xlsx 
 
.EXAMPLE 
".\employees.xlsx" | Import-Xls -Worksheet "Sheet2" 
Import Worksheet "Sheet2" from employees.xlsx 
 
.EXAMPLE 
".\deptA.xslx", ".\deptB.xlsx" | Import-Xls -Worksheet 3 
Import Worksheet 3 from deptA.xlsx and deptB.xlsx. 
Make sure that the worksheets have the same headers, or have some headers in common, or that it works the way you expect. 
 
.EXAMPLE 
Get-ChildItem *.xlsx | Import-Xls -Worksheet "Employees" 
Import Worksheet "Employees" from all .xlsx files in the current directory. 
Make sure that the worksheets have the same headers, or have some headers in common, or that it works the way you expect. 
 
.LINK 
Import-Xls 
http://gallery.technet.microsoft.com/scriptcenter/17bcabe7-322a-43d3-9a27-f3f96618c74b 
Export-Xls 
http://gallery.technet.microsoft.com/scriptcenter/d41565f1-37ef-43cb-9462-a08cd5a610e2 
Import-Csv 
Export-Csv 
 
.NOTES 
Author: Francis de la Cerna 
Created: 2011-03-27 
Modified: 2011-04-09 
#Requires –Version 2.0 
#> 
 
    [CmdletBinding(SupportsShouldProcess=$true)] 
     
    Param( 
        [parameter( 
            mandatory=$true,  
            position=1,  
            ValueFromPipeline=$true,  
            ValueFromPipelineByPropertyName=$true)] 
        [String[]] 
        $Path, 
     
        [parameter(mandatory=$false)] 
        $Worksheet = 1, 
         
        [parameter(mandatory=$false)] 
        [switch] 
        $Force 
    ) 
 
    Begin 
    { 
        function GetTempFileName($extension) 
        { 
            $temp = [io.path]::GetTempFileName(); 
            $params = @{ 
                Path = $temp; 
                Destination = $temp + $extension; 
                Confirm = $false; 
                Verbose = $VerbosePreference; 
            } 
            Move-Item @params; 
            $temp += $extension; 
            return $temp; 
        } 
             
        # since an extension like .xls can have multiple formats, this 
        # will need to be changed 
        # 
        $xlFileFormats = @{ 
            # single worksheet formats 
            '.csv'  = 6;        # 6, 22, 23, 24 
            '.dbf'  = 11;       # 7, 8, 11 
            '.dif'  = 9;        #  
            '.prn'  = 36;       #  
            '.slk'  = 2;        # 2, 10 
            '.wk1'  = 31;       # 5, 30, 31 
            '.wk3'  = 32;       # 15, 32 
            '.wk4'  = 38;       #  
            '.wks'  = 4;        #  
            '.xlw'  = 35;       #  
             
            # multiple worksheet formats 
            '.xls'  = -4143;    # -4143, 1, 16, 18, 29, 33, 39, 43 
            '.xlsb' = 50;       # 
            '.xlsm' = 52;       # 
            '.xlsx' = 51;       # 
            '.xml'  = 46;       # 
            '.ods'  = 60;       # 
        } 
         
        $xl = New-Object -ComObject Excel.Application; 
        $xl.DisplayAlerts = $false; 
        $xl.Visible = $false; 
    } 
 
    Process 
    { 
        $Path | ForEach-Object { 
             
            if ($Force -or $psCmdlet.ShouldProcess($_)) { 
             
                $fileExist = Test-Path $_ 
 
                if (-not $fileExist) { 
                    Write-Error "Error: $_ does not exist" -Category ResourceUnavailable;             
                } else { 
                    # create temporary .csv file from excel file and import .csv 
                    # 
                    $_ = (Resolve-Path $_).toString(); 
                    $wb = $xl.Workbooks.Add($_); 
                    if ($?) { 
                        $csvTemp = GetTempFileName(".csv"); 
                        $ws = $wb.Worksheets.Item($Worksheet); 
                        $ws.SaveAs($csvTemp, $xlFileFormats[".csv"]); 
                        $wb.Close($false); 
                        Remove-Variable -Name ('ws', 'wb') -Confirm:$false; 
                        Import-Csv $csvTemp; 
                        Remove-Item $csvTemp -Confirm:$false -Verbose:$VerbosePreference; 
                    } 
                } 
            } 
        } 
    } 
    
    End 
    { 
        $xl.Quit(); 
        Remove-Variable -name xl -Confirm:$false; 
        [gc]::Collect();
    } 
} 
Function Update_CMM_Failures{
	$Total = 0
	$sqlQuery =		"SELECT [40_CMM_LPT5].[File Name], "
	$sqlQuery += 	"[40_CMM_LPT5].[Dim 1_1], [40_CMM_LPT5].[Dim 1_2], [40_CMM_LPT5].[Dim 2_1], [40_CMM_LPT5].[Dim 2_2], [40_CMM_LPT5].[Dim 3_1], [40_CMM_LPT5].[Dim 3_2], "
	$sqlQuery += 	"[40_CMM_LPT5].[Dim 4_1], [40_CMM_LPT5].[Dim 4_2], [40_CMM_LPT5].[Dim 5_1], [40_CMM_LPT5].[Dim 5_2], [40_CMM_LPT5].[Dim 9_1], [40_CMM_LPT5].[Dim 9_2], "
	$sqlQuery += 	"[40_CMM_LPT5].[Dim 10_1], [40_CMM_LPT5].[Dim 10_2], [40_CMM_LPT5].[Dim 11 Max], [40_CMM_LPT5].[Dim 11 Min], [40_CMM_LPT5].[Dim 12 Max], [40_CMM_LPT5].[Dim 12 Min] "
	$sqlQuery += 	"FROM [40_CMM_LPT5] WHERE [Failures] IS NULL;"
	$script:DataSet = New-Object System.Data.DataSet
	$sqlUpdate = New-Object System.Collections.ArrayList
	$AccessLoaded = Access_SQL($SqlQuery)
	ForEach ($row in $DataSet.Tables[0]) {
	$fileName = $row[0]
		$toleranceArray = New-Object System.Collections.ArrayList
		$n = 1
		ForEach ($tol in $minTol) {
			$toleranceArray.Add("$($row[$n])") > $null
			$n++
		}
		$sqlUpdate.Add("UPDATE [40_CMM_LPT5] SET [Failures]='$(toleranceCheck($toleranceArray))' WHERE  [File Name]='$fileName';") > $null
	}
	ForEach ($sql in $sqlUpdate) {$AccessLoaded = Access_SQL($sql)}
}
Function toleranceCheck($toleranceArray) {
	$toleranceCheck = 0
	$colX = 0
	ForEach ($tol in $toleranceArray) {
		If (!$tol) {
		} ElseIf ([math]::Round([float]$tol,3) -lt $minTol[$colX] -or [math]::Round([float]$tol,3) -gt $maxTol[$colX]){
			$toleranceCheck += 1
			#write-host "$($minTol[$colX])|$($tol)|$($maxTol[$colX])"
		}
		$colX += 1
	}
	return $toleranceCheck
}
If (SQL_Connect -eq $true) {
	write-host "start CMM file search"
	Start_Search
	write-host "start CMM failures"
	Update_CMM_Failures
	write-host "CMM failures complete"
}