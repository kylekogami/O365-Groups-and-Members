Import-Module PSExcel
Import-Module AzureADPreview
mkdir C:\scripts

#Connect to Exchange Online
Connect-ExchangeOnline -ShowBanner:$False




#MICROSOFT365 GROUPS

$O365GroupMembers = "C:\scripts\O365GroupMembers.xlsx"

#Remove the XLSX file if exists
If(Test-Path $O365GroupMembers) { Remove-Item $O365GroupMembers}

$M365 = New-Object -ComObject excel.application
$workbook = $M365.Workbooks.Add(1)

#Get All Office 365 Groups
$O365Groups=Get-UnifiedGroup
$i=0

$Workbook = $M365.Workbooks.Open($O365GroupMembers)

ForEach ($Group in $O365Groups) 
{ 
    $i++
    $CSVPath = "c:\scripts\Group" + $i.ToString() + ".csv"

    #Remove the CSV file if exists
    If(Test-Path $CSVPath) { Remove-Item $CSVPath}

    Write-Host "Group Name:" $Group.DisplayName -ForegroundColor Green
    Get-UnifiedGroupLinks -Identity $Group.Id -LinkType Members | Select DisplayName,PrimarySmtpAddress
 
    #Get Group Members and export to CSV
    Get-UnifiedGroupLinks -Identity $Group.Id -LinkType Members | Select-Object @{Name="Group Name";Expression={$Group.DisplayName}},`
         @{Name="User Name";Expression={$_.DisplayName}},`
         @{Name="Email Address";Expression={$_.PrimarySmtpAddress}} | Export-CSV $CSVPath -NoTypeInformation -Append


    #$xlsxPath = "c:\scripts\Group" + $i.ToString() + ".xlsx"
    #$csv = Import-Csv $CSVPath
    #$csv | Export-Excel -Path $xlsxPath -AutoSize
    #Remove-Item $CSVPath

    
    $worksheet = $Workbook.worksheets.add()
    $worksheet.name = $Group.DisplayName
    
}

$j=0
while ($j -le $i) {
    $j++
    $CSVPath = "c:\scripts\Group" + $j.ToString() + ".csv"
    $worksheet = $workbook.worksheets.Item(($i+1)-$j)
    $TxtConnector = ("TEXT;" + $CSVPath)
    $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
    $query = $worksheet.QueryTables.item($Connector.name)
    $query.TextFileOtherDelimiter = $Excel.Application.International(5)
    $query.TextFileParseType  = 1
    $query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1
    $query.Refresh()
    $query.Delete()

    #Remove the already copied CSV file
    If(Test-Path $CSVPath) { Remove-Item $CSVPath}
} 


$Workbook.SaveAs($O365GroupMembers,51)  

$M365.Quit()




#DISRIBUTION LIST

$O365DistMembers = "C:\scripts\O365DistributionListMembers.xlsx"

#Remove the XLSX file if exists
If(Test-Path $O365DistMembers) { Remove-Item $O365DistMembers}

$DistList = New-Object -ComObject excel.application
$workbook = $DistList.Workbooks.Add(1)

#Get All Office 365 Groups
$O365Groups=Get-DistributionGroup
$i=0

$Workbook = $DistList.Workbooks.Open($O365DistMembers)

ForEach ($Group in $O365Groups) 
{ 
    $i++
    $CSVPath = "c:\scripts\Group" + $i.ToString() + ".csv"

    #Remove the CSV file if exists
    If(Test-Path $CSVPath) { Remove-Item $CSVPath}

    Write-Host "Distribution List Name:" $Group.DisplayName -ForegroundColor Green
    Get-DistributionGroupMember -Identity $Group.Id | Select DisplayName,PrimarySmtpAddress
 
    #Get Group Members and export to CSV
    Get-DistributionGroupMember -Identity $Group.Id | Select-Object @{Name="Distribution List Name";Expression={$Group.DisplayName}},`
        @{Name="User Name";Expression={$_.DisplayName}},`
        @{Name="Email Address";Expression={$_.PrimarySmtpAddress}} | Export-CSV $CSVPath -NoTypeInformation -Append

    $worksheet = $Workbook.worksheets.add()
    $worksheet.name = $Group.DisplayName
    
}

$j=0
while ($j -le $i) {
    $j++
    $CSVPath = "c:\scripts\Group" + $j.ToString() + ".csv"
    $worksheet = $workbook.worksheets.Item(($i+1)-$j)
    $TxtConnector = ("TEXT;" + $CSVPath)
    $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
    $query = $worksheet.QueryTables.item($Connector.name)
    $query.TextFileOtherDelimiter = $Excel.Application.International(5)
    $query.TextFileParseType  = 1
    $query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1
    $query.Refresh()
    $query.Delete()

    #Remove the already copied CSV file
    If(Test-Path $CSVPath) { Remove-Item $CSVPath}
} 


$Workbook.SaveAs($O365DistMembers,51)  

$DistList.Quit()




#Connect to AzureAD
Connect-AzureAD
 
$O365SecurityMembers = "C:\scripts\O365SecurityMembers.xlsx"

#Remove the XLSX file if exists
If(Test-Path $O365SecurityMembers) { Remove-Item $O365SecurityMembers}

$Security = New-Object -ComObject excel.application
$workbook = $Security.Workbooks.Add(1)

#Get All Office 365 Groups
$O365Groups = Get-AzureADGroup -Filter "SecurityEnabled eq true"
$i=0

$Workbook = $Security.Workbooks.Open($O365SecurityMembers)

ForEach ($Group in $O365Groups) 
{ 
    $Group.GetType()
    $i++
    $CSVPath = "c:\scripts\Group" + $i.ToString() + ".csv"

    #Remove the CSV file if exists
    If(Test-Path $CSVPath) { Remove-Item $CSVPath}

    Write-Host "Security Group Name:" $Group.DisplayName -ForegroundColor Green
    Get-AzureADGroupMember -ObjectId $Group.ObjectId | Select DisplayName,UserPrincipalName
 
    #Get Group Members and export to CSV
    Get-AzureADGroupMember -ObjectId $Group.ObjectId | Select-Object @{Name="Security Group Name";Expression={$Group.DisplayName}}, @{Name="User Name";Expression={$_.DisplayName}}, @{Name="Email Address";Expression={$_.UserPrincipalName}}`
    | Export-CSV $CSVPath -NoTypeInformation -Append

    
    $worksheet = $Workbook.worksheets.add()
    $worksheet.name = $Group.DisplayName
    
}

$j=0
while ($j -le $i) {
    $j++
    $CSVPath = "c:\scripts\Group" + $j.ToString() + ".csv"
    $worksheet = $workbook.worksheets.Item(($i+1)-$j)
    $TxtConnector = ("TEXT;" + $CSVPath)
    $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
    $query = $worksheet.QueryTables.item($Connector.name)
    $query.TextFileOtherDelimiter = $Excel.Application.International(5)
    $query.TextFileParseType  = 1
    $query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1
    $query.Refresh()
    $query.Delete()

    #Remove the already copied CSV file
    If(Test-Path $CSVPath) { Remove-Item $CSVPath}
} 


$Workbook.SaveAs($O365SecurityMembers,51)  

$Security.Quit()




#Disconnect Exchange Online
Disconnect-ExchangeOnline -Confirm:$False