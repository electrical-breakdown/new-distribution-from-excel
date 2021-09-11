<#

    .SYNOPSIS
    Creates Exchange distribution groups from data in an Excel spreadsheet

    .DESCRIPTION
    Reads in data from an Excel sheet and creates Exchange distribution groups using the data provided in the spreadsheet. 
    If there are any issues creating a group, that group will not be created by the script and must be created manually.

    .INPUTS
    Accepts an Excel file as an input.

    .OUTPUTS
    Outputs an Excel file with a status update and details for each group.

#>

#-------------------------------------- Connect to Exchange Online --------------------------------------#

try {             

    # Check to see if there is already an active Exchange Online session before connecting a new session
    
    $exchangeSession = Get-PSSession | Where-Object {$_.Name -like 'ExchangeOnlineInternalSession*' -and $_.State -eq 'Opened'}


    if ($null -eq $exchangeSession) {
        
        Write-Verbose "Connecting to Exchange Online. This might take a moment..." -Verbose

        
        # Force TLS 1.2 encryption for compatibility with PowerShell Gallery

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12



        # Install Exchange Online module if it isn't already installed

        if ($null -eq (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {

            Write-Verbose "Installing Exchange Online module..." -Verbose
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -Confirm 

        }    
                        
        # Connect to Exchange if the module is installed and there are no active sessions

        # Connect-ExchangeOnline -UserPrincipalName $env:username@electricalbreakdown.com -ShowBanner:$false 
        Connect-ExchangeOnline -UserPrincipalName mike@electricalbreakdown.com -ShowBanner:$false 
        
        Write-Verbose "Connected to Exhange Online." -Verbose
    }   
} 

# Halt script execution if the connection to Exchange fails

catch {

    Write-Host "There was a problem connecting to Exchange. Please reload the script and try again`n" -ForegroundColor Red        
    throw

}

#---------------------------- Load assemblies and generate file browser ---------------------------------#


Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms


$errorActionPreference = "Stop"

$browser = New-Object System.Windows.Forms.OpenFileDialog
$browser.Filter = "Excel Files (*.xlsx; *xls) | *.xlsx; *xls"
$browser.Title = "Please Select an Excel File"


#---------------------------------- Open workbook and define variables  --------------------------------------#

# Make sure a file is selected 

if ($browser.ShowDialog() -eq "OK") {

    $filePath = $browser.FileName 
    $currentFolder = Split-Path $filePath
    $currentFileName = (Split-Path $filePath -Leaf).Split('.') 
    $exportFilePath = "$($currentFolder)\$($currentFileName[0])_processed.$($currentFileName[-1])"

    Write-Verbose "Saving new copy of workbook and processing data..."  -Verbose

    $excelObject = New-Object -ComObject Excel.Application
    $workbook = $excelObject.Workbooks.Open($filePath)
    $workbook.SaveAs($exportFilePath)

    $sourceSheet = $workbook.Worksheets.Item(1)
    $dataRange = $sourceSheet.UsedRange   

    # Find the last column in the used data range

    $lastCol = $dataRange.SpecialCells(11).Column

    # The column after the last column will be used for a status message

    $statusCol = $lastCol + 1
    $sourceSheet.Cells(1, $statusCol).Value2 = "Status"

    # The column after the status colum will be for notes

    $detailsCol = $lastCol + 2
    $sourceSheet.Cells(1, $detailsCol).Value2 = "Details"
    
    # Define color codes to be used in the exported Excel file

    $successBGColor = [System.Drawing.Color]::FromArgb(169, 208, 142)
    $warningBGColor = [System.Drawing.Color]::FromArgb(255, 217, 102)
    $dangerBGColor = [System.Drawing.Color]::FromArgb(255, 121, 121)

    # Create empty arrays to hold various data

    $groupsCreated = @()
    $groupsNotCreated = @()    
    
    
    #---------------------------------- Process sheet and create groups  --------------------------------------#

    # Iterate through each row in the sheeet. Start at 2 because the first row contains column headers

    for($row = 2; $row -le $dataRange.Rows.Count; $row ++){

        $error.Clear()
        $issuesFound = ""


        $groupAddress = $sourceSheet.Cells($row, 1).Text
        $groupName = $sourceSheet.Cells($row, 2).Text
        $groupOwner = $sourceSheet.Cells($row, 3).Text

        $statusCell = $sourceSheet.Cells($row, $statusCol)
        $detailsCell = $sourceSheet.Cells($row, $detailsCol)

        $groupMembers = @()

        # Loop through remaining columns and gather up all the group members
        
        for($col = 4; $col -le $dataRange.Columns.Count; $col ++){
            
            $member = $sourceSheet.Cells($row, $col).Text

            if(![string]::IsNullOrWhitespace($member)){
                $groupMembers += $member
            }
        }   
        
        
        # Create new distribution group and add members
        
        try {

            Write-Verbose "Creating group $($groupAddress)..." -Verbose            
            
            New-DistributionGroup -Name $groupName `
                -PrimarySMTPAddress $groupAddress `
                -RequireSenderAuthenticationEnabled $false `
                -MemberJoinRestriction "Closed" `
                -MemberDepartRestriction "Closed" `
                -ModeratedBy $groupOwner `
                -Members $groupMembers `
                | Out-Null
                
            $groupsCreated += $groupAddress            
            $groupCreated = $true     
            $statusCell.Value2 = "Created successfully"
            $detailsCell.Value2 = "Group was created and all members were added succesfully."
            $statusCell.EntireRow.Interior.Color = $successBGColor                     
        }
        catch {
            
            Write-Warning "There was a problem creating group $groupAddress" 

            $groupsNotCreated += $groupAddress    
            $issuesFound = $error[0].Exception.Message                         
            $groupCreated = $false            

        }
        
        if($groupCreated -eq $true){            
            
            # Hide group from global address list 

            try {
                
                Set-DistributionGroup -Identity $groupAddress -HiddenFromAddressListsEnabled $true

            }

            catch {

                Write-Warning "Couldn't hide $groupAddress from address list" 
                $issuesFound = $error[0].Exception.Message

                $statusCell.Value2 = "Created - With Issues"        
                $detailsCell.Value2 = $issuesFound                                
                $detailsCell.EntireRow.Interior.Color = $warningBGColor     
            }      
            
            
        } # Close if block

        # If group was not created...

        else {

            $statusCell.Value2 = "Not Created"               
           
            $detailsCell.Value2 += $issuesFound
            $detailsCell.EntireRow.Interior.Color = $dangerBGColor 

        }                   

        Write-Host "`n-------------------------------------------------------------------`n"   


    } # Cloose outer for loop


    #---------------------------- Close Excel file and output status messages  ---------------------------------#
    
    # Select the columns containing new data and resize them        

    $sourceSheet.Cells(1, $statusCol).EntireColumn.AutoFit() | Out-Null
    $sourceSheet.Cells(1, $detailsCol).EntireColumn.AutoFit() | Out-Null

    # Save and close workbook once all groups have been created

    $workbook.Save()  
    $workbook.Close()
    $excelObject.Quit()
    
    Clear-Host   
   
    # Display the groups that were created successfuly 

    Write-Host "`n-------------------------------------------------------------------`n"
    Write-Host "`n$($groupsCreated.Count) groups were created succesfully:`n" -ForegroundColor Green
    $groupsCreated
    Write-Host "`n-------------------------------------------------------------------`n"


    # Display a message if some groups could not be created

    if($groupsNotCreated){

        Write-Host "$($groupsNotCreated.Count) groups could not be created. Please review and create manually:`n" -ForegroundColor Red
        $groupsNotCreated
        Write-Host "`n-------------------------------------------------------------------`n"

    }
       
    
    Write-Host "Detailed results have been exported to: $exportFilePath" -ForegroundColor Yellow
    Write-Host "`n-------------------------------------------------------------------`n"


} #Close main if block

# Display a message if the file browser was closed without selecting a file

else {

    Write-Host "Operation Cancelled" -ForegroundColor Red

}


Read-Host "Press any key to exit"