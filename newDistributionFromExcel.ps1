<#

    .SYNOPSIS
    Creates Exchange distribution groups from data in an Excel spreadsheet

    .DESCRIPTION
    Reads in data from an Excel sheet and creates Excel distribution groups using the data provided in the spreadsheet.

    .INPUTS
    Accepts an Excel file as an input.

    .OUTPUTS
    Does not generate output.


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

        #Connect-ExchangeOnline -UserPrincipalName $env:username@electricalbreakdown.com -ShowBanner:$false 
        Connect-ExchangeOnline -UserPrincipalName mike@electricalbreakdown.onmicrosoft.com -ShowBanner:$false 
       
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



$browser = New-Object System.Windows.Forms.OpenFileDialog
$browser.Filter = "Excel Files (*.xlsx; *xls) | *.xlsx; *xls"
$browser.Title = "Please Select an Excel File"


# Make sure a file is selcted 

if ($browser.ShowDialog() -eq "OK") {

    $filePath = $browser.FileName 

    $excelObject = New-Object -ComObject Excel.Application
    $workbook = $excelObject.Workbooks.Open($filePath)
    $sourceSheet = $workbook.Worksheets.Item(1)
    $dataRange = $sourceSheet.UsedRange     

    # $emailCol = ($dataRange.Rows.Find("Email Address")).Column
    # $displayNameCol = ($dataRange.Rows.Find("Display Name")).Column
    
    #$groupMembers = @()
    $numCreated = 0
    $numNotCreated = 0
    $groupsNotCreated = @()

    for($row = 2; $row -le $dataRange.Rows.Count; $row ++){

        $groupAddress = $sourceSheet.Cells($row, 1).Text
        $groupName = $sourceSheet.Cells($row, 2).Text

        <#
        for($col = 3; $col -le $dataRange.Columns.Count; $col ++){
            
            $groupMember = $sourceSheet.Cells($row, $col).Text
            

            if(![string]::IsNullOrWhitespace($groupMember)){

                $groupMembers += $groupMember
            }

        }
        #>
        try {

            $numCreated += 1
            New-DistributionGroup -Name $groupName -PrimarySMTPAddress $groupAddress -RequireSenderAuthenticationEnabled $false -MemberJoinRestriction "Closed" -MemberDepartRestriction "Closed"
                        
        }
        catch {

            $numNotCreated += 1
            $groupsNotCreated += $groupAddress
            throw
        }

        try {

            Set-DistributionGroup -Identity $groupAddress -HiddenFromAddressListsEnabled $true

        }

        catch {

            Write-Host "Couldn't hide from address list" -ForegroundColor Red
            $err[0]
        }

        
    } # Cloose outer for loop



    # Close workbook once all groups have been created

    $workbook.Close()
    $excelObject.Quit()

    Write-Host "$numCreated groups were created succesfully." -ForegroundColor Green
    Write-Host "`n-------------------------------------------------------------------`n"
    Write-Host "$numNotCreated groups could not be created." -ForegroundColor Red
    $groupsNotCreated

}

# Display a message if the file browser was closed without selecting a file

else {


    Write-Host "Operation Cancelled" -ForegroundColor Red

}



Read-Host "Press any key to exit"