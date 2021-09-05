<#

    .SYNOPSIS
    Creates Exchange distribution groups from data in an Excel spreadsheet

    .DESCRIPTION
    Reads in data from an Excel sheet and creates Exchange distribution groups using the data provided in the spreadsheet.

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
        Connect-ExchangeOnline -UserPrincipalName mike@electricalbreakdown.com -ShowBanner:$false 
       
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


#---------------------------- Process sheet and create groups  ---------------------------------#

# Make sure a file is selected 

if ($browser.ShowDialog() -eq "OK") {

    $filePath = $browser.FileName 

    $excelObject = New-Object -ComObject Excel.Application
    $workbook = $excelObject.Workbooks.Open($filePath)
    $sourceSheet = $workbook.Worksheets.Item(1)
    $dataRange = $sourceSheet.UsedRange   
       
    $groupsCreated = @()
    $groupsNotCreated = @()
    $usersNotAdded = @()
    $summary = [ordered]@{}

    for($row = 2; $row -le $dataRange.Rows.Count; $row ++){

        $groupAddress = $sourceSheet.Cells($row, 1).Text
        $groupName = $sourceSheet.Cells($row, 2).Text

        # Create new distribution group
        
        try {

            Write-Verbose "Creating group $($groupAddress)..." -Verbose
            New-DistributionGroup -Name $groupName -PrimarySMTPAddress $groupAddress -RequireSenderAuthenticationEnabled $false -MemberJoinRestriction "Closed" -MemberDepartRestriction "Closed" | Out-Null
            $groupsCreated += $groupAddress            
            $groupCreated = $true
                        
        }
        catch {
            
            $groupsNotCreated += $groupAddress

            Write-Verbose "Couldn't create group $groupAddress" -Verbose
            
            $summary.Add($groupAddress, @($error[0].Exception.Message))
            Write-Host $error[0].Exception.Message  -ForegroundColor Red
            $groupCreated = $false
        }

        
        if($groupCreated -eq $true){
            # Hide group from address list 

            try {
                
                Set-DistributionGroup -Identity $groupAddress -HiddenFromAddressListsEnabled $true

            }

            catch {

                Write-Host "Couldn't hide from address list" -ForegroundColor Red
                $error[0]
            }


            # Add members to group

            for($col = 3; $col -le $dataRange.Columns.Count; $col ++){
                
                $groupMember = $sourceSheet.Cells($row, $col).Text
                
                # Make sure cell wasn't empty
                if(![string]::IsNullOrWhitespace($groupMember)){
                                    
                    try {

                        Write-Verbose "Adding $($groupMember)..." -Verbose
                        Add-DistributionGroupMember -Identity $groupAddress -Member $groupMember
                        # Maybe add a check to make sure user has an Exchange license
                    }
                    catch {
                        
                        if(!($groupMember -in $usersNotAdded)){

                            $usersNotAdded += $groupMember

                        }
                        
                        # Check to see if the group is already in the hash table. If it is, append the new error to the existing value array

                        if($summary.$groupAddress){

                            $summary[$groupAddress] += $error[0].Exception.Message
                            
                        }

                       

                        # If the group isn't already in the hash table, just add it
                        else {
                            
                            $summary.Add($groupAddress, @($error[0].Exception.Message))

                        }
                                        
                        Write-Host $error[0].Exception.Message  -ForegroundColor Red
                    }               

                }           

            } # Close inner for loop

            # If we get to this point and the group address isn't in '$summary' yet, there were no errors and we can add a success message        

            if(!($summary.$groupAddress)){

                $summary.Add($groupAddress, "Group created and all members added succesfully.")
            }

          
            Write-Host "`n-------------------------------------------------------------------`n"
        
        } # Close if block

    } # Cloose outer for loop


    #---------------------------- Closed Excel file and output status messages  ---------------------------------#

    # Close workbook once all groups have been created

    $workbook.Close()
    $excelObject.Quit()
    
    Clear-Host

    # Display the full summary 

    Write-Host "Summary:`n"
    
    $summary.Keys | Select-Object @{l="Group"; e={$_}}, @{l="Status"; e={$summary.$_}} | Out-Host
    
    Write-Host "-------------------------------------------------------------------`n"

    $summary.GetEnumerator() | Select-Object @{l="Group"; e={$_.Key}}, @{l="Status"; e={$_.Value}} | Export-CSV -Path "C:\Users\Administrator\Desktop\log.csv"


    # Display the groups that were created successfuly 
    
    Write-Host "`n$($groupsCreated.Count) groups were created succesfully:`n" -ForegroundColor Green
    $groupsCreated
    Write-Host "`n-------------------------------------------------------------------`n"


    # Display a message if some groups could not be created

    if($groupsNotCreated){

        Write-Host "$($groupsNotCreated.Count) groups could not be created:`n" -ForegroundColor Red
        $groupsNotCreated
        Write-Host "`n-------------------------------------------------------------------`n"

    }
       
    # Display a message if some users couldn't be added

    if($usersNotAdded){

        Write-Host "$($usersNotAdded.Count) users could not be added:`n" -ForegroundColor Red
        $usersNotAdded
        Write-Host "`n-------------------------------------------------------------------`n"

    }
    

} #Close main if block

# Display a message if the file browser was closed without selecting a file

else {

    Write-Host "Operation Cancelled" -ForegroundColor Red

}


Read-Host "Press any key to exit"