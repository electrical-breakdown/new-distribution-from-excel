# Summary
A Powershell script that creates new distribution groups from a collection of users in an Excel spreadsheet 

# Description
This script prompts the user to open an Excel workbook, then saves it under a new name. The groups are created using the email address in column one, the display name in column two, and the owner in column 3. The group members are colleted into array by looping through the remaining columns, and passed into New-DistributionGroup. Once the group is created it is hidden from the global address list. 

The script writes a status message and any errors to the new workbook, formats the data, and saves the workbook.