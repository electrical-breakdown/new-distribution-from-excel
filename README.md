# Summary
A Powershell script that creates new distribution groups from a collection of users in an Excel spreadsheet 

# Description
This script prompts the user to open an Excel workbook, then saves it under a new name. First the groups are created using the email address in column one, the display name in column two, and the owner in column 3. Once the group is created it is hidden from the global address list. 

Next the members are added. In the current version of the script, the members are added one at a time, which isn't very efficient. I did it this way because if there was a problem adding a member (typo, user already in the group, etc.) I wanted to know which one was the issue rather than have the whole thing fail. I later added some better error handling, so in the next iteration I'll probably revisit adding all the members at once by passing them in as an array when the group is created. 

Finally, the script writes a status message and any errors to the new workbook, formats the data, and saves the workbook.