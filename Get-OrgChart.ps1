<#
    .SYNOPSIS
    Exports user and manager information to an Excel file.

    .DESCRIPTION
    This script connects to Microsoft Graph and retrieves a list of all users in the directory with a non-null job title. It then gets the manager of each user and exports the user and manager information to an Excel file named "OrgChart.xlsx" with a table name of "Staff".

#>


# Installs the ImportExcel module and connects to Microsoft Graph
Install-Module ImportExcel
Connect-MgGraph -Scopes Directory.ReadWrite.All

# Initialize an empty array and retrieve a list of all users in the directory with a non-null job title
$Result = @()
$AllUsers = Get-MgUser -AsJob | Where-Object { $_.JobTitle -ne $Null }

# Set the total number of users to be processed and initialize a counter
$TotalUsers = $AllUsers.Count
$i = 1 

# Loop through each user and get the manager of the user
$AllUsers | ForEach-Object {
    $User = $_
    Write-Progress -Activity "Processing $($_.Displayname)" -Status "$i out of $TotalUsers completed" -Verbose
    $managerObj = Get-MgUserManager -UserId "$($User.Id)" -ErrorAction Stop
    $managerIDText = $managerObj | Select-Object -ExpandProperty Id
    $ManagerUser = Get-MgUser -UserId $managerIDText -ErrorAction Stop
    
    # Add an object to the $Result array with the user and manager information
    $Result += New-Object PSObject -property @{ 
        UserName = $User.DisplayName
        UserPrincipalName = $User.UserPrincipalName
        JobTitle = $User.JobTitle
        ManagerName = $ManagerUser.DisplayName
        ManagerMail = $ManagerUser.Mail
    }
    
    # Increment the counter
    $i++
}

# Select the properties of the objects in the $Result array and export them to an Excel file
$Result | Select-Object UserName, UserPrincipalName,JobTitle,ManagerName,ManagerMail |
Export-Excel -Path ".\OrgChart.xlsx" -TableName "Staff"
