#Load following assemblies in order to run the program in PS Console or VC Studio.
try{
    Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase
}
catch{
    Write-Error "Failed to load required assemblies."
    exit
}

#Set location of xaml file.
$path = Get-Location
$xamlFilePath = "$path\MainWindow.xaml"
#Download XAML file contents replacing every x:Name with Name.
$xml = (New-Object -TypeName System.Net.WebClient).DownloadString($xamlFilePath) -replace "x:Name", "Name"
#Ensure that $xml variable is of XML type.
$xml = [xml]$xml
#Remove XAML attributes which can't be processed by Windows.Markup.XamlReader.
$xml.Window.RemoveAttribute('x:Class')
$xml.Window.RemoveAttribute('mc:Ignorable')

#Initialize XmlNodeReader with previously formatted XML.
$xmlNodeReader = New-Object -TypeName System.Xml.XmlNodeReader $xml

#Parse $xmlNodeReader, select all WPF controls and create PS variables.
try{
    $wpf = [Windows.Markup.XamlReader]::Load($xmlNodeReader)
    $xml.SelectNodes("//*[@Name]") | ForEach-Object{
        $controlName = $_.Name
        $control = $wpf.FindName($controlName)
        if($control -ne $null){
            #Created variable holds all properties of given control.
            Set-Variable -Name $controlName -Value $control    
        }
    }
}
catch{
    Write-Host "Unable to parse xmlNodeReader"; 
    exit
}

###############################################################################################################                                                
################################################## VARIABLES ##################################################
###############################################################################################################
$userGroupsDataTable = [System.Data.DataTable]::New()
$userGroupsDataTable.Columns.AddRange(@(
        'groupName'
    )
)

$userGposDataTable = [System.Data.DataTable]::New()
$userGposDataTable.Columns.AddRange(@(
        'gpoName', 
        'trusteeName'
    )
)

$groupMembersDataTable = [System.Data.DataTable]::New()
$groupMembersDataTable.Columns.AddRange(@(
        'displayName',
        'objectClass'
    )
)

$userDataTable = [System.Data.DataTable]::New()
$userDataTable.Columns.AddRange(@(
        'displayName',
        'enabled',
        'samAccountName'
    )
)

$groupDataTable = [System.Data.DataTable]::New()
$groupDataTable.Columns.AddRange(@(
        'samAccountName'
    )
)

$reportDataTable = [System.Data.DataTable]::New()

###############################################################################################################                                                
################################################## /VARIABLES #################################################
###############################################################################################################

###############################################################################################################                                                
################################################## FUNCTIONS ##################################################
###############################################################################################################

#Function creating a filterString for given column and text. 
#Used in TxtBox handlers.
function filterTxtBox{
    param(
        [string]$dataColumn,
        [string]$filterText
    )

    if($filterText){
        $filterString = "$dataColumn LIKE '%$filterText%'"
    }else{
        $filterString = "$dataColumn LIKE '%%'" 
    }
    return $filterString
}


# Function to fetch user data based on a specific filter.
# Used in reportOptions ListBox.
function Get-UserDataByFilter {
    param(
        [string]$filter,
        [string]$propertyName
    )

    # Query the AD for given report option.
    $userQuery = Get-ADUser -Filter $filter -Properties $propertyName, samAccountName |
        Select-Object -Property samAccountName, $propertyName

    #Save retrived data to reportDataTable.
    $reportDataTable.Columns.AddRange(@(
        'samAccountName',
        $propertyName
    ))
    foreach ($user in $userQuery) {
        [void]$reportDataTable.Rows.Add(
            $user.samAccountName,
            $user.$propertyName
        )
    }
}

function Export-DataToCSV {
    param (
        [Parameter(Mandatory=$true)]
        [Object] $DataGrid
    )

    $data = $DataGrid.ItemsSource
    $columns = $DataGrid.Columns

    $csvContent = @()

    #Add header with column names
    $header = $columns.ForEach({$_.Header})
    $csvContent += $header -join ','

    #Add data rows
    foreach ($row in $data) {
        $rowData = $columns.ForEach({
            $cellValue = $_.GetCellContent($row).Text
            $cellValue = $cellValue -replace '"', '""'
            $cellValue = if ($cellValue -match ',') { "`"$cellValue`"" } else { $cellValue }
            $cellValue
        })
        $csvContent += $rowData -join ','
    }

    #Save to CSV file
    $csvContent | Out-File -FilePath "ExportedData.csv"

    Write-Host "Data exported to ExportedData.csv"
}

###############################################################################################################                                                
################################################## /FUNCTIONS #################################################
###############################################################################################################

###############################################################################################################
############################################### EVENT  HANDLERS ###############################################
###############################################################################################################

############################################### TxtBox handlers ###############################################
$txtFilterGroupMembers.Add_TextChanged({
    #Create filterString based on txtbox input using filterTxtBox function.
    $filterString = filterTxtBox -dataColumn "displayName" -filterText $txtFilterGroupMembers.Text
    #Filter the datatable and update the datagrid.
    $groupMembersDataTable.DefaultView.RowFilter = $filterString
    $groupMembersDataGrid.ItemsSource = $groupMembersDataTable.DefaultView
}
)

$txtBoxFilterUserGPOs.Add_TextChanged({
    $filterString = filterTxtBox -dataColumn "gpoName" -filterText $txtBoxFilterUserGPOs.Text
    $userGposDataTable.DefaultView.RowFilter = $filterString
    $userGposDataGrid.ItemsSource = $userGposDataTable.DefaultView
}
)

$txtBoxFilterUserGroups.Add_TextChanged({
    $filterString = filterTxtBox -dataColumn "groupName" -filterText $txtBoxFilterUserGroups.Text
    $userGroupsDataTable.DefaultView.RowFilter = $filterString
    $userGroupsDataGrid.ItemsSource = $userGroupsDataTable.DefaultView
}
)

$txtBoxFilterGroup.Add_TextChanged({
    $filterString = filterTxtBox -dataColumn "samAccountName" -filterText $txtBoxFilterGroup.Text
    $groupDataTable.DefaultView.RowFilter = $filterString
    $groupDataGrid.ItemsSource = $groupDataTable.DefaultView
}
)

$txtBoxFilterUsername.Add_TextChanged({
    $filterString = filterTxtBox -dataColumn "displayName" -filterText $txtBoxFilterUsername.Text
    $userDataTable.DefaultView.RowFilter = $filterString
    $userDataGrid.ItemsSource = $userDataTable.DefaultView
}
)
###############################################################################################################
############################################## DataGrid handlers ##############################################

#Handler for selectionChanged action in userDataGrid.
$userDataGrid.Add_SelectionChanged({
        #Clear datatables.
        $userGroupsDataTable.Clear()
        $userGposDataTable.Clear()
        
        if ($userDataGrid.SelectedItem -ne $null) {
            #Assing selected user to variable.
            $selectedUser = $userDataGrid.SelectedItem 

            #Retrive further information for selected user.
            $userInfo = Get-ADUser -Identity $selectedUser.samAccountName -Properties emailAddress, title, officePhone, streetAddress, office, department, postalCode, city, state, country, 
                canonicalName, created, modified, lastLogonDate, lastBadPasswordAttempt, passwordLastSet, passwordExpired, passwordNeverExpires, cannotChangePassword, logonCount, distinguishedName |
                Select-Object -Property emailAddress, title, officePhone, streetAddress, office, department, postalCode, city, state, country, 
                canonicalName, created, modified, lastLogonDate, lastBadPasswordAttempt, passwordLastSet, passwordExpired, passwordNeverExpires, cannotChangePassword, logonCount, distinguishedName

            #Retrive groups of selected user and add them do data table.
            $userGroups = Get-ADPrincipalGroupMembership -Identity $selectedUser.samAccountName | Select-Object -Property name
            foreach($group in $userGroups){
                [void]$userGroupsDataTable.Rows.Add(
                    $group.name
                )
            }

            #Retrive GPOs assigned to selected user based on their OU.
            $regex= "CN=[^,]+,"
            $ouPath = $userInfo.distinguishedName -replace $regex,""
            $userGpos = Get-GPInheritance -Target $ouPath | Select-Object -ExpandProperty InheritedGpoLinks | Where-Object -Property DisplayName -ne $null | Select-Object DisplayName
            #Create data table for user GPOs.
            foreach ($gpo in $usergpos) {
                #Get permissions of given GPO.
                $gpoName = $gpo.DisplayName
                $gpoId = (Get-GPO -Name $gpoName).Id
                $gpoPermissions = Get-GPPermissions -Guid $gpoId -All
                $trustees = $GPOPermissions | Select-Object -ExpandProperty Trustee | Select-Object -Property Name
                #Check if GPO is applied to selected user either by OU, user account or any of the groups user is a member of. 
                foreach($trustee in $trustees){
                    foreach($group in $userGroups){
                        #Add GPO to userGposDataTable. 
                        if ($trustee.Name -eq "Authenticated Users" -or $trustee.Name -eq $selectedUser.samAccountName -or $group.Name -eq $trustee.Name) {
                            [void]$userGposDataTable.Rows.Add(
                                $gpo.DisplayName,
                                $trustee.Name
                            )
                            break
                        }
                    }
                }
            }
            #Display user details.
            $txtBlock_displayName.Text = $selectedUser.displayName
            $txtBlock_email.Text = $userInfo.emailAddress
            $txtBlock_title.Text = $userInfo.title
            $txtBlock_officePhone.Text = $userInfo.officePhone  
            $txtBlock_streetAddress.Text = $userInfo.streetAddress
            $txtBlock_office.Text = $userInfo.office
            $txtBlock_department.Text = $userInfo.department
            $txtBlock_postalCode.Text = $userInfo.postalCode
            $txtBlock_city.Text = $userInfo.city
            $txtBlock_state.Text = $userInfo.state
            $txtBlock_country.Text = $userInfo.country

            #Display account details.
            $txtBlock_canonicalName.Text = $userInfo.canonicalName
            $txtBlock_created.Text = $userInfo.created
            $txtBlock_modified.Text = $userInfo.modified
            $txtBlock_lastLogonDate.Text = $userInfo.lastLogonDate 
            $txtBlock_lastBadPasswordAttempt.Text = $userInfo.lastBadPasswordAttempt
            $txtBlock_passwordLastSet.Text = $userInfo.passwordLastSet
            $txtBlock_passwordExpired.Text = $userInfo.passwordExpired
            $txtBlock_passwordNeverExpires.Text = $userInfo.passwordNeverExpires
            $txtBlock_cannotChangePassword.Text = $userInfo.cannotChangePassword
            $txtBlock_logonCount.Text = $userInfo.logonCount
            
            #Display user groups in a data grid.
            $userGroupsDataGrid.ItemsSource = $userGroupsDataTable.DefaultView
            #Display user GPOs in a data grid.
            $userGposDataGrid.ItemsSource = $userGposDataTable.DefaultView
        }
    }
)

#Handler for selectionChanged action in groupDataGrid.
$groupDataGrid.Add_SelectionChanged({
        #Clear datatables.
        $groupMembersDataTable.Clear()

        if ($groupDataGrid.SelectedItem -ne $null) {
            $selectedGroup = $groupDataGrid.SelectedItem
            $groupInfo = Get-ADGroup -Identity $selectedGroup.samAccountName -Properties groupCategory, groupScope, canonicalName, created, modified, protectedFromAccidentalDeletion |
            Select-Object -Property groupCategory, groupScope, canonicalName, created, modified, protectedFromAccidentalDeletion
            $groupMemberCount = (($selectedGroup.samAccountName | Get-ADGroupMember) | Measure-Object).Count
            $groupMembers = Get-ADGroupMember -Identity $selectedGroup.samAccountName
            foreach($member in $groupMembers){
                [void]$groupMembersDataTable.Rows.Add(
                    $member.name,
                    $member.objectClass
                )
            }

            $txtBlock_groupDetails.Text = $selectedGroup.samAccountName
            $txtBlock_memberCount.Text = $groupMemberCount
            $txtBlock_groupScope.Text = $groupInfo.groupScope
            $txtBlock_groupCategory.Text = $groupInfo.groupCategory
            $txtBlock_groupCanonicalName.Text = $groupInfo.canonicalName
            $txtBlock_groupCreated.Text = $groupInfo.created
            $txtBlock_groupModified.Text = $groupInfo.modified
            $txtBlock_deletionProtected.Text = $groupInfo.protectedFromAccidentalDeletion

            $groupMembersDataGrid.ItemsSource = $groupMembersDataTable.DefaultView
        }
    }
)
###############################################################################################################
############################################## ListBox  handlers ##############################################

#Event handler for SelectionChanged action in reportOptions ListBox.
$reportOptions.Add_SelectionChanged(
    {
        $selectedItem = $reportOptions.SelectedItem
        $reportDataTable.Dispose()
        $reportDataTable = [System.Data.DataTable]::New()

        $currentDate = (Get-Date)
        $next7days = $currentDate.AddDays(7)
        $next30days = $currentDate.AddDays(30)
        $last24h = $currentDate.AddDays(-1)
        $last7days = $currentDate.AddDays(-7)
        $last30days = $currentDate.AddDays(-30)
        $last90days = $currentDate.AddDays(-90)
        $last180days = $currentDate.AddDays(-180)

        
        #Retrive data based on the selected report option
        switch ($selectedItem) {
            "Users created in Last 30 days" {
                Get-UserDataByFilter -filter {created -ge $last30days} -propertyName 'created'
            }
            "Users created in Last 90 days" {
                Get-UserDataByFilter -filter {created -ge $last90days} -propertyName 'created'
            }
            "Users created in Last 180 days" {
                Get-UserDataByFilter -filter {created -ge $last180days} -propertyName 'created'
            }
            "Users modified in Last 30 days" {
                Get-UserDataByFilter -filter {modified -ge $last30days} -propertyName 'modified'
            }
            "Users modified in Last 90 days" {
                Get-UserDataByFilter -filter {modified -ge $last90days} -propertyName 'modified'
            }
            "Users modified in Last 180 days" {
                Get-UserDataByFilter -filter {modified -ge $last180days} -propertyName 'modified'
            }
            "Disabled user accounts" {
                Get-UserDataByFilter -filter {enabled -eq $false} -propertyName 'enabled'
            }
            "Users with expiring account" {
                $userQuery = Get-ADUser -Filter {accountExpirationDate -like '*'} -Properties samAccountName, accountExpirationDate | 
                    Select-Object -Property samAccountName, accountNeverExpires
                $reportDataTable.Columns.AddRange(@(
                        'samAccountName',
                        'accountNeverExpires'
                    )
                )
                foreach($user in $userQuery){
                    $user.accountNeverExpires = $false
                    [void]$reportDataTable.Rows.Add(
                        $user.samAccountName,
                        $user.accountNeverExpires
                    )
                } 
            }
            "Users with expired account" {
                $userQuery = Get-ADUser -Filter {accountExpirationDate -lt $currentDate} -Properties samAccountName, enabled | 
                    Select-Object -Property samAccountName, enabled, isExpired
                $reportDataTable.Columns.AddRange(@(
                        'samAccountName',
                        'enabled',
                        'isExpired'
                    )
                )
                foreach($user in $userQuery){
                    $user.isExpired = $true
                    [void]$reportDataTable.Rows.Add(
                        $user.samAccountName,
                        $user.enabled,
                        $user.isExpired
                    )
                } 
            }
            "Users with account that is set to expire" {
                Get-UserDataByFilter -filter {accountExpirationDate -gt $currentDate} -propertyName 'accountExpirationDate'
            }
            "Users with account that is set to expire in 7 days" {
                Get-UserDataByFilter -filter {(accountExpirationDate -le $next7days) -and (accountExpirationDate -gt $currentDate)} -propertyName 'accountExpirationDate'
            }
            "Users with account that is set to expire in 30 days" {
                Get-UserDataByFilter -filter {(accountExpirationDate -le $next30days) -and (accountExpirationDate -gt $currentDate)} -propertyName 'accountExpirationDate'
            }
            "Users inactive for 7 days" {
                Get-UserDataByFilter -filter {lastLogonDate -lt $last7days} -propertyName 'lastLogonDate'
            }
            "Users inactive for 30 days" {
                Get-UserDataByFilter -filter {lastLogonDate -lt $last30days} -propertyName 'lastLogonDate'
            }
            "Users inactive for 90 days" {
                Get-UserDataByFilter -filter {lastLogonDate -lt $last90days} -propertyName 'lastLogonDate'
            }
            "Users inactive for 180 days" {
                Get-UserDataByFilter -filter {lastLogonDate -lt $last180days} -propertyName 'lastLogonDate'
            }
            "Users that logged on in last 24h" {
                Get-UserDataByFilter -filter {lastLogonDate -ge $last24h} -propertyName 'lastLogonDate'
            }
            "Users that logged on in last 7 days" {
                Get-UserDataByFilter -filter {lastLogonDate -ge $last7days} -propertyName 'lastLogonDate'
            }
            "Users that logged on in last 30 days" {
                Get-UserDataByFilter -filter {lastLogonDate -ge $last30days} -propertyName 'lastLogonDate'
            }
            "Users that must change their password at next logon" {
                $userQuery = Get-ADUser -Filter {pwdLastSet -eq 0} -Properties samAccountName  | Select-Object -Property samAccountName, pwdChangeRequired
                $reportDataTable.Columns.AddRange(@(
                        'samAccountName',
                        'pwdChangeRequired'
                    )
                )
                foreach($user in $userQuery){
                    $user.pwdChangeRequired = $true
                    [void]$reportDataTable.Rows.Add(
                        $user.samAccountName,
                        $user.pwdChangeRequired
                    )
                } 
            }
            "Users that must change their password in 1 day" {
                $userQuery = Get-ADUser -LDAPFilter "(msDS-UserPasswordExpiryTimeComputed >= $last24h.ToFileTime())" -Properties samAccountName, passwordLastSet  | 
                    Select-Object -Property samAccountName, passwordLastSet
                $reportDataTable.Columns.AddRange(@(
                        'samAccountName',
                        'passwordLastSet'
                    )
                )
                foreach($user in $userQuery){
                    [void]$reportDataTable.Rows.Add(
                        $user.samAccountName,
                        $user.passwordLastSet
                    )
                } 
            }
            "Users that must change their password in 7 days" {
                $userQuery = Get-ADUser -LDAPFilter "(msDS-UserPasswordExpiryTimeComputed >= $last7days.ToFileTime())" -Properties samAccountName, passwordLastSet  | 
                    Select-Object -Property samAccountName, passwordLastSet
                $reportDataTable.Columns.AddRange(@(
                        'samAccountName',
                        'passwordLastSet'
                    )
                )
                foreach($user in $userQuery){
                    [void]$reportDataTable.Rows.Add(
                        $user.samAccountName,
                        $user.passwordLastSet
                    )
                } 
            }
            "Users with expired password" {
                $userQuery = Get-AdUser -Filter {enabled -eq $true} -Properties samAccountName, passwordExpired, passwordLastSet | 
                Where-Object {$_.passwordExpired -eq $true}  | 
                Select-Object -Property samAccountName, passwordExpired, passwordLastSet
                $reportDataTable.Columns.AddRange(@(
                        'samAccountName',
                        'passwordExpired',
                        'passwordLastSet'
                    )
                )
                foreach($user in $userQuery){
                    [void]$reportDataTable.Rows.Add(
                        $user.samAccountName,
                        $user.passwordExpired,
                        $user.passwordLastSet
                    )
                } 
            }
            "Users with non-expiring password" {
                Get-UserDataByFilter -filter {passwordNeverExpires -eq $true} -propertyName 'passwordNeverExpires'
            }
            "Users that can't change their password" {
                $userQuery = Get-ADUser -Filter * -Properties samAccountName, cannotChangePassword |
                Where-Object { $_.cannotChangePassword -eq $true } |
                Select-Object -Property samAccountName, cannotChangePassword
                $reportDataTable.Columns.AddRange(@(
                        'samAccountName',
                        'cannotChangePassword'
                    )
                )
                foreach($user in $userQuery){
                    [void]$reportDataTable.Rows.Add(
                        $user.samAccountName,
                        $user.cannotChangePassword
                    )
                } 
            }
            "Users that changed their password in last 7 days" {
                Get-UserDataByFilter -filter {passwordLastSet -ge $last7days} -propertyName 'passwordLastSet'
            }
            "Users that changed their password in last 30 days" {
                Get-UserDataByFilter -filter {passwordLastSet -ge $last30days} -propertyName 'passwordLastSet'
            }
            "Users that changed their password in last 90 days" {
                Get-UserDataByFilter -filter {passwordLastSet -ge $last90days} -propertyName 'passwordLastSet'
            }
        }
         #Display report data in reportDataGrid.
        $reportDataGrid.ItemsSource = $reportDataTable.DefaultView
    }   
)
###############################################################################################################
############################################### Button handlers ###############################################

# Event handler for the Export to CSV button click
$buttonExportCSV.Add_Click({ Export-DataToCSV -DataGrid $reportDataGrid })
###############################################################################################################

###############################################################################################################
############################################### /EVENT HANDLERS ###############################################
###############################################################################################################

#Select all AD users.
$users = Get-ADUser -Filter * -Properties displayName, enabled, samAccountName | Select-Object -Property displayName, enabled, samAccountName 


#Populate the DataTable with selected user data.
foreach($user in $users){
    [void]$userDataTable.Rows.Add(
        $user.displayName,
        $user.enabled,
        $user.samAccountName
    )
}

#Select all AD groups. 
$groups = Get-ADGroup -Filter * -Properties samAccountName | Select-Object -Property samAccountName

foreach($group in $groups){
    [void]$groupDataTable.Rows.Add(
        $group.samAccountName
    )
}



#Display data in the datagrids.
$userDataGrid.ItemsSource = $userDataTable.DefaultView
$groupDataGrid.ItemsSource = $groupDataTable.DefaultView

#Load GUI.
$wpf.ShowDialog() | Out-Null