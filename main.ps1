#Load following assemblies in order to run the program in PS Console or VC Studio.
try{
    Add-Type -AssemblyName PresentationCore, PresentationFramework
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

#Function creating a filterString for given column and text. 
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

#Handler for txtFilterGroupMembers  textbox.
$txtFilterGroupMembers.Add_TextChanged({
        #Create filterString based on textbox input using filterTxtBox function.
        $filterString = filterTxtBox -dataColumn "displayName" -filterText $txtFilterGroupMembers.Text
        #Filter the datatable and update the datagrid.
        $groupMembersDataTable.DefaultView.RowFilter = $filterString
        $groupMembersDataGrid.ItemsSource = $groupMembersDataTable.DefaultView
    }
)

#Handler for txtBoxFilterUserGPOs textbox.
$txtBoxFilterUserGPOs.Add_TextChanged({
        #Create filterString based on textbox input using filterTxtBox function.
        $filterString = filterTxtBox -dataColumn "gpoName" -filterText $txtBoxFilterUserGPOs.Text
        #Filter the datatable and update the datagrid.
        $userGposDataTable.DefaultView.RowFilter = $filterString
        $userGposDataGrid.ItemsSource = $userGposDataTable.DefaultView
    }
)

#Handler for txtBoxFilterUserGroups textbox.
$txtBoxFilterUserGroups.Add_TextChanged({
        #Create filterString based on textbox input using filterTxtBox function.
        $filterString = filterTxtBox -dataColumn "groupName" -filterText $txtBoxFilterUserGroups.Text
        #Filter the datatable and update the datagrid.
        $userGroupsDataTable.DefaultView.RowFilter = $filterString
        $userGroupsDataGrid.ItemsSource = $userGroupsDataTable.DefaultView
    }
)

#Handler for txtBoxFilterGroup textbox.
$txtBoxFilterGroup.Add_TextChanged({
        #Create filterString based on textbox input using filterTxtBox function.
        $filterString = filterTxtBox -dataColumn "samAccountName" -filterText $txtBoxFilterGroup.Text
        #Filter the datatable and update the datagrid.
        $groupDataTable.DefaultView.RowFilter = $filterString
        $groupDataGrid.ItemsSource = $groupDataTable.DefaultView
    }
)

#Handler for txtBoxFilterUsername textbox.
$txtBoxFilterUsername.Add_TextChanged({
        #Create filterString based on textbox input using filterTxtBox function.
        $filterString = filterTxtBox -dataColumn "displayName" -filterText $txtBoxFilterUsername.Text
        #Filter the datatable and update the datagrid.
        $userDataTable.DefaultView.RowFilter = $filterString
        $userDataGrid.ItemsSource = $userDataTable.DefaultView
    }
)

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
                        if ($trustee.Name -eq "Authenticated Users" -or $trustee.Name -eq $selectedUser.displayName -or $group.Name -eq $trustee.Name) {
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

#Select all AD users.
$users = Get-ADUser -Filter * -Properties displayName, enabled, samAccountName | Select-Object -Property displayName, enabled, samAccountName 

#Declare a DataTable and add columns.
$userDataTable = [System.Data.DataTable]::New()
$userDataTable.Columns.AddRange(@(
        'displayName',
        'enabled',
        'samAccountName'
    )
)
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
$groupDataTable = [System.Data.DataTable]::New()
$groupDataTable.Columns.AddRange(@(
        'samAccountName'
    )
)
foreach($group in $groups){
    [void]$groupDataTable.Rows.Add(
        $group.samAccountName
    )
}



#Display data in the datagrids.
$userDataGrid.ItemsSource = $userDataTable.DefaultView
$groupDataGrid.ItemsSource = $groupDataTable.DefaultView

#Show GUI.
$wpf.ShowDialog() | Out-Null