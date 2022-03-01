##########################################
#    ███████╗███████╗██████╗ ███████╗    #
#    ██╔════╝╚══███╔╝██╔══██╗╚══███╔╝    #
#    █████╗    ███╔╝ ██████╔╝  ███╔╝     #
#    ██╔══╝   ███╔╝  ██╔═══╝  ███╔╝      #
#    ███████╗███████╗██║     ███████╗    #
#    ╚══════╝╚══════╝╚═╝     ╚══════╝    #               
##########################################

####################################################
#            SECTION ONE: INITIALIZATION           #
####################################################

function Get-ConnectionStatus 
{
    # Returns boolean indicating connection status to Office 365.
    Get-MsolDomain -ErrorAction SilentlyContinue | Out-Null
    $Result = $?
    if (!$Result)
    {
        Write-Host "You've become disconnected from Office 365 due to inactivty."
        Write-Host "Please wait while we reconnect you."
        Connect-Office365
    }
}


function Connect-Office365
{
    # Connect to Office 365.
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $global:Cred -Authentication Basic -AllowRedirection
    Import-Module (Import-PSSession $Session -AllowClobber -DisableNameChecking) -DisableNameChecking 3>$null

    # Connect to Licensing Service.
    Connect-MsolService -Credential $global:Cred | Out-Null
}


function Add-StoredCredential
{
    $Credential = Get-Credential -Message "Enter the [redacted] password" -UserName redacted@maritimetravel.ca

    $Credential.Password | ConvertFrom-SecureString | Out-File "$($global:KeyPath)\$($Credential.Username).cred" -Force
}


function Get-StoredCredentials 
{
    param
    (
        [Parameter(Mandatory=$false, ParameterSetName="Get")]
        [string]$UserName,
        [Parameter(Mandatory=$false, ParameterSetName="List")]
        [switch]$List
    )

    if ($List) 
    {
        try 
        {
            $CredentialList = @(Get-ChildItem -Path $global:KeyPath -Filter *.cred -ErrorAction STOP)
            foreach ($Cred in $CredentialList) 
            {
                Write-Host "Username: $($Cred.BaseName)"
            }
        }
        catch 
        {
            Write-Warning $_.Exception.Message
        }
    }

    if ($UserName) 
    {
        if (Test-Path "$($global:KeyPath)\$($Username).cred") 
        {
            $PwdSecureString = Get-Content "$($global:KeyPath)\$($Username).cred" | ConvertTo-SecureString
            $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $PwdSecureString
        }
        else
        {
            return $False
        }
        return $Credential
    }
}


function Test-Credentials
{     
    [CmdletBinding()]
    Param
    ( 
        [Parameter( 
            Mandatory = $false, 
            ValueFromPipeLine = $true, 
            ValueFromPipelineByPropertyName = $true
        )] 
        [Alias( 
            'PSCredential'
        )] 
        [ValidateNotNull()] 
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()] 
        $Credentials
    )
    $Domain = $null
    $Root = $null
    $Username = $null
    $Password = $null
            
    # Checking module
    Try
    {
        # Split username and password
        $Username = $credentials.username
        $Password = $credentials.GetNetworkCredential().password
  
        # Get Domain
        $Root = "LDAP://" + ([ADSI]'').distinguishedName
        $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$UserName,$Password)
    }
    Catch
    {
        $_.Exception.Message
        Continue
    }
  
    If(!$domain)
    {
        Write-Warning "Something went wrong"
    }
    Else
    {
        If ($domain.name -ne $null)
        {
            return $true
        }
        Else
        {
            return $false
        }
    }
}


function Confirm-Credentials
{
    param
    (
        [string]$CredentialName
    )
    while(1)
    {
        if (Test-Path -Path "$global:KeyPath\$CredentialName.cred")
        {
            $global:Cred = Get-StoredCredentials -UserName "$CredentialName"
            if (Test-Credentials -Credentials $global:Cred)
            {
                break
            }
            else
            {
                Write-Host "The stored credentials for [redacted]@maritimetravel.ca did not work.`nPlease try again." -ForegroundColor Red
                $global:Cred = Add-StoredCredential
            }
        }
        else
        {
            Write-Host "Credentials for [redacted]@maritimetravel.ca were not found.`nPlease enter the credentials." -ForegroundColor Red
            $global:Cred = Add-StoredCredential
        }
    }
}

function Get-BranchHash
{
    $Hash = @{}
    $Query = "SELECT BranchId, internalBranchName, Status FROM [redacted] WHERE (Status = 'Valid' OR internalBranchName = 'CSO') ORDER BY internalBranchName"
    $QueryResults = Invoke-Sqlcmd2 -Query $Query -Database SalesDesk -ServerInstance redacted

    foreach ($Item in $QueryResults)
    {
        $Hash.Add($Item.internalBranchName, $Item.BranchId) 
    }

    $Hash = $Hash.GetEnumerator() | Sort-Object -Property key
    return $Hash
}

function Start-ADSync
{
    Invoke-Command -ComputerName mti-azuread {start-adsyncsynccycle -policytype delta} | Out-Null
}


####################################################
#              SECTION TWO: MENU LOGIC             #
####################################################

# Removes symbols from given string.
function Remove-SpecialCharacters
{
    param
    (
        [Parameter(Mandatory)][string]$InputString
    )
    $SpecialCharacters = "!?*%#@$&^"
    foreach ($Character in $SpecialCharacters.ToCharArray())
    {
        $InputString = $InputString.Replace("$Character","")
    }
    return $InputString
}

function New-ApolloAccount 
{
    param
    (
        [Parameter(Mandatory)][string]$FirstName,
        [Parameter(Mandatory)][string]$LastName,
        [Parameter(Mandatory)][bool]$Emulating,
        [Parameter(Mandatory)][int]$BranchNumber
    )

    if ($Emulating)
    {
        # Picks a random letter from C to Z, this is used to help randomize the username and prevent a duplicate.
        $Initial = Get-Random -InputObject "C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"
        $ApolloUsername = ("ZDMM"+$Initial+$FirstName[0]+$LastName[0]).ToUpper()
    }
    elseif (!$Emulating)
    {
        $ApolloUsername = ("Z"+$FirstName[0]+$LastName[0]).ToUpper()
        $Initial = ($FirstName[0]+$LastName[0]).ToUpper()

        # Gets PCC from SQL database
        $PseudoCityCode = Invoke-SqlCmd2 -query "select PseudoCityCode from [redacted] where TramsBranchNo = '$BranchNumber'" -ServerInstance redacted -Database redacted | Select-Object -ExpandProperty PseudoCityCode
    }
    return $ApolloUsername
}

# Queries Trams DB for lowest unused Interface ID.
function Get-InterfaceID
{
    # Queries Trams DB to find lowest number unused Interface ID in a given range.
    param
    (
        [int]$MinimumValue = 2000,
        [int]$MaximumValue = 2500
    )

    $TramsQuery = "
    SELECT name, case when isnumeric(interfaceid) =1 then interfaceid else 0 end as interfaceid
    FROM [redacted]
    WHERE 
    PROFILETYPE_LINKCODE = 'A'
    and len(interfaceid) = 4
    and (case when isnumeric(interfaceid) =1 then interfaceid else 0 end) between $MinimumValue and $MaximumValue
    order by interfaceid
"
    # Gets starting point.
    $TramsResult = ((Invoke-Sqlcmd2 -Query $TramsQuery -ServerInstance redacted -Database redacted | Select-Object -ExpandProperty InterfaceID -Last 1) + 1)

    do
    {
        # Trams DB takes 24 hours to update, so double-checks employee table to see if number is already used.
        $Employee = Invoke-Sqlcmd2 -Query "SELECT ConsNo FROM [redacted] WHERE (ConsNo = '$TramsResult')" -ServerInstance redacted -Database redacted

        if (!$Employee)
        {
            return $TramsResult
        }
        else
        {
            $TramsResult += 1
        }
    }
    while ($Employee)
}

# Disables Outlook automatically adding events to calendar.
function Disable-EventsFromEmail
{
    param
    (
        [string]$Identity
    )

    Set-MailboxCalendarConfiguration -Identity $Identity -DiningEventsFromEmailEnabled $False -EntertainmentEventsFromEmailEnabled $False -EventsFromEmailEnabled $False -FlightEventsFromEmailEnabled $False -HotelEventsFromEmailEnabled $False -InvoiceEventsFromEmailEnabled $False -PackageDeliveryEventsFromEmailEnabled $False -RentalCarEventsFromEmailEnabled $False
}

# Creates new employee from template user.
function New-Employee
{
    param
    (
        [string] $TemplateUser = $(throw "You must specify a user to copy."),
        [string] $NewUsername = $(throw "You must specify a new username."),
        [string] $FirstName = $(throw "You must specify a first name."),
        [string] $LastName = $(throw "You must specify a last name."),
        [string] $BranchName = $(throw "You must specify a branch."),
        [string] $JobTitle = $(throw "You must specify a job."),
        [bool] $ApolloEmulating = $False,
        [bool] $EmailInsurance = $False,
        [bool] $EmailWelcome = $False
    )

    $Password = New-ADPassword
    $ConsultantNumber = Get-InterfaceID
    $FullName = "$Firstname $Lastname"

    # Convert Branch Name to ID/TramsBranchNumber and Job Title to Job ID.
    Write-Host "Preparing to create employee..."
    $BranchID = $global:BranchHash | Where-Object Name -eq $BranchName | Select-Object -ExpandProperty Value
    $TramsBranchNumber = Invoke-Sqlcmd2 "SELECT tramsbranchno FROM [redacted] WHERE BranchId = $BranchId" -server redacted -Database redacted | Select-Object -ExpandProperty tramsbranchno
    $JobID = $global:JobHash.Item($JobTitle)

    # Preparing user data for copying.
    $TemplateUserData = Get-ADUser -Identity $TemplateUser -Properties *
    switch -Wildcard ($TemplateUserData.mail)
    {
        "*@maritimetravel.ca"
        {
            $PrimarySMTP = "@maritimetravel.ca"
            $SecondarySMTP = "@legrowstravel.ca"
            $TertiarySMTP = "@voyagesmaritime.ca"
        }

        "*@legrowstravel.ca"
        {
            $PrimarySMTP = "@legrowstravel.ca"
            $SecondarySMTP = "@maritimetravel.ca"
            $TertiarySMTP = "@voyagesmaritime.ca"
        }

        "*@voyagesmaritime.ca"
        {
            $PrimarySMTP = "@voyagesmaritime.ca"
            $SecondarySMTP = "@maritimetravel.ca"
            $TertiarySMTP = "@legrowstravel.ca"
        }
    }
    $FullEmail = "$NewUsername$PrimarySMTP"

    # Create new AD account and copy group memberships.
    Write-Host "Creating AD account and group memberships..."
    $NewAccountDetails = 
    @{
        SamAccountName = $NewUsername
        Name = $FullName
        GivenName = $FirstName
        Surname = $LastName
        DisplayName = $FullName
        Instance = $TemplateUserData.DistinguishedName
        Path = ([ADSI](([ADSI]"LDAP://$($TemplateUserData.DistinguishedName)").Parent)) | Select-Object -ExpandProperty distinguishedName
        AccountPassword = (ConvertTo-SecureString -String $Password -AsPlainText -Force)
        UserPrincipalName = "$NewUsername@maritimetravel.ca"
        Company = $TemplateUserData.Company
        Title = $TemplateUserData.Title
        Office = $TemplateUserData.Office
        City = $TemplateUserData.City
        PostalCode = $TemplateUserData.PostalCode
        State =$TemplateUserData.State
        StreetAddress = $TemplateUserData.StreetAddress
        EmailAddress = "$NewUsername$PrimarySMTP"
        Enabled = $true
    }
    New-ADUser @NewAccountDetails
    $TemplateUserData.MemberOf | Add-ADGroupMember -Members $NewUsername -ErrorAction SilentlyContinue
    Set-ADUser -Identity $NewUsername -Add @{Proxyaddresses="SMTP:"+$NewUsername+$PrimarySMTP}
    Set-ADUser -Identity $NewUsername -Add @{Proxyaddresses="smtp:"+$NewUsername+$SecondarySMTP}
    Set-ADUser -Identity $NewUsername -Add @{Proxyaddresses="smtp:"+$NewUsername+$TertiarySMTP}

    # Sync AD to Office 365.
    Write-Host "Syncing Active Directory to Office 365..."
    Start-ADSync

    # Waits until user is seen in Office 365, then adds license.
    while(1)
    {
        $Synced = Get-MsolUser -UserPrincipalName "$NewUsername@maritimetravel.ca" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        If (!$Synced)
        {
            Start-Sleep 5
        }
        else
        {
            Set-MsolUser -UserPrincipalName "$NewUsername@maritimetravel.ca" -UsageLocation "CA"
            Set-MsolUserLicense -UserPrincipalName "$NewUsername@maritimetravel.ca" -AddLicenses "weknowtravelbest:EXCHANGESTANDARD"
            break
        }
    }

    # Waits until mailbox is created, then adds [redacted] send-as rights.
    Write-Host "Waiting for mailbox to be created. This may take some time..."
    while(1) 
    {
        $MailboxCreated = (Get-Mailbox -Identity $NewUsername -EA SilentlyContinue -WA SilentlyContinue).isvalid

        if (!$MailboxCreated)
        {
            Start-Sleep 5
        }
        else
        {
            Add-RecipientPermission -Identity $NewUsername -AccessRights SendAs -Trustee redacted -Confirm:$False -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
            Add-RecipientPermission -Identity $NewUsername -AccessRights SendAs -Trustee redacted -Confirm:$False -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
            Disable-EventsFromEmail -Identity $NewUsername
            break
        }
    }

    # Submit to Jose's new employee tool, which populates SQL tables with employee information.
    Write-Host "Populating SQL tables with employee information..."
    $CBPassword = "$TramsBranchNumber"+$Lastname.subString(0, [System.Math]::Min(4, $Lastname.Length)).ToLower()
    $Company = Invoke-Sqlcmd2 -Query "select company from [redacted] where tramsbranchno = $TramsBranchNumber" -ServerInstance redacted -Database redacted | Select-Object -ExpandProperty company
    $AreaCode = Invoke-SqlCmd2 -Query "select areacode from [redacted] where tramsbranchno = $TramsBranchNumber" -ServerInstance redacted -Database redacted | Select-Object -ExpandProperty areacode
    $PhoneNumber = Invoke-SqlCmd2 -Query "select br_ph from [redacted] where tramsbranchno = $TramsBranchNumber" -ServerInstance redacted -Database redacted | Select-Object -ExpandProperty br_ph
    $PhoneNumber = "($AreaCode) $($PhoneNumber.Insert(3,"-"))"
    $Params = @{fname=$FirstName;lname=$LastName;username=$NewUsername;password=$Password;wemail=$FullEmail;JobId=$JobID;BranchId=$BranchID;wphone=$PhoneNumber;TempLeaveDate="";Company=$Company;fpusername=$ApolloUsername;consno=$ConsultantNumber;cbuname=$NewUsername;cbpword=$CBPassword;travelbotsignon=$NewUsername;betateam=$False;sendemail=$EmailWelcome}
    Invoke-WebRequest -Uri "[redacted]" -Method POST -Body $Params -UseDefaultCredentials | Out-Null
    
    # ORI Insurance email template. Emails their department if specified.
    $ORIInsuranceEmail = "<p>Hello,</p>
    <p>This is an automatically generated email. If there are any issues with the formatting or information below, please alert IT Support.</p>
    <p>Please create an Old Republic account for this new employee.</p>
    <p>$FullName<br />$FullEmail<br />Branch $TramsBranchNumber - $BranchName</p>
    <p>Username: $NewUsername&nbsp;<br />Password: Maritime$TramsBranchNumber`2017<br />Agency Code: $($TramsBranchNumber.ToString("MT000"))</p>
    <p>Thank you,<br />IT Support</p>"
    if ($EmailInsurance)
    {
        Write-Host "Sending email to ORI Insurance..."
        Send-MailMessage -From $global:Cred.UserName -To "[redacted]@maritimetravel.ca" -Cc "$env:USERNAME@maritimetravel.ca" -Subject "New Employee $FullName" -Body $ORIInsuranceEmail -BodyAsHtml -SmtpServer "smtp.office365.com" -UseSsl -Port 587 -Credential $global:Cred
    }

    # TIS insurance email template. Emails their department if specified.
    $EmployeeNumber = Invoke-SqlCmd2 -Query "SELECT EmployeeNo FROM [redacted] WHERE (username = '$NewUsername')" -ServerInstance redacted -Database redacted | Select-Object -ExpandProperty EmployeeNo
    $Province = Invoke-Sqlcmd2 -Query "SELECT Province FROM [redacted] WHERE (TramsBranchNo = '$TramsBranchNumber')" -ServerInstance redacted -Database redacted | Select-Object -ExpandProperty Province
    $TISInsuranceEmail = "<p>Hello,<br>
    This is an automatically generated email. If there are any issues with the formatting or information below, please alert <a href=`"[redacted]@maritimetravel.ca`">[redacted]@maritimetravel.ca</a>.<br>
    <br>
    Please create an account for this new employee:<br>
    <br>
    Full Name:<br>
    $FullName<br>
    <br>
    Email Address:<br>
    $FullEmail<br>
    <br>
    Employee Number:<br>
    $EmployeeNumber<br>
    <br>
    Branch Information:<br>
    #$TramsBranchNumber - $BranchName, $Province<br>
    $TramsBranchNumber@maritimetravel.ca<br>
    <br>
    Thank you,<br>
    IT Support<br>
    Maritime Travel"
    if ($EmailInsurance)
    {
        Write-Host "Sending email to TIS insurance"
        Send-MailMessage -From $global:Cred.UserName -To "[redacted]" -Cc "$env:USERNAME@maritimetravel.ca","[redacted]" -Subject "Maritime Travel - New Employee $FullName" -Body $TISInsuranceEmail -BodyAsHtml -SmtpServer "smtp.office365.com" -UseSsl -Port 587 -Credential $global:Cred    
    }

    # Done, write output.
    Write-Host "`nEmployee Created"
    Write-Host "=================="
    Write-Host "First name: $FirstName"
    Write-Host "Last name: $LastName"
    Write-Host "AD Username: $NewUsername"
    Write-Host "AD Password: $Password"
    Write-Host "ConsNo: $ConsultantNumber"
    Write-Host "Branch number: $TramsBranchNumber"
    Write-Host "Apollo username: $(New-ApolloAccount -FirstName $FirstName -LastName $LastName -Emulating $ApolloEmulating -BranchNumber $TramsBranchNumber)"
}

function Disable-Employee
{
    param
    (
        $Username,
        $BranchNumber,
        $Manager,
        $AutoReply
    )

    # Check if user exists.
    try
    {
        Get-ADuser -Identity $Username | Out-Null
    }
    catch
    {
        Write-Output "$Username not found."
        return
    }

    # Give manager full permissions to user's mailbox.
    if ($Manager)
    {
        Write-Host "Giving mailbox permissions to $Manager..."
        Edit-MailboxPermissions -Add -Master $Username -Slave $Manager | Out-Null
    }

    # Setting up automatic replies. 
    if ($AutoReply)
    {
        Write-Host "Setting out-of-office message..."
        # Getting phone number from SQL.
        $PhoneNumber = Invoke-SqlCmd2 -Query "SELECT areacode+br_ph as phonenumber FROM [redacted] WHERE (tramsbranchno = $BranchNumber)" -ServerInstance redacted -Database redacted | Select-Object -ExpandProperty phonenumber

        # Standardized message.
        $AutoReplyMessage = ConvertTo-Html -Body "**ALERT**
        <br>
        <br>
        Please be advised that I am no longer working at Maritime Travel and your email has been forwarded to my colleagues.
        <br>
        <br>
        For any further assistance please contact them by email at $BranchNumber@maritimetravel.ca or call the branch directly at $PhoneNumber.
        <br>
        <br>
        If your request is an emergency please contact our After Hours Emergency line at [redacted].
        <br>
        <br>
        <br>
        Thank You."

        # Enable message for internal and external replies.
        Set-MailboxAutoReplyConfiguration -Identity $Username -AutoReplyState enabled -InternalMessage $AutoReplyMessage -ExternalMessage $AutoReplyMessage
    }

    # Set mailbox as shared.
    Write-Host "Setting mailbox to type 'shared'..."
    Set-Mailbox -Identity $Username -Type Shared

    # Removing Office 365 License.
    Write-Host "Removing Office 365 license..."
    Set-MsolUserLicense -UserPrincipalName "$Username@maritimetravel.ca" -RemoveLicenses "weknowtravelbest:EXCHANGESTANDARD" -ErrorAction SilentlyContinue
    Set-MsolUserLicense -UserPrincipalName "$Username@maritimetravel.ca" -RemoveLicenses "weknowtravelbest:TEAMS_EXPLORATORY" -ErrorAction SilentlyContinue

    # Removing group memberships.
    Write-Host "Removing group memberships..."
    $Groups = (Get-ADUser -Identity $Username -Properties MemberOf | Select-Object MemberOf).MemberOf

    foreach ($Group in $Groups)
    {
        if (($Group -ne "Domain Users") -and ($Group -ne "Branch Users"))
        {
            Remove-ADGroupMember -Identity $Group -Members $Username -Confirm:$False
        }
    }

    # Disable AD Account.
    Write-Host "Disabling AD account..."
    Disable-ADAccount -Identity $Username

    # Disable using Jose's employee tool.
    Write-Host "Disabling in SQL..."
    $EmployeeNumber = Invoke-Sqlcmd2 -Query "SELECT EmployeeNo FROM [redacted] WHERE (username = '$Username')" -ServerInstance redacted -Database redacted | Select-Object -ExpandProperty EmployeeNo
    $IE = New-Object -ComObject "InternetExplorer.Application"
    $IE.Navigate("[redacted]")
    
    Write-Host "$Username has been disabled." -ForegroundColor Green
}


function Create-DistributionGroup
{
    param
    (
        [string]$GroupName,
        [string]$EmailAddress,
        [string]$Members
    )

    # Switch primary addresses
    if ($EmailAddress -match "@maritimetravel.ca")
    {
        $EmailAddress = $EmailAddress.Substring(0, $EmailAddress.IndexOf("@"))

        $PrimaryCompany = "$EmailAddress@maritimetravel.ca"
        $SecondaryCompany = "$EmailAddress@legrowstravel.ca"
        $TertiaryCompany = "$EmailAddress@voyagesmaritime.ca"
    }
    elseif ($EmailAddress -match "@legrowstravel.ca")
    {
        $EmailAddress = $EmailAddress.Substring(0, $EmailAddress.IndexOf("@"))

        $PrimaryCompany = "$EmailAddress@legrowstravel.ca"
        $SecondaryCompany = "$EmailAddress@maritimetravel.ca"
        $TertiaryCompany = "$EmailAddress@voyagesmaritime.ca"
    }
    elseif ($EmailAddress -match "@voyagesmaritime.ca")
    {

        $EmailAddress = $EmailAddress.Substring(0, $EmailAddress.IndexOf("@"))

        $PrimaryCompany = "$EmailAddress@voyagesmaritime.ca"
        $SecondaryCompany = "$EmailAddress@maritimetravel.ca"
        $TertiaryCompany = "$EmailAddress@legrowstravel.ca"
    }

    New-ADGroup -GroupCategory 0 -GroupScope 2 -Name $GroupName -Path "OU=MMT Distribution - Security Groups,DC=maritimetravel,DC=ca" -OtherAttributes @{'Mail'="$PrimaryCompany"; 'ProxyAddresses'="SMTP:$PrimaryCompany","smtp:$SecondaryCompany","smtp:$TertiaryCompany"}
    Write-Host "Creating group $GroupName..."
    Add-ADGroupMember -Identity $GroupName -Members $Members.Split(',')
    Write-Host "Adding members to group..."
    Sync
}

function Edit-MailboxPermissions
{
    param
    (
        [switch]$Add,
        [switch]$Remove,
        [string]$Master = $(throw "Please specify the mailbox owner."),
        [string]$Slave = $(throw "Please specify the mailbox permission holder.")
    )
    [void]$PSBoundParameters.ContainsKey('Add')
    [void]$PSBoundParameters.ContainsKey('Remove')

    # Check for add/remove arg.
    if ((($Add -eq $False) -and ($Remove -eq $False)) -or (($Add) -and ($Remove)))
    {
        throw "You must specify whether you are adding or removing permissions"
    }

    # Does one or the other.
    if ($Add)
    {
        Add-MailboxPermission $Master -User $Slave -AccessRights FullAccess -Confirm:$False
        Add-RecipientPermission $Master -Trustee $Slave -AccessRights SendAs -Confirm:$False
    }

    if ($Remove)
    {
        Remove-MailboxPermission $Master -User $Slave -AccessRights FullAccess -Confirm:$False
        Remove-RecipientPermission $Master -Trustee $Slave -AccessRights SendAs -Confirm:$False
    }

}

function New-ADPassword
{
    $Colours = @("Red","Green","Blue","Grey","White","Pink")
    $Things = @("cats","goats","bats","hats","boats","pigs","dogs","books","bikes","ships","lions","mice","birds")
    $Symbols = @("!","?","#","$","&","@","%","*")
    
    $Password = "$(Get-Random -Minimum 2 -Maximum 10)" + "$(Get-Random $Colours)" + "$(Get-Random $Things)" + "$(Get-Random $Symbols)"

    return $Password
}

function Reset-ADPassword
{
    param
    (
        [string]$Username,
        [string]$Password,
        [string]$Temporary
    )
    [void]$PSBoundParameters.ContainsKey('Temporary')

    # If password is not defined, generates a random password.
    if (!$Password)
    {
        $Password = New-ADPassword
    }

    try
    {
        Set-ADAccountPassword -Identity $Username -Reset -NewPassword (Convertto-Securestring -AsPlainText $Password -Force)
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
    {
        Write-Host "$Username not found." -ForegroundColor Red
    }
    catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException]
    {
        Write-Host "Password is not complex enough." -ForegroundColor Red
    }

    Write-Host "$Username was given password $Password." -ForegroundColor Green
      
    if ($Temporary)
    {
        Set-ADUser $Username -ChangePasswordAtLogon $True
        Write-Host "$Username must change password at logon." -ForegroundColor Yellow
    }
}


function Unlock-User
{
    param
    (
        [string]$Username
    )

    
    # Checks if user exists, unlocks them if they do. Terminates if not.
    try 
    {
        $Name = Get-ADUser -Identity $Username | Select-Object -ExpandProperty name
        Unlock-ADAccount -Identity $Username
        Write-Host "$Name has been unlocked." -ForegroundColor Green
   
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
    {
        Write-Host "$Username not found." -ForegroundColor Red
    }
}

function Unlock-AllUsers
{
    $Names = Search-ADAccount -LockedOut | Select-Object -ExpandProperty samaccountname
    foreach ($Name in $Names)
    {
        Unlock-ADAccount -Identity $Name
        Write-Host "Unlocked $Name." -ForegroundColor Green
    }
}

function Get-Uptime
{
    param
    (
        [string]$ComputerName
    )

    try
    {
        $BootDate = Invoke-Command -ComputerName $ComputerName -ScriptBlock `
        {
            [System.Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem).LastBootUpTime)
        }

        $Uptime = ((Get-Date) - $BootDate)
        Write-Host "Boot date: $($BootDate.DateTime)"
        Write-Host "Total uptime: $($Uptime.Days) Days, $($Uptime.Hours) Hours, $($Uptime.Minutes) Minutes."
    }
    catch
    {
        Write-Host "Could not get uptime for specified computer." -ForegroundColor Red
    }
}

function Rebuild-Index
{
    param
    (
        [string]$Computer
    )

    Write-Host "Please wait..."

    # Get free disk space before rebuilding index.
    $Disk = Get-WmiObject Win32_LogicalDisk -ComputerName $Computer -Filter "DeviceID='C:'" | Select-Object Size,FreeSpace
    $FreeSpaceBefore = [math]::Round(($Disk.FreeSpace / 1GB), 3)

    # Remotely running script on target computer.
    Invoke-Command -ComputerName $Computer -ScriptBlock `
    {
        # Stops Windows Search Service.
        Stop-Service wsearch

        # Removes Search Index file.
        $Directory = "$ENV:ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"
        Remove-Item -LiteralPath $Directory

        # While windows is rebuilding the index, loop until Windows Search service is successfully started.
        while (!$Success)
        {
            $Service = Get-Service wsearch

            if ($Service.Status -eq "Running")
            {
                $Success = $True
            }
            else
            {
                Start-Service wsearch -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Start-Sleep 3
            }
        }
    }

    # Get free disk space after rebuilding index.
    $Disk = Get-WmiObject Win32_LogicalDisk -ComputerName $Computer -Filter "DeviceID='C:'" | Select-Object Size,FreeSpace
    $FreeSpaceAfter = [math]::Round(($Disk.FreeSpace / 1GB), 3)

    Write-Output "Free space before: $FreeSpaceBefore GB"
    Write-Output "Free space after: $FreeSpaceAfter GB"
}

# Does a lookup on computer's public ip address.
function Get-PublicIP
{
    param 
    (
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName { Invoke-RestMethod http://ipinfo.io/json | Select-Object -ExpandProperty ip }
}

# Kills Smartpoint/Apollo tasks and restarts SSL service.
function Fix-Smartpoint 
{
    param
    (
        [string]$Computer
    )

    Invoke-Command -ComputerName $Computer -ScriptBlock `
    {
        Write-Host "Killing Apollo/Smartpoint tasks..."
        taskkill /im hcmmux.exe /im viewpoint.exe /im viewpointlistener.exe /im travelport.smartpoint.app.exe /im travelport.smartpoint.startup.exe /f
        if (Get-Service -name "Galileo SSL Tunnel" -ea SilentlyContinue) 
        {
            Write-Host "Restarting Galileo SSL Serice..."
            Restart-Service -name "Galileo SSL Tunnel"
        } 
    }
}

function Fix-Printer
{
    param
    (
        [int]$BranchNumber
    )

    $LocationID = Invoke-Sqlcmd2 -Query "SELECT LocationID FROM [redacted] WHERE Realstart = '172.17.$BranchNumber.1'" -ServerInstance redacted -Database redacted | Select-Object -ExpandProperty LocationID
    $ComputerArray = @(Invoke-Sqlcmd2 -Query "SELECT AssetName FROM [redacted] WHERE LocationID = '$LocationID' AND AssetType = '-1'" -ServerInstance redacted -Database redacted | Select-Object -ExpandProperty AssetName)
    Invoke-Command -ComputerName $ComputerArray -ErrorAction SilentlyContinue `
    {
        $Path = "C:\Windows\System32\Spool\Printers"

        Stop-Service spooler -Force

        if (Test-Path -Path $Path)
        {
            Remove-Item -path c:\windows\system32\spool\printers -Force -Recurse
        }
        Start-Service spooler
        return "Clearing queues for $ENV:COMPUTERNAME."
    }

    Write-Output "Cleared queues and restarted printers for everyone in Branch $BranchNumber.`nYou may need to run this script again if it doesn't work the first time."
}

function Send-Mergeback 
{
    param
    (
        [array]$BookingNumber,
        [string]$Username
    )

    $BookingNumbers = $BookingNumber.Replace(' ','').Split(',')

    foreach ($BookingNumber in $BookingNumbers)
    {
        $BookingExists = Invoke-Sqlcmd2 -Query "
        Select * from [redacted] where entryId 
        in
        (
              Select xmlEntryId from [redacted] where BookingNumber = '$BookingNumber'
        )
        " -ServerInstance redacted -Database redacted 

        if ($BookingExists)
        {
            Invoke-Sqlcmd2 -Query "[redacted]" -ServerInstance redacted -Database redacted | Out-Null

            Write-Host "Booking number $BookingNumber has been sent to $Username." -ForegroundColor Green
        }
        else
        {
            Write-Host "Booking number $BookingNumber not found." -ForegroundColor Red
        }
    }
}

####################################################
#              SECTION THREE: MENUS                #
####################################################

# Creates a new employee by copying existing employee.
function Start-menuNewEmployee
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Creates a new employee from a template."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "Template" -Message "Template's username:" -ValidateNotEmpty
        New-AnyBoxPrompt -Group "Agent Information" -InputType Text -Name "Username" -Message "Username:" -ValidateNotEmpty
        New-AnyBoxPrompt -Group "Agent Information" -InputType Text -Name "FirstName" -Message "First name:" -ValidateNotEmpty
        New-AnyBoxPrompt -Group "Agent Information" -InputType Text -Name "LastName" -Message "Last name:" -ValidateNotEmpty
        New-AnyBoxPrompt -Group "Agent Information" -InputType Text -Name "Branch" -Message "Branch:" -ValidateSet $global:BranchHash.Key
        New-AnyBoxPrompt -Group "Agent Information" -InputType Text -Name "Career" -Message "Job title:" -ValidateSet $global:JobHash.Keys
        New-AnyBoxPrompt -Group "Agent Information" -InputType Checkbox -Name "EmailInsurance" -Message "Send email to ORI and TIS"
        New-AnyBoxPrompt -Group "Agent Information" -InputType Checkbox -Name "EmailWelcome" -Message "Send welcome email" 
        New-AnyBoxPrompt -Group "Apollo" -InputType Text -Name "Apollo" -ValidateSet "Emulating", "Non-Emulating" -ShowSetAs Radio_Wide -ValidateNotEmpty
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $UserInput = $Box | Show-AnyBox
    
    if ($UserInput['Submit'])
    {
        Get-ConnectionStatus

        if ($UserInput.Apollo -eq "Emulating")
        {
            $UserInput.Apollo = $true
        }
        else
        {
            $UserInput.Apollo = $false
        }

        # Unchecked checkbox returns an empty string, not null or false. This should be improved.
        if ($UserInput.EmailInsurance -ne $true)
        {
            $UserInput.EmailInsurance = $false
        }

        if ($UserInput.EmailHolidayEscapes -ne $true)
        {
            $UserInput.EmailHolidayEscapes = $false
        }

        if ($UserInput.EmailWelcome -ne $true)
        {
            $UserInput.EmailWelcome = $false
        }
        
        # Splatting.
        $EmployeeDetails =
        @{
            TemplateUser = $UserInput.Template
            NewUsername = $UserInput.Username
            FirstName = $UserInput.FirstName
            LastName = $UserInput.LastName
            BranchName = $UserInput.Branch
            JobTitle = $UserInput.Career
            ApolloEmulating = $UserInput.Apollo
            EmailHolidayEscapes = $UserInput.EmailHolidayEscapes
            EmailInsurance = $UserInput.EmailInsurance
            EmailWelcome = $UserInput.EmailWelcome
        }
        New-Employee @EmployeeDetails
    }

    if ($UserInput['Cancel'])
    {
        Start-menuUserAccountTools
    }
}

# Disables an existing employee.
function Start-menuDisableEmployee
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Disables employee account."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "Username" -Message "Username:" -ValidateNotEmpty
        New-AnyBoxPrompt -InputType Text -Name "BranchNumber" -Message "Branch Number:" -ValidateNotEmpty
        New-AnyBoxPrompt -InputType Text -Name "Manager" -Message "Forward Emails To:"
        New-AnyBoxPrompt -InputType Checkbox -Name "Autoreply" -Message "Generic Autoreply Message"
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $Box.Comment = "Leave forwarding field blank if emails are not to be forwarded."
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Get-ConnectionStatus

        if ($UserInput.Manager.Length -eq 0)
        {
            $UserInput.Manager = $false
        }

        if ($UserInput.Autoreply -ne $true)
        {
            $UserInput.Autoreply = $false
        }

        Disable-Employee -Username $UserInput.Username -BranchNumber $UserInput.BranchNumber -Manager $UserInput.Manager -Autoreply $UserInput.Autoreply
    }

    if ($UserInput['Cancel'])
    {
        Start-menuUserAccountTools
    }
}

# Adds or removes mailbox privileges to specified user(s).
function Start-menuMailboxPermissions
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Grants or removes mailbox permissions for a given user."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "Master" -Message "Mailbox owner:" -ValidateNotEmpty
        New-AnyBoxPrompt -InputType Text -Name "Slave" -Message "Mailbox permission holder:" -ValidateNotEmpty
        New-AnyBoxPrompt -InputType Text -Name "Type" -ValidateSet "Add", "Remove" -ShowSetAs Radio_Wide -ValidateNotEmpty
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Get-ConnectionStatus

        if ($UserInput.Type -eq "Add")
        {
            Write-Host "Please wait..."
            Edit-MailboxPermissions -Add -Master $UserInput.Master -Slave $UserInput.Slave
            Write-Host "$($UserInput.Slave) has been given full permissions to $($UserInput.Master)'s mailbox." -ForegroundColor Green
        }
        elseif ($UserInput.Type -eq "Remove")
        {
            Write-Host "Please wait..."
            Edit-MailboxPermissions -Remove -Master $UserInput.Master -Slave $UserInput.Slave
            Write-Host "$($UserInput.Slave) has had their permissions to $($UserInput.Master)'s mailbox revoked." -ForegroundColor Green
        }
        else
        {
            Write-Host "You must specify whether you are adding or removing permissions." -ForegroundColor Red
        }
    }

    if ($UserInput['Cancel'])
    {
        Start-menuEmailTools
    }
}

# Sets mailbox to either "Regular" or "Shared".
function Start-menuMailboxType
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Sets mailbox type as either Regular or Shared."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "Mailbox" -Message "Mailbox Username:" -ValidateNotEmpty
        New-AnyBoxPrompt -InputType Text -Name "Type" -ValidateSet "Regular", "Shared" -ShowSetAs Radio_Wide -ValidateNotEmpty
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Get-ConnectionStatus
        Set-Mailbox -Identity $UserInput.Mailbox -Type $UserInput.Type -ErrorAction SilentlyContinue
        Write-Host "$($UserInput.Mailbox) has been set to a $($UserInput.Type.ToLower()) mailbox."
    }

    if ($UserInput['Cancel'])
    {
        Start-menuEmailTools
    }
}

function Start-menuCreateDistributionGroup
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Creates a new AD group and populates with given users."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "GroupName" -Message "Group name:" -ValidateNotEmpty
        New-AnyBoxPrompt -InputType Text -Name "EmailAddress" -Message "Primary email address:" -ValidateNotEmpty
        New-AnyBoxPrompt -InputType Text -Name "Members" -Message "Group members (comma separated):" -ValidateNotEmpty
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Create-DistributionGroup -GroupName $UserInput.GroupName -EmailAddress $UserInput.EmailAddress -Members $UserInput.Members
    }

    if ($UserInput['Cancel'])
    {
        Start-menuUserAccountTools
    }
}

# Resets AD password.
function Start-menuResetPassword
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Resets a given user's AD password."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "Username" -Message "Username:" -ValidateNotEmpty
        New-AnyBoxPrompt -InputType Text -Name "Password" -Message "Password:"
        New-AnyBoxPrompt -InputType Checkbox -Name Temporary -Message "Temporary"
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Reset" -Text "Reset" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $Box.Comment = "If password is left blank, a random password will be used."
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Reset'])
    {
        Reset-ADPassword -Username $UserInput.Username -Password $UserInput.Password -Temporary $UserInput.Temporary
    }

    if ($UserInput['Cancel'])
    {
        Start-menuUserAccountTools
    }
}

# Unlocks a single user.
function Start-menuUnlockUser
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Unlocks a given account, or all AD accounts."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "Username" -Message "Username:"
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "UnlockAll" -Text "Unlock All"
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Unlock-User -Username $UserInput.Username
    }

    if ($UserInput['UnlockAll'])
    {
        Unlock-AllUsers
    }

    if ($UserInput['Cancel'])
    {
        Start-menuUserAccountTools
    }
}

# Gets uptime for a single computer. Exposes liars.
function Start-menuUptime
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Checks how long a given computer has been powered on for. Exposes liars."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "Computer" -Message "Computer name:" -ValidateNotEmpty
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Get-Uptime -ComputerName $UserInput.Computer
    }

    if ($UserInput['Cancel'])
    {
        Start-menuDiagnostics
    }
}

# Rebuilds search index on C drive. Often frees up much space.
function Start-menuRebuildIndex
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Rebuilds search index on C:/ drive."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "Computer" -Message "Computer name:" -ValidateNotEmpty
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Rebuild-Index -Computer $UserInput.Computer
    }

    if ($UserInput['Cancel'])
    {
        Start-menuDiagnostics
    }
}

# Checks if given credentials are valid.
function Start-menuGetPublicIP
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Uses a web service to look up public IP of given computer."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "Computer" -Message "Computer Name:" -ValidateNotEmpty
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Get-PublicIP -ComputerName $UserInput.Computer
    }

    if ($UserInput['Cancel'])
    {
        Start-menuDiagnostics
    }
}

function Start-menuFixSmartpoint
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "For a given computer, kills Apollo/Smartpoint tasks and restarts SSL service."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "Computer" -Message "Computer Name:" -ValidateNotEmpty
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Fix-Smartpoint -Computer $UserInput.Computer
    }

    if ($UserInput['Cancel'])
    {
        Start-menuDiagnostics
    }
}

function Start-menuFixPrinter
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "For a given branch, clears print queues and restarts spooler service."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "BranchNumber" -Message "Branch number:" -ValidateNotEmpty
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Fix-Printer -BranchNumber $UserInput.BranchNumber
    }

    if ($UserInput['Cancel'])
    {
        Start-menuDiagnostics
    }
}


# Renames a computer in AD.
function Start-menuRenameComputer
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Renames a given computer."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "OldComputerName" -Message "Old Computer Name:" -ValidateNotEmpty
        New-AnyboxPrompt -InputType Text -Name "NewComputerName" -Message "New Computer Name:" -ValidateNotEmpty
        New-AnyBoxPrompt -InputType Checkbox -Name "RebootImmediately" -Message "Reboot Immediately"
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $UserInput = $Box | Show-AnyBox  

    if ($UserInput['Submit'])
    {
        Write-Host "Renaming computer..."
        Rename-Computer -ComputerName $UserInput.OldComputerName -NewName $UserInput.NewComputerName -DomainCredential $global:Cred -Force -WarningAction SilentlyContinue

        if ($UserInput.RebootImmediately)
        {
            Write-Host "Restarting computer..."
            Restart-Computer -ComputerName $UserInput.OldComputerName -Force
        }

        Write-Host "Renaming complete."
    }

    if ($UserInput['Cancel'])
    {
        Start-menuMiscellaneous
    }
}

# Manually sends Galileo Vacations "mergebacks" to a specified user.
function Start-menuSendMergeback
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Checks if mergeback exists and re-sends to given user."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "BookingNumber" -Message "Booking Number(s):" -ValidateNotEmpty
        New-AnyboxPrompt -InputType Text -Name "Username" -Message "Username" -ValidateNotEmpty
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $Box.Comment = "For multiple bookings, separate by comma."
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Send-Mergeback -BookingNumber $UserInput.BookingNumber -Username $UserInput.Username
    }

    if ($UserInput['Cancel'])
    {
        Start-menuMiscellaneous
    }
}

# Executes a custom PS command in the console.
function Start-menuCustomCommand
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Executes given code in PowerShell terminal."
    $Box.Prompts = 
    @(
        New-AnyBoxPrompt -InputType Text -Name "Command" -Message "PowerShell Command:" -ValidateNotEmpty
    )
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "Submit" -Text "Submit" -IsDefault
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $Box.Comment = "Please use this feature with caution, you are executing unchecked code."
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['Submit'])
    {
        Get-ConnectionStatus
        Invoke-Expression -Command $UserInput.Command
    }

    if ($UserInput['Cancel'])
    {
        Start-menuMiscellaneous
    }
}

function Start-menuMain
{
    while(1)
    {
        Get-ConnectionStatus
        $Box = New-Object AnyBox.AnyBox
        $Box.Title = $Title
        $Box.FontSize = $FontSize
        $Box.MinWidth = $MinWidth
        $Box.BackgroundColor = $BackgroundColor
        $Box.FontColor = $FontColor
        $Box.Message = "Welcome to EZPZ!"
        $Box.Buttons =
        @(
            New-AnyBoxButton -Name "UserAccountTools" -Text "User Account Tools"
            New-AnyBoxButton -Name "EmailTools" -Text "Email Tools"
            New-AnyBoxButton -Name "Diagnostics" -Text "Diagnostics & Fixes"
            New-AnyBoxButton -Name "Miscellaneous" -Text "Miscellaneous"

        )
        $Box.ButtonRows = $Box.Buttons.Count
        $UserInput = $Box | Show-AnyBox

        # Menu Choices
        if ($UserInput['UserAccountTools'])
        {
            Start-menuUserAccountTools
        }

        if ($UserInput['EmailTools'])
        {
            Start-menuEmailTools
        }

        if ($UserInput['Diagnostics'])
        {
            Start-menuDiagnostics
        }

        if ($UserInput['Miscellaneous'])
        {
            Start-menuMiscellaneous
        }
    }
}

function Start-menuUserAccountTools
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "User Account Tools"
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "NewEmployee" -Text "New Employee"
        New-AnyBoxButton -Name "DisableEmployee" -Text "Disable Employee"
        New-AnyBoxButton -Name "CreateDistributionGroup" -Text "Create Distribution Group"
        New-AnyBoxButton -Name "ResetADPassword" -Text "Reset AD Password"
        New-AnyBoxButton -Name "UnlockADAccount" -Text "Unlock AD Account"
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $Box.ButtonRows = $Box.Buttons.Count
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['NewEmployee'])
    {
        Start-menuNewEmployee
    }

    if ($UserInput['DisableEmployee'])
    {
        Start-menuDisableEmployee
    }

    if ($UserInput['CreateDistributionGroup'])
    {
        Start-menuCreateDistributionGroup
    }

    if ($UserInput['ResetADPassword'])
    {
        Start-menuResetPassword
    }

    if ($UserInput['UnlockADAccount'])
    {
        Start-menuUnlockUser
    }

    if ($UserInput['Cancel'])
    {
        Start-menuMain
    }
}

function Start-menuEmailTools
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Email Tools"
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "MailboxPermissions" -Text "Mailbox Permissions"
        New-AnyBoxButton -Name "MailboxType" -Text "Mailbox Type"
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $Box.ButtonRows = $Box.Buttons.Count
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['MailboxPermissions'])
    {
        Start-menuMailboxPermissions
    }

    if ($UserInput['MailboxType'])
    {
        Start-menuMailboxType
    }

    if ($UserInput['Cancel'])
    {
        Start-menuMain
    }
}

function Start-menuDiagnostics
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Diagnostics & Fixes"
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "ComputerUptime" -Text "Computer Uptime"
        New-AnyBoxButton -Name "RebuildIndex" -Text "Rebuild Index"
        New-AnyBoxButton -Name "GetPublicIP" -Text "Get Public IP"
        New-AnyBoxButton -Name "FixSmartpoint" -Text "Fix Smartpoint"
        New-AnyBoxButton -Name "FixPrinter" -Text "Fix Printer"
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $Box.ButtonRows = $Box.Buttons.Count
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['ComputerUptime'])
    {
        Start-menuUptime
    }

    if ($UserInput['RebuildIndex'])
    {
        Start-menuRebuildIndex
    }

    if ($UserInput['GetPublicIP'])
    {
        Start-menuGetPublicIP
    }

    if ($UserInput['FixSmartpoint'])
    {
        Start-menuFixSmartpoint
    }
    
    if ($UserInput['FixPrinter'])
    {
        Start-menuFixPrinter
    }
    
    if ($UserInput['Cancel'])
    {
        Start-menuMain
    }
}

function Start-menuMiscellaneous
{
    $Box = New-Object AnyBox.AnyBox
    $Box.Title = $Title
    $Box.FontSize = $FontSize
    $Box.MinWidth = $MinWidth
    $Box.BackgroundColor = $BackgroundColor
    $Box.FontColor = $FontColor
    $Box.Message = "Miscellaneous"
    $Box.Buttons =
    @(
        New-AnyBoxButton -Name "ADSync" -Text "AD Sync"
        New-AnyBoxButton -Name "RenameComputer" -Text "Rename Computer"
        New-AnyBoxButton -Name "SendMergeback" -Text "Send Mergeback"
        New-AnyBoxButton -Name "CustomCommand" -Text "Custom Command"
        New-AnyBoxButton -Name "Cancel" -Text "Cancel" -IsCancel
    )
    $Box.ButtonRows = $Box.Buttons.Count
    $UserInput = $Box | Show-AnyBox

    if ($UserInput['ADSync'])
    {
        Start-ADSync
    }

    if ($UserInput['RenameComputer'])
    {
        Start-menuRenameComputer
    }

    if ($UserInput['SendMergeback'])
    {
        Start-menuSendMergeback
    }

    if ($UserInput['CustomCommand'])
    {
        Start-menuCustomCommand
    }

    if ($UserInput['Cancel'])
    {
        Start-menuMain
    }
}

function Main
{
    # Make sure credentials are properly stored.
    Write-Host "Verifying credentials..."
    #Confirm-Credentials -CredentialName "[redacted]@maritimetravel.ca"

    # Load modules and connect to Office 365.
    $global:WarningPreference = "SilentlyContinue"
    Write-Host "Loading AnyBox..."
    Import-Module AnyBox
    Write-Host "Loading SQL..."
    Import-module SQLPS
    Write-Host "Loading Active Directory..."
    Import-Module ActiveDirectory
    Write-Host "Connecting to Office 365..."
    #Connect-Office365
    $global:WarningPreference = "Continue"

    # Welcome screen.
    Write-Host "`n...`n...`n...`n`n"
    Write-Host "Welcome to EZPZ."
    Write-Host "Version 0.2 Alpha`n`n"
    Start-menuMain
}

# AnyBox constants.
$Title = "EZPZ - Maritime Travel"
$FontSize = 18
$MinWidth = 400
$BackgroundColor = "#141414"
$FontColor = "#ffffff"

# System constants.
$global:Cred = Get-StoredCredentials -UserName z_ews@maritimetravel.ca
$global:KeyPath = ".\Credentials"
$global:BranchHash = Get-BranchHash
$global:JobHash = 
@{
    "Accountant"          = "37"
    "Branch Admin"        = "46"
    "Branch Manager"      = "12"
    "Coordinator, BTM"    = "20"
    "Helpdesk Specialist" = "87"
    "Outside Sales"       = "84"
    "Student"             = "106"
    "Travel Consultant"   = "8"
}

Main