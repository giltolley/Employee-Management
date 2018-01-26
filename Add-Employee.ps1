#========================================================================
# _______________________________________________
# |Author:      |
# |Gil Tolley                                   |
# |7/19/2017 								    |
# |												|
# |Last Updated:               			        |
# |1/26/2018 		   						    |
# |_____________________________________________|
#========================================================================
#----------------------------------------------
#region Application Functions
#----------------------------------------------

Function OnApplicationLoad {
$CreateXML = @"
<?xml version="1.0" standalone="no"?>
<OPTIONS Product="Employee Management Tool" Version="1.8.5">
 <Settings>
  <sAMAccountName Generate="True">
   <Style Format="FirstName.LastName" Enabled="False" />
   <Style Format="FirstInitialLastName" Enabled="True" />
   <Style Format="LastNameFirstInitial" Enabled="False" />
  </sAMAccountName>
  <UPN Generate="True">
   <Style Format="FirstName.LastName" Enabled="False" />
   <Style Format="FirstInitialLastName" Enabled="True" />
   <Style Format="LastNameFirstInitial" Enabled="False" />
  </UPN>
  <DisplayName Generate="True">
   <Style Format="FirstName LastName" Enabled="True" />
   <Style Format="LastName, FirstName" Enabled="False" />
  </DisplayName>
  <EmailAddress Generate="True">
   <Style Format="FirstNameLastName" Enabled="True" />
   <Style Format="FirstInitialLastName" Enabled="False" />
   <Style Format="LastNameFirstInitial" Enabled="False" />
  </EmailAddress>
  <AccountStatus Enabled="True" />
  <Password ChangeAtLogon="False" />
  <Description Generate="True" />
 </Settings>
 <Default>
  <Domain>company.com</Domain>
  <Path>CN=Users,DC=company,DC=com</Path>
  <FirstName></FirstName>
  <LastName></LastName>
  <Office>Corp Office</Office>
  <Title></Title>
  <Description></Description>
  <Department></Department>
  <Company>Example Company LLC</Company>
  <Phone></Phone>
  <Site>TX</Site>
  <StreetAddress>1 Highway Street</StreetAddress>
  <City>Kingwood</City>
  <State>TX</State>
  <PostalCode>77339</PostalCode>
  <Password></Password>
  <MemberOf></MemberOf>
  <UPNSuffix>company.com</UPNSuffix>
 </Default>
 <Locations>
  <Location Site="NH">
   <StreetAddress>1 Central Avenue</StreetAddress>
   <City>Dover</City>
   <State>NH</State>
   <Office>Dover, NH</Office>
   <PostalCode>03820</PostalCode>
   <Groups>
    <Group>Dover Office</Group>
	<Group>Remote Office Employees</Group>
   </Groups>
  </Location>
  <Location Site="TX">
   <StreetAddress>1 Highway Street</StreetAddress>
   <City>Kingwood</City>
   <State>TX</State>
   <Office>Corp Office</Office>
   <PostalCode>77339</PostalCode>
   <Groups>
    <Group>Home Office Employees</Group>
   </Groups>
  </Location>
  <Location Site="Custom">
   <StreetAddress></StreetAddress>
   <City></City>
   <State></State>
   <Office></Office>
   <PostalCode></PostalCode>
   <Groups>
   </Groups>
  </Location>
 </Locations>
 <Domains>
  <Domain Name="company.com">
   <Path>CN=Users,DC=Company,DC=Com</Path>
   <Path>OU=Accounting,OU=Kingwood Users,OU=Kingwood,DC=Company,DC=Co</Path>
   <Path>OU=Asset Management,OU=Kingwood Users,OU=Kingwood,DC=Company,DC=Co</Path>
   <Path>OU=Credit Processing,OU=Kingwood Users,OU=Kingwood,DC=Company,DC=Co</Path>
   <Path>OU=General Admin,OU=Kingwood Users,OU=Kingwood,DC=Company,DC=Co</Path>
   <Path>OU=Human Resources,OU=Kingwood Users,OU=Kingwood,DC=Company,DC=Co</Path>
   <Path>OU=Sales Center,OU=Kingwood Users,OU=Kingwood,DC=Company,DC=Co</Path>
   <Path>OU=SOHO Users,OU=SOHO,DC=Company,DC=Co</Path>
   <Path>OU=Merchant Finance Users,OU=Merchant Finance,DC=Company,DC=Co</Path>
   <Path>OU=Technology Users,OU=Technology,DC=Company,DC=Co</Path>
  </Domain>
 </Domains>
 <UPNSuffixes>
  <UPNSuffix>company.com</UPNSuffix>
  <UPNSuffix>onmicrosoft.company.com</UPNSuffix>
 </UPNSuffixes>
 <Titles>
  <Title>Account Representative</Title>
  <Title>Asset Manager</Title>
  <Title>VP Sales</Title>
  <Title>Finance Manager</Title>
  <Title>Service Account</Title>
  <Title>Contract Administrator</Title>
  <Title>Funding Manager</Title>
  <Title>Sales and Funding Coordinator</Title>
  <Title>Customer Service Representative</Title>
  <Title>Credit Analyst</Title>
  <Title>Loan Administrator</Title>
  <Title>Records Administrator</Title>
  <Title>Regional Sales Manager</Title>
  <Title>Software Engineer</Title>
  <Title>Systems and Network Administrator</Title>
  <Title>Vendor Services Representative</Title>
 </Titles>
 <Descriptions>
  <Description>Account Representative</Description>
  <Description>Asset Manager</Description>
  <Description>VP Sales</Description>
  <Description>Finance Manager</Description>
  <Description>Service Account</Description>
  <Description>Contract Administrator</Description>
  <Description>Funding Manager</Description>
  <Description>Sales and Funding Coordinator</Description>
  <Description>Customer Service Representative</Description>
  <Description>Credit Analyst</Description>
  <Description>Loan Administrator</Description>
  <Description>Records Administrator</Description>
  <Description>Regional Sales Manager</Description>
  <Description>Software Engineer</Description>
  <Description>Systems and Network Administrator</Description>
  <Description>Vendor Services Representative</Description>
 </Descriptions>
 <Departments>
  <Department Name="Accounting">
   <DefaultGroups>
    <DefaultGroup>Accounting Team</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
	<SuggestedGroup>VPN Users</SuggestedGroup>
	<SuggestedGroup>Controllership Team</SuggestedGroup>
	<SuggestedGroup>Funding Team</SuggestedGroup>
	<SuggestedGroup>Contract Administration Team</SuggestedGroup>
	<SuggestedGroup>Accounting Management</SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Asset Management">
   <DefaultGroups>
    <DefaultGroup>Collections Team</DefaultGroup>
	<DefaultGroup>Images Group</DefaultGroup>
	<DefaultGroup>LeasePlus Users</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
	<SuggestedGroup>LeaseCompleteUsers</SuggestedGroup>
	<SuggestedGroup>VPN Users</SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Capital Markets">
   <DefaultGroups>
    <DefaultGroup>LeasePlus Users</DefaultGroup>
	<DefaultGroup>Marketplace Sales</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup>DIS Scanners</SuggestedGroup>
	<SuggestedGroup>VPN Users</SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Collections">
   <DefaultGroups>
    <DefaultGroup>Collections Team</DefaultGroup>
	<DefaultGroup>Images Group</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup>DIS View-Print</SuggestedGroup>
	<SuggestedGroup>VPN Users</SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Credit">
   <DefaultGroups>
    <DefaultGroup>Credit Team</DefaultGroup>
	<DefaultGroup>VPN Users</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup>Mid Ticket Credit Team</SuggestedGroup>
	<SuggestedGroup>Small Ticket Credit Team</SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Customer Service">
   <DefaultGroups>
    <DefaultGroup>Customer Service Team</DefaultGroup>
	<DefaultGroup>LeasePlus Users</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup>Contract Administration Team</SuggestedGroup>
	<SuggestedGroup>DIS Scanners</SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Direct Sales">
   <DefaultGroups>
    <DefaultGroup>Direct Sales Team</DefaultGroup>
	<DefaultGroup>VPN Users</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup></SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Executive">
   <DefaultGroups>
    <DefaultGroup></DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup></SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Human Resources">
   <DefaultGroups>
    <DefaultGroup>HR Department</DefaultGroup>
	<DefaultGroup>VPN Users</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup></SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Independent Contractor">
   <DefaultGroups>
    <DefaultGroup></DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup></SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Information Technology">
   <DefaultGroups>
    <DefaultGroup>Technology Team</DefaultGroup>
	<DefaultGroup>VPN Users</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup>Network Operations Team</SuggestedGroup>
	<SuggestedGroup>Development Team</SuggestedGroup>
	<SuggestedGroup>RODC Admins</SuggestedGroup>
	<SuggestedGroup>Not Moderated</SuggestedGroup>
	<SuggestedGroup>Business Analysts</SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Legal">
   <DefaultGroups>
    <DefaultGroup>VPN Users</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup></SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Marketing">
   <DefaultGroups>
    <DefaultGroup>Marketing Team</DefaultGroup>
	<DefaultGroup>Sales Team</DefaultGroup>
	<DefaultGroup>VPN Users</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup>Not Moderated</SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Merchant Finance">
   <DefaultGroups>
    <DefaultGroup>Merchant Finance</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
	<SuggestedGroup>Images Group</SuggestedGroup>
	<SuggestedGroup>VPN Users</SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Risk Management">
   <DefaultGroups>
    <DefaultGroup>VPN Users</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup>Technology Team</SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Sales">
   <DefaultGroups>
    <DefaultGroup>Sales Team</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup></SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Vendor Originations">
   <DefaultGroups>
    <DefaultGroup>Sales Team</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
    <SuggestedGroup>VPN Users</SuggestedGroup>
   </SuggestedGroups>
  </Department>
  <Department Name="Vendor Services">
   <DefaultGroups>
    <DefaultGroup>VSR Team</DefaultGroup>
   </DefaultGroups>
   <SuggestedGroups>
	<SuggestedGroup>VPN Users</SuggestedGroup>
	<SuggestedGroup>Tilting Team</SuggestedGroup>
	<SuggestedGroup>Ops Team</SuggestedGroup>
   </SuggestedGroups>
  </Department>
 </Departments>
 <Groups>
  <Group>Accounting Team</Group>
  <Group>VPN Users</Group>
  <Group>Controllership Team</Group>
  <Group>Funding Team</Group>
  <Group>Contract Administration Team</Group>
  <Group>Accounting Management</Group>
  <Group>Collections Team</Group>
  <Group>Images Group</Group>
  <Group>Marketplace Sales</Group>
  <Group>DIS Scanners</Group>
  <Group>Credit Team</Group>
  <Group>Mid Ticket Credit Team</Group>
  <Group>Small Ticket Credit Team</Group>
  <Group>Direct Sales Team</Group>
  <Group>HR Department</Group>
  <Group>Marketing Team</Group>
  <Group>Sales Team</Group>
  <Group>Not Moderated</Group>
  <Group>Merchant Finance</Group>
  <Group>Technology Team</Group>
  <Group>Tilting Team</Group>
  <Group>Ops Team</Group>
  <Group>Remote Desktop Users</Group>
 </Groups>
</OPTIONS>
"@
	
	# Load the XML file and gather global info, such as the filepath
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$XMLOptions = "data/EmployeeManagement.Options.xml"
	$Script:ParentFolder = Split-Path (Get-Variable MyInvocation -scope 1 -ValueOnly).MyCommand.Definition
	$XMLFile = Join-Path $ParentFolder $XMLOptions
	
	$XMLMsg = "Configuration file $XMLOptions not detected in folder $ParentFolder.  Would you like to create one now?"
	if(!(Test-Path $XMLFile)){
 	   if([System.Windows.Forms.MessageBox]::Show($XMLMsg,"Warning",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq "Yes")
            {
				$CreateXML | Out-File $XMLFile
				$TemplateMsg = "Opening XML configuration file for editing ($XMLFile).  Please relaunch the script when the configuration is complete."
				[System.Windows.Forms.MessageBox]::Show($TemplateMsg,"Information",[System.Windows.Forms.MessageBoxButtons]::Ok) | Out-Null
				notepad $XMLFile
				Exit
	 	    }
        else{Exit}
	}
	else{[XML]$Script:XML = Get-Content $XMLFile}
    if($XML.Options.Version -ne ([xml]$CreateXML).Options.Version)
        {
			$VersionMsg = "You are using an older version of the Options file.  Please generate a new Options file and transfer your settings.`r`nIn Use: $($XML.Options.Version) `r`nLatest: $(([xml]$CreateXML).Options.Version)"
			[System.Windows.Forms.MessageBox]::Show($VersionMsg,"Warning",[System.Windows.Forms.MessageBoxButtons]::Ok)
        }
	else{}

	return $true #return true for success or false for failure
}

Function Connect-ManagementPrograms {
	Param([pscredential]$UserCredential)

	# Import AD and connect to Office 365, rather than doing it during the user creation process
	Import-Module ActiveDirectory
	$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
									 -Credential $UserCredential -Authentication Basic -AllowRedirection
	Import-PsSession $ExchangeSession -AllowClobber
	Connect-MsolService -Credential $UserCredential
}
Function OnApplicationExit {
	Remove-Module ActiveDirectory
	Remove-PSSession *
	Show-Console
	$script:ExitCode = 0 #Set the exit code for the Packager
}


Function Get-RandomPassword() {
    Param ([Int32] $Length = 15, [String] $Encoding = "UTF16-LE")

    If ($length -lt 4) { $length = 4 }   #Password must be at least 4 characters long in order to satisfy complexity requirements.

    #Use the .NET crypto random number generator, not the weaker System.Random class with Get-Random:
    $RngProv = [System.Security.Cryptography.RNGCryptoServiceProvider]::Create()
    [byte[]] $onebyte = @(255)
    [Int32] $x = 0

    Do {
        [byte[]] $password = @() 
        
        $hasupper =     $false    #Has uppercase letter character flag.
        $haslower =     $false    #Has lowercase letter character flag.
        $hasnumber =    $false    #Has number character flag.
        $hasnonalpha =  $false    #Has non-alphanumeric character flag.
        $isstrong =     $false    #Assume password is not complex until tested otherwise.
        
        For ($i = $length; $i -gt 0; $i--)
        {                                                         
            While ($true)
            {   
                #Generate a random US-ASCII code point number.
                $RngProv.GetNonZeroBytes( $onebyte ) 
                [Int32] $x = $onebyte[0]                  
                if ($x -ge 32 -and $x -le 126){ break }   
            }
            
            # Even though it reduces randomness, eliminate problem characters to preserve sanity while debugging.
            # If you're worried, increase the length of the password or comment out the desired line(s):
            If ($x -eq 32) { $x++ }    #Eliminates the space character; causes problems for other scripts/tools.
            If ($x -eq 34) { $x-- }    #Eliminates double-quote; causes problems for other scripts/tools.
            If ($x -eq 39) { $x-- }    #Eliminates single-quote; causes problems for other scripts/tools.
            If ($x -eq 47) { $x-- }    #Eliminates the forward slash; causes problems for net.exe.
            If ($x -eq 96) { $x-- }    #Eliminates the backtick; causes problems for PowerShell.
            If ($x -eq 48) { $x++ }    #Eliminates zero; causes problems for humans who see capital O.
            If ($x -eq 79) { $x++ }    #Eliminates capital O; causes problems for humans who see zero. 
            
            $password += [Byte] $x  

            If ($x -ge 65 -And $x -le 90)  { $hasupper = $true }   #Non-USA users may wish to customize the code point numbers by hand,
            If ($x -ge 97 -And $x -le 122) { $haslower = $true }   #which is why we don't use functions like IsLower() or IsUpper() here.
            If ($x -ge 48 -And $x -le 57)  { $hasnumber = $true } 
            If (($x -ge 32 -And $x -le 47) -Or ($x -ge 58 -And $x -le 64) -Or ($x -ge 91 -And $x -le 96) -Or ($x -ge 123 -And $x -le 126)) { $hasnonalpha = $true } 
            If ($hasupper -And $haslower -And $hasnumber -And $hasnonalpha) { $isstrong = $true } 
        } 
    } While ($isstrong -eq $false)

    #$RngProv.Dispose() #Not compatible with PowerShell 2.0.

    #Output as a string with the desired encoding:
    Switch -Regex ( $Encoding.ToUpper().Trim() )
    {
        'ASCII' 
            { ([System.Text.Encoding]::ASCII).GetString($password) ; continue }
        'UTF8'     
            { ([System.Text.Encoding]::UTF8).GetString($password) ; continue } 
        'UNICODE|UTF16-LE|^UTF16$'  
            {
                $password = [System.Text.AsciiEncoding]::Convert([System.Text.Encoding]::ASCII, [System.Text.Encoding]::Unicode, $password )  
                ([System.Text.Encoding]::Unicode).GetString($password) 
                continue
            } 
        'UTF32'    
            { 
                $password = [System.Text.AsciiEncoding]::Convert([System.Text.Encoding]::ASCII, [System.Text.Encoding]::UTF32, $password )  
                ([System.Text.Encoding]::UTF32).GetString($password) 
                continue
            }
        '^UTF16-BE$' 
            { 
                $password = [System.Text.AsciiEncoding]::Convert([System.Text.Encoding]::ASCII, [System.Text.Encoding]::BigEndianUnicode, $password )  
                ([System.Text.Encoding]::BigEndianUnicode).GetString($password)
                continue 
            } 
        default #UTF16-LE Unicode
            {
                $password = [System.Text.AsciiEncoding]::Convert([System.Text.Encoding]::ASCII, [System.Text.Encoding]::Unicode, $password )  
                ([System.Text.Encoding]::Unicode).GetString($password) 
                continue
            } 
    }
}

Function Set-Email {
	# Default Email Suffix really should be moved to the XML file options for
	#   reusability, but hardcoding it for now...
	$DefaultSuffix = 'outlook.com'
	$Email = "{0}{1}@{2}" -f ($txtFirstName.Text,$txtLastName.Text,$DefaultSuffix)
	return $Email
}

Function Add-Office365User {
	param([string]$ADUserUPN)

	# Force ADAzure Sync and then check to see if user exists
	# Best way to do this is invoke the command on our AD Azure Connect server, KWDTECH01
	Invoke-Command -ComputerName "KWDTECH01" -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Initial}

	# Check if the user has synced
	do 
	{
		try {$Office365User = Get-MsolUser -UserPrincipalName $ADUserUPN -ErrorAction Stop}
		catch{Start-Sleep -Seconds 60}
	}
	while($Office365User -eq $null)

	# Set Office365 and Intune Licenses
	Set-MsolUser -UserPrincipalName $Office365User.userPrincipalName -UsageLocation US
	$OfficeLicense = "Company:O365_BUSINESS_PREMIUM"
	$InTuneLicense = "Company:INTUNE_A"

	# Check Availability of InTune Licenses
	$InTuneIsAvailable = Get-LicenseAvailability($InTuneLicense)
	while (!$InTuneIsAvailable) {
		$UserResponse = [System.Windows.Forms.MessageBox]::Show("No InTune Licenses Available! Please purchase more before continuing...", `
			"No Licenses Available","RetryCancel","Error")
		if($UserResponse.ToString() -like "*Cancel*"){break}
		else{$InTuneIsAvailable = Get-LicenseAvailability($InTuneLicense)}}
	if($InTuneIsAvailable)
	{Set-MsolUserLicense -UserPrincipalName $Office365User.userPrincipalName -AddLicenses $InTuneLicense}
	else
	{[System.Windows.Forms.MessageBox]::Show("No available licenses. InTune license was not set.", "No Licenses Available","Ok","Error")}

	# Check Availability of Office365 Licenses
	$OfficeIsAvailable = Get-LicenseAvailability($OfficeLicense)
	while (!$OfficeIsAvailable) {
		$UserResponse = [System.Windows.Forms.MessageBox]::Show("No Office365 Licenses Available! Please purchase more before continuing...", `
			"No Licenses Available","RetryCancel","Error")
		if($UserResponse.ToString() -like "*Cancel*"){break}
		else{$OfficeIsAvailable = Get-LicenseAvailability($OfficeLicense)}}
	if($OfficeIsAvailable)
	{Set-MsolUserLicense -UserPrincipalName $Office365User.userPrincipalName -AddLicenses $OfficeLicense}
	else
	{[System.Windows.Forms.MessageBox]::Show("No available licenses. Office365 license was not set.", "No Licenses Available","Ok","Error")}


	#TODO Set Email aliases
}

Function Get-LicenseAvailability {
	param([string]$LicenseName)

	$LicenseObject = Get-MsolAccountSku | Where-Object {$_.AccountSkuID -like "*$LicenseName*"}

	if($LicenseObject.ConsumedUnits -lt $LicenseObject.ActiveUnits)
	{return $true}else{return $false}
}

Function Set-sAMAccountName {
    Param([Switch]$Csv=$false)
    if(!$Csv)
	{
		$GivenName = $txtFirstName.text
		$SurName = $txtLastName.text
	}
    else{}
    Switch($XML.Options.Settings.sAMAccountName.Style | Where-Object{$_.Enabled -eq $True} | Select-Object -ExpandProperty Format)
	{
		"FirstName.LastName"    {"{0}.{1}" -f $GivenName.ToLower(),$Surname.ToLower()}
		"FirstInitialLastName"  {"{0}{1}" -f ($GivenName).ToLower()[0],$SurName.ToLower()}
		"LastNameFirstInitial"  {"{0}{1}" -f $SurName.ToLower(),($GivenName).ToLower()[0]}
		Default                 {"{0}.{1}" -f $GivenName.ToLower(),$Surname.ToLower()}
	}
}

Function Set-UPN {
    Param([Switch]$Csv=$false)
    if(!$Csv)
	{
		$GivenName = $txtFirstName.text
		$SurName = $txtLastName.text
		$UPNSuffix = $cboUPNSuffix.Text
	}
    else{}
    Switch($XML.Options.Settings.UPN.Style | Where-Object{$_.Enabled -eq $True} | Select-Object -ExpandProperty Format)
	{
		"FirstName.LastName"    {"{0}.{1}@{2}" -f $GivenName.ToLower(),$Surname.ToLower(),$UPNSuffix}
		"FirstInitialLastName"  {"{0}{1}@{2}" -f ($GivenName).ToLower()[0],$SurName.ToLower(),$UPNSuffix}
		"LastNameFirstInitial"  {"{0}{1}@{2}" -f $SurName.ToLower(),($GivenName).ToLower()[0],$UPNSuffix}
		Default                 {"{0}.{1}@{2}" -f $GivenName.ToLower(),$Surname.ToLower(),$UPNSuffix}
	}
}

Function Set-DisplayName {
    Param([Switch]$Csv=$false)
    if(!$Csv)
	{
		$GivenName = $txtFirstName.text
		$SurName = $txtLastName.text
	}
    else{}
    Switch($XML.Options.Settings.DisplayName.Style | Where-Object{$_.Enabled -eq $True} | Select-Object -ExpandProperty Format)
	{
		"FirstName LastName"    {"{0} {1}" -f $GivenName,$Surname}
		"LastName, FirstName"   {"{0}, {1}" -f $SurName, $GivenName}
		Default                 {"{0} {1}" -f $GivenName,$Surname}
	}
}
	
Function Set-Description {
	if ($cboTitle.Text -ne "")
	{
		$CurrentDate = Get-Date -UFormat "%m/%d/%Y"
		$DateFormat = " - " + $CurrentDate
		$txtDescription.Text = -join $cboTitle.Text,$DateFormat
	}
	else{$txtDescription.Text = ""}
}
	
Function Set-GroupMemberships {
	#First thing first, clear the listbox of old group names
	$listboxGroups.Items.Clear()
	
	Write-Verbose "Gathering Default Groups based on location"
	$SelectedSite = $cboSite.Text
	$XML.Options.Locations.Location | Where-Object{$_.Site -eq $SelectedSite} | Select-Object -ExpandProperty Groups | Select-Object -ExpandProperty Group | ForEach-Object {$listboxGroups.Items.Add($_); $listboxGroups.SelectedItems.Add($_)}
	
	Write-Verbose "Gathering Default Groups based on Department"
	$SelectedDepartment = $cboDepartment.Text
	$XML.Options.Departments.Department | Where-Object{$_.Name -eq $SelectedDepartment} | Select-Object -ExpandProperty DefaultGroups | Select-Object -ExpandProperty DefaultGroup | ForEach-Object {$listboxGroups.Items.Add($_); $listboxGroups.SelectedItems.Add($_)}
	
	Write-Verbose "Gathering Suggested Groups based on Department"
	$XML.Options.Departments.Department | Where-Object{$_.Name -eq $SelectedDepartment} | Select-Object -ExpandProperty SuggestedGroups | Select-Object -ExpandProperty SuggestedGroup | ForEach-Object {$listboxGroups.Items.Add($_)}
}

Function RetrieveGroups{
    param($selectedItems)
	
	Write-Verbose "Parsing the Selected Items and Returning a List of Their Distinguished Names"
    $returnList = foreach($group in $selectedItems){(Get-ADGroup -Filter {Name -eq $group}).distinguishedName}
    return $returnList
}

Function Set-Path {
	Param($Site,$Department)

	# This function is only partially finished, I plan on having the OU be
	#   set automatically based on what location and department you select.
	#   Functionality is there, just need to add options.
	# TODO Add more sites and departments, expanding the Switch statement
	Switch($Site)
	{
		"NH" {$cboPath.SelectedItem = "OU=Dover Users,OU=Dover Office,DC=Company,DC=com"}
		"TX" {Switch($Department)
			{
				"Credit" {$cbo.SelectedItem = "OU=Credit Processing,OU=Kingwood Users,OU=Kingwood,DC=Company,DC=com"}
				Default {}
			}		
		}
		Default{}
	}
}

Function Send-ConfirmationEmail{
Param([string]$DisplayName,[string]$EmailAddress,`
	  [string]$PhoneNumber,[string]$FaxNumber,`
	  [string]$CellNumber,[bool]$CompanyCell,[pscredential]$Creds)

$CompanyCellMessage = ""
if ($CompanyCell)
{
	$CompanyCellMessage = "will be using their own cell phone, please reimburse."
} else {
	$CompanyCellMessage = "will be using a company cell phone, no need to reimburse."
}

$To = $Creds.UserName.ToString()
$From = $Creds.UserName.ToString()
$Subject = "New Hire - $DisplayName"
$SMTP = "smtp.office365.com"
$Body = @"
$DisplayName

Phone: $PhoneNumber
Fax: $FaxNumber
Cell: $CellNumber

Email: $EmailAddress

$DisplayName $CompanyCellMessage


-Company Technology Department

"@

Send-MailMessage -To $To -From $From -Subject $Subject -Body $Body -Credential $Creds -SmtpServer $SMTP -UseSsl -Port 587
}

function Show-Console {
	$consolePtr = [Console.Window]::GetConsoleWindow()
	[Console.Window]::ShowWindow($consolePtr, 5)
}

function Hide-Console {
	$consolePtr = [Console.Window]::GetConsoleWindow()
	[Console.Window]::ShowWindow($consolePtr, 0)
}

#endregion Application Functions

#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function Invoke-EmployeeManagementForm {
	Param([pscredential]$UserCredential)

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::LoadWithPartialName("Microsoft.VisualBasic")
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formMain = New-Object System.Windows.Forms.Form
	$btnSubmitAll = New-Object System.Windows.Forms.Button
	$btnLast = New-Object System.Windows.Forms.Button
	$btnNext = New-Object System.Windows.Forms.Button
	$btnPrev = New-Object System.Windows.Forms.Button
	$btnFirst = New-Object System.Windows.Forms.Button
	$btnImportCSV = New-Object System.Windows.Forms.Button
	$lvCSV = New-Object System.Windows.Forms.ListView
	$txtUPN = New-Object System.Windows.Forms.TextBox
	$txtsAM = New-Object System.Windows.Forms.TextBox
	$txtDN = New-Object System.Windows.Forms.TextBox
	$cboDepartment = New-Object System.Windows.Forms.ComboBox
	$labelUserPrincipalName = New-Object System.Windows.Forms.Label
	$labelUPNSuffix = New-Object System.Windows.Forms.Label
    $cboUPNSuffix = New-Object System.Windows.Forms.ComboBox
	$labelSamAccountName = New-Object System.Windows.Forms.Label
	$labelDisplayName = New-Object System.Windows.Forms.Label
	$SB = New-Object System.Windows.Forms.StatusBar
	$listboxGroups = New-Object System.Windows.Forms.ListBox
	$labelGroups = New-Object System.Windows.Forms.Label
	$labelEmails = New-Object System.Windows.Forms.Label
	$labelEmailFormat = New-Object System.Windows.Forms.Label
	$btnRemoveAlias = New-Object System.Windows.Forms.Button
	$btnAddAlias = New-Object System.Windows.Forms.Button
	$listboxEmails = New-Object System.Windows.Forms.ListBox
	$txtEmails = New-Object System.Windows.Forms.TextBox
	$cboSite = New-Object System.Windows.Forms.ComboBox
	$labelSite = New-Object System.Windows.Forms.Label
	$txtDescription = New-Object System.Windows.Forms.TextBox
	$txtPassword = New-Object System.Windows.Forms.TextBox
	$labelPassword = New-Object System.Windows.Forms.Label
	$cboDomain = New-Object System.Windows.Forms.ComboBox
	$labelCurrentDomain = New-Object System.Windows.Forms.Label
	$txtPostalCode = New-Object System.Windows.Forms.TextBox
	$txtState = New-Object System.Windows.Forms.TextBox
	$txtCity = New-Object System.Windows.Forms.TextBox
	$txtStreetAddress = New-Object System.Windows.Forms.TextBox
	$txtOffice = New-Object System.Windows.Forms.TextBox
	$txtCompany = New-Object System.Windows.Forms.TextBox
	$cboTitle = New-Object System.Windows.Forms.ComboBox
	$txtOfficePhone = New-Object System.Windows.Forms.TextBox
	$txtLastName = New-Object System.Windows.Forms.TextBox
	$cboPath = New-Object System.Windows.Forms.ComboBox
	$labelOU = New-Object System.Windows.Forms.Label
	$txtFirstName = New-Object System.Windows.Forms.TextBox
	$labelPostalCode = New-Object System.Windows.Forms.Label
	$labelState = New-Object System.Windows.Forms.Label
	$labelCity = New-Object System.Windows.Forms.Label
	$labelStreetAddress = New-Object System.Windows.Forms.Label
	$labelOffice = New-Object System.Windows.Forms.Label
	$labelCompany = New-Object System.Windows.Forms.Label
	$labelDepartment = New-Object System.Windows.Forms.Label
	$labelTitle = New-Object System.Windows.Forms.Label
	$btnSubmit = New-Object System.Windows.Forms.Button
	$labelDescription = New-Object System.Windows.Forms.Label
	$labelOfficePhone = New-Object System.Windows.Forms.Label
	$labelLastName = New-Object System.Windows.Forms.Label
	$labelFirstName = New-Object System.Windows.Forms.Label
	$menustrip1 = New-Object System.Windows.Forms.MenuStrip
	$fileToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$formMode = New-Object System.Windows.Forms.ToolStripMenuItem
	$CSVTemplate = New-Object System.Windows.Forms.SaveFileDialog
	$OFDImportCSV = New-Object System.Windows.Forms.OpenFileDialog
	$CreateCSVTemplate = New-Object System.Windows.Forms.ToolStripMenuItem
	$MenuExit = New-Object System.Windows.Forms.ToolStripMenuItem
	$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	
	
	$formMain_Load={
		# Import the DLLs necessary to hide the console window
		Add-Type -Name Window -Namespace Console -MemberDefinition '
		[DllImport("Kernel32.dll")]
		public static extern IntPtr GetConsoleWindow();
	
		[DllImport("user32.dll")]
		public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
		'
		# Custom Icon. Cool, right?
		$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSScriptRoot + "\data\icon.ico")
		$formMain.Icon = $Icon	
		
		# Fix the size of the window and disable size adjusting
		#   The forms would get all screwy when the size was adjustable. Less programming
		#   and headaches if we just permanently fix the size.
		$formMain.Size = "325,1030"
		$formMain.FormBorderStyle = "FixedSingle"
		$formMain.MaximizeBox = $false
		$formMain.MinimizeBox = $false
		
		# Set the Window Title using current version number
		$formMain.Text = $formMain.Text + " " + $XML.Options.Version
		
		Write-Verbose "Adding domains to combo box"
		$XML.Options.Domains.Domain | ForEach-Object{$cboDomain.Items.Add($_.Name)}
		
		Write-Verbose "Adding OUs to combo box"
	    $XML.Options.Domains.Domain | Where-Object{$_.Name -match $cboDomain.Text} | Select-Object -ExpandProperty Path | ForEach-Object{$cboPath.Items.Add($_)}
		
		Write-Verbose "Adding titles to combo box"
		$XML.Options.Titles.Title | ForEach-Object{$cboTitle.Items.Add($_)}
		
		Write-Verbose "Adding sites to combo box"
		$XML.Options.Locations.Location | ForEach-Object{$cboSite.Items.Add($_.Site)}
		
		Write-Verbose "Adding departments to combo box"
		$XML.Options.Departments.Department.Name | ForEach-Object{$cboDepartment.Items.Add($_)}
		
		Write-Verbose "Adding UPN suffixes to combo box"
		$XML.Options.UPNSuffixes.UPNSuffix | ForEach-Object{$cboUPNSuffix.Items.Add($_)}
		
		Write-Verbose "Setting default fields"
		$cboDomain.SelectedItem = $XML.Options.Default.Domain
	    $cboPath.SelectedItem = $XML.Options.Default.Path
		$txtFirstName.Text = $XML.Options.Default.FirstName
		$txtLastName.Text = $XML.Options.Default.LastName
		$txtOffice.Text = $XML.Options.Default.Office
		$cboTitle.Text = $XML.Options.Default.Title
		$cboDepartment.SelectedItem = $XML.Options.Default.Department
		$txtCompany.Text = $XML.Options.Default.Company
		$txtOfficePhone.Text = $XML.Options.Default.Phone
		$cboSite.SelectedItem = $XML.Options.Default.Site
		$txtStreetAddress.Text = $XML.Options.Default.StreetAddress
		$txtCity.Text = $XML.Options.Default.City
		$txtState.Text = $XML.Options.Default.State
		$txtPostalCode.Text = $XML.Options.Default.PostalCode
		$cboUPNSuffix.Text = $XML.Options.Default.UPNSuffix
		#$txtPassword.Text = $XML.Options.Default.Password
		$txtPassword.Text = Get-RandomPassword
		
		#If-Else statement to determine if we need to autogenerate the description
		if ($XML.Options.Settings.Description.Generate -eq "True")
		{
			Write-Verbose "Generating Automatic Description"
			Set-Description
		} else { $txtDescription.Text = $XML.Options.Default.Description }
		
		Write-Verbose "Creating CSV Headers"
		$Headers = @('ID','Domain','Path','FirstName','LastName','Office','Title','Description','Department','Company','Phone','StreetAddress','City','State','PostalCode','Password','sAMAccountName','userPrincipalName','UPNSuffix','DisplayName', 'Site')
		$Headers | ForEach-Object{[Void]$lvCSV.Columns.Add($_)}

		Hide-Console # hide the console at start
	}
	
	$btnSubmit_Click={

		# Local active directory user creation
		$Path=$cboPath.Text
		$GivenName = $txtFirstName.Text
		$Surname = $txtLastName.Text
		$OfficePhone = $txtOfficePhone.Text
		$Description = $txtDescription.Text
		$Title = $cboTitle.Text
		$Department = $cboDepartment.Text
		$Company = $txtCompany.Text
		$Office = $txtOffice.Text
		$StreetAddress = $txtStreetAddress.Text
		$City = $txtCity.Text
		$State = $txtState.Text
		$PostalCode = $txtPostalCode.Text
		$Emails = $txtEmails.Text
	
		if($XML.Options.Settings.Password.ChangeAtLogon -eq "True"){$ChangePasswordAtLogon = $True}
        else{$ChangePasswordAtLogon = $false}
		
        if($XML.Options.Settings.AccountStatus.Enabled -eq "True"){$Enabled = $True}
        else{$Enabled = $false}
	
		$Name="$GivenName $Surname"
		
        if($XML.Options.Settings.sAMAccountName.Generate -eq $True){$sAMAccountName = Set-sAMAccountName}
		else{$sAMAccountName = $txtsAM.Text}

        if($XML.Options.Settings.uPN.Generate -eq $True){$userPrincipalName = Set-UPN}
        else{$userPrincipalName = $txtuPN.Text}
		
        if($XML.Options.Settings.DisplayName.Generate -eq $True){$DisplayName = Set-DisplayName}
        else{$DisplayName = $txtDN.Text}

		$AccountPassword = $txtPassword.text | ConvertTo-SecureString -AsPlainText -Force
		
		$GroupList = RetrieveGroups($listboxGroups.SelectedItems)
	
		$User = @{
		    Name = $Name
		    GivenName = $GivenName
		    Surname = $Surname
		    Path = $Path
		    samAccountName = $samAccountName
		    userPrincipalName = $userPrincipalName
		    DisplayName = $DisplayName
		    AccountPassword = $AccountPassword
		    ChangePasswordAtLogon = $ChangePasswordAtLogon
		    Enabled = $Enabled
		    OfficePhone = $OfficePhone
		    Description = $Description
		    Title = $Title
		    Department = $Department
		    Company = $Company
		    Office = $Office
		    StreetAddress = $StreetAddress
		    City = $City
		    State = $State
			PostalCode = $PostalCode
			EmailAddress = $Emails
		}
			
		$SB.Text = "Creating new user $sAMAccountName"
        $ADError = $Null
		New-ADUser @User -ErrorVariable ADError
        if ($ADerror){$SB.Text = "[$sAMAccountName] $ADError"}
        else{$SB.Text = "$sAMAccountName created successfully."}

		$SB.Text = "Adding $sAMAccountName to selected Security Groups"
		$ADGroupError = $Null
		foreach ($group in $GroupList){Add-ADGroupMember -Identity $group -Members $sAMAccountName -ErrorVariable $ADGroupError}
        if ($ADGroupError){$SB.Text = "[$sAMAccountName] $ADGroupError"}
		else{$SB.Text = "$sAMAccountName added to groups successfully."}

		# Office 365 user creation
		$SB.Text = "Waiting for user to sync with ADAzure..."
		Add-Office365User($userPrincipalName)

		$SB.Text = "$sAMAccountName added to Office 365 successfully."

		Send-ConfirmationEmail -DisplayName $DisplayName -EmailAddress $Emails `
							   -PhoneNumber $OfficePhone -FaxNumber $OfficePhone -CellNumber $OfficePhone `
							   -CompanyCell $True -Creds $UserCredential
		$SB.Text = "Account Created Successfully!"

		[System.Windows.Forms.MessageBox]::Show("$sAMAccountName was created successfully!","Account Created","Ok","Information") | Out-Null
	}
	
	$btnAddAlias_Click={
		$EmailInputTitle = "Add New Email Alias"
		$EmailInputMessage = "Please enter the desired email alias in the box below:"

		$NewEmailAlias = [Microsoft.VisualBasic.Interaction]::InputBox($EmailInputMessage,$EmailInputTitle)

		$listboxEmails.Items.Add($NewEmailAlias)
	}

	$btnRemoveAlias_Click={
		$toBeRemoved=@()
		foreach($item in $listboxEmails.SelectedItems){$toBeRemoved += $item}
		foreach($item in $toBeRemoved){$listboxEmails.Items.Remove($item)}
	}

	$txtEmails_TextChanged={
		#TODO Determine if this function is necessary, and if so,
		#     what code needs to go here.
	}

	$cboDomain_SelectedIndexChanged={
		$cboPath.Items.Clear()
		Write-Verbose "Adding OUs to combo box"
	    $XML.Options.Domains.Domain | Where-Object{$_.Name -match $cboDomain.Text} | Select-Object -ExpandProperty Path | ForEach-Object{$cboPath.Items.Add($_)}	
		Write-Verbose "Creating required account fields"
		
        if ($XML.Options.Settings.DisplayName.Generate) {$txtDN.Text = Set-DisplayName}
        if ($XML.Options.Settings.sAMAccountName.Generate) {$txtsAM.Text = Set-sAMAccountName}
        if ($XML.Options.Settings.UPN.Generate) {$txtUPN.Text = Set-UPN}	
	}
	
	$cboTitle_SelectedIndexChanged={
		Set-Description
	}
	
	$cboDepartment_SelectedIndexChanged={
		Set-GroupMemberships
		Set-Path
	}
	
	$cboTitle_TextChanged={
		Set-Description
	}
	
	$cboUPNSuffix_SelectedIndexChanged={
		$txtUPN.Text = Set-UPN
	}	
	
	$cboSite_SelectedIndexChanged={
		Write-Verbose "Updating site fields with address information"
	    $Site = $XML.Options.Locations.Location | Where-Object{$_.Site -match $cboSite.Text}
		$txtStreetAddress.Text = $Site.StreetAddress
		$txtCity.Text = $Site.City
		$txtState.Text = $Site.State
		$txtOffice.Text = $Site.Office
		$txtPostalCode.Text = $Site.PostalCode
		
		Set-GroupMemberships
		Set-Path
	}	
	
	$txtName_TextChanged={
		Write-Verbose "Creating required account fields"
        
        if ($XML.Options.Settings.DisplayName.Generate -eq $True) {$txtDN.Text = Set-DisplayName}
        if ($XML.Options.Settings.sAMAccountName.Generate -eq $True) {$txtsAM.Text = (Set-sAMAccountName)}
		if ($XML.Options.Settings.UPN.Generate -eq $True) {$txtUPN.Text = Set-UPN}
		if ($XML.Options.Settings.EmailAddress.Generate -eq $True) {$txtEmails.Text = Set-Email}
	}
	
	$createTemplateToolStripMenuItem_Click={
		$CSVTemplate.ShowDialog()
	}
	
	$CSVTemplate_FileOk=[System.ComponentModel.CancelEventHandler]{
		"" |
		Select-Object Domain,Path,UPNSuffix,FirstName,LastName,Office,Title,Description,Department,Company,Phone,StreetAddress,City,State,PostalCode,Password,sAMAccountName,userPrincipalName,DisplayName,Site |
		Export-CSV $CSVTemplate.FileName -NoTypeInformation	
	}
	
	$formMode_Click={
		if($formMode.Text -eq 'CSV Mode'){
			$formMode.Text = "Single-User Mode"
			Get-Variable | Where-Object{$_.Name -match "txt"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left'}catch{}}
			Get-Variable | Where-Object{$_.Name -match "cbo"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left'}catch{}}
			Get-Variable | Where-Object{$_.Name -match "btn"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left'}catch{}}
			Get-Variable | Where-Object{$_.Name -match "listbox"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left'}catch{}}
			$formMain.Size = '1550,890'
			$btnFirst.Visible = $True
			$btnPrev.Visible = $True
			$btnNext.Visible = $True
			$btnLast.Visible = $True
			$btnImportCSV.Visible = $True
			$btnSubmitAll.Visible = $True
			$lvCSV.Visible = $True
			}
		else{
			$formMode.Text = "CSV Mode"
			$formMain.Size = '325,890'
			Get-Variable | Where-Object{$_.Name -match "txt"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left,Right'}catch{}}
			Get-Variable | ForEach-Object{$_.Name -match "cbo"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left,Right'}catch{}}
			Get-Variable | ForEach-Object{$_.Name -match "btn"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left,Right'}catch{}}
			$btnFirst.Visible = $False
			$btnPrev.Visible = $False
			$btnNext.Visible = $False
			$btnLast.Visible = $False
			$btnImportCSV.Visible = $False
			$btnSubmitAll.Visible = $False
			$lvCSV.Visible = $False
			}
	}
	
	$btnImportCSV_Click={
		$OFDImportCSV.ShowDialog()
		$CSV = Import-Csv $OFDImportCSV.FileName
		$i = 0
		ForEach ($Entry in $CSV){
			$User = New-Object System.Windows.Forms.ListViewItem($i)
			ForEach ($Col in ($lvCSV.Columns | Where-Object{$_.Text -ne "ID"})){
                $Field = $Col.Text
                $SubItem = "$($Entry.$Field)"
                if($Field -eq 'FirstName'){$Script:GivenName = $SubItem}
				if($Field -eq 'LastName'){$Script:Surname = $SubItem}
                if($Field -eq 'sAMAccountName' -AND $SubItem -eq ""){$SubItem = Set-sAMAccountName -Csv}
                if($Field -eq 'userPrincipalName' -AND $SubItem -eq ""){$SubItem = Set-UPN -Csv}
				if($Field -eq 'DisplayName' -AND $SubItem -eq ""){$SubItem = Set-DisplayName -Csv}
				if($Field -eq 'Site'){$Script:Site = $Subitem}
                $User.SubItems.Add($SubItem)
                }
			$lvCSV.Items.Add($User)
			$i++
		}
	}
	
	$lvCSV_SelectedIndexChanged={
		<# 
			UPDATE: This whole thing was more trouble than it was worth, so I abandoned the whole
					CSV mode for the time being.
					
					All the comments and code below is left exactly as it was when I stopped working
					on it, so no guarantees that any of this function works properly...
		#>


		# If statement to check that the name spaces aren't blank.
		#   This usually means that the program is just loading, and the
		#   default blank spaces are still present.
		#   Might be a better way to handle that, but....

		# Object to hold information in the forms controls while we switch out the users
		#   Declaring it now so it is available in scope later, if necessary
		$previousUser = New-Object PSobject | Select-Object Domain,Path,FirstName,LastName,Office,Title,Description,Department,Company,OfficePhone,StreetAddress,City,State,PostalCode,Password,sAM,UPNSuffix,DN,Site,UPN,Groups
		
		if(($txtFirstName.Text -ne '') -And ($txtLastName.Text -ne ''))
		{
			#Store the information entered into the form controls in a temporary spot
			try{[string]$previousUser.Domain = $cboDomain.SelectedItem}catch{}
			try{[string]$previousUser.Path = $cboPath.SelectedItem}catch{}				
			try{$previousUser.FirstName = $txtFirstName.Text}catch{}
			try{$previousUser.LastName = $txtLastName.Text}catch{}
			try{$previousUser.Office = $txtOffice.Text}catch{}
			try{$previousUser.Title = $cboTitle.Text}catch{}
			try{$previousUser.Description = $txtDescription.SelectedItem}catch{}
			try{$previousUser.Department = $cboDepartment.SelectedItem}catch{}
			try{$previousUser.Company = $txtCompany.Text}catch{}
			try{$previousUser.OfficePhone = $txtOfficePhone.Text}catch{}
			try{$previousUser.StreetAddress = $txtStreetAddress.Text}catch{}
			try{$previousUser.City = $txtCity.Text}catch{}
			try{$previousUser.State = $txtState.Text}catch{}
			try{$previousUser.PostalCode = $txtPostalCode.Text}catch{}
			try{$previousUser.Password = $txtPassword.Text}catch{}
			try{$previousUser.sAM = $txtsAM.Text}catch{}
			try{$previousUser.UPNSuffix = $cboUPNSuffix.Text}catch{}
			try{$previousUser.DN = $txtDN.Text}catch{}
			try{$previousUser.Site = $cboSite.Text}catch{}
			try{$previousUser.UPN = $txtUPN.Text}catch{}
			try{$previousUser.Groups = $listboxGroups.SelectedItems}catch{}
		}
		

		#Populate the form controls with the newly selected user information
		try{$cboDomain.SelectedItem = $lvCSV.SelectedItems[0].SubItems[1].Text}catch{}
		try{$cboPath.SelectedItem = $lvCSV.SelectedItems[0].SubItems[2].Text}catch{}
		try{$txtFirstName.Text = $lvCSV.SelectedItems[0].SubItems[3].Text}catch{}
		try{$txtLastName.Text = $lvCSV.SelectedItems[0].SubItems[4].Text}catch{}
		try{$txtOffice.Text = $lvCSV.SelectedItems[0].SubItems[5].Text}catch{}
		try{$cboTitle.Text = $lvCSV.SelectedItems[0].SubItems[6].Text}catch{}
		try{$txtDescription.SelectedItem = $lvCSV.SelectedItems[0].SubItems[7].Text}catch{}
		try{$cboDepartment.SelectedItem = $lvCSV.SelectedItems[0].SubItems[8].Text}catch{}
		try{$txtCompany.Text = $lvCSV.SelectedItems[0].SubItems[9].Text}catch{}
		try{$txtOfficePhone.Text = $lvCSV.SelectedItems[0].SubItems[10].Text}catch{}
		try{$txtStreetAddress.Text = $lvCSV.SelectedItems[0].SubItems[11].Text}catch{}
		try{$txtCity.Text = $lvCSV.SelectedItems[0].SubItems[12].Text}catch{}
		try{$txtState.Text = $lvCSV.SelectedItems[0].SubItems[13].Text}catch{}
		try{$txtPostalCode.Text = $lvCSV.SelectedItems[0].SubItems[14].Text}catch{}
		try{$txtPassword.Text = $lvCSV.SelectedItems[0].SubItems[15].Text}catch{}
        try{$txtsAM.Text = $lvCSV.SelectedItems[0].SubItems[16].Text}catch{}
        try{$txtuPN.Text = $lvCSV.SelectedItems[0].SubItems[17].Text}catch{}
		try{$cboUPNSuffix.Text = $lvCSV.SelectedItems[0].SubItems[18].Text}catch{}
		try{$txtDN.Text = $lvCSV.SelectedItems[0].SubItems[19].Text}catch{}
		try{$cboSite.Text = $lvCSV.SelectedItems[0].SubItems[20].Text}catch{}
		try{$txtUPN.Text = Set-UPN}catch{}
		try{Set-GroupMemberships}catch{}
	}
	
	$btnFirst_Click={
		$lvCSV.Items | ForEach-Object{$_.Selected = $False}
		$lvCSV.Items[0].Selected = $True	
	}
	
	$btnLast_Click={
		$LastRow = ($lvCSV.Items).Count - 1
		$lvCSV.Items | ForEach-Object{$_.Selected = $False}
		$lvCSV.Items[$LastRow].Selected = $True	
	}
	
	$btnNext_Click={
		$LastRow = ($lvCSV.Items).Count - 1
		[Int]$Index = $lvCSV.SelectedItems[0].Index
		if($LastRow -gt $Index){
			$lvCSV.Items | ForEach-Object{$_.Selected = $False}
			$lvCSV.Items[$Index+1].Selected = $True	
		}
	}
	
	$btnPrev_Click={
		[Int]$Index = $lvCSV.SelectedItems[0].Index
		if($Index -gt 0){
			$lvCSV.Items | ForEach-Object{$_.Selected = $False}
			$lvCSV.Items[$Index-1].Selected = $True	
		}
	}
	
	$MenuExit_Click={
		$formMain.Close()
	}
	
	$btnSubmitAll_Click={
		$lvCSV.Items | ForEach-Object{
			
			$Path = $_.Subitems[2].Text
			$GivenName = $_.Subitems[3].Text
			$Surname = $_.Subitems[4].Text
			$OfficePhone = $_.Subitems[5].Text
			$Title = $_.Subitems[6].Text
			$Description = $_.Subitems[7].Text
			$Department = $_.Subitems[8].Text
			$Company = $_.Subitems[9].Text
			$Office = $_.Subitems[10].Text
			$StreetAddress = $_.Subitems[11].Text
			$City = $_.Subitems[12].Text
			$State = $_.Subitems[13].Text
			$PostalCode = $_.Subitems[14].Text
	
			$Name = "$GivenName $Surname"

		    if($XML.Options.Settings.Password.ChangeAtLogon -eq "True"){$ChangePasswordAtLogon = $True}
            else{$ChangePasswordAtLogon = $false}
		
            if($XML.Options.Settings.AccountStatus.Enabled -eq "True"){$Enabled = $True}
            else{$Enabled = $false}
	
		    if($_.Subitems[16].Text -eq $null){$sAMAccountName = Set-sAMAccountName}
		    else{$sAMAccountName = $_.Subitems[16].Text}

            if($_.Subitems[17].Text -eq $null){$userPrincipalName = Set-UPN}
            else{$userPrincipalName = $_.Subitems[17].Text}
		
            if($_.Subitems[19].Text -eq $null){$DisplayName = Set-DisplayName}
            else{$DisplayName = $_.Subitems[19].Text}

			$AccountPassword = $_.Subitems[15].Text | ConvertTo-SecureString -AsPlainText -Force
	
			$GroupList = RetrieveGroups($listboxGroups.SelectedItems)
	
			$User = @{
			    Name = $Name
			    GivenName = $GivenName
			    Surname = $Surname
			    Path = $Path
			    samAccountName = $samAccountName
			    userPrincipalName = $userPrincipalName
			    DisplayName = $DisplayName
			    AccountPassword = $AccountPassword
			    ChangePasswordAtLogon = $ChangePasswordAtLogon
			    Enabled = $Enabled
			    OfficePhone = $OfficePhone
			    Description = $Description
			    Title = $Title
			    Department = $Department
			    Company = $Company
			    Office = $Office
			    StreetAddress = $StreetAddress
			    City = $City
			    State = $State
				PostalCode = $PostalCode
			    }
			$SB.Text = "Creating new user $sAMAccountName"
            $ADError = $Null
			New-ADUser @User -ErrorVariable ADError
            if ($ADerror){$SB.Text = "[$sAMAccountName] $ADError"}
            else{$SB.Text = "$sAMAccountName created successfully."}
			
			$SB.Text = "Adding $sAMAccountName to selected Security Groups"
			$ADGroupError = $Null
			foreach ($group in $GroupList){Add-ADGroupMember -Identity $group -Members $sAMAccountName -ErrorVariable $ADGroupError}
			if ($ADGroupError){$SB.Text = "[$sAMAccountName] $ADGroupError"}
			else{$SB.Text = "$sAMAccountName added to groups successfully."}
		}
	}
	
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$formMain.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$btnSubmitAll.remove_Click($btnSubmitAll_Click)
			$btnLast.remove_Click($btnLast_Click)
			$btnNext.remove_Click($btnNext_Click)
			$btnPrev.remove_Click($btnPrev_Click)
			$btnFirst.remove_Click($btnFirst_Click)
			$btnImportCSV.remove_Click($btnImportCSV_Click)
			$lvCSV.remove_SelectedIndexChanged($lvCSV_SelectedIndexChanged)
			$cboSite.remove_SelectedIndexChanged($cboSite_SelectedIndexChanged)
			$cboDomain.remove_SelectedIndexChanged($cboDomain_SelectedIndexChanged)
			$txtLastName.remove_TextChanged($txtName_TextChanged)
			$txtFirstName.remove_TextChanged($txtName_TextChanged)
			$btnSubmit.remove_Click($btnSubmit_Click)
			$formMain.remove_Load($formMain_Load)
			$formMode.remove_Click($formMode_Click)
			$CSVTemplate.remove_FileOk($CSVTemplate_FileOk)
			$CreateCSVTemplate.remove_Click($createTemplateToolStripMenuItem_Click)
			$MenuExit.remove_Click($MenuExit_Click)
			$formMain.remove_Load($Form_StateCorrection_Load)
			$formMain.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch [Exception]
		{ }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	#
	# formMain
	#
	$formMain.Controls.Add($btnSubmitAll)
	$formMain.Controls.Add($btnLast)
	$formMain.Controls.Add($btnNext)
	$formMain.Controls.Add($btnPrev)
	$formMain.Controls.Add($btnFirst)
	$formMain.Controls.Add($btnImportCSV)
	$formMain.Controls.Add($lvCSV)
	$formMain.Controls.Add($txtUPN)
	$formMain.Controls.Add($txtsAM)
	$formMain.Controls.Add($txtDN)
	$formMain.Controls.Add($cboDepartment)
	$formMain.Controls.Add($labelUserPrincipalName)
	$formMain.Controls.Add($labelSamAccountName)
	$formMain.Controls.Add($labelDisplayName)
	$formMain.Controls.Add($SB)
	$formMain.Controls.Add($listboxGroups)
	$formMain.Controls.Add($listboxEmails)
	$formMain.Controls.Add($labelGroups)
	$formMain.Controls.Add($labelEmails)
	#$formMain.Controls.Add($labelEmailFormat)
	$formMain.Controls.Add($txtEmails)
	$formMain.Controls.Add($btnAddAlias)
	$formMain.Controls.Add($btnRemoveAlias)
	$formMain.Controls.Add($cboUPNSuffix)
	$formMain.Controls.Add($labelUPNSuffix)
	$formMain.Controls.Add($cboSite)
	$formMain.Controls.Add($labelSite)
	$formMain.Controls.Add($txtDescription)
	$formMain.Controls.Add($txtPassword)
	$formMain.Controls.Add($labelPassword)
	$formMain.Controls.Add($cboDomain)
	$formMain.Controls.Add($labelCurrentDomain)
	$formMain.Controls.Add($txtPostalCode)
	$formMain.Controls.Add($txtState)
	$formMain.Controls.Add($txtCity)
	$formMain.Controls.Add($txtStreetAddress)
	$formMain.Controls.Add($txtOffice)
	$formMain.Controls.Add($txtCompany)
	$formMain.Controls.Add($cboTitle)
	$formMain.Controls.Add($txtOfficePhone)
	$formMain.Controls.Add($txtLastName)
	$formMain.Controls.Add($cboPath)
	$formMain.Controls.Add($labelOU)
	$formMain.Controls.Add($txtFirstName)
	$formMain.Controls.Add($labelPostalCode)
	$formMain.Controls.Add($labelState)
	$formMain.Controls.Add($labelCity)
	$formMain.Controls.Add($labelStreetAddress)
	$formMain.Controls.Add($labelOffice)
	$formMain.Controls.Add($labelCompany)
	$formMain.Controls.Add($labelDepartment)
	$formMain.Controls.Add($labelTitle)
	$formMain.Controls.Add($btnSubmit)
	$formMain.Controls.Add($labelDescription)
	$formMain.Controls.Add($labelOfficePhone)
	$formMain.Controls.Add($labelLastName)
	$formMain.Controls.Add($labelFirstName)
	$formMain.Controls.Add($menustrip1)
	# Disabling the AcceptButton property to eliminate accidental user creation
	#$formMain.AcceptButton = $btnSubmit
	$formMain.ClientSize = '304, 980'
	$System_Windows_Forms_MenuStrip_1 = New-Object System.Windows.Forms.MenuStrip
	$System_Windows_Forms_MenuStrip_1.Location = '0, 0'
	$System_Windows_Forms_MenuStrip_1.Name = ""
	$System_Windows_Forms_MenuStrip_1.Size = '279, 24'
	$System_Windows_Forms_MenuStrip_1.Visible = $False
	$formMain.MainMenuStrip = $System_Windows_Forms_MenuStrip_1
	$formMain.Name = "formMain"
	$formMain.ShowIcon = $True
	$formMain.StartPosition = 'Manual'
	$formMain.Location = "0,0"
	$formMain.Text = "Employee Management - Add User"
	$formMain.add_Load($formMain_Load)
	#
	# btnSubmitAll
	#
	$btnSubmitAll.Location = '513, 0'
	$btnSubmitAll.Name = "btnSubmitAll"
	$btnSubmitAll.Size = '75, 25'
	$btnSubmitAll.TabIndex = 47
	$btnSubmitAll.Text = "Submit All"
	$btnSubmitAll.UseVisualStyleBackColor = $True
	$btnSubmitAll.Visible = $False
	$btnSubmitAll.add_Click($btnSubmitAll_Click)
	#
	# btnLast
	#
	$btnLast.Location = '481, 0'
	$btnLast.Name = "btnLast"
	$btnLast.Size = '30, 25'
	$btnLast.TabIndex = 46
	$btnLast.Text = ">>"
	$btnLast.UseVisualStyleBackColor = $True
	$btnLast.Visible = $False
	$btnLast.add_Click($btnLast_Click)
	#
	# btnNext
	#
	$btnNext.Location = '450, 0'
	$btnNext.Name = "btnNext"
	$btnNext.Size = '30, 25'
	$btnNext.TabIndex = 45
	$btnNext.Text = ">"
	$btnNext.UseVisualStyleBackColor = $True
	$btnNext.Visible = $False
	$btnNext.add_Click($btnNext_Click)
	#
	# btnPrev
	#
	$btnPrev.Location = '419, 0'
	$btnPrev.Name = "btnPrev"
	$btnPrev.Size = '30, 25'
	$btnPrev.TabIndex = 44
	$btnPrev.Text = "<"
	$btnPrev.UseVisualStyleBackColor = $True
	$btnPrev.Visible = $False
	$btnPrev.add_Click($btnPrev_Click)
	#
	# btnFirst
	#
	$btnFirst.Location = '388, 0'
	$btnFirst.Name = "btnFirst"
	$btnFirst.Size = '30, 25'
	$btnFirst.TabIndex = 43
	$btnFirst.Text = "<<"
	$btnFirst.UseVisualStyleBackColor = $True
	$btnFirst.Visible = $False
	$btnFirst.add_Click($btnFirst_Click)
	#
	# btnImportCSV
	#
	$btnImportCSV.Location = '312, 0'
	$btnImportCSV.Name = "btnImportCSV"
	$btnImportCSV.Size = '75, 25'
	$btnImportCSV.TabIndex = 42
	$btnImportCSV.Text = "Import CSV"
	$btnImportCSV.UseVisualStyleBackColor = $True
	$btnImportCSV.Visible = $False
	$btnImportCSV.add_Click($btnImportCSV_Click)
	#
	# lvCSV
	#
	$lvCSV.FullRowSelect = $True
	$lvCSV.GridLines = $True
	$lvCSV.Location = '312, 35'
	$lvCSV.Name = "lvCSV"
	$lvCSV.Size = '1200, 770'
	$lvCSV.TabIndex = 41
	$lvCSV.UseCompatibleStateImageBehavior = $False
	$lvCSV.View = 'Details'
	$lvCSV.Visible = $False
	$lvCSV.add_SelectedIndexChanged($lvCSV_SelectedIndexChanged)
	#
	# labelUPNSuffix
	#
	$labelUPNSuffix.Location = '10, 505'
	$labelUPNSuffix.Name = "labelUPNSuffix"
	$labelUPNSuffix.Size = '100, 20'
	$labelUPNSuffix.TabIndex = 49
	$labelUPNSuffix.Text = "UPNSuffix"
	$labelUPNSuffix.TextAlign = 'MiddleLeft'
	#
	# cboUPNSuffix
	#
	$cboUPNSuffix.Anchor = 'Top, Left, Right'
	$cboUPNSuffix.FormattingEnabled = $True
	$cboUPNSuffix.Location = '118, 505'
	$cboUPNSuffix.Name = "cboUPNSuffix"
	$cboUPNSuffix.Size = '175, 21'
	$cboUPNSuffix.TabIndex = 51
	$cboUPNSuffix.add_SelectedIndexChanged($cboUPNSuffix_SelectedIndexChanged)
	#
	# txtUPN
	#
	$txtUPN.Anchor = 'Top, Left, Right'
	$txtUPN.Location = '118, 530'
	$txtUPN.Name = "txtUPN"
	$txtUPN.Size = '175, 20'
	$txtUPN.TabIndex = 18
	$txtUPN.ReadOnly = $true
	#
	# txtsAM
	#
	$txtsAM.Anchor = 'Top, Left, Right'
	$txtsAM.Location = '118, 480'
	$txtsAM.Name = "txtsAM"
	$txtsAM.Size = '175, 20'
	$txtsAM.TabIndex = 17
	#
	# txtDN
	#
	$txtDN.Anchor = 'Top, Left, Right'
	$txtDN.Location = '118, 455'
	$txtDN.Name = "txtDN"
	$txtDN.Size = '175, 20'
	$txtDN.TabIndex = 16
	#
	# cboDepartment
	#
	$cboDepartment.Anchor = 'Top, Left, Right'
	$cboDepartment.FormattingEnabled = $True
	$cboDepartment.Location = '118, 235'
	$cboDepartment.Name = "cboDepartment"
	$cboDepartment.Size = '175, 21'
	$cboDepartment.TabIndex = 8
	$cboDepartment.add_SelectedIndexChanged($cboDepartment_SelectedIndexChanged)
	#
	# labelUserPrincipalName
	#
	$labelUserPrincipalName.Location = '10, 530'
	$labelUserPrincipalName.Name = "labelUserPrincipalName"
	$labelUserPrincipalName.Size = '100, 23'
	$labelUserPrincipalName.TabIndex = 40
	$labelUserPrincipalName.Text = "userPrincipalName"
	$labelUserPrincipalName.TextAlign = 'MiddleLeft'
	#
	# labelSamAccountName
	#
	$labelSamAccountName.Location = '10, 480'
	$labelSamAccountName.Name = "labelSamAccountName"
	$labelSamAccountName.Size = '100, 23'
	$labelSamAccountName.TabIndex = 39
	$labelSamAccountName.Text = "samAccountName"
	$labelSamAccountName.TextAlign = 'MiddleLeft'
	#
	# labelDisplayName
	#
	$labelDisplayName.Location = '10, 455'
	$labelDisplayName.Name = "labelDisplayName"
	$labelDisplayName.Size = '100, 23'
	$labelDisplayName.TabIndex = 38
	$labelDisplayName.Text = "Display Name"
	$labelDisplayName.TextAlign = 'MiddleLeft'
	#
	# SB
	#
	$SB.Location = '0, 1000'
	$SB.Name = "SB"
	$SB.Size = '304, 22'
	$SB.TabIndex = 48
	$SB.Text = "Ready"
	#
	# labelEmails
	#
	$labelEmails.Location = '10, 585'
	$labelEmails.Name = "labelEmails"
	$labelEmails.Size = '100, 23'
	$labelEmails.TabIndex = 53
	$labelEmails.Text = "Email Address"
	$labelEmails.TextAlign = 'MiddleLeft'
	#
	# labelEmailFormat
	#
	$labelEmailFormat.Location = '10, 925'
	$labelEmailFormat.Name = "labelEmails"
	$labelEmailFormat.Size = '250, 23'
	$labelEmailFormat.TabIndex = 55
	$labelEmailFormat.Text = "Please enter each full address on it's own line"
	$labelEmailFormat.TextAlign = 'MiddleLeft'
	#
	# txtEmails
	#
	$txtEmails.Anchor = 'Top, Left, Right'
	$txtEmails.Location = '118, 585'
	$txtEmails.Name = "txtEmails"
	$txtEmails.Size = '175, 21'
	$txtEmails.TabIndex = 54
	$txtEmails.Multiline = $True
	$txtEmails.add_TextChanged($txtEmails_TextChanged)
	#
	# listboxEmails
	#
	$listboxEmails.Anchor = 'Top, Left, Right'
	$listboxEmails.Location = '10, 650'
	$listboxEmails.Name = "listboxEmails"
	$listboxEmails.Size = '283, 100'
	$listboxEmails.TabIndex = 52
	$listboxEmails.Font = "Courier New,10"
	$listboxEmails.SelectionMode = "MultiSimple"
	#
	# btnAddAlias
	#
	$btnAddAlias.Location = '9, 751'
	$btnAddAlias.Name = "btnAddAlias"
	$btnAddAlias.Size = '142, 25'
	$btnAddAlias.TabIndex = 56
	$btnAddAlias.Text = "Add New Email Alias"
	$btnAddAlias.UseVisualStyleBackColor = $True
	$btnAddAlias.add_Click($btnAddAlias_Click)
	#
	# btnRemoveAlias
	#
	$btnRemoveAlias.Location = '157, 751'
	$btnRemoveAlias.Name = "btnRemoveAlias"
	$btnRemoveAlias.Size = '142, 25'
	$btnRemoveAlias.TabIndex = 57
	$btnRemoveAlias.Text = "Remove Selected Aliases"
	$btnRemoveAlias.UseVisualStyleBackColor = $True
	$btnRemoveAlias.add_Click($btnRemoveAlias_Click)
	#
	# labelGroups
	#
	$labelGroups.Location = '10, 780'
	$labelGroups.Name = "labelGroups"
	$labelGroups.Size = '50, 20'
	$labelGroups.TabIndex = 51
	$labelGroups.Text = "Groups:"
	$labelGroups.TextAlign = 'MiddleLeft'
	#
	# listboxGroups
	#
	$listboxGroups.Anchor = 'Top, Left, Right'
	$listboxGroups.Location = '10, 800'
	$listboxGroups.Name = "listboxGroups"
	$listboxGroups.Size = '283, 170'
	$listboxGroups.TabIndex = 52
	$listboxGroups.Font = "Courier New,10"
    $listboxGroups.SelectionMode = "MultiSimple"
	#
	# cboSite
	#
	$cboSite.Anchor = 'Top, Left, Right'
	$cboSite.FormattingEnabled = $True
	$cboSite.Location = '118, 320'
	$cboSite.Name = "cboSite"
	$cboSite.Size = '175, 21'
	$cboSite.TabIndex = 11
	$cboSite.add_SelectedIndexChanged($cboSite_SelectedIndexChanged)
	#
	# labelSite
	#
	$labelSite.Location = '10, 320'
	$labelSite.Name = "labelSite"
	$labelSite.Size = '100, 23'
	$labelSite.TabIndex = 32
	$labelSite.Text = "Site"
	$labelSite.TextAlign = 'MiddleLeft'
	#
	# txtDescription
	#
	$txtDescription.Anchor = 'Top, Left, Right'
	$txtDescription.Location = '118, 210'
	$txtDescription.Name = "txtDescription"
	$txtDescription.Size = '175, 21'
	$txtDescription.TabIndex = 7
	#
	# txtPassword
	#
	$txtPassword.Anchor = 'Top, Left, Right'
	$txtPassword.Location = '118, 555'
	$txtPassword.Name = "txtPassword"
	$txtPassword.Size = '175, 20'
	$txtPassword.TabIndex = 19
	#$txtPassword.UseSystemPasswordChar = $True
	#
	# labelPassword
	#
	$labelPassword.Location = '10, 555'
	$labelPassword.Name = "labelPassword"
	$labelPassword.Size = '100, 20'
	$labelPassword.TabIndex = 37
	$labelPassword.Text = "Password"
	$labelPassword.TextAlign = 'MiddleLeft'
	#
	# cboDomain
	#
	$cboDomain.Anchor = 'Top, Left, Right'
	$cboDomain.FormattingEnabled = $True
	$cboDomain.Location = '118, 35'
	$cboDomain.Name = "cboDomain"
	$cboDomain.Size = '175, 21'
	$cboDomain.TabIndex = 1
	$cboDomain.add_SelectedIndexChanged($cboDomain_SelectedIndexChanged)
	#
	# labelCurrentDomain
	#
	$labelCurrentDomain.Location = '10, 35'
	$labelCurrentDomain.Name = "labelCurrentDomain"
	$labelCurrentDomain.Size = '100, 23'
	$labelCurrentDomain.TabIndex = 22
	$labelCurrentDomain.Text = "Current Domain"
	$labelCurrentDomain.TextAlign = 'MiddleLeft'
	#
	# txtPostalCode
	#
	$txtPostalCode.Anchor = 'Top, Left, Right'
	$txtPostalCode.Location = '118, 420'
	$txtPostalCode.Name = "txtPostalCode"
	$txtPostalCode.Size = '175, 20'
	$txtPostalCode.TabIndex = 15
	#
	# txtState
	#
	$txtState.Anchor = 'Top, Left, Right'
	$txtState.Location = '118, 395'
	$txtState.Name = "txtState"
	$txtState.Size = '175, 20'
	$txtState.TabIndex = 14
	#
	# txtCity
	#
	$txtCity.Anchor = 'Top, Left, Right'
	$txtCity.Location = '118, 370'
	$txtCity.Name = "txtCity"
	$txtCity.Size = '175, 20'
	$txtCity.TabIndex = 13
	#
	# txtStreetAddress
	#
	$txtStreetAddress.Anchor = 'Top, Left, Right'
	$txtStreetAddress.Location = '118, 345'
	$txtStreetAddress.Name = "txtStreetAddress"
	$txtStreetAddress.Size = '175, 20'
	$txtStreetAddress.TabIndex = 12
	#
	# txtOffice
	#
	$txtOffice.Anchor = 'Top, Left, Right'
	$txtOffice.Location = '118, 160'
	$txtOffice.Name = "txtOffice"
	$txtOffice.Size = '175, 20'
	$txtOffice.TabIndex = 5
	#
	# txtCompany
	#
	$txtCompany.Anchor = 'Top, Left, Right'
	$txtCompany.Location = '118, 260'
	$txtCompany.Name = "txtCompany"
	$txtCompany.Size = '175, 20'
	$txtCompany.TabIndex = 9
	#
	# cboTitle
	#
	$cboTitle.Anchor = 'Top, Left, Right'
	$cboTitle.Location = '118, 185'
	$cboTitle.Name = "cboTitle"
	$cboTitle.Size = '175, 20'
	$cboTitle.TabIndex = 6
	$cboTitle.add_SelectedIndexChanged($cboTitle_SelectedIndexChanged)
	$cboTitle.add_TextChanged($cboTitle_TextChanged)
	#
	# txtOfficePhone
	#
	$txtOfficePhone.Anchor = 'Top, Left, Right'
	$txtOfficePhone.Location = '118, 285'
	$txtOfficePhone.Name = "txtOfficePhone"
	$txtOfficePhone.Size = '175, 20'
	$txtOfficePhone.TabIndex = 10
	#
	# txtLastName
	#
	$txtLastName.Anchor = 'Top, Left, Right'
	$txtLastName.Location = '118, 135'
	$txtLastName.Name = "txtLastName"
	$txtLastName.Size = '175, 20'
	$txtLastName.TabIndex = 4
	$txtLastName.add_TextChanged($txtName_TextChanged)
	#
	# cboPath
	#
	$cboPath.Anchor = 'Top, Left, Right'
	$cboPath.FormattingEnabled = $True
	$cboPath.Location = '45, 65'
	$cboPath.Name = "cboPath"
	$cboPath.Size = '249, 21'
	$cboPath.TabIndex = 2
	#
	# labelOU
	#
	$labelOU.Location = '10, 65'
	$labelOU.Name = "labelOU"
	$labelOU.Size = '36, 23'
	$labelOU.TabIndex = 23
	$labelOU.Text = "OU"
	$labelOU.TextAlign = 'MiddleLeft'
	#
	# txtFirstName
	#
	$txtFirstName.Anchor = 'Top, Left, Right'
	$txtFirstName.Location = '118, 110'
	$txtFirstName.Name = "txtFirstName"
	$txtFirstName.Size = '175, 20'
	$txtFirstName.TabIndex = 3
	$txtFirstName.add_TextChanged($txtName_TextChanged)
	#
	# labelPostalCode
	#
	$labelPostalCode.Location = '10, 420'
	$labelPostalCode.Name = "labelPostalCode"
	$labelPostalCode.Size = '100, 23'
	$labelPostalCode.TabIndex = 36
	$labelPostalCode.Text = "Postal Code"
	$labelPostalCode.TextAlign = 'MiddleLeft'
	#
	# labelState
	#
	$labelState.Location = '10, 395'
	$labelState.Name = "labelState"
	$labelState.Size = '100, 23'
	$labelState.TabIndex = 35
	$labelState.Text = "State"
	$labelState.TextAlign = 'MiddleLeft'
	#
	# labelCity
	#
	$labelCity.Location = '10, 370'
	$labelCity.Name = "labelCity"
	$labelCity.Size = '100, 23'
	$labelCity.TabIndex = 34
	$labelCity.Text = "City"
	$labelCity.TextAlign = 'MiddleLeft'
	#
	# labelStreetAddress
	#
	$labelStreetAddress.Location = '10, 345'
	$labelStreetAddress.Name = "labelStreetAddress"
	$labelStreetAddress.Size = '100, 23'
	$labelStreetAddress.TabIndex = 33
	$labelStreetAddress.Text = "Street Address"
	$labelStreetAddress.TextAlign = 'MiddleLeft'
	#
	# labelOffice
	#
	$labelOffice.Location = '10, 160'
	$labelOffice.Name = "labelOffice"
	$labelOffice.Size = '100, 23'
	$labelOffice.TabIndex = 26
	$labelOffice.Text = "Office"
	$labelOffice.TextAlign = 'MiddleLeft'
	#
	# labelCompany
	#
	$labelCompany.Location = '10, 260'
	$labelCompany.Name = "labelCompany"
	$labelCompany.Size = '100, 23'
	$labelCompany.TabIndex = 30
	$labelCompany.Text = "Company"
	$labelCompany.TextAlign = 'MiddleLeft'
	#
	# labelDepartment
	#
	$labelDepartment.Location = '10, 235'
	$labelDepartment.Name = "labelDepartment"
	$labelDepartment.Size = '100, 23'
	$labelDepartment.TabIndex = 29
	$labelDepartment.Text = "Department"
	$labelDepartment.TextAlign = 'MiddleLeft'
	#
	# labelTitle
	#
	$labelTitle.Location = '10, 185'
	$labelTitle.Name = "labelTitle"
	$labelTitle.Size = '100, 23'
	$labelTitle.TabIndex = 27
	$labelTitle.Text = "Title"
	$labelTitle.TextAlign = 'MiddleLeft'
	#
	# btnSubmit
	#
	$btnSubmit.Location = '224, 0'
	#$btnSubmit.Anchor = "Right"
	$btnSubmit.Name = "btnSubmit"
	$btnSubmit.Size = '75, 25'
	$btnSubmit.TabIndex = 21
	$btnSubmit.Text = "Submit"
	$btnSubmit.UseVisualStyleBackColor = $True
	$btnSubmit.add_Click($btnSubmit_Click)
	#
	# labelDescription
	#
	$labelDescription.Location = '10, 210'
	$labelDescription.Name = "labelDescription"
	$labelDescription.Size = '100, 23'
	$labelDescription.TabIndex = 28
	$labelDescription.Text = "Description"
	$labelDescription.TextAlign = 'MiddleLeft'
	#
	# labelOfficePhone
	#
	$labelOfficePhone.Location = '10, 285'
	$labelOfficePhone.Name = "labelOfficePhone"
	$labelOfficePhone.Size = '100, 23'
	$labelOfficePhone.TabIndex = 31
	$labelOfficePhone.Text = "Office Phone"
	$labelOfficePhone.TextAlign = 'MiddleLeft'
	#
	# labelLastName
	#
	$labelLastName.Location = '10, 135'
	$labelLastName.Name = "labelLastName"
	$labelLastName.Size = '100, 23'
	$labelLastName.TabIndex = 25
	$labelLastName.Text = "Last Name"
	$labelLastName.TextAlign = 'MiddleLeft'
	#
	# labelFirstName
	#
	$labelFirstName.Location = '10, 110'
	$labelFirstName.Name = "labelFirstName"
	$labelFirstName.Size = '100, 23'
	$labelFirstName.TabIndex = 24
	$labelFirstName.Text = "First Name"
	$labelFirstName.TextAlign = 'MiddleLeft'
	#
	# menustrip1
	#
	[void]$menustrip1.Items.Add($fileToolStripMenuItem)
	$menustrip1.Location = '0, 0'
	$menustrip1.Name = "menustrip1"
	$menustrip1.Size = '371, 24'
	$menustrip1.TabIndex = 20
	$menustrip1.Text = "menustrip1"
	#
	# fileToolStripMenuItem
	#
	[void]$fileToolStripMenuItem.DropDownItems.Add($formMode)
	[void]$fileToolStripMenuItem.DropDownItems.Add($CreateCSVTemplate)
	[void]$fileToolStripMenuItem.DropDownItems.Add($MenuExit)
	$fileToolStripMenuItem.Name = "fileToolStripMenuItem"
	$fileToolStripMenuItem.Size = '37, 20'
	$fileToolStripMenuItem.Text = "File"
	#
	# formMode
	#
	$formMode.Name = "formMode"
	$formMode.Size = '185, 22'
	$formMode.Text = "CSV Mode"
	$formMode.add_Click($formMode_Click)
	$formMode.Enabled = $False
	#
	# CSVTemplate
	#
	$CSVTemplate.CheckPathExists = $False
	$CSVTemplate.DefaultExt = "csv"
	$CSVTemplate.FileName = "ADusers.csv"
	$CSVTemplate.Filter = "CSV Files|*.csv|All Files|*.*"
	$CSVTemplate.ShowHelp = $True
	$CSVTemplate.Title = "Create CSV Template For Company LLC"
	$CSVTemplate.add_FileOk($CSVTemplate_FileOk)
	#
	# OFDImportCSV
	#
	$OFDImportCSV.FileName = "C:\ADUserAdd\ADusers.csv"
	$OFDImportCSV.ShowHelp = $True
	#
	# CreateCSVTemplate
	#
	$CreateCSVTemplate.Name = "CreateCSVTemplate"
	$CreateCSVTemplate.Size = '185, 22'
	$CreateCSVTemplate.Text = "Create CSV Template"
	$CreateCSVTemplate.add_Click($createTemplateToolStripMenuItem_Click)
	$CreateCSVTemplate.Enabled = $False
	#
	# MenuExit
	#
	$MenuExit.Name = "MenuExit"
	$MenuExit.Size = '185, 22'
	$MenuExit.Text = "Exit"
	$MenuExit.add_Click($MenuExit_Click)
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $formMain.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formMain.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formMain.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $formMain.ShowDialog()

} #End Function

#Call OnApplicationLoad to initialize
if((OnApplicationLoad) -eq $true)
{
	# Gater credentials now so I can pass them to the main function
	$AdminCredentials = Get-Credential
	# Connect to AD and O365
	Connect-ManagementPrograms -UserCredential $AdminCredentials
	# Call the form
	Invoke-EmployeeManagementForm -UserCredential $AdminCredentials | Out-Null
	# Perform cleanup
	OnApplicationExit
}