<#
    .SYNOPSIS
    Retreive user accounts with Office 365 license in Azure and adds them to a local SCCM Collection
    
    .DESCRIPTION
    Connects to a tennant ID and retrieves user accounts scoped to a local AD Organization unit. Connects to Configuration Manager and adds them to a collection.
    
    .PARAMETER LocalADBase
    Specifies OU strucutre to limit accounts searched to.

    .PARAMETER TargetCollection
    Specifies the collection for users to be added to.

    .PARAMETER UserGroupFilter
    Used for testing to filter down scope of users searched for.

    .PARAMETER USCSiteServer
    If Specified, attempts to connect to a Config Manager site here. Otherwise prompted.

    .PARAMETER LicenseListFilter
    Specify an array of licenses to search on. Otherwise all licenses are specified.

    .EXAMPLE
    C:\>
    
    .EXAMPLE
    C:\PS>Get-CfgCollectionMembers "VMware vSphere Client 4.1 MSI WKS"
    ComputerName                                                Collection
	------------                                                ----------
	GMNQ12S                                                     VMware vSphere Client 4.1 MSI WKS
    
       
    .NOTES
    Author: Jesse Harris
    For: University of Sunshine Coast
    Date Created: 09 Jan 2012        
    ChangeLog:
#>
[CmdLetBinding()]
Param($LocalADBase="OU=Staff,DC=USC,DC=INTERNAL",$TargetCollection,$UserGroupFilter,$USCSiteServer,$LicenseListFilter = 'PROJECTCLIENT_FACULTY,VISIOCLIENT_FACULTY,PROJECTPROFESSIONAL_FACULTY')
Start-Transcript -Path C:\Scripts\AzureToSCCM.log
$VerbosePreference = 'Continue'
$Modules = 'AzureAD','ActiveDirectory','ConfigurationManager'
ForEach ($Module in $Modules) {
    If (-Not (Get-Module -Name $Module -EA SilentlyContinue)) {
        Import-Module $Module -EA Stop
    }
}

function Import-GlobalVars {
[CmdLetBinding()]
Param()
    $VarFile = "$env:LOCALAPPDATA\USC\OfficeLicenseToSCCMCollection.xml"
    #Try import Secure Username and Password
    If (Test-Path $VarFile) {
        Import-Clixml -Path $VarFile | %{ Set-Variable $_.Name $_.Value -Scope Global }
    
    }
    If ($USCusername -and $USCpassword) { 
        $SecureString = $USCpassword | ConvertTo-SecureString -AsPlainText -Force
        $credential = New-Object -TypeName System.Management.Automation.PSCredential $USCusername,$SecureString
        $connectionTestSuccess = Connect-AzureAD -Credential $credential -Verbose
    }
    while (-Not $connectionTestSuccess -or (-Not $USCusername)) {
        $Global:USCusername = Read-Host -Prompt "Enter your AzureAD Username: "
        $Global:USCpassword = Read-Host -Prompt "Enter your AzureAD Password: "
        $SecureString = $USCpassword | ConvertTo-SecureString -AsPlainText -Force
        $credential = New-Object -TypeName System.Management.Automation.PSCredential $USCusername,$SecureString
        $connectionTestSuccess = Connect-AzureAD -Credential $credential -Verbose
        $SavePassword = $True
    }
    Write-Verbose "Succefully connected to AzureAD"
    If ($Global:USCSiteServer) { $connectionTestSuccess = Test-connection -ComputerName $Global:USCSiteServer -Count 1 -Quiet }
    while ($connectionTestSuccess -ne $True -and (-Not $USCSiteServer)) {
        $Global:USCSiteServer = Read-Host -Prompt "Enter your site server hostname: "
        $connectionTestSuccess = Test-connection -ComputerName $USCSiteServer -Count 1 -Quiet 
    }

    If ($SavePassword -or ($USCSiteCode -eq $null)) {
        $Global:USCSiteCode = (Get-WmiObject -ComputerName $USCSiteServer -Namespace root\sms -Class SMS_ProviderLocation -EA SilentlyContinue).SiteCode
        If (-Not $USCSiteCode) {
            Write-Error "Could not obtain sitecode from $USCSiteServer"
            return
        } Else {
            If (-Not (Test-Path -Path (Split-Path -Path $VarFile -Parent))) {
                New-Item -Path (Split-Path -Path $VarFile -Parent) -ItemType Directory
            }
            Get-Variable USC* -Scope Global | Export-Clixml -Path $VarFile
        }
    }
}

Import-GlobalVars -Verbose
#Get Config Manager Site Server
If (-Not (Get-PSDrive -Name $USCSiteCode -EA SilentlyContinue)) {
    New-PSDrive -Name $USCSiteCode -PSProvider CMSite -Root $USCSiteServer
}
#Get AD Domain
$Domain = (Get-ADDomain).Name
#Lets make sure we only get objects from the Staff OU and not Students
If ($UserGroupFilter) {
    $LocalUsers = Get-ADGroupMember -Identity $UserGroupFilter | Get-ADUser
} Else {
    $LocalUsers = Get-ADUser -SearchBase $LocalADBase -Filter {Enabled -eq 'True'}
}
$i=0
$LicenseHash=@()
$LocalUsers | ForEach-Object {
    $i++
    [int]$Percent = $i*100/$LocalUsers.Count
    Write-Progress -Activity "Getting user licenses" -Status "$Percent% Complete out of $($LocalUsers.Count)" -PercentComplete $Percent
    $Licenses = (Get-AzureADUser -ObjectId $_.UserPrincipalName -ErrorAction SilentlyContinue | Get-AzureADUserLicenseDetail).SkuPartNumber
    ForEach ($License in $Licenses) {
        If ($LicenseListFilter) {
            $ListOfLicenses = $LicenseListFilter.Split(',')
            If ($License -in $ListOfLicenses) {
                Write-Verbose "Add user $($_.UserPrincipalName) to collection $License"
                #Setup collections initially
                $LicenseHash+=@{Name = $($_.UserPrincipalName); License = $License}
            }
        } Else {
            Write-Verbose "Add user $($_.UserPrincipalName) to collection $License"
            #Setup collections initially
            $LicenseHash+=@{Name = $($_.UserPrincipalName); License = $License}
        }
    }
}
Write-Progress -Activity "Getting user licenses" -Status "Complete." -Completed
$i = 0
ForEach ($UniqueLicense in ($LicenseHash.License|Select-Object -Unique | Where-Object {$LicenseHash.IndexOf($_)})) {
    $UserList = $LicenseHash | ForEach-Object{ 
        $UserLicense = $_
        If ($UserLicense.License -eq $UniqueLicense ) {
            $UserLicense.Name
        }
    }
    Push-Location
    Set-Location -Path "$($USCSiteCode):\"

    #Get User ResourceID
    $UserList | ForEach-Object {
        $i++
        [int]$Percent = $i*100/$UserList.Count
        Write-Progress -Activity "Adding users to collections" -Status "Adding $UserDomainName to $UniqueLicense" -PercentComplete $Percent
        $LicenseCollection = Get-CMCollection -Name $UniqueLicense -CollectionType User
        If (-Not $LicenseCollection) {
            $LicenseCollection = New-CMCollection -CollectionType User -Name $UniqueLicense -LimitingCollectionName 'All Users'
        }
        $UserDomainName = "$Domain\$($_.split('@')[0])"
        $UserResource = Get-CMUser -Name $UserDomainName
        Add-CMUserCollectionDirectMembershipRule -InputObject $LicenseCollection -ResourceId $UserResource.ResourceID -ErrorAction SilentlyContinue
    }
    Write-Progress -Activity "Adding users to collection" -Status "Complete." -Completed

    #Reconcile Memberships
    <#Get-CMCollection -Name $UniqueLicense -CollectionType User
    Get-CMCollectionMember -CollectionName 'PROJECTCLIENT_FACULTY'
    #>
    Pop-Location


}
Stop-Transcript
