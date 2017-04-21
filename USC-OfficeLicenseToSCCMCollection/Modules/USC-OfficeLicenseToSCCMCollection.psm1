[CmdLetBinding()]
Param($LocalADBase="OU=Staff,DC=USC,DC=INTERNAL",$TargetCollection)

Import-Module AzureAD
Import-Module ActiveDirectory

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

    If ($SavePassword) {
        If (-Not (Test-Path -Path (Split-Path -Path $VarFile -Parent))) {
            New-Item -Path (Split-Path -Path $VarFile -Parent) -ItemType Directory
        }
        Get-Variable USC* -Scope Global | Export-Clixml -Path $VarFile
    }
}

Import-GlobalVars -Verbose
#Lets make sure we only get objects from the Staff OU and not Students
Get-ADUser -SearchBase "