<#
User activation and setup
Verion 0.4
OCWS
CSV file path C:\Users\deploylnc\Desktop\users.csv
#>
<#
param (
	[Parameter(Position=0,
				HelpMessage="Path to input CSV file")]
	[alias("CSVPath")]
	[String]$InputCSVPath="C:\Users\NXLX8474\Desktop\users.csv"
)
#>

#Dialin policy
#$DialinPolicy = "Generali_Conferencing"

#Email account credentials:
$secpasswd = ConvertTo-SecureString "Weconne2016" -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ("Groupe\Weconnect", $secpasswd)

#User CSV loading
$users = $null
$users = Import-CSV C:\Sources\users.csv
$count = $users.count

write-host "User count within CSV file=" $count
write-host ""

ForEach ($user in $users) 
{
	
    if ($user.Tactical -eq 'Y'){
        write-host 'User ' -nonewline; Write-Host $user.upn -nonewline -ForegroundColor Cyan; Write-Host ' is Tactical, skipping'

	}
    elseif ($user.Tactical -eq 'N'){
        write-host 'User ' -nonewline; Write-Host $user.upn -nonewline -ForegroundColor Cyan; Write-Host ' is NOT Tactical, enabling user for enterprise voice and setting properties'

#Computing LineUri from CSV file
        if ($user.Extension -match '^\d+$'){
            $LineUri_complete = "tel:" + $user.TelURI + ";ext=" + $user.Extension
        }
        else {
            $LineUri_complete = "tel:" + $user.TelURI
        }
        write-host "Set-CsUser -Identity " -nonewline; Write-Host $user.upn -nonewline -ForegroundColor Cyan; Write-Host -NoNewline " -EnterpriseVoiceEnabled $true -LineUri "; Write-Host -ForegroundColor Green $LineUri_complete
		Set-CsUser -Identity $user.upn -EnterpriseVoiceEnabled $true -LineUri $LineUri_complete

	}

#Generating Random non trivial PIN
    do{
        $PIN=Get-random -Maximum 1000000 -Minimum 100000
    }until($PIN -ne 123456 -And $PIN -ne 234567 -And $PIN -ne 345678 -And $PIN -ne 456789 -And $PIN -ne 111111 -And $PIN -ne 222222 -And $PIN -ne 333333 -And $PIN -ne 444444 -And $PIN -ne 555555 -And $PIN -ne 666666 -And $PIN -ne 777777 -And $PIN -ne 888888 -And $PIN -ne 999999 -And $PIN -ne 000000)

#Setting PIN
	Set-CsPinSendCAWelcomeMail -UserUri $user.upn -From "weconnect@generali.fr" -Subject "Votre nouveau PIN Lync" -UserEmailAddress $users.EmailAddress -Pin $PIN -Force -SmtpServer rapport.groupe.generali.fr -Credential $mycreds
	if ($?=$true){
		write-host -NoNewline "Dial-in conferencing PIN set for user: "; Write-Host -ForegroundColor Cyan $user.upn
	}    
#Granting voice policy
    Grant-CsVoicePolicy -Identity $user.upn -PolicyName $user.VoicePolicy
	if ($?=$true){
		write-host -NoNewLine "Voice policy "; Write-Host -NoNewline -ForegroundColor Green $user.VoicePolicy; Write-Host -NoNewline " granted to useruser: "; Write-Host -ForegroundColor Cyan $user.upn
	}
#Granting Conferencing policy
<#
    if ($user.dialin -eq 'Y'){
        Grant-CsConferencingPolicy -Identity $user.upn -PolicyName $DialinPolicy
	    if ($?=$true){
		    write-host -NoNewLine "Conferencing policy $DialinPolicy granted to user "; Write-Host -NoNewline -ForegroundColor Cyan $user.upn ; Write-Host " Enabling dial-in conferencing"
	    }
    }
#>

	write-host "****************************User " -nonewline; Write-Host -foregroundcolor Cyan $user.upn -nonewline; Write-Host  " set****************************"
	write-host ""
		
}