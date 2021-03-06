<#
User activation and setup
Verion 0.7
OCWS
BSD 3-Clause License
#>

param([string]$userCsv = "C:\Sources\users.csv",[string]$PINfilePath = "C:\Sources\ProvisioningPIN.csv",[string]$mailUser = "Groupe\Weconnect",[string]$mailPassword = "Weconne2016")

#Email account credentials:
$secpasswd = ConvertTo-SecureString $mailPassword -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ($mailUser,$secpasswd)

#PIN file initialise
Add-Content "upn,FirstName,LastName,EmailAddress,Desk,Extension,PIN" -Path $PINfilePath

#User CSV loading
$usersList = $null
$usersList = Import-Csv $userCsv
$count = $usersList.count
Write-Host "User count within CSV file=" $count
Write-Host ""

foreach ($user in $usersList)
{

  if ($user.Tactical -eq 'True') {
    Write-Host 'User ' -NoNewline; Write-Host $user.upn -NoNewline -ForegroundColor Cyan; Write-Host ' is Tactical, skipping'

  }
  elseif ($user.Tactical -eq 'False') {
    Write-Host 'User ' -NoNewline; Write-Host $user.upn -NoNewline -ForegroundColor Cyan; Write-Host ' is NOT Tactical, enabling user for enterprise voice and setting properties'

    #Computing LineUri from CSV file
    if ($user.Extension -match '^\d+$') {
      $LineUri_complete = "tel:" + $user.TelURI + ";ext=" + $user.Extension
    }
    else {
      $LineUri_complete = "tel:" + $user.TelURI
    }
    Write-Host "Set-CsUser -Identity " -NoNewline; Write-Host $user.upn -NoNewline -ForegroundColor Cyan; Write-Host -NoNewline " -EnterpriseVoiceEnabled $true -LineUri "; Write-Host -ForegroundColor Green $LineUri_complete
    Set-CsUser -identity $user.upn -EnterpriseVoiceEnabled $true -LineUri $LineUri_complete

  }

  #Generating Random non trivial PIN
  do {
    $PIN = Get-Random -Maximum 1000000 -Minimum 100000
  } until ($PIN -ne 123456 -and $PIN -ne 012345 -and $PIN -ne 234567 -and $PIN -ne 345678 -and $PIN -ne 456789 -and $PIN -ne 111111 -and $PIN -ne 222222 -and $PIN -ne 333333 -and $PIN -ne 444444 -and $PIN -ne 555555 -and $PIN -ne 666666 -and $PIN -ne 777777 -and $PIN -ne 888888 -and $PIN -ne 999999 -and $PIN -ne 000000 -and $PIN -ne 987654 -and $PIN -ne 876543 -and $PIN -ne 765432 -and $PIN -ne 654321 -and $PIN -ne 543210)

  #Setting PIN
  Set-CsPinSendCAWelcomeMail -UserUri $user.upn -From "weconnect@generali.fr" -Subject "Votre code pin pour le nouveau téléphone" -UserEmailAddress $user.EmailAddress -Bcc "prodWC@generali.fr" -Pin $PIN -Extension $user.Extension -Force -SmtpServer rapport.groupe.generali.fr -Credential $mycreds
  if ($? -eq $true) {
    Write-Host -NoNewline "Dial-in conferencing PIN set for user: "; Write-Host -ForegroundColor Cyan $user.upn
  }

  #Writing PIN to CSV

  $line = $user.upn + "," + $user.FirstName + "," + $user.LastName + "," + $user.EmailAddress + "," + $user.Desk + "," + $user.Extension + "," + $PIN
  Add-Content $line -Path $PINfilePath

  #Granting voice policy
  Grant-CsVoicePolicy -identity $user.upn -PolicyName $user.VoicePolicy
  if ($? -eq $true) {
    Write-Host -NoNewline "Voice policy "; Write-Host -NoNewline -ForegroundColor Green $user.VoicePolicy; Write-Host -NoNewline " granted to useruser: "; Write-Host -ForegroundColor Cyan $user.upn
  }

  Write-Host "****************************User " -NoNewline; Write-Host -ForegroundColor Cyan $user.upn -NoNewline; Write-Host " set****************************"
  Write-Host ""

}
