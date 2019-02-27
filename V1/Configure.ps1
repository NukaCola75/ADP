Write-Host "###### Debut de la configuration ######"

New-Item -Name "config.txt" -Force
Set-Content -Path ".\config.txt" -Value $null
Clear-Host

$USERNAME = Read-Host "Veuillez renseigner votre identifiant ADP"
Set-Content -Path ".\config.txt" -Value $USERNAME
Clear-Host

New-Item -Path "HKCU:\Software\CLS\APP\ADP" -Force -ErrorAction 'SilentlyContinue'
Clear-Host

$PASSWORD = Read-Host "Veuillez renseigner votre mot de passe ADP"
New-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_Magic_Word" -Value $PASSWORD -ErrorAction 'SilentlyContinue'

Clear-Host
Write-Host "###### Fin de la configuration ######"
pause