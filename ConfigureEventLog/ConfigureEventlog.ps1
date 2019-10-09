New-EventLog -LogName "CLS_Script" -Source "Pointage" -ErrorAction 'SilentlyContinue'
new-item -path "HKLM:\Software\CLS\INVENTORY\Packages\EventLogPointage" -Force -ErrorAction 'SilentlyContinue'
new-itemproperty -path "HKLM:\Software\CLS\INVENTORY\Packages\EventLogPointage" -name "InstallDate" -value (Get-Date) -ErrorAction 'SilentlyContinue'