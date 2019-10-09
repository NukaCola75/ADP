Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
 
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
function HIDE-CONSOLE($hide)
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, $hide)
}

# Hide Powershell console host
Hide-Console 0


function msgbox
{
    param
    (
        [string]$Message,
        [string]$Title = 'Message box title',   
        [string]$buttons = 'OKCancel',
        [string]$icon = 'Exclamation'
    )
    # This function displays a message box by calling the .Net Windows.Forms (MessageBox class)
     
    # Load the assembly
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
     
    # Define the button types
    switch ($buttons)
    {
       'ok' {$btn = [System.Windows.Forms.MessageBoxButtons]::OK; break}
       'okcancel' {$btn = [System.Windows.Forms.MessageBoxButtons]::OKCancel; break}
       'AbortRetryIgnore' {$btn = [System.Windows.Forms.MessageBoxButtons]::AbortRetryIgnore; break}
       'YesNoCancel' {$btn = [System.Windows.Forms.MessageBoxButtons]::YesNoCancel; break}
       'YesNo' {$btn = [System.Windows.Forms.MessageBoxButtons]::yesno; break}
       'RetryCancel'{$btn = [System.Windows.Forms.MessageBoxButtons]::RetryCancel; break}
       default {$btn = [System.Windows.Forms.MessageBoxButtons]::RetryCancel; break}
    }

    $displayType = [System.Windows.Forms.MessageBoxOptions]"ServiceNotification"
     
    # Display the message box
    $Return=[System.Windows.Forms.MessageBox]::Show($Message,$Title,$btn,$icon, 'Button2', $displayType)
    $Return
}

function Pointage {
    ##################################### Authentification Bloc #####################################
    $TARGET = "-SM-https%3A%2F%2Fpointage.adp.com%2Figested%2F2_02_01%2Fpointage"
    $USERNAME = Get-Content -Path "$PathExecute\config.txt" -ErrorAction 'SilentlyContinue'
    $PASSWORD = (Get-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_Magic_Word" -ErrorAction 'SilentlyContinue').ADP_Magic_Word

    If (($PASSWORD) -AND ($USERNAME))
    {
        $loginUrl = "https://hr-services.fr.adp.com/ipclogin/1/loginform.fcc"

        $formFields = "TARGET=" + $TARGET + "&USER=" + $USERNAME + "&PASSWORD=" + $PASSWORD
        $AuthRequest = Invoke-WebRequest -Uri $loginUrl -Method Post -Body $formFields -ContentType "application/x-www-form-urlencoded" -SessionVariable websession
        $cookies = $websession.Cookies.GetCookies($loginUrl) 

        if ($cookies)
        {
            foreach ($cookie in $cookies)
            {
                if (($cookie.name -eq "SMSESSION") -AND ($cookie.value))
                {
                    ##################################### Pointage Bloc #####################################
                    $pointageUrl = "https://pointage.adp.com/igested/2_02_01/pointage"

                    # R�cup�ration heure GMT
                    $PointageRequest = Invoke-WebRequest -Uri $pointageUrl -Method Get -WebSession $websession

                    foreach ($line in ($PointageRequest).Content)
                    {
                        if (($line).Contains("<input id='GMT_DATE' name='GMT_DATE' type='hidden' value="))
                        {
                            # Pointage
                            $GMTdate = $PointageRequest.Forms["FORM_POINTAGE"].Fields["GMT_DATE"]
                            $FormatedGMTdate = ($GMTdate).Replace("`/","%2F")
                            $FormatedGMTdate = ($FormatedGMTdate).Replace(":","%3A")
                            $FormatedGMTdate = ($FormatedGMTdate).Replace(" ","+")

                            $hourToCompare = ($GMTdate).Substring(11, 2)
                            $sysHour =  (Get-Date -Format HH)
                            $resultCompareHour = $sysHour - $hourToCompare
                            $Global:OFFSET = ""

                            if ($resultCompareHour -eq 1)
                            {
                                $OFFSET = "NjA%3D"
                            }
                            elseif ($resultCompareHour -eq 2)
                            {
                                $OFFSET = "MTIw"
                            }

                            ##################################### Enregistrement Bloc #####################################
                            $formFields = "ACTION=POI_PRES&FONCTION=&GMT_DATE=" + $FormatedGMTdate + "&USER_OFFSET=" + $OFFSET
                            $SendPointage = Invoke-WebRequest -Uri $pointageUrl -Method Post -Body $formFields  -WebSession $websession
                            if (($SendPointage).Content)
                            {
                                ##################################### Validation Bloc #####################################
                                Write-Host "Validation Pointage"
                                $GMTdate = $SendPointage.Forms["FORM_POINTAGE"].Fields["GMT_DATE"]
                                $FormatedGMTdate = ($GMTdate).Replace("`/","%2F")
                                $FormatedGMTdate = ($FormatedGMTdate).Replace(":","%3A")
                                $FormatedGMTdate = ($FormatedGMTdate).Replace(" ","+")

                                $formFields = "ACTION=ENR_PRES&FONCTION=&GMT_DATE=" + $FormatedGMTdate + "&USER_OFFSET=" + $OFFSET
                                $SendPointageValidation = Invoke-WebRequest -Uri $pointageUrl -Method Post -Body $formFields  -WebSession $websession
                                if (($SendPointageValidation).Content.Contains("Votre saisie a bien �t� enregistr�e"))
                                {
                                    $today = Get-Date -Format 'dd/MM/yyyy'
                                    New-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_pointageDate" -Value $today -Force -ErrorAction 'SilentlyContinue'
                                    $res = msgbox "Pointage r�ussi !" "Succ�s" ok "Exclamation"
                                    Exit
                                }
                                else
                                {
                                    $res = msgbox "Echec du pointage !" "Attention" ok "Error"
                                    Exit
                                }
                            }
                            else
                            {
                                $res = msgbox "Une erreur s'est produite. Veuillez v�rifier votre connectivit� r�seau. Si le probl�me persiste, veuillez tester une connexion manuelle � ADP." "Attention" ok "Error"
                                Exit
                            }
                        }
                        else
                        {
                            $res = msgbox "Une erreur s'est produite. Veuillez v�rifier votre connectivit� r�seau. Si le probl�me persiste, veuillez tester une connexion manuelle � ADP." "Attention" ok "Error"
                            Exit
                        }
                    }
                }
                else 
                {
                    $res = msgbox "Une erreur s'est produite. Veuillez v�rifier votre connectivit� r�seau. Si le probl�me persiste, veuillez tester une connexion manuelle � ADP." "Attention" ok "Error"
                    Exit
                }
            }
        }
        else
        {
            $res = msgbox "Une erreur s'est produite. Veuillez v�rifier votre connectivit� r�seau. Si le probl�me persiste, veuillez tester une connexion manuelle � ADP." "Attention" ok "Error"
            Exit
        }
    }
    else
    {
        $res = msgbox "L'application n'est pas configur�e, veuillez lancer l'outil de configuration." "Attention" ok "Exclamation"
        Exit
    }
}

function testADP {
    If (!(Test-Connection "adp.com" -Quiet))
    {
        $res = msgbox "Le site adp.com n'est pas op�rationnel ou vous n'�tes pas connect� au r�seau." "Attention" ok "Error"
        Exit
    }
}

# Current Path
$CurrentPath = Get-Location
$PathExecute = (Convert-Path $CurrentPath)

$today = Get-Date -Format 'dd/MM/yyyy'
$storedDate = (Get-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_pointageDate" -ErrorAction 'SilentlyContinue').ADP_pointageDate
$cadre = (Get-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_cadre" -ErrorAction 'SilentlyContinue').ADP_cadre

testADP

if (($today -eq $storedDate) -AND ($cadre -eq $true))
{
    Exit
}
elseif (($today -ne $storedDate) -AND ($cadre -eq $true))
{
    $res = msgbox "Vous n'avez pas point� aujourd'hui. Souhaitez vous le faire ?" "Attention" YesNo "Question"
    if ($res -eq "Yes")
    {
        Pointage
    }
    else
    {
        Exit
    }
}
elseif ($cadre -eq $false)
{
    Pointage
}
else
{
    Pointage
}