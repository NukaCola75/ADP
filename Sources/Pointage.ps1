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

[Net.ServicePointManager]::SecurityProtocol = [net.SecurityProtocolType]::Tls12 # Active & force TLS 1.2 support

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


function WriteInEventLog($EventMessage, $ErrorLevel)
{
	Write-EventLog –LogName "CLS_Script" –Source "Pointage" –EntryType $ErrorLevel –EventID 0 –Message $EventMessage -ErrorAction 'SilentlyContinue'
}

function URLEncode($pass)
{
    Add-Type -Assembly System.Web
    $pass = [System.Web.HttpUtility]::UrlEncode($pass)
    return $pass
}

function Pointage {
    ##################################### Authentification Bloc #####################################
    $TARGET = "-SM-https://pointage.adp.com/igested/pointage"
    $FormatedTarget = URLEncode $TARGET
    $USERNAME = Get-Content -Path "$PathExecute\config.txt" -ErrorAction 'SilentlyContinue'
    $PASSWORD = (Get-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_Magic_Word" -ErrorAction 'SilentlyContinue').ADP_Magic_Word
    $PASSWORD = URLEncode $PASSWORD

    If (($PASSWORD) -AND ($USERNAME))
    {
        $loginUrl = "https://hr-services.fr.adp.com/ipclogin/1/loginform.fcc"

        $formFields = "TARGET=" + $FormatedTarget + "&USER=" + $USERNAME + "&PASSWORD=" + $PASSWORD
        $AuthRequest = Invoke-WebRequest -Uri $loginUrl -Method Post -Body $formFields -ContentType "application/x-www-form-urlencoded" -SessionVariable websession
        $cookies = $websession.Cookies.GetCookies($loginUrl) 

        if ($cookies)
        {
            foreach ($cookie in $cookies)
            {
                # Write-Host $cookie
                if (($cookie.name -eq "SMSESSION") -AND ($cookie.value))
                {
                    ##################################### Pointage Bloc #####################################
                    $formPointage = $AuthRequest.ParsedHtml.getElementsByTagName('form') | Where-Object {($_.name) -match 'FORM_POINTAGE'}
                    if ($formPointage.action -ne $null) {
                        $pointageUrl = "https://pointage.adp.com" + $formPointage.action
                        # Write-Host $pointageUrl
                        # Récupération heure GMT
                        $PointageRequest = Invoke-WebRequest -Uri $pointageUrl -Method Get -WebSession $websession

                        foreach ($line in ($PointageRequest).Content) {
                            if (($line).Contains("<input id='GMT_DATE' name='GMT_DATE' type='hidden' value=")) {
                                # Pointage
                                $GMTdate = $PointageRequest.Forms["FORM_POINTAGE"].Fields["GMT_DATE"]
                                $FormatedGMTdate = ($GMTdate).Replace("`/", "%2F")
                                $FormatedGMTdate = ($FormatedGMTdate).Replace(":", "%3A")
                                $FormatedGMTdate = ($FormatedGMTdate).Replace(" ", "+")

                                $hourToCompare = ($GMTdate).Substring(11, 2)
                                $sysHour = (Get-Date -Format HH)
                                $resultCompareHour = $sysHour - $hourToCompare
                                $Global:OFFSET = ""

                                if ($resultCompareHour -eq 1) {
                                    $OFFSET = "NjA%3D"
                                }
                                elseif ($resultCompareHour -eq 2) {
                                    $OFFSET = "MTIw"
                                }

                                ##################################### Enregistrement Bloc #####################################
                                $formFields = "ACTION=POI_PRES&FONCTION=&GMT_DATE=" + $FormatedGMTdate + "&USER_OFFSET=" + $OFFSET
                                $SendPointage = Invoke-WebRequest -Uri $pointageUrl -Method Post -Body $formFields  -WebSession $websession
                                if (($SendPointage).Content) {
                                    ##################################### Validation Bloc #####################################
                                    Write-Host "Validation Pointage"
                                    $GMTdate = $SendPointage.Forms["FORM_POINTAGE"].Fields["GMT_DATE"]
                                    $FormatedGMTdate = ($GMTdate).Replace("`/", "%2F")
                                    $FormatedGMTdate = ($FormatedGMTdate).Replace(":", "%3A")
                                    $FormatedGMTdate = ($FormatedGMTdate).Replace(" ", "+")

                                    $formFields = "ACTION=ENR_PRES&FONCTION=&GMT_DATE=" + $FormatedGMTdate + "&USER_OFFSET=" + $OFFSET
                                    $SendPointageValidation = Invoke-WebRequest -Uri $pointageUrl -Method Post -Body $formFields  -WebSession $websession
                                    if (($SendPointageValidation).Content.Contains("Votre saisie a bien été enregistrée")) {
                                        $today = Get-Date -Format 'dd/MM/yyyy'
                                        WriteInEventLog "Pointage: $today" "Information"
                                        $res = msgbox "Pointage réussi !" "Succès" ok "Information"
                                        Exit
                                    }
                                    else {
                                        $res = msgbox "Echec du pointage !" "Attention" ok "Error"
                                        WriteInEventLog "Error while pointing." "Error"
                                        Exit
                                    }
                                }
                                else {
                                    $res = msgbox "Une erreur s'est produite. Veuillez vérifier votre mot de passe ainsi que votre connectivité réseau. Si le problème persiste, veuillez tester une connexion manuelle à ADP." "Attention" ok "Error"
                                    WriteInEventLog "An error has occured. Pointage request." "Error"
                                    Exit
                                }
                            }
                            else {
                                $res = msgbox "Une erreur s'est produite. Veuillez vérifier votre mot de passe ainsi que votre connectivité réseau. Si le problème persiste, veuillez tester une connexion manuelle à ADP." "Attention" ok "Error"
                                WriteInEventLog "An error has occured. No date provided by ADP." "Error"
                                Exit
                            }
                        }
                    }
                    else {
                        $res = msgbox "Une erreur s'est produite. Veuillez vérifier votre mot de passe ainsi que votre connectivité réseau. Si le problème persiste, veuillez tester une connexion manuelle à ADP." "Attention" ok "Error"
                        WriteInEventLog "An error has occured. No pointage url provided by ADP." "Error"
                        Exit
                    }
                }
                elseif ($cookie.name -eq "SMTRYNO")
                {
                    $res = msgbox "Une erreur s'est produite. Votre mot de passe est peut-être incorrect ou est sur le point d'expirer. Veuillez vous connecter manuellement sur ADP." "Attention" ok "Error"
                    WriteInEventLog "An error has occured. Bad password or password must be changed." "Error"
                    Exit
                }
            }
            $res = msgbox "Une erreur s'est produite. Veuillez vérifier votre mot de passe ainsi que votre connectivité réseau. Si le problème persiste, veuillez tester une connexion manuelle à ADP." "Attention" ok "Error"
            WriteInEventLog "An error has occured. Bad cookies." "Error"
            Exit
        }
        else
        {
            $res = msgbox "Une erreur s'est produite. Veuillez vérifier votre mot de passe ainsi que votre connectivité réseau. Si le problème persiste, veuillez tester une connexion manuelle à ADP." "Attention" ok "Error"
            WriteInEventLog "An error has occured. No cookies." "Error"
            Exit
        }
    }
    else
    {
        $res = msgbox "L'application n'est pas configurée, souhaitez vous lancer l'outil de configuration ?" "Attention" YesNo "Exclamation"
        WriteInEventLog "The application is not configured." "Error"
        if ($res -eq "Yes")
        {
            $form.Hide();
            $form.Close();
            . $PathExecute\Configure.ps1
            Exit
        }
        else
        {
            Exit
        }
        # Exit
    }
}

function cadrePointage {
    ##################################### Authentification Bloc #####################################
    $TARGET = "-SM-https://pointage.adp.com/igested/pointage"
    $FormatedTarget = URLEncode $TARGET
    $USERNAME = Get-Content -Path "$PathExecute\config.txt" -ErrorAction 'SilentlyContinue'
    $PASSWORD = (Get-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_Magic_Word" -ErrorAction 'SilentlyContinue').ADP_Magic_Word
    $PASSWORD = URLEncode $PASSWORD

    If (($PASSWORD) -AND ($USERNAME))
    {
        $loginUrl = "https://hr-services.fr.adp.com/ipclogin/1/loginform.fcc"
        $formFields = "TARGET=" +  $FormatedTarget + "&USER=" + $USERNAME + "&PASSWORD=" + $PASSWORD
        $AuthRequest = Invoke-WebRequest -Uri $loginUrl -Method Post -Body $formFields -ContentType "application/x-www-form-urlencoded" -SessionVariable websession
        $cookies = $websession.Cookies.GetCookies($loginUrl)

        if ($cookies)
        {
            foreach ($cookie in $cookies)
            {
                if (($cookie.name -eq "SMSESSION") -AND ($cookie.value))
                {
                    ##################################### Check Bloc #####################################
                    $formPointage = $AuthRequest.ParsedHtml.getElementsByTagName('form') | Where-Object {($_.name).toUpper() -match 'FORM_POINTAGE'}
                    if ($formPointage.action -ne $null) {
                        $pointageUrl = "https://pointage.adp.com" + $formPointage.action
                        $PointageRequest = Invoke-WebRequest -Uri $pointageUrl -Method Get -WebSession $websession
                        $formFields = "ACTION=POI_CONS&FONCTION=&GMT_DATE=" + $FormatedGMTdate + "&USER_OFFSET=NjA%3D"
                        $GetPointageDate = Invoke-WebRequest -Uri $pointageUrl -Method Post -Body $formFields  -WebSession $websession
                        $datepointage = ($GetPointageDate.ParsedHtml.getElementsByTagName("td") | Where-Object { $_.className -eq "cel_liste col_date" }).innertext
                        if ($datepointage -match $today) {
                            WriteInEventLog "User has already pointed today." "Information"
                            $res = msgbox "Vous avez déjà pointé aujourd'hui." "Succès" ok "Exclamation"
                            Exit
                        }
                        else {
                            $res = msgbox "Vous n'avez pas pointé aujourd'hui. Souhaitez vous le faire ?" "Attention" YesNo "Question"
                            WriteInEventLog "Ask for user pointage." "Warning"
                            if ($res -eq "Yes") {
                                Pointage
                            }
                            else {
                                WriteInEventLog "User doesn't want to point for the moment." "Warning"
                                Exit
                            }
                        }
                    }
                }
                elseif ($cookie.name -eq "SMTRYNO")
                {
                    $res = msgbox "Une erreur s'est produite. Votre mot de passe est peut-être incorrect ou est sur le point d'expirer. Veuillez vous connecter manuellement sur ADP." "Attention" ok "Error"
                    WriteInEventLog "An error has occured. Bad password or password must be changed." "Error"
                    Exit
                }
            }
            $res = msgbox "Une erreur s'est produite. Veuillez vérifier votre mot de passe ainsi que votre connectivité réseau. Si le problème persiste, veuillez tester une connexion manuelle à ADP." "Attention" ok "Error"
            WriteInEventLog "An error has occured. Bad cookies." "Error"
            Exit
        }
    }
}

function testADP {
    If (!(Test-Connection "adp.com" -Quiet))
    {
        $res = msgbox "Le site adp.com n'est pas opérationnel ou vous n'êtes pas connecté au réseau." "Attention" ok "Error"
        WriteInEventLog "ADP is not reachable or network connectivity is down." "Error"
        Exit
    }
}

# Current Path
$CurrentPath = Get-Location
$PathExecute = (Convert-Path $CurrentPath)
$today = Get-Date -Format 'dd/MM/yyyy'
$cadre = (Get-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_cadre" -ErrorAction 'SilentlyContinue').ADP_cadre

$lastEvent = Get-EventLog -LogName CLS_Script -Source pointage -Newest 1 -EntryType Information -ErrorAction 'SilentlyContinue'

testADP

if (($lastEvent -ne $null) -AND ($cadre -eq $true))
{
    if ((($lastEvent.TimeGenerated).ToShortDateString() -eq (get-date).ToShortDateString()) -AND ($cadre -eq $true))
    {
        WriteInEventLog "User has already pointed today." "Warning"
        Exit
    }
    elseif ((($lastEvent.TimeGenerated).ToShortDateString() -ne (get-date).ToShortDateString()) -AND ($cadre -eq $true))
    {
        cadrePointage
    }
}
elseif (($lastEvent -eq $null) -and ($cadre -eq $true)) 
{
    cadrePointage
}
elseif ($cadre -eq $false)
{
    Pointage
}
else
{
    Pointage
}