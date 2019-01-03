##################################### Authentification Bloc #####################################
$TARGET = "-SM-https%3A%2F%2Fpointage.adp.com%2Figested%2F2_02_01%2Fpointage"
$USERNAME = Get-Content -Path ".\config.txt"

$PASSWORD = (Get-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_Magic_Word" -ErrorAction 'SilentlyContinue').ADP_Magic_Word
If ($PASSWORD) 
{
    $loginUrl = "https://hr-services.fr.adp.com/ipclogin/1/loginform.fcc"

    $formFields = "TARGET=" + $TARGET + "&USER=" + $USERNAME + "&PASSWORD=" + $PASSWORD
    $AuthRequest = Invoke-WebRequest -Uri $loginUrl -Method Post -Body $formFields -ContentType "application/x-www-form-urlencoded" -SessionVariable websession
    #Write-Host ($AuthRequest).Content
    $cookies = $websession.Cookies.GetCookies($loginUrl) 
    #Write-Host $cookies
    #$cookies = $websession.Cookies.GetCookies($loginUrl)

    Write-Host "Authentification ADP"

    if ($cookies)
    {
        foreach ($cookie in $cookies)
        {
            if (($cookie.name -eq "SMSESSION") -AND ($cookie.value))
            {
                ##################################### Pointage Bloc #####################################

                $pointageUrl = "https://pointage.adp.com/igested/2_02_01/pointage"

                Write-Host "Récupération heure GMT"
                $PointageRequest = Invoke-WebRequest -Uri $pointageUrl -Method Get -WebSession $websession

                #Write-Host  ($PointageRequest).Content
                foreach ($line in ($PointageRequest).Content)
                {
                    if (($line).Contains("<input id='GMT_DATE' name='GMT_DATE' type='hidden' value="))
                    {
                        Write-Host "Pointage"
                        $GMTdate = $PointageRequest.Forms["FORM_POINTAGE"].Fields["GMT_DATE"]
                        $FormatedGMTdate = ($GMTdate).Replace("`/","%2F")
                        $FormatedGMTdate = ($FormatedGMTdate).Replace(":","%3A")
                        $FormatedGMTdate = ($FormatedGMTdate).Replace(" ","+")

                        ##################################### Enregistrement Bloc #####################################
                        $formFields = "ACTION=POI_PRES&FONCTION=&GMT_DATE=" + $FormatedGMTdate + "&USER_OFFSET=NjA%3D"
                        #Write-Host $formFields
                        $SendPointage = Invoke-WebRequest -Uri $pointageUrl -Method Post -Body $formFields  -WebSession $websession
                        #Write-Host ($SendPointage).Content
                        if (($SendPointage).Content)
                        {
                            ##################################### Validation Bloc #####################################
                            Write-Host "Validation Pointage"
                            $GMTdate = $SendPointage.Forms["FORM_POINTAGE"].Fields["GMT_DATE"]
                            $FormatedGMTdate = ($GMTdate).Replace("`/","%2F")
                            $FormatedGMTdate = ($FormatedGMTdate).Replace(":","%3A")
                            $FormatedGMTdate = ($FormatedGMTdate).Replace(" ","+")

                            $formFields = "ACTION=ENR_PRES&FONCTION=&GMT_DATE=" + $FormatedGMTdate + "&USER_OFFSET=NjA%3D"
                            $SendPointageValidation = Invoke-WebRequest -Uri $pointageUrl -Method Post -Body $formFields  -WebSession $websession

                            if (($SendPointageValidation).Content.Contains("Votre saisie a bien été enregistrée"))
                            {
                                Write-Host "Pointage réussi !"
                                pause
                            }
                            else
                            {
                                Write-Error "Une erreur s'est produite."    
                                pause
                            }
                        }
                        else
                        {
                            Write-Error "Une erreur s'est produite."
                            pause
                        }
                    }
                    else
                    {
                        Write-Error "Une erreur s'est produite."
                        pause
                    }
                }
            }
        }
    }
    else
    {
        Write-Error "Une erreur s'est produite."
        pause
    }
}
else
{
    Write-Host "Aucun mot de passe configuré, veuillez lancer le script de configuration."
    pause
}