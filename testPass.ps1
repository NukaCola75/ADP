function URLEncode($pass)
{
    Add-Type -Assembly System.Web
    $pass = [System.Web.HttpUtility]::UrlEncode($pass)
    return $pass
}

function testCredentials($username, $userpass) {
    $userpass = URLEncode $userpass
    $loginUrl = "https://hr-services.fr.adp.com/ipclogin/1/loginform.fcc"
    $TARGET = "-SM-https%3A%2F%2Fpointage.adp.com%2Figested%2F2_02_01%2Fpointage"
    $formFields = "TARGET=" + $TARGET + "&USER=" + $username + "&PASSWORD=" + $userpass
    $AuthRequest = Invoke-WebRequest -Uri $loginUrl -Method Post -Body $formFields -ContentType "application/x-www-form-urlencoded" -SessionVariable websession
    $cookies = $websession.Cookies.GetCookies($loginUrl) 

    if ($cookies)
    {
        foreach ($cookie in $cookies)
        {
            if (($cookie.name -eq "SMSESSION") -AND ($cookie.value))
            {
                return $true
                break
            }
        }
        return $false
    }
    else
    {
        return $false
    }
}

Write-Host (testCredentials "hblanc-n63" "Ys8QZvXS4uG96H6")