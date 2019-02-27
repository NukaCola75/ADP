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

# Current Path
$CurrentPath = Get-Location
$PathExecute = (Convert-Path $CurrentPath)

# Chargement des assemblies
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")


# Provide MSGBOX support
function msgbox
{
    param
    (
        [string]$Message,
        [string]$Title = 'Message box title',   
        [string]$buttons = 'OKCancel'
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
     
    # Display the message box
    $Return=[System.Windows.Forms.MessageBox]::Show($Message,$Title,$btn)
    $Return
}


# Creation de la form principale
$form = New-Object Windows.Forms.Form
# Pour bloquer le resize du form et supprimer les icones Minimize and Maximize
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.ControlBox = $true
$form.MaximizeBox = $false
$form.MinimizeBox = $false
# Choix du titre
$form.Text = "ADP Pointage Assist Configuration"
# Choix de la taille
$form.Size = New-Object System.Drawing.Size(300,400)

# IMG File
$file = (get-item $PathExecute'\IMG\CLS.png')
$img = [System.Drawing.Image]::Fromfile($file);

# IMG box
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width =  $img.Size.Width;
$pictureBox.Height =  $img.Size.Height;
$pictureBox.Location = New-Object System.Drawing.Point(90,30);
$pictureBox.Image = $img;

# Text
$label_text_nom = New-Object System.Windows.Forms.Label
$label_text_nom.Location = New-Object System.Drawing.Point(5,150)
$label_text_nom.Size = New-Object System.Drawing.Size(300,25)
$label_text_nom.Font = 'Segoe UI, 14pt'
$label_text_nom.Text = "Nom d'utilisateur ADP: "

# TextBox
$textbox_nom = New-Object System.Windows.Forms.TextBox
$textbox_nom.Location = New-Object System.Drawing.Point(8,180)
$textbox_nom.Size = New-Object System.Drawing.Size(264,20)
$textbox_nom.Font = 'Segoe UI, 14pt'

# Text
$label_text_pass = New-Object System.Windows.Forms.Label
$label_text_pass.Location = New-Object System.Drawing.Point(5,215)
$label_text_pass.Size = New-Object System.Drawing.Size(300,25)
$label_text_pass.Font = 'Segoe UI, 14pt'
$label_text_pass.Text = "Mot de passe ADP: "

# TextBox
$textbox_pass = New-Object System.Windows.Forms.TextBox
$textbox_pass.Location = New-Object System.Drawing.Point(8,245)
$textbox_pass.Size = New-Object System.Drawing.Size(264,20)
$textbox_pass.PasswordChar = '*'
$textbox_pass.Font = 'Segoe UI, 14pt'

# CheckBox
$checkBox_cadre = New-Object System.Windows.Forms.CheckBox
$checkBox_cadre.Location = New-Object System.Drawing.Point(8,285)
$checkBox_cadre.Size = New-Object System.Drawing.Size(250,20)
$checkBox_cadre.Text = "Je suis cadre"
$checkBox_cadre.Font = 'Segoe UI, 14pt'

# Validation button
$button_validate = New-Object System.Windows.Forms.Button
$button_validate.Text = "Valider"
$button_validate.Font = 'Segoe UI, 14pt, style=Bold'
$button_validate.Location = New-Object System.Drawing.Point(90,310)
$button_validate.Size = New-Object System.Drawing.Size(100, 40)

# Create Window
$form.Controls.Add($pictureBox)
$form.Controls.Add($label_text_nom)
$form.Controls.Add($textbox_nom)
$form.Controls.Add($label_text_pass)
$form.Controls.Add($textbox_pass)
$form.Controls.Add($checkBox_cadre)
$form.Controls.Add($button_validate)


$button_validate.Add_Click(
{
    $nom = $textbox_nom.Text
    $nom = $nom.Trim()
    $pass = $textbox_pass.Text
    $pass = $pass.Trim()

    if (!$nom -or !$pass)
    {
        $res = msgbox "L'identifiant et le mot de passe sont obligatoires." "Attention" ok
    }
    else 
    {
        # Config file creation
        New-Item -Name "config.txt" -Force
        Set-Content -Path ".\config.txt" -Value $null

        # Set username in config file
        Set-Content -Path ".\config.txt" -Value $nom
        
        # Config key creation
        New-Item -Path "HKCU:\Software\CLS\APP\ADP" -Force -ErrorAction 'SilentlyContinue'

        # Set value
        New-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_Magic_Word" -Value $pass -ErrorAction 'SilentlyContinue'

        if ($checkBox_cadre.Checked)
        {
            New-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_Cadre" -Value $true -Force -ErrorAction 'SilentlyContinue'
            # Configure scheduled task
            $triggers = @()
            $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "& `"$PathExecute\Pointage.ps1`""
            $triggers += New-ScheduledTaskTrigger -AtLogOn -User "PC-CLS\$env:USERNAME"
            $triggers += New-ScheduledTaskTrigger -Daily -At "10:00"
            $triggers += New-ScheduledTaskTrigger -Daily -At "14:00"
            $config = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
            Register-ScheduledTask -TaskName "ADP Pointage $env:USERNAME" -Trigger $triggers -Action $action -Settings $config -TaskPath "\" -Description "Hey ! Don't forget to point today !"
        }
        else
        {
            New-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_Cadre" -Value $false -Force -ErrorAction 'SilentlyContinue'
        }

        New-ItemProperty -Path "HKCU:\Software\CLS\APP\ADP" -Name "ADP_pointageDate" -Value "01/01/1970" -Force -ErrorAction 'SilentlyContinue'

        # Show success message
        $res = msgbox "Configuration terminée. Vous pouvez relancer cet outil si besoin." "Succès" ok

        # Close the window
        $form.Close()
    }
})

# Open the window
$form.ShowDialog()