<#Load required Assembliees
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")  
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 

#Draw forms and controls
$Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Create Desktop Shortcut"
    $Form.Font = "Arial"
    $Form.Size = New-Object System.Drawing.Size(800,600)
    $Form.FormBorderStyle = "FixedDialog"
    $Form.StartPosition = "CenterScreen"
    $Form.ControlBox = $true
    $Font = New-Object System.Drawing.Font("Arial Narrow",15,[System.Drawing.FontStyle]::Regular)
    $form.Font = $Font 

#Add a label to the form
$label_shortcutLocation = New-Object System.Windows.Forms.Label
    $label_shortcutLocation.Location = New-Object System.Drawing.Size(192, 120)
    $label_shortcutLocation.AutoSize = $true
    $label_shortcutLocation.TextAlign = "MiddleCenter"
    $label_shortcutLocation.Text = "Shortcut Location"
    $Form.Controls.Add($label_shortcutLocation)

#Add a label to the form
$label_appLocation = New-Object System.Windows.Forms.Label
    $label_appLocation.Location = New-Object System.Drawing.Size(192, 160)
    $label_appLocation.AutoSize = $true
    $label_appLocation.TextAlign = "MiddleCenter"
    $label_appLocation.Text = "Target"
    $Form.Controls.Add($label_appLocation)

#Add a label to the form
$label_shortcutLocation = New-Object System.Windows.Forms.Label
    $label_Args.Location = New-Object System.Drawing.Size(192, 200)
    $label_Args.AutoSize = $true
    $label_Args.TextAlign = "MiddleCenter"
    $label_Args.Text = "Arguments"
    $Form.Controls.Add($label_Args)

#Add a label to the form
$label_icon = New-Object System.Windows.Forms.Label
    $label_icon.Location = New-Object System.Drawing.Size(192, 240)
    $label_icon.AutoSize = $true
    $label_icon.TextAlign = "MiddleCenter"
    $label_icon.Text = "Icon"
    $Form.Controls.Add($label_icon)

#Add a label to the form
$label_description = New-Object System.Windows.Forms.Label
    $label_description.Location = New-Object System.Drawing.Size(192, 280)
    $label_description.AutoSize = $true
    $label_description.TextAlign = "MiddleCenter"
    $label_description.Text = "Description"
    $Form.Controls.Add($label_description)

#Add a button to the form not working
$button_create = New-Object System.Windows.Forms.Button
    $button_create.Location = New-Object System.Drawing.Size(10, 10)
    $button_create.AutoSize = $true
    $button_create.Text = "Create"
    $button_create.TextAlign = "MiddleCenter"
    $button_create.Add_Click({
        $button_create.Text = "Created"
    })
    $Form.Controls.Add($button_create)
#>   

#Create a desktop shortcut
    $linkLocation = "$Home\Desktop\Job_21002.lnk"
    $AppLocation = "C:\TeklaStructures\21.1\nt\bin\TeklaStructures.exe"

#Create new shortcut object
    $WshShell = New-Object -ComObject WScript.Shell

#Test if the shortcut already exists, if not then create it
    if (!(Test-Path $linkLocation) ) {
        $Shortcut = $WshShell.CreateShortcut($linkLocation)
        $Shortcut.TargetPath = $AppLocation
        $Shortcut.Arguments = '" -I "C:\mycustomfiles\imperial.ini"'
        $Shortcut.IconLocation = "C:\TeklaStructures\21.1\nt\TeklaStructures.ico"
        $Shortcut.WorkingDirectory ="Start In"
        $Shortcut.Description ="Start Tekla Structures Model"
        $Shortcut.Save()

        #Log a success message to the terminal
        Write-Host "**Shorcut Created**" -ForegroundColor green
    } 
    else {
        #Log a failed message to the terminal
        Write-Host "!!Shortcut Already Exists!!" -ForegroundColor Red
    }


<#
#Show Form
    #$Form.Add_Shown($label_shortcutLocation)
    [void] $Form.ShowDialog()
#>
#Prompt the user to hit enter to close the terminal window
    Read-Host -Prompt "Press Enter to exit"
