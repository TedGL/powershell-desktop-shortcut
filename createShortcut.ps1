#Create a desktop shortcut
    #Set some variables to be used
    $linkLocation = "$Home\Desktop\Job_21002.lnk"
    $AppLocation = "C:\TeklaStructures\21.1\nt\bin\TeklaStructures.exe"

#Create new shortcut object
    $WshShell = New-Object -ComObject WScript.Shell

#Test if the shortcut already exists, if not then create it
    if (!(Test-Path $linkLocation) ) {
        $Shortcut = $WshShell.CreateShortcut($linkLocation)
        $Shortcut.TargetPath = $AppLocation
        $Shortcut.Arguments = '"ARGS"'
        $Shortcut.IconLocation = "<LOCATION FOR NEW SHORTCUT>"
        $Shortcut.WorkingDirectory ="Start In"
        $Shortcut.Description =""
        $Shortcut.Save()

        #Log a success message to the terminal
        Write-Host "**Shorcut Created**" -ForegroundColor green
    } 
    else {
        #Log a failed message to the terminal
        Write-Host "!!Shortcut Already Exists!!" -ForegroundColor Red
    }

#Prompt the user to hit enter to close the terminal window
    Read-Host -Prompt "Press Enter to exit"
