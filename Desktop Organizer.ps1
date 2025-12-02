$currentUsername = $Env:USERNAME
$currentUserDesktopPath = "$Env:USERPROFILE\Desktop"
$desktopItems = $null

Write-Output "Currently logged in user is: $currentUsername"

Write-Output "Changing working directory to $currentUserDesktopPath"

Set-Location $currentUserDesktopPath | Out-Null

$desktopItems = Get-ChildItem .

function CheckForAndCreateFolder($folderName) {
    if (-not (Test-Path -PathType Container -Path $folderName)){
        New-Item -Path $folderName -ItemType Directory | Out-Null
    }
}

CheckForAndCreateFolder("Documents")
CheckForAndCreateFolder("Presentations")
CheckForAndCreateFolder("Spreadsheets")
CheckForAndCreateFolder("PDFs")
CheckForAndCreateFolder("Pictures")
CheckForAndCreateFolder("Videos")
CheckForAndCreateFolder("Music")
CheckForAndCreateFolder("Shortcuts")

Write-Output "Looping through and organizing files on desktop:"
Write-Output ""

foreach ($item in $desktopItems){
    if($item.PSIsContainer -eq $false){
    $fileDestination = $null
    switch ($item.Extension){
        {$_ -eq ".docx" -or $_ -eq ".doc" -or $_ -eq ".txt"}{
            Move-Item $item -Destination .\Documents
            $fileDestination = "$currentUserDesktopPath\Documents"
        }
        {$_ -eq ".pptx" -or $_ -eq ".pptm"}{
            Move-Item $item -Destination .\Presentations
            $fileDestination = "$currentUserDesktopPath\Presentations"
        }
        {$_ -eq ".xlsx" -or $_ -eq ".xls"}{
            Move-Item $item -Destination .\Spreadsheets
            $fileDestination = "$currentUserDesktopPath\Spreadsheets"
        }
        {$_ -eq ".pdf" -or $_ -eq ".xps"}{
            Move-Item $item -Destination .\PDFs
            $fileDestination = "$currentUserDesktopPath\PDFs"
        }
        {$_ -eq ".png" -or $_ -eq ".jpg" -or $_ -eq ".jpeg" -or $_ -eq ".gif"}{
            Move-Item $item -Destination .\Pictures
            $fileDestination = "$currentUserDesktopPath\Pictures"
        }
        {$_ -eq ".webm" -or $_ -eq ".mkv" -or $_ -eq ".mov" -or $_ -eq ".mp4" }{
            Move-Item $item -Destination .\Videos
            $fileDestination = "$currentUserDesktopPath\Videos"
        }
        {$_ -eq ".mp3" -or $_ -eq ".wav"}{
            Move-Item $item -Destination .\Music
            $fileDestination = "$currentUserDesktopPath\Music"
        }
        {$_ -eq ".lnk" -or $_ -eq ".url"}{
            Move-Item $item -Destination .\Shortcuts
            $fileDestination = "$currentUserDesktopPath\Shortcuts"
        }
    }
    Write-Output "$item moved to $fileDestination\$item"
}
}

Write-Output ""
Write-Output "All files with supported formats have been organized!"
Pause

