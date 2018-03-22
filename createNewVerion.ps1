function sz($args) {
     -ArgumentList $args
}

Write-Host "     ===========================" -ForegroundColor Yellow
Write-Host "     ~~~~ TGH Build Creator ~~~~" -ForegroundColor Yellow
Write-Host "     ===========================" -ForegroundColor Yellow

$remoteScriptPath = "C:\Users\ThomasGH\Documents\Project\Playground\RemoteTool"
$localScriptPath = "C:\Users\ThomasGH\Documents\Project\Playground\MyTool Single Conf"
$logViewerPath = "C:\Users\ThomasGH\source\repos\TGH-Log-Viewer\TGH-Log-Viewer\bin\Debug"
$buildFolderPath = "C:\Users\ThomasGH\Documents\Project\Builds"
$executablePath = "C:\Users\ThomasGH\Documents\Project\Builds\Reference\OpenLogViewer.exe.lnk"
$guidePath = "C:\Users\ThomasGH\Google Drive\Belangrijk\Semester 6\Project\Documentatie\Guide V0.1.docx"

#Print the current detected version
$mostRecentZip = Get-ChildItem -Path $buildFolderPath -File "*.zip" | Sort-Object -Property CreationTime | Select-Object -Last 1 
Write-Host " > Most recent version is"$mostRecentZip.BaseName -ForegroundColor Red

#Ask version and check if the string is ok
$accepted = $true
while($accepted){
    $version = Read-Host " > Enter version number (e.g. 0.3.5) "
    if($version -match "\d.\d.\d"){ 
        Write-Host "   > Version ok!" -ForegroundColor Green
        $accepted = $false
    }
}

#Ask if there must be a git tag added, do if ok
$gittagged = Read-Host " > Does the current git commit needs to be tagged? (Y/n)"
if($gittagged -ne "n"){
    Write-Host "   > Tagging git repository..." -ForegroundColor Green
}

#Create a folder in de build folder with the version name
Write-Host " > Creating version folder..." -NoNewline
New-Item -ItemType Directory -Path ($buildFolderPath + "\V" + $version) -ErrorAction SilentlyContinue | Out-Null
Write-Host "DONE!" -ForegroundColor Green


#Copy over the content of the debug folder of the logviewer app
Write-Host " > Copying Log Viewer..." -NoNewline
New-Item -ItemType Directory -Path ($buildFolderPath + "\V" + $version + "\LogViewer") -ErrorAction SilentlyContinue | Out-Null
Copy-Item -Path ($logViewerPath + "\*") -Destination ($buildFolderPath + "\V" + $version + "\LogViewer") -Recurse
Write-Host "DONE!" -ForegroundColor Green

#Copy over de remote script (do not copy git files)
Write-Host " > Copying remote script..." -NoNewline
New-Item -ItemType Directory -Path ($buildFolderPath + "\V" + $version + "\RemoteLogStash") -ErrorAction SilentlyContinue | Out-Null
Copy-Item -Path ($remoteScriptPath + "\*") -Destination ($buildFolderPath + "\V" + $version + "\RemoteLogStash") -Recurse
Write-Host "DONE!" -ForegroundColor Green

#Copy over the local script (do not copy git files)
#Write-Host " > Copying local script..." -NoNewline
#New-Item -ItemType Directory -Path ($buildFolderPath + "\V" + $version + "\LocalLogStash") -ErrorAction SilentlyContinue | Out-Null
#Robocopy $localScriptPath ($buildFolderPath + "\V" + $version + "\LocalLogStash") /E | Out-Null
#Remove-Item -Path ($buildFolderPath + "\V" + $version + "\LocalLogStash\.*") -Recurse -Force
#Write-Host "DONE!" -ForegroundColor Green

#Open the word file and save as PDF
Write-Host " > Edit the release notes, save and close the file."
Invoke-Item $guidePath
Read-Host "Press enter to continue..."

#Create PDF and put it in the build folder
Write-Host " > Creating PDF..." -NoNewline
$wordApp = New-Object -ComObject Word.Application
$document =  $wordApp.Documents.Open($guidePath)
$filename = $buildFolderPath + "\V" + $version + "\" + $(Get-Item $guidePath).BaseName + ".pdf"
$document.SaveAs([ref] $filename, [ref] 17)
$document.Close();
$wordApp.Quit();
Write-Host "DONE!" -ForegroundColor Green

#Copy over the executable in reference
Write-Host " > Copying remote script..." -NoNewline
Copy-Item -Path $executablePath -Destination ($buildFolderPath + "\V" + $version)
Write-Host "DONE!" -ForegroundColor Green

#Zip all
Write-Host " > Zipping folder..." -NoNewline
cd "C:\Users\ThomasGH\Documents\Project\Builds"
Start-Process -FilePath "C:\Program Files\7-Zip\7z.exe" -ArgumentList "a V$version.zip V$version\"
Write-Host "DONE!" -ForegroundColor Green

#DONE 
Write-Host " ~~~~ FINISHED ~~~~ " -BackgroundColor Black -ForegroundColor Green
Read-Host


