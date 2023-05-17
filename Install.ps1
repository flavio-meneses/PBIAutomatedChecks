# =================================================================================
# PBI Automated Checks
# This script installs the necessary files in the user's machine
# It should be ran as an Admin as it needs to copy files to "C:\Program Files (x86)" 

# Flavio Meneses
# https://uk.linkedin.com/in/flaviomeneses
# https://github.com/flavio-meneses
# ===================================================================================

# Download and extract files
$repoSourceUrl = "https://github.com/flavio-meneses/PBIAutomatedChecks/archive/main.zip"
$repoDownloadTarget = Join-Path $HOME "Downloads\PBIAutomatedChecks"
try {
    Invoke-RestMethod -Uri $repoSourceUrl -OutFile "$repoDownloadTarget.zip"
    Expand-Archive -Path "$repoDownloadTarget.zip" -DestinationPath $repoDownloadTarget -Force
    Write-Host "Files downloaded and extracted successfully"
}
catch {
    Write-Host "Download failed. Error: $($_.Exception.Message)"
}

# Define paths
$currentLocalPath = "$($repoDownloadTarget)\PBIAutomatedChecks-main" #local path of this script
$externalToolTarget = "C:\Program Files (x86)\Common Files\Microsoft Shared\Power BI Desktop\External Tools"
$automatedChecksTarget = "C:\PowerBI_AutomatedChecks"
    
function runInstall {  
      
    #Copy Settings file
    Write-Host "Copying Settings.json"
            
    # Create the target folder if it doesn't already exist
    if (-not (Test-Path $automatedChecksTarget)) {
        New-Item -ItemType Directory -Path $automatedChecksTarget | Out-Null
    }

    # Delete any existing files, including sub-folders, before starting to copy new files     
    Remove-Item -Path "$($automatedChecksTarget)\*" -Recurse -Force

    #Overwrites file if there's changes

    Copy-Item "$($currentLocalPath)\Settings.json" -Destination $automatedChecksTarget -Force

    # Loop through each subfolder in the source root folder
    Get-ChildItem -Path $currentLocalPath -Directory | ForEach-Object {
        $sourceFolderPath = $_.FullName
        $sourceFolderName = Split-Path -Path $_ -Leaf

        #Copy External Tool Registration
        if ("External Tool Registration" -contains $sourceFolderName) {

            $targetFolderPath = $externalToolTarget

            # Create the target folder if it doesn't already exist
            if (-not (Test-Path $targetFolderPath)) {
                New-Item -ItemType Directory -Path $targetFolderPath | Out-Null
            }

            #Copy all files from the source folder to the target folder, including sub-folders
            #Overwrite file if it already exists
            Write-Host "Copying files to $($targetFolderPath)"
            Copy-Item -Path "$($sourceFolderPath)\*" -Destination $targetFolderPath -Recurse -Force
        }

        #Copy PBITools, Scripts - requires exact matches.
        #Copy any folders that contain 'tests' in name, in case multiple reports are present 
        if ($sourceFolderName -eq "PBITools" -or $sourceFolderName -eq "Scripts" -or $sourceFolderName -clike "*Tests*") {

            $targetFolderPath = "$($automatedChecksTarget)\$($sourceFolderName)"

            # Create the target folder if it doesn't already exist
            if (-not (Test-Path $targetFolderPath)) {
                New-Item -ItemType Directory -Path $targetFolderPath | Out-Null
            }

            #Copy all files from the source folder to the target folder, including sub-folders
            Write-Host "Copying files to $($targetFolderPath)"
            Copy-Item -Path "$($sourceFolderPath)\*" -Destination $targetFolderPath -Recurse
        }
    }
    Write-Host "" 
    Write-Host "--------------" 
    Write-Host "Installation complete" -ForegroundColor Green -BackgroundColor Black
    Write-Host "You can now close this window" 
}

try {
    #if a previous installation already exists inform user this will overwrite any existing Settings or Tests on the local machine  
    if (Test-Path $automatedChecksTarget) {
        Write-Host "This will install the Automated Checks external tool on your machine using default settings and tests" -ForegroundColor Yellow
        Write-Host "Any existing settings or tests will be overwritten and lost. You should back them up before proceeding" -ForegroundColor Yellow
        $userResponse = Read-Host "Do you want to continue? (Y/N)"
        
        if ($userResponse -eq "Y" -or $userResponse -eq "y") {
            runInstall
        }
        elseif ($userResponse -eq "N" -or $userResponse -eq "n") {
            exit
        }
        else {
            Read-Host -Prompt  "Invalid input. Only Y or N are accepted"
            exit
        }
    }
    #if there's no previous installation
    else {
        runInstall
    }
}
catch {
    Read-Host -Prompt $_
    exit
}

