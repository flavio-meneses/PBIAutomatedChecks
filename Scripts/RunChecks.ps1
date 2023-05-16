# =================================================================================
# PBI Automated Checks
# This script orchestrates all the checks to run on the .pbix file

# Flavio Meneses
# https://uk.linkedin.com/in/flaviomeneses
# https://github.com/flavio-meneses
# ===================================================================================

[CmdletBinding()] 
param 
( 
    [Parameter(Mandatory = $true)]         
    [string] $Server, 
    [Parameter(Mandatory = $true)]         
    [string] $Database  
) 

$portNumber = $Server.Split(":")[1]
        
Write-Host "---" 
Write-Host "Analysis Services Port Number: " -NoNewline 
Write-Host "$portNumber" -ForegroundColor Yellow 
Write-Host "Database: " -NoNewline 
Write-Host "$Database" -ForegroundColor Green 
Write-Host "---" 

#Set Location to current script's
Set-Location -Path $PSScriptRoot
Write-host "Current Location: $(Get-Location)"

#Imports Auxiliary Functions and Settings
Import-Module -Name .\AuxFunctions.psm1 -Force
if (Test-Path ..\LocalSettings.json) { $settings = Get-Content -Path ..\LocalSettings.json -Raw -Force | ConvertFrom-Json }
else { $settings = Get-Content -Path ..\Settings.json -Raw -Force | ConvertFrom-Json }

try {
    #-----------------------------
    #region Extract .pbix into Json

    #set the .pbix export folder to a temporary location on the user's machine. This will be deleted at the end of the script
    $exportFolder = Join-Path $env:TEMP "AutomatedChecksTempFolder"
    
    # Get path of .pbix file that called external tool
    $pbiSessions = (& ..\PBITools\pbi-tools.exe info | ConvertFrom-Json).pbiSessions
    $activePbix = $pbiSessions | Where-Object { $_.Port -eq $portNumber }
    $PbiFilePath = $activePbix.PbixPath

    if ($PbiFilePath -eq $null) {
        Write-host "Can't find current .pbix file name."
        $PbiFilePath = Read-Host "Please enter the path of the file you want to check without quotation marks"
    }
    
    # Output the file path
    Write-Output "Power BI File path: $PbiFilePath"

    #Extract Power BI file into its Json components
    & ..\PBITools\pbi-tools.exe extract $PbiFilePath -extractFolder $exportFolder

    #endregion
    
    #initialize check results table
    $checkResultsTable = @()

    #-----------------------------
    #region File structure checks

    #import settings for current .pbix
    $countReportSettings = ($settings.reportSettings | Where-Object { $_.reportPath -eq $PbiFilePath }).Count
    if ($countReportSettings -eq 0) {
        Write-Host "No report tests found. Please check the report is present in the Settings file and the paths are correct" -ForegroundColor Red
        Read-Host -Prompt 'Press [Enter] to close this window'
        exit
    }
    if ($countReportSettings -gt 1) {
        Write-Host "Multiple tests found for the same report. Please ensure each report only has one entry under 'reportSettings' on the Settings file" -ForegroundColor Red
        Read-Host -Prompt 'Press [Enter] to close this window'
        exit
    }
    else {
        foreach ($reportSetting in $settings.reportSettings) {
            if ($reportSetting.reportPath -eq $PbiFilePath) {
                $workspaceId = $reportSetting.publishWorkspaceId
                $testsFolderPath = $reportSetting.testsFolderPath
                Import-Module -Name "$($testsFolderPath)\FileStructureTests.psm1" -Force
                $DAXTests = Get-Content -Path "$($testsFolderPath)\DAXTests.dax" -Raw -Force
                $ignoreTestsPath = "$($testsFolderPath)\ignoreTests.json"
                $runDAXtest = [bool]$reportSetting.runDAXtest
            }
        }
    }

    #Get all Sections (i.e. Pages) in report
    $pages = Get-ChildItem -Path "$exportFolder\Report\Sections" -Directory -Depth 0

    #Check report landing page is first page and add to Check Results Table
    $checkResultsTable += CheckLandingPage -exportFolder $exportFolder -ignoreTestsPath $ignoreTestsPath

    $checkResultsTable += CheckReportFilters -exportFolder $exportFolder -ignoreTestsPath $ignoreTestsPath

    #loop through pages
    foreach ($page in $pages) {
    
        #Check page dimensions and add to Check Results Table
        $checkResultsTable += CheckPageDimensions -page $page -ignoreTestsPath $ignoreTestsPath

        #Check filters applied to each page
        $checkResultsTable += CheckPageFilters -page $page -ignoreTestsPath $ignoreTestsPath

        #Get all visuals in page
        $visuals = Get-ChildItem -Path "$($page.FullName)\visualContainers" -Directory -Depth 0

        #loop through visuals inside page
        foreach ($visual in $visuals) {
        
            #Check if visual is broken and add to Check Results Table
            #Currently not in use https://github.com/pbi-tools/pbi-tools/issues/261
            #$checkResultsTable += CheckBrokenVisual -visual $visual -page $page

            #Check if there's filters applied to visual and add to Check Results Table
            $checkResultsTable += CheckVisualFilters -visual $visual -page $page -ignoreTestsPath $ignoreTestsPath
        }
    }

    #endregion

    #region DAX Checks
    #To use the Power BI REST API, follow these steps

    #Requirements: This section requires a Power BI Pro account
    #1. Register an App on Azure Active Directory: https://app.powerbi.com/embedsetup 
    #1.1 Pick "Embed for your customers"
    #1.2 Sign in to Power BI
    #1.3 Give your app a name
    #1.4 Select access required - Select all for read/write access
    #1.5 Skip creation of workspaces/importing content
    #1.6 On Step 5, grant permissions to app, click "Accept"

    #Requirements: This section requires Global Administrator/Application Administrator role in Azure AD 
    #2 Create a secret for your app
    #2.1 head over to https://portal.azure.com/
    #2.2 search for "App Registrations" then "All applications" and select the App just created in step 1
    #2.3 Copy the "Application (client) ID" and "Directory (tenant) ID" and keep them somewhere safe, so you can refer to them later
    #2.5 click on "Certificates & secrets" and add a client secret, giving it a description so you can know where this secret will be used
    #2.6 copy the secret and keep it somewhere safe, so you can refer to them later
    #2.7 click on "Authentication" on the left menu and scroll to "Implicit grant and hybrid flows". Make sure both "Access tokens (used for implicit flows)" and "ID tokens (used from implicit and hybrid flows)" options are checked

    #3 Create security group, if one doens't exist already
    #3.1 head over to https://portal.azure.com/
    #3.2 search for "Groups" and click "New Group"
    #3.3 on "Group Type" select "Security" and add a group name, for example "PowerBI_REST_API".
    #3.4 on "Owners" add the group admins and on "Members" add the App created in step 1

    #Requirements: This section requires a Power BI Admin role 
    #4 Configure Power BI Tenant Settings
    #4.1 Head over to https://app.powerbi.com/admin-portal/tenantSettings
    #4.2 On "Tenant Settings" -> "Developer Settings" make sure "embed content in apps" is active
    #4.3 In the same section, make sure "Allow service principals to use Power BI APIs" is active. Give access to the Security Group created in step 3
    #4.3 On "Tenant Settings" -> "Admin API Settings" make sure "Allow service principals to use read-only admin APIs" is active. Give access to the Security Group created in step 3
    #4.4 On "Integration Settings", make sure "Dataset Execute Queries REST API" is enabled. Give access to the Security Group created in step 3

    #5 Assign Service Principal to Power BI Workspace
    #5.1 Head over to https://app.powerbi.com/
    #5.2 On the workspace you want to read/write, go to "Access"
    #5.3 Add the App created on Step 1 as an Admin

    if ($runDAXtest) {
        #Only runs this if report settings are set to run DAX Tests, which use the Power BI REST API
        
        #Import System Settings
        $tenantId = $settings.systemSettings.tenantId
        $appId = $settings.systemSettings.appId
        $appSecret = $settings.systemSettings.appSecret
    
        #Acquire access token to use API
        $accessToken = getAccessToken -tenantID $tenantId -appID $appId -appSecret $appSecret

        #Publish .pbix to Power BI Service so DAX queries can be done against dataset
        $reportName = "temp $([System.IO.Path]::GetFileNameWithoutExtension($PbiFilePath))"
        $publishPBIXFile = publishPBIX -pbixLocalPath $PbiFilePath -accessToken $accessToken -workspaceId $workspaceId -reportName $reportName -importMode "Ignore"
    
        #Run DAX Queries against published dataset
        $datasetId = $publishPBIXFile.DatasetId
        $checkResultsTable += executeQueries -datasetId $datasetId -DAXquery $DAXTests -accessToken $accessToken
    }
    #endregion

    Write-Host "" 
    Write-Host "--------------" 
    Write-Host "Checks completed" -ForegroundColor Green -BackgroundColor Black
    Read-Host -Prompt "Press [Enter] to see results" 
    
    # Set file path to current user's desktop and include timestamp
    $userDesktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    $fileExportPath = "$userDesktop\$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss') Check Results.csv"
    
    # Export to CSV and open
    $checkResultsTable | Export-Csv -Path $fileExportPath -NoTypeInformation
    Invoke-Item $fileExportPath
    exit
}

catch {
    Write-Host "" 
    Write-Host "--------------" -ForegroundColor Red
    Write-host "An error has occured:" -ForegroundColor Red
    Write-Error $_
    Read-Host -Prompt 'Press [Enter] to close this window'
    exit
}

# Execute at the end of the script, even if there's errors
finally {
    #Remove the temporary local folder, if it exists. 
    if (Test-Path -Path $exportFolder) {
        Remove-Item -Recurse -Force $exportFolder
    }
    
    if ($runDAXtest) {
        #Only runs this if report settings are set to run DAX Tests, which use the Power BI REST API
        
        #Remove published dataset
        deleteDataSets -workspaceId $workspaceId -datasetId $datasetId -accessToken $accessToken
    }
}