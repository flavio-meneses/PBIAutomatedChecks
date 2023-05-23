# =================================================================================
# PBI Automated Checks
# This script contains the file structure tests to run on the .pbix file

# Flavio Meneses
# https://uk.linkedin.com/in/flaviomeneses
# https://github.com/flavio-meneses
# ===================================================================================
function CheckLandingPage {
    param (
        [String] $exportFolder,
        [String] $ignoreTestsPath
    )

    try {

        Write-Host "Checking landing page:"
                
        #Load "config" file
        $visualConfig = Get-ChildItem "$exportFolder\Report\config.json" -File
        #convert to Json
        $visualConfigJson = Get-Content $visualConfig -Raw | ConvertFrom-Json
        #get landing page index
        $landingPageIndex = $visualConfigJson.activeSectionIndex

        #to verify if this specific check has been ignored by user
        $checkType = "File Structure"
        $checkName = "Report landing page is first page"

        #verify if check should be ignored. If not, proceed with check
        $ignoreCheck = IgnoreCheck -checkType $checkType -checkName $checkName -ignoreTestsPath $ignoreTestsPath
            
        if ($ignoreCheck) {
            $detail = $null
            $result = "Ignored"
        }

        elseif (-not $ignoreCheck) {

            if ($landingPageIndex -eq 0) {
                $detail = $null
                $result = $true
            }
            else {
                $detail = "Landing page isn't first page. Save report in the first page to ensure report users see it first"
                $result = $false
            }

        }
        #create table row
        $rowData = [PSCustomObject]@{
            "Check Type" = $checkType
            "Check Name" = $checkName
            "Detail"     = $detail
            "Pass?"      = $result
        }
        #return result as table row
        Write-Output $rowData
    }

    catch {
        Write-host "Error checking landing page:"
        Write-Error $_
    }
}
function CheckPageDimensions {
    param(
        [System.IO.DirectoryInfo] $page,
        [String] $ignoreTestsPath
    )
    try {
        #the "section" file captures the page name and dimensions
        $sectionFile = Get-ChildItem "$($page.FullName)\section.json" -File
        #convert to Json
        $sectionJson = Get-Content $sectionFile -Raw | ConvertFrom-Json

        #read attributes
        $pageName = $sectionJson.displayName
        $height = $sectionJson.height
        $width = $sectionJson.width

        Write-Host "Checking page '$pageName' dimensions"

        #to verify if this specific check has been ignored by user
        $checkType = "File Structure"
        $checkName = "Page dimensions for page '$pageName'"

        #verify if check should be ignored. If not, proceed with check
        $ignoreCheck = IgnoreCheck -checkType $checkType -checkName $checkName -ignoreTestsPath $ignoreTestsPath
            
        if ($ignoreCheck) {
            $detail = $null
            $result = "Ignored"
        }
        elseif (-not $ignoreCheck) {
            #test page dimensions
            if ($height -eq 720 -and $width -eq 1280) {
                $detail = $null
                $result = $true
                
            }
            else {
                $detail = "Expecting default page size of 1280x720 but page is $($width)x$($height)"
                $result = $false
            } 
        }
        #create table row
        $rowData = [PSCustomObject]@{
            "Check Type" = $checkType
            "Check Name" = $checkName
            "Detail"     = $detail
            "Pass?"      = $result
        }
        #return result as table row
        Write-Output $rowData 
    }
    catch {
        Write-host "Error checking page dimensions:"
        Write-Error $_
    }
}
function CheckReportFilters {    
    param (
        [String] $exportFolder,
        [String] $ignoreTestsPath
    )
    try {
        Write-Host "Checking report filters"
        
        $reportFiltersPath = "$($exportFolder)\Report\filters.json"

        if (Test-Path -Path $reportFiltersPath) {
            #Load "config" file
            $reportFilters = Get-ChildItem $reportFiltersPath -File
            #convert to Json
            $reportFiltersJson = Get-Content $reportFilters -Raw | ConvertFrom-Json
        }
        else {
            Write-Host "No report filters file found"
            return $null
        }

        #init results
        $resultDataRows = @()

        $checkType = "File Structure"
        $checkName = "No filters applied to report (filters on all pages)"

        $countActiveFilters = 0 
        foreach ($filter in $reportFiltersJson) {             
            if ($filter.filter) {                            
                $countActiveFilters++ 
            }
        }
        #no filters applied to report
        if ($countActiveFilters -eq 0) {
            # Verify if check should be ignored. No need to pass $detail as we already know there's no filters
            $ignoreCheck = IgnoreCheck -checkType $checkType -checkName $checkName -ignoreTestsPath $ignoreTestsPath

            if ($ignoreCheck) { $result = "Ignored" }
            else { $result = $true }

            $resultDataRows = [PSCustomObject]@{
                "Check Type" = $checkType
                "Check Name" = $checkName
                "Detail"     = $null
                "Pass?"      = $result
            }
        }
        else {
            #there's filters applied to report
            foreach ($filter in $reportFiltersJson) { 
                $filterType = $filter.type
                $filteredColumn = $filter.expression.Column.Property
                $filterIsActive = if ($filter.filter) { $true } else { $false }
                
                #detail message
                if ($filterIsActive) {
                    $detail = "$filterType filter applied on '$filteredColumn'"   
                }
                else {
                    $detail = "$filterType filter present but not applied on '$filteredColumn'"   
                }
                 
                # Verify if check should be ignored
                $ignoreCheck = IgnoreCheck -checkType $checkType -checkName $checkName -checkDetail $detail -ignoreTestsPath $ignoreTestsPath

                if ($ignoreCheck) { 
                    $resultDataRows += [PSCustomObject]@{
                        "Check Type" = $checkType
                        "Check Name" = $checkName
                        "Detail"     = $detail
                        "Pass?"      = "Ignored"
                    }
                }
                #if check isn't ignored
                else {
                    #Filter is Active
                    if ($filterIsActive) {
                        $resultDataRows += [PSCustomObject]@{
                            "Check Type" = $checkType
                            "Check Name" = $checkName
                            "Detail"     = $detail
                            "Pass?"      = $false
                        }
                    }
                    else {
                        #Filter is present but not active
                        $resultDataRows += [PSCustomObject]@{
                            "Check Type" = $checkType
                            "Check Name" = $checkName
                            "Detail"     = $detail
                            "Pass?"      = $true
                        }
                    }
                } 
            }
        }
        #return result as table row
        Write-Output $resultDataRows            
    }

    catch {
        Write-host "Error checking report filters:"
        Write-Error $_
    }
}
function CheckPageFilters {    
    param(
        [System.IO.DirectoryInfo] $page,
        [String] $ignoreTestsPath
    )
    try {
        Write-Host "Checking page filters"
        #The "filter" file captures the page filers
        $pageFilters = Get-ChildItem "$($page.FullName)\filters.json" -File
        #convert to Json
        $pageFiltersJson = Get-Content $pageFilters -Raw | ConvertFrom-Json
        $pageName = $($page.Name).Substring($($page.Name).IndexOf("_") + 1)

        #init results
        $resultDataRows = @()

        $checkType = "File Structure"
        $checkName = "No filters applied on this page ($($pageName))"

        $countActiveFilters = 0 
        foreach ($filter in $reportFiltersJson) {             
            if ($filter.filter) {                            
                $countActiveFilters++ 
            }
        }

        #no filters applied to page
        if ($countActiveFilters -eq 0) {
            # Verify if check should be ignored. No need to pass $detail as we already know there's no filters
            $ignoreCheck = IgnoreCheck -checkType $checkType -checkName $checkName -ignoreTestsPath $ignoreTestsPath

            if ($ignoreCheck) { $result = "Ignored" }
            else { $result = $true }

            $resultDataRows = [PSCustomObject]@{
                "Check Type" = $checkType
                "Check Name" = $checkName
                "Detail"     = $null
                "Pass?"      = $result
            }
        }
        else {
            #there's filters applied to page
            foreach ($filter in $pageFiltersJson) { 
                $filterType = $filter.type
                $filteredColumn = $filter.expression.Column.Property
                $filterIsActive = if ($filter.filter) { $true } else { $false }
                
                #detail message
                if ($filterIsActive) {
                    $detail = "$filterType filter applied on '$filteredColumn'"   
                }
                else {
                    $detail = "$filterType filter present but not applied on '$filteredColumn'"   
                }
                 
                # Verify if check should be ignored
                $ignoreCheck = IgnoreCheck -checkType $checkType -checkName $checkName -checkDetail $detail -ignoreTestsPath $ignoreTestsPath

                if ($ignoreCheck) { 
                    $resultDataRows += [PSCustomObject]@{
                        "Check Type" = $checkType
                        "Check Name" = $checkName
                        "Detail"     = $detail
                        "Pass?"      = "Ignored"
                    }
                }
                #if check isn't ignored
                else {
                    #Filter is Active
                    if ($filterIsActive) {
                        $resultDataRows += [PSCustomObject]@{
                            "Check Type" = $checkType
                            "Check Name" = $checkName
                            "Detail"     = $detail
                            "Pass?"      = $false
                        }
                    }
                    else {
                        #Filter is present but not active
                        $resultDataRows += [PSCustomObject]@{
                            "Check Type" = $checkType
                            "Check Name" = $checkName
                            "Detail"     = $detail
                            "Pass?"      = $true
                        }
                    }
                } 
            }
        }
        #return result as table row
        Write-Output $resultDataRows            
    }

    catch {
        Write-host "Error checking page filters:"
        Write-Error $_
    }
}
function CheckVisualFilters {
    param (
        [System.IO.DirectoryInfo] $page,
        [System.IO.DirectoryInfo] $visual,
        [String] $ignoreTestsPath
    )

    try {
        #Load "config" file
        $visualConfig = Get-ChildItem "$($visual.FullName)\config.json" -File
        #convert to Json
        $visualConfigJson = Get-Content $visualConfig -Raw | ConvertFrom-Json
        #get visual type & group
        $visualType = $visualConfigJson.singleVisual.visualType
        $visualGroup = $visualConfigJson.singleVisualGroup.displayName

        #only check for filters if visual isn't one of the types below
        if ((("shape", "textbox", "actionButton") -notcontains $visualType) -and (-not $visualGroup)) {

            #get page & visual names user friendly
            $pageName = $($page.Name).Substring($($page.Name).IndexOf("_") + 1)
            $visualName = $($visual.Name).Substring($($visual.Name).IndexOf("_") + 1)

            Write-Host "Checking filters applied to visual '$visualName' on page '$pageName'"
            
            #init results
            $resultDataRows = @()

            #Get filters file and convert to Json
            $visualFilters = Get-ChildItem "$($visual.FullName)\filters.json" -File -ErrorAction Stop #throw error, so this can be handled by catch block
            $visualFiltersJson = Get-Content $visualFilters -Raw | ConvertFrom-Json
            $checkType = "File Structure"
            $checkName = "No filters applied to visual '$visualName' on page '$pageName'"

            $countActiveFilters = 0 
            foreach ($filter in $visualFiltersJson) {             
                if ($filter.filter) {                            
                    $countActiveFilters++ 
                }
            }

            #no filters applied to visual
            if ($countActiveFilters -eq 0) {
                # Verify if check should be ignored. No need to pass $detail as we already know there's no filters
                $ignoreCheck = IgnoreCheck -checkType $checkType -checkName $checkName -ignoreTestsPath $ignoreTestsPath

                if ($ignoreCheck) { $result = "Ignored" }
                else { $result = $true }

                $resultDataRows = [PSCustomObject]@{
                    "Check Type" = $checkType
                    "Check Name" = $checkName
                    "Detail"     = $null
                    "Pass?"      = $result
                }
            }
            else {
                #filters applied to visual
                foreach ($filter in $visualFiltersJson) { 
                    $filterType = $filter.type

                    # Checks if filter is applied to a Column, Measure or Aggregation
                    $filterAppliedTo = if ($filter.expression.Column.Property) {
                        $filter.expression.Column.Property
                    }
                    elseif ($filter.expression.Measure.Property) {
                        $filter.expression.Measure.Property 
                    }
                    elseif ($filter.expression.Aggregation.Expression.Column.Property) {
                        $filter.expression.Aggregation.Expression.Column.Property 
                    }
                    else { $null }
                    $filterIsActive = if ($filter.filter) { $true } else { $false }

                    #detail message
                    if ($filterIsActive) {
                        $detail = "$filterType filter applied on '$filterAppliedTo'"   
                    }
                    else {
                        $detail = "$filterType filter present but not applied on '$filterAppliedTo'"   
                    }
                
                    # Verify if check should be ignored
                    $ignoreCheck = IgnoreCheck -checkType $checkType -checkName $checkName -checkDetail $detail -ignoreTestsPath $ignoreTestsPath

                    if ($ignoreCheck) { 
                        $resultDataRows += [PSCustomObject]@{
                            "Check Type" = $checkType
                            "Check Name" = $checkName
                            "Detail"     = $detail
                            "Pass?"      = "Ignored"
                        }
                    }
                    #if check isn't ignored
                    else {
                        #Filter is Active
                        if ($filterIsActive) {
                            $resultDataRows += [PSCustomObject]@{
                                "Check Type" = $checkType
                                "Check Name" = $checkName
                                "Detail"     = $detail
                                "Pass?"      = $false
                            }
                        }
                        else {
                            #Filter is present but not active
                            $resultDataRows += [PSCustomObject]@{
                                "Check Type" = $checkType
                                "Check Name" = $checkName
                                "Detail"     = $detail
                                "Pass?"      = $true
                            }
                        }
                    } 
                }
            }
            #return result as table row
            Write-Output $resultDataRows            
        }
    }

    catch {
        Write-host "Error checking visual filters:"
        Write-Error $_
    }
}
#currently not in use https://github.com/pbi-tools/pbi-tools/issues/261
function CheckBrokenVisual {
    param(
        [System.IO.DirectoryInfo] $page,
        [System.IO.DirectoryInfo] $visual
    )

    #Load "config" file
    $visualConfig = Get-ChildItem "$($visual.FullName)\config.json" -File
    #convert to Json
    $visualConfigJson = Get-Content $visualConfig -Raw | ConvertFrom-Json
    #get visual type & group
    $visualType = $visualConfigJson.singleVisual.visualType
    $visualGroup = $visualConfigJson.singleVisualGroup.displayName

    try {

        #only check if visual isn't one of the types below
        if ((("shape", "textbox", "actionButton") -notcontains $visualType) -and (-not $visualGroup)) {

            #look for the "dataTransforms" file. If it doesn't exist the visual is broken
            $visualExists = Test-Path "$($visual.FullName)\dataTransforms.json"

            if ($visualExists) {
                $result = $true
            }
            else {
                $result = $false
            }

            #create table row
            $rowData = [PSCustomObject]@{
                "Check name" = "Broken visual checks for page '$($page.Name)', $($visual.Name)"
                "Pass?"      = $result
            }

            #return result as table row
            Write-Output $rowData
        }

    }
    catch {
        Write-host "Error checking broken visuals:"
        Write-Error $_
    }
}
function IgnoreCheck {
    param (
        [String] $checkType,
        [String] $checkName,
        [String] $checkDetail = $null, #optional
        [String] $ignoreTestsPath
    )

    try {
        $ignoreChecks = Get-Content -Path $ignoreTestsPath -Raw -Force | ConvertFrom-Json

        $ignoreChecksarray = @()

        #check input against list to ignore. If there's a match, check should be ignored
        foreach ($ignore in $ignoreChecks) { 
            if ($ignore."Check Type" -eq $checkType -and $ignore."Check Name" -eq $checkName -and ($ignore."Detail" -eq $checkDetail -or [string]::IsNullOrEmpty($checkDetail))) {
                $ignoreChecksarray += $true
            }
            else {
                $ignoreChecksarray += $false
            }
        }
        
        #tests all results in array. If there's any true (i.e. ignore check), this returns false (i.e. not all results are false) 
        $allFalse = [bool]($ignoreChecksarray | ForEach-Object { $_ -eq $false } | Measure-Object -Minimum).Minimum

        return !$allFalse

    }
    catch {
        Write-host "Error verifying if check should be ignored:"
        Write-Error $_
    }
}