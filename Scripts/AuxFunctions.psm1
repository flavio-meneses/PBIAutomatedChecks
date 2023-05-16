# =================================================================================
# PBI Automated Checks
# This script contains auxiliary functions used to run the data validity tests on the .pbix file
# It uses the Power BI REST API  

# Flavio Meneses
# https://uk.linkedin.com/in/flaviomeneses
# https://github.com/flavio-meneses
# ===================================================================================
function getAccessToken {
    [CmdletBinding()]
    param (
        [string] $tenantId,
        [string] $appID,
        [string] $appSecret
    )
    try {
        Write-Host "Getting access token"

        # Construct the authentication endpoint URL
        $authUrl = "https://login.microsoftonline.com/$tenantId/oauth2/token"

        # Construct the authentication request body
        $body = @{
            grant_type    = "client_credentials"
            client_id     = $appId
            client_secret = $appSecret
            resource      = "https://analysis.windows.net/powerbi/api"
        }

        # Send the authentication request and capture the response
        $response = Invoke-RestMethod -Uri $authUrl -Method Post -Body $body

        # Extract the access token from the response
        $accessToken = $response.access_token
        
        Write-host "Access Token acquired successfully:"
        Write-Output $accessToken
    
    }
    catch {
        Write-host "Error getting Access Token:"
        Write-Error $_
    }
}

function publishPBIX {
    #based on https://gitlab.com/Lieben/assortedFunctions/blob/master/Import-PBIXToPowerBI.ps1
    [CmdletBinding()]
    Param(
        [string] $pbixLocalPath, #local path to .PBIX file
        [string] $accessToken,
        [string] $workspaceId,
        [string] $reportName,
        [string] $importMode  #valid values: https://docs.microsoft.com/en-us/rest/api/power-bi/imports/postimportingroup#importconflicthandlermode
    )
    try {

        Write-Host "Publishing report '$reportName' to Power BI Service to run DAX Tests"

        #region publish .pbix

        $uri = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/imports?datasetDisplayName=$reportName"
        
        $boundary = "---------------------------" + (Get-Date).Ticks.ToString("x")
        $boundarybytes = [System.Text.Encoding]::ASCII.GetBytes("`r`n--" + $boundary + "`r`n")
        $request = [System.Net.WebRequest]::Create($uri)
        $request.ContentType = "multipart/form-data; boundary=" + $boundary
        $request.Method = "POST"
        $request.KeepAlive = $true
        $request.Headers.Add("Authorization", "Bearer $accessToken")
        $rs = $request.GetRequestStream()
        $rs.Write($boundarybytes, 0, $boundarybytes.Length);
        $header = "Content-Disposition: form-data; filename=`"temp.pbix`"`r`nContent-Type: application / octet - stream`r`n`r`n"
        $headerbytes = [System.Text.Encoding]::UTF8.GetBytes($header)
        $rs.Write($headerbytes, 0, $headerbytes.Length);
        $fileContent = [System.IO.File]::ReadAllBytes($pbixLocalPath)
        $rs.Write($fileContent, 0, $fileContent.Length)
        $trailer = [System.Text.Encoding]::ASCII.GetBytes("`r`n--" + $boundary + "--`r`n");
        $rs.Write($trailer, 0, $trailer.Length);
        $rs.Flush()
        $rs.Close()
        $response = $request.GetResponse()
        $stream = $response.GetResponseStream()
        $streamReader = [System.IO.StreamReader]($stream)
        $content = $streamReader.ReadToEnd() | convertfrom-json
        $jobId = $content.id
        $streamReader.Close()
        $response.Close()

        #endregion

        #region check publish status
        $header2 = @{
            "Authorization" = "Bearer $accessToken"
        }

        while ($true) {
            $res = Invoke-RestMethod -Method GET -uri "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/imports/$jobId" -Headers $header2

            if ($res.ImportState -ne "Publishing") {
                Write-Host "" 
                Write-Host "--------------" 
                Write-Host "Report publishing status: $($res.importState)"
                $output = [PSCustomObject] @{
                    "PublishingStatus" = $($res.importState)
                    "DatasetId"        = $($res.datasets.id)
                }
                return $output
            }
            Start-Sleep -s 2
        }
        #endregion
    }
    catch {
        Write-host "Error Publishing .pbix:"
        Write-Error $_
    }
}

function deleteDataSets {
    param (
        [string] $workspaceId,
        [string] $datasetId,
        [string] $accessToken
    )

    try {

        $deleteUrl = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets/$datasetId"
        Invoke-RestMethod -Method DELETE -Uri $deleteUrl -Headers @{Authorization = "Bearer $accessToken" }
        Write-Output "Cleanup: temporary dataset $datasetId has been deleted."
    }

    catch {
        Write-host "Error deleting temp dataset:"
        Write-Error $_
    }
}
        
function executeQueries {
    param (
        [string] $datasetId,
        [string] $DAXquery,
        [string] $accessToken
    )
    try {
        #This requires the "Dataset Execute Queries REST API" setting enabled on "Tenant Settings"
        #Maximum of 120 requests per user per minute
        #Maximum of 100,000 rows or 1,000,000 values per query (whichever is hit first)

        Write-Host "Running DAX Tests"

        $uri = "https://api.powerbi.com/v1.0/myorg/datasets/$datasetId/executeQueries"

        $headers = @{
            "Content-Type"  = "application/json"
            "Authorization" = "Bearer $accessToken"
        }
        $body = @{
            queries            = @(
                @{
                    query = "$DAXquery"
                }
            )
            serializerSettings = @{
                includeNulls = $true
            }
        } | ConvertTo-Json

        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body 

        $results = $response.results
        
        $resultsTable = @()
    
        foreach ($table in $results.tables) {
            foreach ($row in $table.rows) {
                $checkName = $row.'[TestName]'
                $details = "Expected value: $($row.'[ExpectedValue]'), Actual value: $($row.'[ActualValue]')"
                $pass = $row.'[Passed]'
                $rowData = [PSCustomObject] @{
                    "Check Type" = "Data"
                    "Check Name" = $checkName
                    "Detail"     = $details
                    "Pass?"      = $pass
                }
                $resultsTable += $rowData
            }
        }

        Write-Output $resultsTable
        
    }
    catch {
        $errorResponse = $_.Exception.Response
        $reader = New-Object System.IO.StreamReader($errorResponse.GetResponseStream())
        $errorMessage = $reader.ReadToEnd() | ConvertFrom-Json
        Write-Host "" 
        Write-Host "--------------" 
        Write-host "Error calling 'Execute Queries API': $($_.Exception.Message)"
        Write-host $errorMessage.error."pbi.error".details[0].detail.value  
    }
}