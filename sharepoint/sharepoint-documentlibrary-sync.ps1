#region Functions
function Sync-SharepointLocation {
    param (
        [string]$siteId,
        [string]$webId,
        [string]$listId,
        [mailaddress]$userEmail,
        [string]$webUrl,
        [string]$webTitle,
        [string]$listTitle,
        [string]$syncPath,
        [string]$copiedLibraryId
    )
    try {
        Add-Type -AssemblyName System.Web
        #Encode email address.
        [string]$userEmail = [System.Web.HttpUtility]::UrlEncode($userEmail)
        #Parse copied library ID for URI.
        [string]$siteId = [regex]::Match($copiedLibraryId, 'siteId=([^&]+)').Groups[1].Value
        [string]$webId  = [regex]::Match($copiedLibraryId, 'webId=([^&]+)').Groups[1].Value
        [string]$listId = [regex]::Match($copiedLibraryId, 'listId=([^&]+)').Groups[1].Value
        [string]$webUrl = [regex]::Match($copiedLibraryId, 'webUrl=([^&]+)').Groups[1].Value
        #build the URI
        $uri = New-Object System.UriBuilder
        $uri.Scheme = "odopen"
        $uri.Host = "sync"
        $uri.Query = "siteId=$siteId&webId=$webId&listId=$listId&userEmail=$userEmail&webUrl=$webUrl&listTitle=$listTitle&webTitle=$webTitle"
        #launch the process from URI
        Write-Host $uri.ToString()
        start-process -filepath $($uri.ToString())
    }
    catch {
        $errorMsg = $_.Exception.Message
    }
    if ($errorMsg) {
        Write-Warning "Sync failed."
        Write-Warning $errorMsg
    }
    else {
        Write-Host "Sync completed."
        while (!(Get-ChildItem -Path $syncPath -ErrorAction SilentlyContinue)) {
            Start-Sleep -Seconds 2
        }
        return $true
    }    
}
function Test-RegistryValueInSubkeys {
    param (
        [string]$BasePath,
        [string]$ValueName
    )
    try {
        # Get all subkeys under the base path
        $subkeys = Get-ChildItem -Path $BasePath -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }

        # Iterate through each subkey
        foreach ($subkey in $subkeys) {
            $value = Get-ItemProperty -Path $subkey.PSPath -Name $ValueName -ErrorAction SilentlyContinue
            if ($null -ne $value) {
                return $false
            }
        }

        # If no matching value found in any subkey
        return $true
    } catch {
        return $true
    }
}
#endregion
#region Main Process
try {
    #region Sharepoint Sync
    [mailaddress]$userUpn = cmd /c "whoami /upn"
    $params = @{
        #replace with data captured from your sharepoint site.
        copiedLibraryID = "Copied Library ID from clicking sync button."
        userEmail       = $userUpn
        webTitle        = "SharePoint Site Title"
        listTitle       = "Document Library Folder Name"
    }
    $baseRegistryPath   = "HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts\Business1\Tenants"
    $registryValueName  = "$(split-path $env:onedrive)\$($userUpn.Host)\$($params.webTitle) - $($Params.listTitle)"
    Write-Host "SharePoint params:"
    $params | Format-Table
    if (!(Test-RegistryValueInSubkeys -BasePath $baseRegistryPath -ValueName $registryValueName)) {
        Write-Host "Sharepoint folder not found locally, will now sync.." -ForegroundColor Yellow
        $sp = Sync-SharepointLocation @params
        if (!($sp)) {
            Throw "Sharepoint sync failed."
        }
    }
    else {
        Write-Host "Location already syncronized: $($params.syncPath)" -ForegroundColor Yellow
    }
    #endregion
}
catch {
    $errorMsg = $_.Exception.Message
}
finally {
    if ($errorMsg) {
        Write-Warning $errorMsg
        Throw $errorMsg
    }
    else {
        Write-Host "Completed successfully.."
    }
}
#endregion       