function Get-M365GraphCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter()]
        [ValidateRange(1, 10)]
        [int]$MaxRetryCount = 5,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]$DefaultRetryDelaySeconds = 3
    )

    function Get-M365GraphStatusCode {
        param(
            [Parameter(Mandatory)]
            [System.Management.Automation.ErrorRecord]$ErrorRecord
        )

        try {
            if ($ErrorRecord.Exception -and $ErrorRecord.Exception.ResponseStatusCode) {
                return [int]$ErrorRecord.Exception.ResponseStatusCode
            }
        }
        catch {
        }

        try {
            if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.StatusCode) {
                return [int]$ErrorRecord.Exception.Response.StatusCode
            }
        }
        catch {
        }

        $message = [string]$ErrorRecord.Exception.Message
        if ($message -match '\b429\b') {
            return 429
        }
        if ($message -match '\b503\b') {
            return 503
        }

        return $null
    }

    function Get-M365GraphRetryDelaySeconds {
        param(
            [Parameter(Mandatory)]
            [System.Management.Automation.ErrorRecord]$ErrorRecord,

            [Parameter(Mandatory)]
            [int]$FallbackDelaySeconds
        )

        try {
            $exception = $ErrorRecord.Exception
            if ($exception -and $exception.Response -and $exception.Response.Headers) {
                $retryAfter = $exception.Response.Headers['Retry-After']
                if ($retryAfter) {
                    $retryAfterValue = [string]($retryAfter | Select-Object -First 1)
                    $parsedSeconds = 0
                    if ([int]::TryParse($retryAfterValue, [ref]$parsedSeconds)) {
                        return [Math]::Max($parsedSeconds, 1)
                    }
                }
            }
        }
        catch {
        }

        return $FallbackDelaySeconds
    }

    $results = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = $null
        $attempt = 0

        while ($true) {
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
                break
            }
            catch {
                $attempt++
                $statusCode = Get-M365GraphStatusCode -ErrorRecord $_
                $isRetryable = ($statusCode -in @(429, 503))

                if (-not $isRetryable -or $attempt -ge $MaxRetryCount) {
                    throw
                }

                $delaySeconds = Get-M365GraphRetryDelaySeconds -ErrorRecord $_ -FallbackDelaySeconds ($DefaultRetryDelaySeconds * $attempt)
                Start-Sleep -Seconds $delaySeconds
            }
        }

        if ($null -eq $response) {
            break
        }

        if ($response.PSObject.Properties.Name -contains 'value' -and $null -ne $response.value) {
            foreach ($item in @($response.value)) {
                [void]$results.Add($item)
            }
        }
        elseif ($response -is [System.Collections.IEnumerable] -and -not ($response -is [string])) {
            foreach ($item in $response) {
                [void]$results.Add($item)
            }
        }
        else {
            [void]$results.Add($response)
        }

        $nextUri = $null
        if ($response.PSObject.Properties.Name -contains '@odata.nextLink') {
            $nextUri = [string]$response.'@odata.nextLink'
        }
    }

    return @($results)
}