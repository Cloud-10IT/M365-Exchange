function Get-M365GraphCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $results = @()
    $nextUri = $Uri

    while ($nextUri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
        if ($response.value) {
            $results += @($response.value)
        }

        $nextUri = $response.'@odata.nextLink'
    }

    return $results
}
