function Get-M365AIApplicationUsageReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateRange(1, 180)]
        [int]$Days = 30,

        [Parameter()]
        [ValidateRange(100, 50000)]
        [int]$ResultSize = 5000,

        [Parameter()]
        [string]$ExportPath
    )

    Assert-ExchangeOnlineConnected

    function ConvertTo-LowerSafe {
        param(
            [Parameter()]
            [AllowNull()]
            [object]$Value
        )

        if ($null -eq $Value) {
            return ''
        }

        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) {
            return ''
        }

        return $text.ToLowerInvariant()
    }

    function Ensure-StringArray {
        param([object]$Item)
        if ($null -eq $Item) { return @() }
        if ($Item -is [string]) { return @($Item) }
        return @($Item)
    }

    function Test-AnyMatch {
        param(
            [Parameter()]
            [object]$Values,

            [Parameter()]
            [object]$Terms
        )

        $valueList = @(
            foreach ($v in @($Values)) {
                if ($null -eq $v) { continue }
                [string]$s = [string]$v
                if (-not [string]::IsNullOrWhiteSpace($s)) {
                    $s
                }
            }
        )

        $termList = @(
            foreach ($t in @($Terms)) {
                if ($null -eq $t) { continue }
                [string]$s = [string]$t
                if (-not [string]::IsNullOrWhiteSpace($s)) {
                    $s
                }
            }
        )

        if ($valueList.Count -lt 1 -or $termList.Count -lt 1) {
            return $false
        }

        foreach ($value in $valueList) {
            $candidate = ConvertTo-LowerSafe -Value $value
            if ([string]::IsNullOrWhiteSpace($candidate)) {
                continue
            }

            foreach ($term in $termList) {
                $needle = ConvertTo-LowerSafe -Value $term
                if (-not [string]::IsNullOrWhiteSpace($needle) -and $candidate.Contains($needle)) {
                    return $true
                }
            }
        }

        return $false
    }

    function Get-GraphSubset {
        param(
            [Parameter(Mandatory)]
            [string]$Uri,

            [Parameter(Mandatory)]
            [int]$MaxItems
        )

        $results = [System.Collections.Generic.List[object]]::new()
        $nextUri = $Uri

        while ($nextUri -and $results.Count -lt $MaxItems) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
            foreach ($item in @($response.value)) {
                if ($results.Count -ge $MaxItems) {
                    break
                }

                $results.Add($item)
            }

            $nextUri = if ($results.Count -ge $MaxItems) { $null } else { $response.'@odata.nextLink' }
        }

        return @($results)
    }

    $catalog = @(
        [pscustomobject]@{
            AIApplication   = 'Microsoft Copilot'
            MatchTerms      = @('copilot', 'github copilot', 'microsoft copilot', 'copilot studio')
            MatchDomains    = @('copilot.microsoft.com', 'copilotstudio.microsoft.com', 'github.com/features/copilot')
            RiskObservation = 'AI assistant activity detected. Validate approved use cases, data handling boundaries, plugin access, and DLP / sensitivity label enforcement.'
        }
        [pscustomobject]@{
            AIApplication   = 'ChatGPT / OpenAI'
            MatchTerms      = @('chatgpt', 'openai')
            MatchDomains    = @('chat.openai.com', 'openai.com')
            RiskObservation = 'External generative AI activity detected. Review tenant approval, OAuth consent, uploaded content, prompt data handling, and egress controls.'
        }
        [pscustomobject]@{
            AIApplication   = 'Claude / Anthropic'
            MatchTerms      = @('claude', 'anthropic')
            MatchDomains    = @('claude.ai', 'anthropic.com')
            RiskObservation = 'External generative AI activity detected. Confirm whether business data is being shared outside approved Microsoft 365 controls.'
        }
        [pscustomobject]@{
            AIApplication   = 'Gemini / Google AI'
            MatchTerms      = @('gemini', 'bard', 'google ai studio')
            MatchDomains    = @('gemini.google.com', 'bard.google.com', 'aistudio.google.com')
            RiskObservation = 'External Google AI activity detected. Validate enterprise approval, retention, and whether SSO / data access has been sanctioned.'
        }
        [pscustomobject]@{
            AIApplication   = 'Grok / xAI'
            MatchTerms      = @('grok', 'xai', 'x.ai')
            MatchDomains    = @('grok.com', 'x.ai')
            RiskObservation = 'External AI activity detected. Review whether access is sanctioned and whether any browser or OAuth-based data transfer is occurring.'
        }
        [pscustomobject]@{
            AIApplication   = 'Perplexity'
            MatchTerms      = @('perplexity')
            MatchDomains    = @('perplexity.ai')
            RiskObservation = 'AI search usage detected. Confirm whether prompts include internal data and whether browser/OAuth usage is approved.'
        }
        [pscustomobject]@{
            AIApplication   = 'Poe / Quora AI'
            MatchTerms      = @('poe', 'quora ai')
            MatchDomains    = @('poe.com')
            RiskObservation = 'Multi-model AI aggregator usage detected. Review whether users can pivot between unapproved third-party models.'
        }
        [pscustomobject]@{
            AIApplication   = 'Mistral / Le Chat'
            MatchTerms      = @('mistral', 'le chat')
            MatchDomains    = @('chat.mistral.ai', 'mistral.ai')
            RiskObservation = 'External LLM usage detected. Validate whether enterprise controls and approved data boundaries exist.'
        }
        [pscustomobject]@{
            AIApplication   = 'Meta AI'
            MatchTerms      = @('meta ai', 'llama')
            MatchDomains    = @('meta.ai')
            RiskObservation = 'Consumer AI assistant usage detected. Review whether prompts or linked services expose business data.'
        }
        [pscustomobject]@{
            AIApplication   = 'Character.AI'
            MatchTerms      = @('character.ai', 'character ai')
            MatchDomains    = @('character.ai')
            RiskObservation = 'Consumer AI chat usage detected. This is typically outside sanctioned enterprise AI controls.'
        }
        [pscustomobject]@{
            AIApplication   = 'DeepSeek (China)'
            MatchTerms      = @('deepseek')
            MatchDomains    = @('deepseek.com', 'chat.deepseek.com')
            RiskObservation = 'Chinese-origin AI usage detected. Review regulatory, data residency, vendor risk, and prompt data exposure considerations immediately.'
        }
        [pscustomobject]@{
            AIApplication   = 'Qwen / Tongyi (Alibaba, China)'
            MatchTerms      = @('qwen', 'tongyi', 'tongyi qianwen', 'alibaba cloud ai')
            MatchDomains    = @('tongyi.aliyun.com', 'qianwen.aliyun.com')
            RiskObservation = 'Chinese-origin AI usage detected. Assess vendor risk, approved usage status, and whether enterprise data is being shared externally.'
        }
        [pscustomobject]@{
            AIApplication   = 'ERNIE Bot / Baidu Wenxin (China)'
            MatchTerms      = @('ernie', 'wenxin', 'baidu ai')
            MatchDomains    = @('yiyan.baidu.com', 'wenxin.baidu.com')
            RiskObservation = 'Chinese-origin AI usage detected. Confirm whether access is approved and whether data handling aligns with regulatory requirements.'
        }
        [pscustomobject]@{
            AIApplication   = 'Kimi / Moonshot AI (China)'
            MatchTerms      = @('kimi', 'moonshot ai')
            MatchDomains    = @('kimi.moonshot.cn', 'moonshot.cn')
            RiskObservation = 'Chinese-origin AI usage detected. Review approval status, data handling, and external service usage boundaries.'
        }
        [pscustomobject]@{
            AIApplication   = 'Doubao / ByteDance (China)'
            MatchTerms      = @('doubao', 'bytedance ai')
            MatchDomains    = @('doubao.com', 'doubao.bytedance.com')
            RiskObservation = 'Chinese-origin AI usage detected. Validate vendor risk, approved business justification, and data leakage controls.'
        }
        [pscustomobject]@{
            AIApplication   = 'Tencent Yuanbao / Hunyuan (China)'
            MatchTerms      = @('yuanbao', 'hunyuan', 'tencent ai')
            MatchDomains    = @('yuanbao.tencent.com', 'hunyuan.tencent.com')
            RiskObservation = 'Chinese-origin AI usage detected. Review enterprise approval, data sharing boundaries, and jurisdictional risk.'
        }
    )

    $startDateUtc = (Get-Date).ToUniversalTime().AddDays(-1 * [Math]::Abs($Days))
    $startDateText = $startDateUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')

    Write-Host "Collecting AI application usage signals for the last $Days day(s)..." -ForegroundColor Cyan

    $servicePrincipals = @()
    $signIns = @()
    $servicePrincipalStatus = 'Available'
    $signInStatus = 'Available'

    try {
        $servicePrincipals = @(Get-M365GraphCollection -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals?$top=999&$select=id,appId,displayName,publisherName,homepage,servicePrincipalType')
    }
    catch {
        $servicePrincipalStatus = if ($_.Exception.Message -match 'Forbidden|insufficient privileges') {
            'Unavailable (Directory.Read.All required for service principal inventory)'
        }
        else {
            "Service principal lookup failed: $($_.Exception.Message)"
        }
    }

    try {
        $uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$top=250&`$filter=createdDateTime%20ge%20$startDateText&`$select=createdDateTime,appDisplayName,resourceDisplayName,userDisplayName,userPrincipalName,ipAddress,clientAppUsed,isInteractive,appId"
        $signIns = @(Get-GraphSubset -Uri $uri -MaxItems $ResultSize)
    }
    catch {
        $signInStatus = if ($_.Exception.Message -match 'Forbidden|insufficient privileges|AuditLog') {
            'Unavailable (AuditLog.Read.All required for sign-in activity)'
        }
        else {
            "Sign-in lookup failed: $($_.Exception.Message)"
        }
    }

    $results = foreach ($app in $catalog) {
        $terms = @(
            @(Ensure-StringArray $app.MatchTerms) + @(Ensure-StringArray $app.MatchDomains) |
                ForEach-Object { [string]$_ } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique
        )

        $matchingEnterpriseApps = @(
            $servicePrincipals | Where-Object {
                Test-AnyMatch -Values @($_.displayName, $_.publisherName, $_.homepage) -Terms $terms
            }
        )

        $matchingSignIns = @(
            $signIns | Where-Object {
                Test-AnyMatch -Values @($_.appDisplayName, $_.resourceDisplayName) -Terms $terms
            }
        )

        $status = if ($matchingEnterpriseApps.Count -gt 0 -or $matchingSignIns.Count -gt 0) { 'Detected' } else { 'Not Detected' }
        $detectionSource = if ($matchingEnterpriseApps.Count -gt 0 -and $matchingSignIns.Count -gt 0) {
            'SignInLog + EnterpriseApp'
        }
        elseif ($matchingSignIns.Count -gt 0) {
            'SignInLog'
        }
        elseif ($matchingEnterpriseApps.Count -gt 0) {
            'EnterpriseApp'
        }
        else {
            'None'
        }

        $uniqueUsers = @(
            $matchingSignIns |
                ForEach-Object {
                    if (-not [string]::IsNullOrWhiteSpace([string]$_.userPrincipalName)) {
                        [string]$_.userPrincipalName
                    }
                    elseif (-not [string]::IsNullOrWhiteSpace([string]$_.userDisplayName)) {
                        [string]$_.userDisplayName
                    }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique
        )

        $matchedEnterpriseNames = @($matchingEnterpriseApps | ForEach-Object { [string]$_.displayName } | Where-Object { $_ } | Sort-Object -Unique)
        $matchedPublishers = @($matchingEnterpriseApps | ForEach-Object { [string]$_.publisherName } | Where-Object { $_ } | Sort-Object -Unique)
        $matchedSignInApps = @($matchingSignIns | ForEach-Object { if ($_.appDisplayName) { [string]$_.appDisplayName } else { [string]$_.resourceDisplayName } } | Where-Object { $_ } | Sort-Object -Unique)
        $clientApps = @($matchingSignIns | ForEach-Object { [string]$_.clientAppUsed } | Where-Object { $_ } | Sort-Object -Unique)
        $lastSeen = @($matchingSignIns | ForEach-Object { $_.createdDateTime } | Sort-Object -Descending | Select-Object -First 1)

        [pscustomobject]@{
            AIApplication            = $app.AIApplication
            Status                   = $status
            DetectionSource          = $detectionSource
            EnterpriseAppCount       = @($matchingEnterpriseApps).Count
            SignInCount              = @($matchingSignIns).Count
            UniqueUserCount          = @($uniqueUsers).Count
            LastSeenDateTime         = if ($lastSeen.Count -gt 0) { $lastSeen[0] } else { $null }
            MatchedEnterpriseApps    = $matchedEnterpriseNames -join '; '
            MatchedPublishers        = $matchedPublishers -join '; '
            MatchedSignInApps        = $matchedSignInApps -join '; '
            ExampleUsers             = (@($uniqueUsers | Select-Object -First 10)) -join '; '
            ClientAppTypes           = $clientApps -join '; '
            ServicePrincipalStatus   = $servicePrincipalStatus
            SignInStatus             = $signInStatus
            RiskObservation          = if ($status -eq 'Detected') { $app.RiskObservation } else { '' }
            CoverageNote             = 'Catalog-based detection of major AI platforms using Entra enterprise apps and Entra sign-in logs. Does not prove unmanaged browser use of public AI websites unless traffic is federated through Entra or separately monitored by Defender / proxy tooling.'
        }
    }

    $orderedResults = @($results | Sort-Object @{ Expression = { if ($_.Status -eq 'Detected') { 0 } else { 1 } } }, AIApplication)
    Export-M365ReportData -InputObject $orderedResults -ExportPath $ExportPath
}