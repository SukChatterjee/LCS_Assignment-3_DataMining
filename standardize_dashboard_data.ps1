param(
    [string]$DetailInputPath = ".\clean_coop_records_dashboard.csv",
    [string]$SummaryInputPath = ".\employer_summary_dashboard.csv",
    [string]$DetailOutputPath = ".\clean_coop_records_dashboard_standardized.csv",
    [string]$SummaryOutputPath = ".\employer_summary_dashboard_standardized.csv",
    [string]$AuditOutputPath = ".\company_standardization_audit.csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ExactKeyOverrides = @{
    "PNG" = "PROCTER AND GAMBLE"
    "P AND G" = "PROCTER AND GAMBLE"
    "PANDG" = "PROCTER AND GAMBLE"
    "PROCTOR AND GAMBLE" = "PROCTER AND GAMBLE"
    "PROCTER AND GAMBLE PANDG" = "PROCTER AND GAMBLE"
    "KPMG US" = "KPMG"
}

$DisplayNameOverrides = @{
    "PROCTER AND GAMBLE" = "Procter & Gamble (P&G)"
    "ERNST AND YOUNG" = "Ernst & Young"
    "BLUE AND" = "Blue & Co., LLC"
    "BATH AND BODY WORKS" = "Bath & Body Works"
    "BMW MANUFACTURING" = "BMW Manufacturing Co."
    "JP MORGAN" = "JP Morgan"
    "KOHLS" = "Kohl's"
    "KPMG" = "KPMG"
    "MACYS" = "Macy's"
    "NORTHRUP GRUMMAN" = "Northrop Grumman"
}

$TrailingSuffixTokens = @(
    "CO", "COMPANY", "CO", "CORP", "CORPORATION", "INC", "INCORPORATED",
    "LLC", "LLP", "LP", "LTD", "LIMITED", "PC", "PLC", "PLLC"
)

function Remove-Diacritics {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $normalized = $Text.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder

    foreach ($char in $normalized.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if ($category -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }

    return $builder.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function Normalize-CompanyText {
    param([string]$Text)

    $value = Remove-Diacritics $Text
    $value = $value.ToUpperInvariant()
    $value = $value -replace "\bP\s*&\s*G\b", "PROCTER AND GAMBLE"
    $value = $value -replace "\bP\s*AND\s*G\b", "PROCTER AND GAMBLE"
    $value = $value -replace "&", " AND "
    $value = $value -replace "@", " AT "
    $value = $value -replace "[^A-Z0-9 ]", " "
    $value = $value -replace "\s+", " "
    return $value.Trim()
}

function Get-CanonicalEmployerKey {
    param(
        [string]$EmployerKey,
        [string]$EmployerName
    )

    $source = if (-not [string]::IsNullOrWhiteSpace($EmployerKey)) { $EmployerKey } else { $EmployerName }
    $normalized = Normalize-CompanyText $source

    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return "UNKNOWN EMPLOYER"
    }

    if ($ExactKeyOverrides.ContainsKey($normalized)) {
        return $ExactKeyOverrides[$normalized]
    }

    $tokens = [System.Collections.Generic.List[string]]::new()
    foreach ($token in ($normalized -split " ")) {
        if (-not [string]::IsNullOrWhiteSpace($token)) {
            $tokens.Add($token)
        }
    }

    while ($tokens.Count -gt 0 -and $tokens[0] -eq "THE") {
        $tokens.RemoveAt(0)
    }

    while ($tokens.Count -gt 1 -and $TrailingSuffixTokens -contains $tokens[$tokens.Count - 1]) {
        $tokens.RemoveAt($tokens.Count - 1)
    }

    if ($tokens.Count -gt 1 -and ($tokens.Count % 2 -eq 0)) {
        $half = [int]($tokens.Count / 2)
        $left = ($tokens.GetRange(0, $half) -join " ")
        $right = ($tokens.GetRange($half, $half) -join " ")
        if ($left -eq $right) {
            $tokens = [System.Collections.Generic.List[string]]::new()
            foreach ($token in ($left -split " ")) {
                if (-not [string]::IsNullOrWhiteSpace($token)) {
                    $tokens.Add($token)
                }
            }
        }
    }

    $candidate = ($tokens -join " ").Trim()
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        $candidate = $normalized
    }

    if ($ExactKeyOverrides.ContainsKey($candidate)) {
        return $ExactKeyOverrides[$candidate]
    }

    return $candidate
}

function Test-MeaningfulText {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    $trimmed = $Value.Trim()
    return $trimmed -notin @("NA", "N/A", "NULL")
}

function Get-ModeValue {
    param(
        [object[]]$Values,
        [string]$Default = "NA"
    )

    $filtered = @(
        $Values |
        ForEach-Object { "$_" } |
        Where-Object { Test-MeaningfulText $_ }
    )

    if ($filtered.Count -eq 0) {
        return $Default
    }

    return (
        $filtered |
        Group-Object |
        Sort-Object @{ Expression = "Count"; Descending = $true }, @{ Expression = "Name"; Descending = $false } |
        Select-Object -First 1
    ).Name
}

function Get-RepresentativeEmployerName {
    param(
        [string]$CanonicalKey,
        [object[]]$Rows
    )

    if ($DisplayNameOverrides.ContainsKey($CanonicalKey)) {
        return $DisplayNameOverrides[$CanonicalKey]
    }

    $candidates = @(
        $Rows |
        ForEach-Object { $_.Employer_Clean } |
        Where-Object { Test-MeaningfulText $_ } |
        Group-Object |
        ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Count = $_.Count
                PunctuationPenalty = ([regex]::Matches($_.Name, "[,()]")).Count
                Length = $_.Name.Length
            }
        } |
        Sort-Object @{ Expression = "Count"; Descending = $true },
                    @{ Expression = "PunctuationPenalty"; Descending = $false },
                    @{ Expression = "Length"; Descending = $false },
                    @{ Expression = "Name"; Descending = $false }
    )

    if ($candidates.Count -gt 0) {
        return $candidates[0].Name
    }

    return $CanonicalKey
}

function ConvertTo-NullableDouble {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = "$Value".Trim()
    if (-not (Test-MeaningfulText $text)) {
        return $null
    }

    $parsed = 0.0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }

    if ([double]::TryParse($text, [ref]$parsed)) {
        return $parsed
    }

    return $null
}

if (-not (Test-Path -LiteralPath $DetailInputPath)) {
    throw "Detail dataset not found: $DetailInputPath"
}

if (-not (Test-Path -LiteralPath $SummaryInputPath)) {
    throw "Summary dataset not found: $SummaryInputPath"
}

$detailRows = Import-Csv -LiteralPath $DetailInputPath
$summaryRows = Import-Csv -LiteralPath $SummaryInputPath

$groupedDetail = @{}
foreach ($row in $detailRows) {
    $canonicalKey = Get-CanonicalEmployerKey -EmployerKey $row.Employer_Key -EmployerName $row.Employer_Clean
    if (-not $groupedDetail.ContainsKey($canonicalKey)) {
        $groupedDetail[$canonicalKey] = [System.Collections.Generic.List[object]]::new()
    }
    $groupedDetail[$canonicalKey].Add($row)
}

$canonicalMetadata = @{}
foreach ($canonicalKey in $groupedDetail.Keys) {
    $rows = $groupedDetail[$canonicalKey]
    $displayName = Get-RepresentativeEmployerName -CanonicalKey $canonicalKey -Rows $rows
    $alignedIndustry = Get-ModeValue -Values ($rows | ForEach-Object { $_."Employer Industry" }) -Default "NA"

    $canonicalMetadata[$canonicalKey] = [PSCustomObject]@{
        Employer_Key = $canonicalKey
        Employer_Display_Name = $displayName
        Employer_Industry = $alignedIndustry
    }
}

$standardizedDetail = foreach ($row in $detailRows) {
    $canonicalKey = Get-CanonicalEmployerKey -EmployerKey $row.Employer_Key -EmployerName $row.Employer_Clean
    $metadata = $canonicalMetadata[$canonicalKey]
    $output = [ordered]@{}

    foreach ($property in $row.PSObject.Properties.Name) {
        switch ($property) {
            "Employer_Clean" { $output[$property] = $metadata.Employer_Display_Name }
            "Employer Industry" { $output[$property] = $metadata.Employer_Industry }
            "Employer_Key" { $output[$property] = $metadata.Employer_Key }
            default { $output[$property] = $row.$property }
        }
    }

    $output["Employer_Clean_Original"] = $row.Employer_Clean
    $output["Employer_Industry_Original"] = $row."Employer Industry"
    $output["Employer_Key_Original"] = $row.Employer_Key
    $output["Employer_Display_Name"] = $metadata.Employer_Display_Name
    $output["Employer_Industry_Aligned"] = $metadata.Employer_Industry
    $output["Canonical_Employer_Key"] = $metadata.Employer_Key

    [PSCustomObject]$output
}

$industryAudit = foreach ($canonicalKey in ($groupedDetail.Keys | Sort-Object)) {
    $rows = $groupedDetail[$canonicalKey]
    $metadata = $canonicalMetadata[$canonicalKey]

    $originalNames = @(
        $rows |
        ForEach-Object { $_.Employer_Clean } |
        Where-Object { Test-MeaningfulText $_ } |
        Sort-Object -Unique
    )

    $originalKeys = @(
        $rows |
        ForEach-Object { $_.Employer_Key } |
        Where-Object { Test-MeaningfulText $_ } |
        Sort-Object -Unique
    )

    $originalIndustries = @(
        $rows |
        ForEach-Object { $_."Employer Industry" } |
        Where-Object { Test-MeaningfulText $_ } |
        Sort-Object -Unique
    )

    [PSCustomObject]@{
        Canonical_Employer_Key = $metadata.Employer_Key
        Employer_Display_Name = $metadata.Employer_Display_Name
        Canonical_Industry = $metadata.Employer_Industry
        Record_Count = $rows.Count
        Original_Name_Count = $originalNames.Count
        Original_Key_Count = $originalKeys.Count
        Original_Industry_Count = $originalIndustries.Count
        Original_Names = ($originalNames -join " | ")
        Original_Keys = ($originalKeys -join " | ")
        Original_Industries = ($originalIndustries -join " | ")
    }
}

$summaryCountsByKey = @{}
foreach ($row in $summaryRows) {
    $canonicalKey = Get-CanonicalEmployerKey -EmployerKey $row.Employer_Key -EmployerName $row.Employer_Key

    if (-not $summaryCountsByKey.ContainsKey($canonicalKey)) {
        $summaryCountsByKey[$canonicalKey] = [PSCustomObject]@{
            CPT_Record_Count = 0
            OPT_Record_Count = 0
        }
    }

    $summaryCountsByKey[$canonicalKey].CPT_Record_Count += [int](ConvertTo-NullableDouble $row.CPT_Record_Count)
    $summaryCountsByKey[$canonicalKey].OPT_Record_Count += [int](ConvertTo-NullableDouble $row.OPT_Record_Count)
}

# Rebuild the employer summary from the standardized detail rows so all rolled-up metrics use the same canonical employer.
$standardizedSummary = foreach ($canonicalKey in ($groupedDetail.Keys | Sort-Object)) {
    $rows = $groupedDetail[$canonicalKey]
    $metadata = $canonicalMetadata[$canonicalKey]

    $approvedCompletedRows = @(
        $rows |
        Where-Object { $_.Status_Clean -in @("Approved", "Completed") }
    )

    $wages = @(
        $approvedCompletedRows |
        ForEach-Object { ConvertTo-NullableDouble $_.Hourly_Wage_Imputed } |
        Where-Object { $null -ne $_ }
    )

    $years = @(
        $rows |
        ForEach-Object { ConvertTo-NullableDouble $_.Year } |
        Where-Object { $null -ne $_ } |
        Sort-Object -Unique
    )

    $seasons = @(
        $rows |
        ForEach-Object { $_.Season } |
        Where-Object { Test-MeaningfulText $_ } |
        Sort-Object -Unique
    )

    $majors = @(
        $rows |
        ForEach-Object { $_.Primary_Major } |
        Where-Object { Test-MeaningfulText $_ } |
        Sort-Object -Unique
    )

    $industries = @(
        $rows |
        ForEach-Object { $metadata.Employer_Industry } |
        Where-Object { Test-MeaningfulText $_ } |
        Sort-Object -Unique
    )

    $remoteValues = @(
        $rows |
        ForEach-Object { ConvertTo-NullableDouble $_.Is_Remote } |
        Where-Object { $null -ne $_ }
    )

    $avgHourlyWage = $null
    if ($wages.Count -gt 0) {
        $avgHourlyWage = ($wages | Measure-Object -Average).Average
    }

    $remoteRate = $null
    if ($remoteValues.Count -gt 0) {
        $remoteRate = ($remoteValues | Measure-Object -Average).Average
    }

    $cptCount = 0
    $optCount = 0
    if ($summaryCountsByKey.ContainsKey($canonicalKey)) {
        $cptCount = $summaryCountsByKey[$canonicalKey].CPT_Record_Count
        $optCount = $summaryCountsByKey[$canonicalKey].OPT_Record_Count
    }

    $yearsActive = $years.Count

    [PSCustomObject][ordered]@{
        Employer_Key = $metadata.Employer_Key
        Employer_Display_Name = $metadata.Employer_Display_Name
        Total_Coop_Records = $rows.Count
        Completed_Approved_Records = $approvedCompletedRows.Count
        Avg_Hourly_Wage = $avgHourlyWage
        First_Year = if ($years.Count -gt 0) { ($years | Measure-Object -Minimum).Minimum } else { $null }
        Last_Year = if ($years.Count -gt 0) { ($years | Measure-Object -Maximum).Maximum } else { $null }
        Years_Active = $yearsActive
        Seasons_Active = $seasons.Count
        Majors_Served = $majors.Count
        Industries_Served = $industries.Count
        Remote_Rate = $remoteRate
        Relationship_Type = if ($yearsActive -ge 3) { "Repeated Employer" } else { "New/Occasional Employer" }
        CPT_Record_Count = $cptCount
        CPT_Flag = if ($cptCount -gt 0) { 1 } else { 0 }
        OPT_Record_Count = $optCount
        OPT_Flag = if ($optCount -gt 0) { 1 } else { 0 }
        International_Student_Friendly = if (($cptCount + $optCount) -gt 0) { "Yes" } else { "No" }
    }
}

$standardizedDetail | Export-Csv -LiteralPath $DetailOutputPath -NoTypeInformation -Encoding UTF8
$standardizedSummary | Export-Csv -LiteralPath $SummaryOutputPath -NoTypeInformation -Encoding UTF8
$industryAudit | Sort-Object @{ Expression = "Record_Count"; Descending = $true }, @{ Expression = "Employer_Display_Name"; Descending = $false } | Export-Csv -LiteralPath $AuditOutputPath -NoTypeInformation -Encoding UTF8

Write-Output "Standardized detail file: $DetailOutputPath"
Write-Output "Standardized summary file: $SummaryOutputPath"
Write-Output "Audit file: $AuditOutputPath"
