param(
    [string]$SqlServer = 'VLCCI-HJSQL02',
    [string]$Database = 'VGS_IDASH',
    [string]$OutputDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $PSScriptRoot 'PCA_Schema_Audit_Output'
}

$datasets = @(
    [PSCustomObject]@{
        Dataset     = 'OD_SUMMARY'
        CsvPath     = 'E:\Data_Import\ALMS\Inbound\RPT_PCA_OD_SUMMARY.csv'
        ImportTable = 'dbo.IMPORT_RPT_PCA_OD_SUMMARY_2306'
        TargetTable = 'dbo.METRIX_RPT_PCA_OD_SUMMARY_2306'
    },
    [PSCustomObject]@{
        Dataset     = 'PLANT'
        CsvPath     = 'E:\Data_Import\ALMS\Inbound\RPT_PCA_PLANT.csv'
        ImportTable = 'dbo.IMPORT_RPT_PCA_PLANT_2306'
        TargetTable = 'dbo.METRIX_RPT_PCA_PLANT_2306'
    }
)

function Get-NormalizedName {
    param([AllowNull()][string]$Name)

    if ($null -eq $Name) {
        return $null
    }

    return ($Name -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
}

function Get-MaximumCharacters {
    param([Parameter(Mandatory)]$SchemaColumn)

    $typeName = ([string]$SchemaColumn.TypeName).ToLowerInvariant()
    $maxLength = [int]$SchemaColumn.MaxLength

    if ($typeName -notin @('char', 'varchar', 'nchar', 'nvarchar')) {
        return $null
    }

    if ($maxLength -eq -1) {
        return -1
    }

    if ($typeName -in @('nchar', 'nvarchar')) {
        return [int]($maxLength / 2)
    }

    return $maxLength
}

function Get-DisplayCapacity {
    param([AllowNull()]$Capacity)

    if ($null -eq $Capacity) {
        return $null
    }

    if ([int]$Capacity -eq -1) {
        return 'MAX'
    }

    return [string]$Capacity
}

function Get-TableSchema {
    param(
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][string]$QualifiedTableName
    )

    $command = $Connection.CreateCommand()
    $command.CommandText = @'
SELECT
    c.column_id,
    c.name AS ColumnName,
    t.name AS TypeName,
    c.max_length,
    c.precision,
    c.scale,
    c.is_nullable
FROM sys.columns AS c
INNER JOIN sys.types AS t
    ON t.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID(@QualifiedTableName)
ORDER BY c.column_id;
'@

    $parameter = $command.Parameters.Add(
        '@QualifiedTableName',
        [System.Data.SqlDbType]::NVarChar,
        517
    )
    $parameter.Value = $QualifiedTableName

    $reader = $command.ExecuteReader()
    $columns = New-Object System.Collections.Generic.List[object]

    try {
        while ($reader.Read()) {
            [void]$columns.Add([PSCustomObject]@{
                Ordinal    = [int]$reader['column_id']
                ColumnName = [string]$reader['ColumnName']
                TypeName   = [string]$reader['TypeName']
                MaxLength  = [int]$reader['max_length']
                Precision  = [int]$reader['precision']
                Scale      = [int]$reader['scale']
                IsNullable = [bool]$reader['is_nullable']
            })
        }
    }
    finally {
        $reader.Close()
        $command.Dispose()
    }

    if ($columns.Count -eq 0) {
        throw "Table '$QualifiedTableName' was not found in database '$Database'."
    }

    return @($columns)
}

function Get-CsvStatistics {
    param([Parameter(Mandatory)][string]$CsvPath)

    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
        throw "CSV file not found: $CsvPath"
    }

    $statistics = New-Object System.Collections.Generic.List[object]
    $dataRowCount = 0

    Import-Csv -LiteralPath $CsvPath | ForEach-Object {
        $dataRowCount++
        $csvRow = $dataRowCount + 1
        $properties = @($_.PSObject.Properties)

        if ($statistics.Count -eq 0) {
            for ($index = 0; $index -lt $properties.Count; $index++) {
                [void]$statistics.Add([PSCustomObject]@{
                    Ordinal          = $index + 1
                    Header           = [string]$properties[$index].Name
                    MaxLength        = 0
                    LongestValue     = $null
                    LongestValueRow  = $null
                    NonblankRowCount = 0
                })
            }
        }

        if ($properties.Count -ne $statistics.Count) {
            throw "CSV row $csvRow has $($properties.Count) fields; the header has $($statistics.Count)."
        }

        for ($index = 0; $index -lt $properties.Count; $index++) {
            $value = [string]$properties[$index].Value
            $length = $value.Length
            $stat = $statistics[$index]

            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $stat.NonblankRowCount++
            }

            if ($length -gt $stat.MaxLength) {
                $stat.MaxLength = $length
                $stat.LongestValue = $value
                $stat.LongestValueRow = $csvRow
            }
        }
    }

    if ($dataRowCount -eq 0) {
        throw "CSV contains no data rows: $CsvPath"
    }

    return [PSCustomObject]@{
        RowCount    = $dataRowCount
        ColumnCount = $statistics.Count
        Statistics  = @($statistics)
    }
}

function Test-SqlValue {
    param(
        [AllowNull()][string]$Value,
        [Parameter(Mandatory)]$SchemaColumn
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $trimmed = $Value.Trim()
    $typeName = ([string]$SchemaColumn.TypeName).ToLowerInvariant()
    $integerValue = [long]0
    $decimalValue = [decimal]0
    $doubleValue = [double]0
    $dateValue = [datetime]::MinValue
    $dateOffsetValue = [datetimeoffset]::MinValue
    $timeValue = [timespan]::Zero
    $guidValue = [guid]::Empty
    $integerStyles = [Globalization.NumberStyles]::Integer
    $numberStyles = [Globalization.NumberStyles]::Number -bor [Globalization.NumberStyles]::AllowExponent
    $culture = [Globalization.CultureInfo]::InvariantCulture

    switch ($typeName) {
        { $_ -in @('char', 'varchar', 'nchar', 'nvarchar', 'text', 'ntext', 'xml') } {
            return $null
        }
        'tinyint' {
            if (-not [long]::TryParse($trimmed, $integerStyles, $culture, [ref]$integerValue)) {
                return 'Not a valid integer'
            }
            if ($integerValue -lt 0 -or $integerValue -gt 255) {
                return 'TINYINT overflow (valid range 0 through 255)'
            }
            return $null
        }
        'smallint' {
            if (-not [long]::TryParse($trimmed, $integerStyles, $culture, [ref]$integerValue)) {
                return 'Not a valid integer'
            }
            if ($integerValue -lt -32768 -or $integerValue -gt 32767) {
                return 'SMALLINT overflow (valid range -32768 through 32767)'
            }
            return $null
        }
        'int' {
            if (-not [long]::TryParse($trimmed, $integerStyles, $culture, [ref]$integerValue)) {
                return 'Not a valid integer'
            }
            if ($integerValue -lt -2147483648 -or $integerValue -gt 2147483647) {
                return 'INT overflow'
            }
            return $null
        }
        'bigint' {
            if (-not [long]::TryParse($trimmed, $integerStyles, $culture, [ref]$integerValue)) {
                return 'Not a valid BIGINT'
            }
            return $null
        }
        { $_ -in @('decimal', 'numeric', 'money', 'smallmoney') } {
            if (-not [decimal]::TryParse($trimmed, $numberStyles, $culture, [ref]$decimalValue)) {
                return "Not a valid $($typeName.ToUpperInvariant()) value"
            }
            if ($typeName -eq 'smallmoney' -and
                ($decimalValue -lt -214748.3648D -or $decimalValue -gt 214748.3647D)) {
                return 'SMALLMONEY overflow'
            }
            if ($typeName -eq 'money' -and
                ($decimalValue -lt -922337203685477.5808D -or $decimalValue -gt 922337203685477.5807D)) {
                return 'MONEY overflow'
            }
            return $null
        }
        { $_ -in @('float', 'real') } {
            if (-not [double]::TryParse($trimmed, $numberStyles, $culture, [ref]$doubleValue)) {
                return "Not a valid $($typeName.ToUpperInvariant()) value"
            }
            return $null
        }
        'bit' {
            if ($trimmed.ToLowerInvariant() -notin @('0', '1', 'true', 'false')) {
                return 'Not a valid BIT value'
            }
            return $null
        }
        { $_ -in @('date', 'datetime', 'datetime2', 'smalldatetime') } {
            $parsedInvariant = [datetime]::TryParse(
                $trimmed,
                $culture,
                [Globalization.DateTimeStyles]::AllowWhiteSpaces,
                [ref]$dateValue
            )
            if (-not $parsedInvariant) {
                $parsedCurrent = [datetime]::TryParse(
                    $trimmed,
                    [Globalization.CultureInfo]::CurrentCulture,
                    [Globalization.DateTimeStyles]::AllowWhiteSpaces,
                    [ref]$dateValue
                )
                if (-not $parsedCurrent) {
                    return "Not a valid $($typeName.ToUpperInvariant()) value"
                }
            }
            return $null
        }
        'datetimeoffset' {
            if (-not [datetimeoffset]::TryParse($trimmed, $culture, [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$dateOffsetValue)) {
                return 'Not a valid DATETIMEOFFSET value'
            }
            return $null
        }
        'time' {
            if (-not [timespan]::TryParse($trimmed, $culture, [ref]$timeValue)) {
                return 'Not a valid TIME value'
            }
            return $null
        }
        'uniqueidentifier' {
            if (-not [guid]::TryParse($trimmed, [ref]$guidValue)) {
                return 'Not a valid UNIQUEIDENTIFIER value'
            }
            return $null
        }
        default {
            return $null
        }
    }
}

function New-ValidationState {
    return [PSCustomObject]@{
        InvalidCount = 0
        FirstRow     = $null
        FirstValue   = $null
        FirstReason  = $null
    }
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$connectionString = "Server=$SqlServer;Database=$Database;Integrated Security=True;TrustServerCertificate=True;Application Name=PCA CSV Schema Audit;"
$connection = New-Object System.Data.SqlClient.SqlConnection $connectionString
$allResults = New-Object System.Collections.Generic.List[object]
$datasetSummaries = New-Object System.Collections.Generic.List[object]

try {
    $connection.Open()

    foreach ($dataset in $datasets) {
        Write-Host "Analyzing $($dataset.Dataset)..." -ForegroundColor Cyan

        $csv = Get-CsvStatistics -CsvPath $dataset.CsvPath
        $importSchema = @(Get-TableSchema -Connection $connection -QualifiedTableName $dataset.ImportTable)
        $targetSchema = @(Get-TableSchema -Connection $connection -QualifiedTableName $dataset.TargetTable)

        $importByOrdinal = @{}
        $targetByOrdinal = @{}
        $validation = @{}

        foreach ($column in $importSchema) {
            $importByOrdinal[$column.Ordinal] = $column
        }
        foreach ($column in $targetSchema) {
            $targetByOrdinal[$column.Ordinal] = $column
        }

        foreach ($stat in $csv.Statistics) {
            $validation["Import:$($stat.Ordinal)"] = New-ValidationState
            $validation["Target:$($stat.Ordinal)"] = New-ValidationState
        }

        $dataRowNumber = 0
        Import-Csv -LiteralPath $dataset.CsvPath | ForEach-Object {
            $dataRowNumber++
            $csvRow = $dataRowNumber + 1
            $properties = @($_.PSObject.Properties)

            for ($index = 0; $index -lt $properties.Count; $index++) {
                $ordinal = $index + 1
                $value = [string]$properties[$index].Value

                foreach ($stage in @('Import', 'Target')) {
                    $schemaColumn = if ($stage -eq 'Import') {
                        $importByOrdinal[$ordinal]
                    }
                    else {
                        $targetByOrdinal[$ordinal]
                    }

                    if ($null -eq $schemaColumn) {
                        continue
                    }

                    $reason = Test-SqlValue -Value $value -SchemaColumn $schemaColumn
                    if ($null -ne $reason) {
                        $state = $validation["$stage`:$ordinal"]
                        $state.InvalidCount++

                        if ($null -eq $state.FirstRow) {
                            $state.FirstRow = $csvRow
                            $state.FirstValue = $value
                            $state.FirstReason = $reason
                        }
                    }
                }
            }
        }

        foreach ($stat in $csv.Statistics) {
            $ordinal = [int]$stat.Ordinal
            $importColumn = $importByOrdinal[$ordinal]
            $targetColumn = $targetByOrdinal[$ordinal]
            $importValidation = $validation["Import:$ordinal"]
            $targetValidation = $validation["Target:$ordinal"]
            $issues = New-Object System.Collections.Generic.List[string]

            $importCapacity = $null
            $targetCapacity = $null
            $importNameMatch = $null
            $targetNameMatch = $null

            if ($null -eq $importColumn) {
                [void]$issues.Add('CSV ordinal is missing from import table')
            }
            else {
                $importCapacity = Get-MaximumCharacters -SchemaColumn $importColumn
                $importNameMatch = (Get-NormalizedName $stat.Header) -eq (Get-NormalizedName $importColumn.ColumnName)

                if ($null -ne $importCapacity -and $importCapacity -ne -1 -and $stat.MaxLength -gt $importCapacity) {
                    [void]$issues.Add("Import character overflow: CSV max $($stat.MaxLength), SQL max $importCapacity")
                }
                if ($importValidation.InvalidCount -gt 0) {
                    [void]$issues.Add("Import conversion failures: $($importValidation.InvalidCount); $($importValidation.FirstReason)")
                }
            }

            if ($null -eq $targetColumn) {
                [void]$issues.Add('CSV ordinal is missing from target table')
            }
            else {
                $targetCapacity = Get-MaximumCharacters -SchemaColumn $targetColumn
                $targetNameMatch = (Get-NormalizedName $stat.Header) -eq (Get-NormalizedName $targetColumn.ColumnName)

                if ($null -ne $targetCapacity -and $targetCapacity -ne -1 -and $stat.MaxLength -gt $targetCapacity) {
                    [void]$issues.Add("Target character overflow: CSV max $($stat.MaxLength), SQL max $targetCapacity")
                }
                if ($targetValidation.InvalidCount -gt 0) {
                    [void]$issues.Add("Target conversion failures: $($targetValidation.InvalidCount); $($targetValidation.FirstReason)")
                }
            }

            [void]$allResults.Add([PSCustomObject]@{
                Dataset                 = $dataset.Dataset
                CsvPath                 = $dataset.CsvPath
                DataRows                = $csv.RowCount
                Ordinal                 = $ordinal
                CsvHeader               = $stat.Header
                CsvMaxLength            = $stat.MaxLength
                CsvLongestValue         = $stat.LongestValue
                CsvLongestValueRow      = $stat.LongestValueRow
                CsvNonblankRows         = $stat.NonblankRowCount
                ImportTable             = $dataset.ImportTable
                ImportColumn            = if ($null -eq $importColumn) { $null } else { $importColumn.ColumnName }
                ImportNameMatches       = $importNameMatch
                ImportType              = if ($null -eq $importColumn) { $null } else { $importColumn.TypeName }
                ImportMaxCharacters     = Get-DisplayCapacity $importCapacity
                ImportPrecision         = if ($null -eq $importColumn) { $null } else { $importColumn.Precision }
                ImportScale             = if ($null -eq $importColumn) { $null } else { $importColumn.Scale }
                ImportInvalidCount      = $importValidation.InvalidCount
                ImportFirstInvalidRow   = $importValidation.FirstRow
                ImportFirstInvalidValue = $importValidation.FirstValue
                ImportInvalidReason     = $importValidation.FirstReason
                TargetTable             = $dataset.TargetTable
                TargetColumn            = if ($null -eq $targetColumn) { $null } else { $targetColumn.ColumnName }
                TargetNameMatches       = $targetNameMatch
                TargetType              = if ($null -eq $targetColumn) { $null } else { $targetColumn.TypeName }
                TargetMaxCharacters     = Get-DisplayCapacity $targetCapacity
                TargetPrecision         = if ($null -eq $targetColumn) { $null } else { $targetColumn.Precision }
                TargetScale             = if ($null -eq $targetColumn) { $null } else { $targetColumn.Scale }
                TargetInvalidCount      = $targetValidation.InvalidCount
                TargetFirstInvalidRow   = $targetValidation.FirstRow
                TargetFirstInvalidValue = $targetValidation.FirstValue
                TargetInvalidReason     = $targetValidation.FirstReason
                Issue                   = $issues -join ' | '
            })
        }

        [void]$datasetSummaries.Add([PSCustomObject]@{
            Dataset          = $dataset.Dataset
            CsvRows          = $csv.RowCount
            CsvColumns       = $csv.ColumnCount
            ImportColumns    = $importSchema.Count
            TargetColumns    = $targetSchema.Count
            ExpectedTarget   = $csv.ColumnCount + 1
            ImportCountMatch = $importSchema.Count -eq $csv.ColumnCount
            TargetCountMatch = $targetSchema.Count -eq ($csv.ColumnCount + 1)
        })
    }
}
finally {
    if ($connection.State -ne [System.Data.ConnectionState]::Closed) {
        $connection.Close()
    }
    $connection.Dispose()
}

$fullReportPath = Join-Path $OutputDirectory 'PCA_CSV_SQL_schema_audit_full.csv'
$issueReportPath = Join-Path $OutputDirectory 'PCA_CSV_SQL_schema_audit_issues.csv'
$summaryPath = Join-Path $OutputDirectory 'PCA_CSV_SQL_schema_audit_summary.csv'

$allResults |
    Export-Csv -LiteralPath $fullReportPath -NoTypeInformation -Encoding UTF8

$issues = @($allResults | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Issue) })
$issues |
    Export-Csv -LiteralPath $issueReportPath -NoTypeInformation -Encoding UTF8

$datasetSummaries |
    Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host 'Dataset structure summary:' -ForegroundColor Cyan
$datasetSummaries | Format-Table -AutoSize

Write-Host ''
if ($issues.Count -eq 0) {
    Write-Host 'No capacity or conversion problems were detected.' -ForegroundColor Green
}
else {
    Write-Host "$($issues.Count) column problem(s) detected:" -ForegroundColor Yellow
    $issues |
        Select-Object Dataset, Ordinal, CsvHeader, CsvMaxLength, ImportColumn, ImportType, TargetColumn, TargetType, Issue |
        Format-Table -Wrap -AutoSize
}

Write-Host ''
Write-Host "Full report:    $fullReportPath"
Write-Host "Issues report:  $issueReportPath"
Write-Host "Summary report: $summaryPath"

