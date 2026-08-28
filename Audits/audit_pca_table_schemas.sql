USE [VGS_IDASH];
GO

/*
    Schema-only companion audit for the RPT_PCA job's Steps 2 and 3.

    This script does not modify data or table definitions. It compares the
    import and final tables by physical ordinal, because both stored procedures
    currently depend on ordinal mapping. Run audit_pca_csv_vs_sql.ps1 for the
    CSV-aware length and conversion audit.
*/

SET NOCOUNT ON;

DECLARE @TablePairs TABLE
(
    Dataset     varchar(30)  NOT NULL,
    ImportTable sysname      NOT NULL,
    TargetTable sysname      NOT NULL
);

INSERT INTO @TablePairs (Dataset, ImportTable, TargetTable)
VALUES
    ('OD_SUMMARY', 'dbo.IMPORT_RPT_PCA_OD_SUMMARY_2306', 'dbo.METRIX_RPT_PCA_OD_SUMMARY_2306'),
    ('PLANT',      'dbo.IMPORT_RPT_PCA_PLANT_2306',      'dbo.METRIX_RPT_PCA_PLANT_2306');

;WITH ImportColumns AS
(
    SELECT
        p.Dataset,
        p.ImportTable,
        p.TargetTable,
        c.column_id AS Ordinal,
        c.name AS ImportColumn,
        t.name AS ImportType,
        c.max_length AS ImportStorageBytes,
        CASE
            WHEN c.max_length = -1 THEN -1
            WHEN t.name IN ('nvarchar', 'nchar') THEN c.max_length / 2
            WHEN t.name IN ('varchar', 'char') THEN c.max_length
        END AS ImportMaximumCharacters,
        c.precision AS ImportPrecision,
        c.scale AS ImportScale,
        c.is_nullable AS ImportIsNullable
    FROM @TablePairs AS p
    INNER JOIN sys.columns AS c
        ON c.object_id = OBJECT_ID(p.ImportTable)
    INNER JOIN sys.types AS t
        ON t.user_type_id = c.user_type_id
),
TargetColumns AS
(
    SELECT
        p.Dataset,
        p.ImportTable,
        p.TargetTable,
        c.column_id AS Ordinal,
        c.name AS TargetColumn,
        t.name AS TargetType,
        c.max_length AS TargetStorageBytes,
        CASE
            WHEN c.max_length = -1 THEN -1
            WHEN t.name IN ('nvarchar', 'nchar') THEN c.max_length / 2
            WHEN t.name IN ('varchar', 'char') THEN c.max_length
        END AS TargetMaximumCharacters,
        c.precision AS TargetPrecision,
        c.scale AS TargetScale,
        c.is_nullable AS TargetIsNullable
    FROM @TablePairs AS p
    INNER JOIN sys.columns AS c
        ON c.object_id = OBJECT_ID(p.TargetTable)
    INNER JOIN sys.types AS t
        ON t.user_type_id = c.user_type_id
)
SELECT
    COALESCE(i.Dataset, t.Dataset) AS Dataset,
    COALESCE(i.Ordinal, t.Ordinal) AS Ordinal,
    i.ImportColumn,
    i.ImportType,
    i.ImportMaximumCharacters,
    i.ImportPrecision,
    i.ImportScale,
    i.ImportIsNullable,
    t.TargetColumn,
    t.TargetType,
    t.TargetMaximumCharacters,
    t.TargetPrecision,
    t.TargetScale,
    t.TargetIsNullable,
    CASE
        WHEN i.Ordinal IS NULL THEN 'Target ordinal has no import-table column'
        WHEN t.Ordinal IS NULL THEN 'Import ordinal has no target-table column'
        WHEN LOWER(REPLACE(REPLACE(i.ImportColumn, '_', ''), ' ', ''))
           <> LOWER(REPLACE(REPLACE(t.TargetColumn, '_', ''), ' ', ''))
            THEN 'Column names differ at this ordinal; verify procedure mapping'
        WHEN i.ImportType IN ('nvarchar', 'varchar', 'nchar', 'char')
         AND t.TargetType IN ('nvarchar', 'varchar', 'nchar', 'char')
         AND i.ImportMaximumCharacters <> -1
         AND t.TargetMaximumCharacters <> -1
         AND t.TargetMaximumCharacters < i.ImportMaximumCharacters
            THEN 'Target character capacity is smaller than import capacity'
        WHEN i.ImportType IN ('nvarchar', 'varchar', 'nchar', 'char')
         AND t.TargetType IN ('tinyint', 'smallint', 'int', 'bigint',
                              'decimal', 'numeric', 'money', 'smallmoney',
                              'float', 'real', 'bit', 'date', 'datetime',
                              'datetime2', 'smalldatetime', 'datetimeoffset',
                              'time', 'uniqueidentifier')
            THEN 'Text import requires conversion to the target datatype; inspect CSV values'
        ELSE NULL
    END AS ReviewReason
FROM ImportColumns AS i
FULL OUTER JOIN TargetColumns AS t
    ON t.Dataset = i.Dataset
   AND t.Ordinal = i.Ordinal
ORDER BY
    COALESCE(i.Dataset, t.Dataset),
    COALESCE(i.Ordinal, t.Ordinal);

/* Condensed list of schema mappings that deserve review. */
;WITH Pairs AS
(
    SELECT
        p.Dataset,
        ic.column_id AS Ordinal,
        ic.name AS ImportColumn,
        it.name AS ImportType,
        CASE
            WHEN ic.max_length = -1 THEN -1
            WHEN it.name IN ('nvarchar', 'nchar') THEN ic.max_length / 2
            WHEN it.name IN ('varchar', 'char') THEN ic.max_length
        END AS ImportMaximumCharacters,
        tc.name AS TargetColumn,
        tt.name AS TargetType,
        CASE
            WHEN tc.max_length = -1 THEN -1
            WHEN tt.name IN ('nvarchar', 'nchar') THEN tc.max_length / 2
            WHEN tt.name IN ('varchar', 'char') THEN tc.max_length
        END AS TargetMaximumCharacters
    FROM @TablePairs AS p
    INNER JOIN sys.columns AS ic
        ON ic.object_id = OBJECT_ID(p.ImportTable)
    INNER JOIN sys.types AS it
        ON it.user_type_id = ic.user_type_id
    LEFT JOIN sys.columns AS tc
        ON tc.object_id = OBJECT_ID(p.TargetTable)
       AND tc.column_id = ic.column_id
    LEFT JOIN sys.types AS tt
        ON tt.user_type_id = tc.user_type_id
)
SELECT
    Dataset,
    Ordinal,
    ImportColumn,
    ImportType,
    ImportMaximumCharacters,
    TargetColumn,
    TargetType,
    TargetMaximumCharacters
FROM Pairs
WHERE TargetColumn IS NULL
   OR LOWER(REPLACE(REPLACE(ImportColumn, '_', ''), ' ', ''))
      <> LOWER(REPLACE(REPLACE(TargetColumn, '_', ''), ' ', ''))
   OR
   (
       ImportType IN ('nvarchar', 'varchar', 'nchar', 'char')
       AND TargetType IN ('nvarchar', 'varchar', 'nchar', 'char')
       AND ImportMaximumCharacters <> -1
       AND TargetMaximumCharacters <> -1
       AND TargetMaximumCharacters < ImportMaximumCharacters
   )
   OR
   (
       ImportType IN ('nvarchar', 'varchar', 'nchar', 'char')
       AND TargetType IN ('tinyint', 'smallint', 'int', 'bigint',
                          'decimal', 'numeric', 'money', 'smallmoney',
                          'float', 'real', 'bit', 'date', 'datetime',
                          'datetime2', 'smalldatetime', 'datetimeoffset',
                          'time', 'uniqueidentifier')
   )
ORDER BY Dataset, Ordinal;
GO
