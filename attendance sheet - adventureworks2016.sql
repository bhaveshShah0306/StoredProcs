use AdventureWorksDW2016

-- Set date parameters - now supporting multi-year ranges
DECLARE @StartDate DATE = '2007-01-01';
DECLARE @EndDate DATE = '2009-12-31';

-- Get the range of years to process
DECLARE @StartYear INT = YEAR(@StartDate);
DECLARE @EndYear INT = YEAR(@EndDate);
DECLARE @CurrentYear INT = @StartYear;

-- Process one year at a time
WHILE @CurrentYear <= @EndYear
BEGIN
    -- Set the year boundaries
    DECLARE @YearStart DATE = DATEFROMPARTS(@CurrentYear, 1, 1);
    DECLARE @YearEnd DATE = DATEFROMPARTS(@CurrentYear, 12, 31);
    
    -- Check if our temp results table exists, and drop it if it does
    IF OBJECT_ID('tempdb..#CurrentYearResults', 'U') IS NOT NULL
        DROP TABLE #CurrentYearResults;
    
    -- Create a temp table for this year's results
    CREATE TABLE #CurrentYearResults (
        RowNum INT IDENTITY(1,1),
        Title NVARCHAR(255),
        FullName NVARCHAR(255),
        [Jan] NVARCHAR(10),
        [Feb] NVARCHAR(10),
        [Mar] NVARCHAR(10),
        [Apr] NVARCHAR(10),
        [May] NVARCHAR(10),
        [Jun] NVARCHAR(10),
        [Jul] NVARCHAR(10),
        [Aug] NVARCHAR(10),
        [Sep] NVARCHAR(10),
        [Oct] NVARCHAR(10),
        [Nov] NVARCHAR(10),
        [Dec] NVARCHAR(10)
    );

	    
    -- Insert the report title as the first row
    INSERT INTO #CurrentYearResults (Title, FullName, [Jan], [Feb], [Mar], [Apr], [May], [Jun], [Jul], [Aug], [Sep], [Oct], [Nov], [Dec])
    VALUES ('Attendance Report for Year: ' + CAST(@CurrentYear AS NVARCHAR(4)), NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL);
    
    -- Get titles with ranking and employee data for the current year only
    WITH CleanTitles AS (
    SELECT DISTINCT UPPER(TRIM(Title)) as Title
    FROM DimEmployee
	),	
	TitleGroups AS (
        SELECT DISTINCT 
            Title,
            ROW_NUMBER() OVER (ORDER BY Title) AS TitleRank
        FROM CleanTitles
    ),
    EmployeeRanks AS (
        SELECT 
            e.Title,
            tg.TitleRank,
            e.LastName + ', ' + e.FirstName AS FullName,
            e.StartDate,
            e.EndDate,
            ROW_NUMBER() OVER (PARTITION BY e.Title ORDER BY e.LastName, e.FirstName) AS EmployeeRank
        FROM DimEmployee e
        JOIN TitleGroups tg ON e.Title = tg.Title
        -- Only include employees who were active at any point during this year
        WHERE (e.StartDate <= @YearEnd) 
          AND (e.EndDate IS NULL OR e.EndDate >= @YearStart)
    )
    
    -- Insert the attendance data for the CURRENT YEAR directly
    INSERT INTO #CurrentYearResults (Title, FullName, [Jan], [Feb], [Mar], [Apr], [May], [Jun], [Jul], [Aug], [Sep], [Oct], [Nov], [Dec])
    SELECT 
        CASE WHEN EmployeeRank = 1 THEN Title ELSE NULL END AS Title,
        FullName,
        -- January
        CASE WHEN (StartDate <= DATEFROMPARTS(@CurrentYear, 1, 31) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 1, 1)))
             THEN '1' ELSE NULL END AS [Jan],
        -- February (accounting for leap years)
        CASE WHEN (StartDate <= EOMONTH(DATEFROMPARTS(@CurrentYear, 2, 1)) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 2, 1)))
             THEN '1' ELSE NULL END AS [Feb],
        -- March
        CASE WHEN (StartDate <= DATEFROMPARTS(@CurrentYear, 3, 31) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 3, 1)))
             THEN '1' ELSE NULL END AS [Mar],
        -- April
        CASE WHEN (StartDate <= DATEFROMPARTS(@CurrentYear, 4, 30) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 4, 1)))
             THEN '1' ELSE NULL END AS [Apr],
        -- May
        CASE WHEN (StartDate <= DATEFROMPARTS(@CurrentYear, 5, 31) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 5, 1)))
             THEN '1' ELSE NULL END AS [May],
        -- June
        CASE WHEN (StartDate <= DATEFROMPARTS(@CurrentYear, 6, 30) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 6, 1)))
             THEN '1' ELSE NULL END AS [Jun],
        -- July
        CASE WHEN (StartDate <= DATEFROMPARTS(@CurrentYear, 7, 31) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 7, 1)))
             THEN '1' ELSE NULL END AS [Jul],
        -- August
        CASE WHEN (StartDate <= DATEFROMPARTS(@CurrentYear, 8, 31) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 8, 1)))
             THEN '1' ELSE NULL END AS [Aug],
        -- September
        CASE WHEN (StartDate <= DATEFROMPARTS(@CurrentYear, 9, 30) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 9, 1)))
             THEN '1' ELSE NULL END AS [Sep],
        -- October
        CASE WHEN (StartDate <= DATEFROMPARTS(@CurrentYear, 10, 31) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 10, 1)))
             THEN '1' ELSE NULL END AS [Oct],
        -- November
        CASE WHEN (StartDate <= DATEFROMPARTS(@CurrentYear, 11, 30) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 11, 1)))
             THEN '1' ELSE NULL END AS [Nov],
        -- December
        CASE WHEN (StartDate <= DATEFROMPARTS(@CurrentYear, 12, 31) AND (EndDate IS NULL OR EndDate >= DATEFROMPARTS(@CurrentYear, 12, 1)))
             THEN '1' ELSE NULL END AS [Dec]
    FROM EmployeeRanks
    ORDER BY TitleRank, EmployeeRank;
    
    -- Output the result set for the current year
    SELECT 
        Title,
        FullName,
        [Jan],
        [Feb],
        [Mar],
        [Apr],
        [May],
        [Jun],
        [Jul],
        [Aug],
        [Sep],
        [Oct],
        [Nov],
        [Dec]
    FROM #CurrentYearResults
    ORDER BY RowNum;
    
    -- Move to the next year
    SET @CurrentYear = @CurrentYear + 1;
END;
