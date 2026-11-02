/* =====================================================================
   SQL Wait Monitor - test scenario: recreate a blocking incident on a TEST server
   and watch it in the viewer (statement view, blocking tab, story report).
   Replace YourTestDb / dbo.YourBigTable with a real database and a big table
   (e.g. StackOverflow.dbo.Posts or AdventureWorks Sales.SalesOrderDetail).
   ===================================================================== */

/* ---- Step 1: a demo proc with several statements, one of them dynamic SQL ---- */
USE YourTestDb;
GO
CREATE OR ALTER PROCEDURE dbo.usp_WaitDemo
AS
BEGIN
    SET NOCOUNT ON;

    /* Statement A: plain read - will wait on the schema lock (LCK_M_SCH_S) */
    SELECT COUNT(*) FROM dbo.YourBigTable;

    /* Statement B: dynamic SQL - shows as "Ad hoc", with Called from = EXEC dbo.usp_WaitDemo */
    DECLARE @sql NVARCHAR(MAX) = N'SELECT TOP (100) * FROM dbo.YourBigTable ORDER BY NEWID();';
    EXEC sys.sp_executesql @sql;
END
GO

/* ---- Step 2 (window 1): the blocker - offline rebuild inside an open transaction ---- */
-- BEGIN TRAN;
-- ALTER INDEX ALL ON dbo.YourBigTable REBUILD;
-- ...leave it open 5-10 minutes, then:
-- ROLLBACK;

/* ---- Step 3 (windows 2..N, or SQLQueryStress with 20-40 threads): the victims ---- */
-- EXEC dbo.usp_WaitDemo;

/* ---- Step 4: check the capture by hand (the viewer shows the same) ---- */
USE DBAWait;
GO
SELECT TOP (20) v.Label, w.WaitCategory, SUM(w.WeightSec) AS WaitSec, MAX(v.LastCalledFrom) AS CalledFrom
FROM dbo.WaitSamples w
JOIN dbo.vStatements v ON v.StatementId = w.StatementId
WHERE w.SampleTime >= DATEADD(MINUTE, -15, SYSUTCDATETIME())
GROUP BY v.Label, w.WaitCategory
ORDER BY WaitSec DESC;

SELECT TOP (20) * FROM dbo.BlockerSamples ORDER BY SampleTime DESC;   -- the idle ALTER INDEX session, status 'sleeping'
GO

/* Expected:
   - "#n dbo.usp_WaitDemo line 7"  with LCK_M_SCH_S   (statement A, inside the proc;
     line numbers count from the CREATE line, same as the proc's definition in SSMS)
   - "#m Ad hoc: SELECT TOP (100) * FROM dbo.YourBigTa..." (statement B, the dynamic SQL),
     Called from = EXEC dbo.usp_WaitDemo
   - BlockerSamples: the rebuild session, status sleeping, Open tran 1, Max blocked ~ number of threads
   Then: EXEC DBAWait.dbo.usp_WaitStory;  -- chapter 5 tells the blocking incident in plain English
*/
