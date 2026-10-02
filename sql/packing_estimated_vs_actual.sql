/*
Consumer query for dbo.vw_stlPackingEstimatedVsActual.

Set any filter to NULL to ignore it.
This keeps the core logic reusable in the view while allowing ad hoc filtering here.
*/

SET NOCOUNT ON;

DECLARE @BestDateFrom date = NULL;
DECLARE @BestDateTo date = NULL;
DECLARE @DeliveryDateFrom date = NULL;
DECLARE @DeliveryDateTo date = NULL;
DECLARE @ActualWorkDateFrom date = NULL;
DECLARE @ActualWorkDateTo date = NULL;
DECLARE @CustomerLike nvarchar(100) = NULL;
DECLARE @WorkCenterLike nvarchar(100) = NULL;
DECLARE @EmployeeLike nvarchar(100) = NULL;

SELECT
    [Job ID] = v.JobID,
    [Estimate Date] = v.EstimateDate,
    [Delivery Date] = v.DeliveryDate,
    [Customer] = v.CustomerName,
    [Job Description] = v.JobDescription,
    [Order Quantity] = v.OrderQuantity,
    [Actual Work Date] = v.ActualWorkDate,
    [Work Center] = v.WorkCenterName,
    [Employee] = v.EmployeeName,
    [Estimated Packing Minutes] = v.EstimatedPackingMinutes,
    [Actual Packing Minutes] = v.ActualPackingMinutes,
    [Total Actual Packing Minutes] = v.ActualPackingMinutesTotal,
    [Share of Job Actual Time] = v.ActualShareOfJob,
    [Packing Minutes Variance] = v.PackingMinutesVariance,
    [Packing Time Ratio] = v.PackingTimeRatio,
    [Packing Status] = v.PackingStatus
FROM dbo.vw_stlPackingEstimatedVsActual AS v
WHERE (@BestDateFrom IS NULL OR v.EstimateDate >= @BestDateFrom)
  AND (@BestDateTo IS NULL OR v.EstimateDate < DATEADD(DAY, 1, @BestDateTo))
  AND (@DeliveryDateFrom IS NULL OR v.DeliveryDate >= @DeliveryDateFrom)
  AND (@DeliveryDateTo IS NULL OR v.DeliveryDate < DATEADD(DAY, 1, @DeliveryDateTo))
  AND (@ActualWorkDateFrom IS NULL OR v.ActualWorkDate >= @ActualWorkDateFrom)
  AND (@ActualWorkDateTo IS NULL OR v.ActualWorkDate < DATEADD(DAY, 1, @ActualWorkDateTo))
    AND (
      @CustomerLike IS NULL
      OR (
        (CHARINDEX(N'%', @CustomerLike) > 0 OR CHARINDEX(N'_', @CustomerLike) > 0 OR CHARINDEX(N'[', @CustomerLike) > 0)
        AND v.CustomerName LIKE @CustomerLike
      )
      OR (
        CHARINDEX(N'%', @CustomerLike) = 0
        AND CHARINDEX(N'_', @CustomerLike) = 0
        AND CHARINDEX(N'[', @CustomerLike) = 0
        AND CHARINDEX(N',' + LTRIM(RTRIM(ISNULL(v.CustomerName, N''))) + N',', N',' + REPLACE(@CustomerLike, N', ', N',') + N',') > 0
      )
      )
    AND (
      @WorkCenterLike IS NULL
      OR (
        (CHARINDEX(N'%', @WorkCenterLike) > 0 OR CHARINDEX(N'_', @WorkCenterLike) > 0 OR CHARINDEX(N'[', @WorkCenterLike) > 0)
        AND v.WorkCenterName LIKE @WorkCenterLike
      )
      OR (
        CHARINDEX(N'%', @WorkCenterLike) = 0
        AND CHARINDEX(N'_', @WorkCenterLike) = 0
        AND CHARINDEX(N'[', @WorkCenterLike) = 0
        AND CHARINDEX(N',' + LTRIM(RTRIM(ISNULL(v.WorkCenterName, N''))) + N',', N',' + REPLACE(@WorkCenterLike, N', ', N',') + N',') > 0
      )
      )
    AND (
      @EmployeeLike IS NULL
      OR (
        (CHARINDEX(N'%', @EmployeeLike) > 0 OR CHARINDEX(N'_', @EmployeeLike) > 0 OR CHARINDEX(N'[', @EmployeeLike) > 0)
        AND v.EmployeeName LIKE @EmployeeLike
      )
      OR (
        CHARINDEX(N'%', @EmployeeLike) = 0
        AND CHARINDEX(N'_', @EmployeeLike) = 0
        AND CHARINDEX(N'[', @EmployeeLike) = 0
        AND CHARINDEX(N',' + LTRIM(RTRIM(ISNULL(v.EmployeeName, N''))) + N',', N',' + REPLACE(@EmployeeLike, N', ', N',') + N',') > 0
      )
      )
ORDER BY
    COALESCE(v.EstimateDate, v.DeliveryDate) DESC,
    v.JobID,
    v.ActualWorkDate,
    v.WorkCenterName,
    v.EmployeeName;
