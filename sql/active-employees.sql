SELECT
  -- If NameFirstPreferred is present then use this as first name
    EmpID
  ,CASE
     WHEN NameFirstPreferred <> ''
      THEN NameFirstPreferred
     ELSE NameFirst
     END AS NameFirst
    , NameLast
    , NameMiddle
    , EmailWork
    , JobClassDescr
    , JobCategoryDescr
    , SiteId
    , SiteDescr
    , EmploymentStatusDescr
    , EmploymentStatusCode
    , EmploymentTypeCode
		, (CASE
		WHEN NameFirstPreferred <> ''
		   THEN NameFirstPreferred
		ELSE NameFirst
		END
	 + ' ' + NameLast)       AS FullName
    , BargUnitID
    , PersonTypeId
FROM vwHREmploymentList
WHERE
  -- EmploymentStatusCode NOT IN ('R','T','I','X','L','W')
  EmploymentStatusCode IN ('A','S')
   AND DateTimeEdited >= DATEADD(day,-30, GETDATE())
   -- AND EmpID IN (1234567) -- For testing
   -- AND SiteId = 010
ORDER BY NameLast;