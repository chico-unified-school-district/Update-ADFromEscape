SELECT
  -- If NameFirstPreferred is present then use this as first name
  CASE
     WHEN NameFirstPreferred <> ''
      THEN NameFirstPreferred
     ELSE NameFirst
     END AS NameFirst
    , NameLast
    , NameMiddle
    , EmpID
    , JobClassDescr
    , JobCategoryDescr
    , SiteID
    , SiteDescr
    , EmploymentStatusDescr
    , EmploymentStatusCode
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
   AND DateTimeEdited >= DATEADD(day, -7, GETDATE())
   -- AND empId = 9999999
-- ORDER BY EmploymentStatusCode;
ORDER BY NameLast;