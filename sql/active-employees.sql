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
		, (CASE
		WHEN NameFirstPreferred <> ''
		   THEN NameFirstPreferred
		ELSE NameFirst
		END
	 + ' ' + NameLast)       AS FullName
    , SiteDescr
    , BargUnitID
FROM vwHREmploymentList
WHERE
  -- PersonTypeId IN (1,2,4)
  -- AND
  EmploymentStatusCode NOT IN ('R','T')
AND DateTimeEdited > DATEADD(day,-45,getdate())
ORDER BY empId;