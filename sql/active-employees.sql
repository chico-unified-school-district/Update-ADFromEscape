SELECT
    -- If NameFirstPreferred is present then use this data as GivenName
    CASE
     WHEN vwHREmploymentList.NameFirstPreferred <> ''
      THEN vwHREmploymentList.NameFirstPreferred
     ELSE vwHREmploymentList.NameFirst
     END                                         AS givenname,
    vwHREmploymentList.NameLast                  AS sn,
    vwHREmploymentList.NameMiddle                AS middlename,
    SUBSTRING(vwHREmploymentList.NameMiddle,1,1) AS initials,
    vwHREmploymentList.EmpID                     AS employeeID,
    'Chico Unified School District'              AS company,
    vwHREmploymentList.JobClassDescr             AS title,
    vwHREmploymentList.JobClassDescr             AS description,
    vwHREmploymentList.JobCategoryDescr          AS department,
    vwHREmploymentList.SiteID                    AS departmentnumber,
        -- ( CASE
 --     WHEN vwHREmploymentList.NameFirstPreferred <> ''
 --      THEN vwHREmploymentList.NameFirstPreferred
 --     ELSE vwHREmploymentList.NameFirst
 --     END  ) + ' ' + vwHREmploymentList.NameLast  AS DisplayName,
    vwHREmploymentList.SiteDescr                 AS physicalDeliveryOfficeName,
    vwHREmploymentList.BargUnitID                AS extensionAttribute1

   FROM vwHREmploymentList
   LEFT JOIN HREmployment ON HREmployment.EmpID = vwHREmploymentList.EmpID
   WHERE
        HREmployment.PersonTypeId IN (1,2,4)
    AND HREmployment.EmploymentStatusCode IN ('A','L','W','I','S')
    --AND vwHREmploymentList.DateTimeEdited > DATEADD(day,-5,getdate())
   ORDER BY employeeID