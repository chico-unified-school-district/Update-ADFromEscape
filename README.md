# Update-ADFromEscape
Pull data from Employee Database (Escape Online) and update Active Directory user objects using employeeID as the foreign key.
Accounts with read access to the employee database and modify user access to Active Directory are required as well as an OrgUnitPath
in AD where employee user objects are stored.

Data fields refenced in the project's SQL query file HAVE to be aliased 
as reaplaceable string-based Active Directory attributes.
 - Example: SELECT someDBfield1 AS givenName,someDBfield2 AS sn FROM userDB
 - Both 'givenName' and 'sn'  are both writeable string-based Active Directory Attributes
