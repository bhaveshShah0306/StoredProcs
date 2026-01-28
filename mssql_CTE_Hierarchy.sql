Declare @EmployeeID int = 122

;WITH employee_hierarchy 
AS
(
SELECT	 e.BusinessEntityID,
		 p.FirstName+' '+ p.MiddleName+ ' '+ p.LastName as [Employee Name],
		 e.OrganizationNode.GetLevel() AS [Level],
		CAST(NULL AS nvarchar(200)) AS [Manager Name],
	        0 AS RecursionLevel,
			e.OrganizationNode,
			e.JobTitle
			
			FROM HumanResources.Employee  e
		JOIN Person.Person as p 
		on e.BusinessEntityID = p.BusinessEntityID
	where e.BusinessEntityID = @EmployeeID

UNION ALL

SELECT 
			m.BusinessEntityID,
			mp.FirstName+ ' '+ mp.MiddleName+' '+ mp.LastName  as [Employee Name],
			m.OrganizationNode.GetLevel() AS [Level],
			CAST(null AS nvarchar(200))as [Manager Name] ,
			ct.RecursionLevel + 1,
			m.OrganizationNode,
			m.JobTitle
	FROM employee_hierarchy ct
		JOIN HumanResources.Employee m
	ON ct.OrganizationNode.GetAncestor(1) = m.OrganizationNode
		JOIN Person.Person mp
	ON 
	mp.BusinessEntityID = m.BusinessEntityID
)
SELECT
    eh.BusinessEntityID,
    eh.[Employee Name],
    eh.[Level],
    mgr.[Employee Name] AS [Manager Name],
    eh.RecursionLevel,
    eh.OrganizationNode,
    eh.JobTitle
FROM employee_hierarchy eh
LEFT JOIN employee_hierarchy mgr
    ON mgr.RecursionLevel = eh.RecursionLevel + 1
ORDER BY eh.RecursionLevel
OPTION (MAXRECURSION 25);

