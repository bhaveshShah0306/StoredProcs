DECLARE @EmployeeID int = 122;
DECLARE @StartNode hierarchyid;

SELECT @StartNode = OrganizationNode
FROM HumanResources.Employee
WHERE BusinessEntityID = @EmployeeID;

SELECT
    e.BusinessEntityID,
    p.FirstName + ' ' + p.MiddleName + ' ' + p.LastName AS [Employee Name],
    e.OrganizationNode.GetLevel() AS [Level],
    mp.FirstName + ' ' + mp.MiddleName + ' ' + mp.LastName AS [Manager Name],
    e.OrganizationNode,
    e.JobTitle
FROM HumanResources.Employee e
JOIN Person.Person p
    ON p.BusinessEntityID = e.BusinessEntityID
LEFT JOIN HumanResources.Employee m
    ON e.OrganizationNode.GetAncestor(1) = m.OrganizationNode
LEFT JOIN Person.Person mp
    ON mp.BusinessEntityID = m.BusinessEntityID
WHERE @StartNode.IsDescendantOf(e.OrganizationNode) = 1
ORDER BY e.OrganizationNode.GetLevel() DESC;
