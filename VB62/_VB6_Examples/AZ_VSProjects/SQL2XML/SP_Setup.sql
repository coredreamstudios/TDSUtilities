use NORTHWIND
GO

IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_XML_CategoryProducts' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_XML_CategoryProducts
GO

CREATE PROCEDURE dbo.sp_XML_CategoryProducts

@CategoryID	int,
@CatName		nvarchar(15) OUTPUT

AS
BEGIN
	SET @CatName = (SELECT TOP 1 CategoryName FROM Categories WHERE CategoryID = @CategoryID)

	SELECT Products.ProductName, Products.QuantityPerUnit, Products.UnitsInStock
	FROM Products
	WHERE Products.Discontinued <> 1 And Products.CategoryID = @CategoryID
	FOR XML AUTO, ELEMENTS

END
GO

IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_XML_CategoryList' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_XML_CategoryList
GO

CREATE PROCEDURE dbo.sp_XML_CategoryList 
AS
	SELECT Categories.CategoryID, Categories.CategoryName, Categories.Description
	FROM Categories
	ORDER BY CategoryName
	FOR XML AUTO
GO

