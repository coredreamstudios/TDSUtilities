<%

'Call a SQL Server 2000 XML Returning Stored Procedure, from ASP
'
' ** Note, you will have to enable EXEC permissions on the Stored Procedure in SQL Server 2000 for 
'     IUSR_YourServerName for this ASP to run the query
'
'	- This example uses the cSPXML class compiled in an Active X dll for use within ASP pages. It shows
'	  how to call a stored procedure within an input parameter and an output parameter.
'

dim objSP, objXML, objXSL

set objSP = Server.CreateObject("SQL2KSPEXEC.cSPXML")

'set the connection string
objSP.ConnectString = "Provider=SQLOLEDB;SERVER=SERVER1;Database=Northwind;Trusted_Connection=yes;"

objSP.ParamCount = 2				'set the parameter count (important)

objSP.AddParam "@CategoryID", 3, 1, 4, 4	'set the input parameter (CategoryID 4 = Dairy Products) 	
objSP.AddParam "@CatName", 200, 2, 15		'set the output parameter that we retrive later

'this calls the XML Stored Procedure we created
'we pass it the name of the stored procedure and a root node value to wrap around the XML returned by SQL Server

if objSP.CallSPXML("sp_XML_CategoryProducts", "<Root>") then

	'write the category name from the output parameter
	response.write "<H1> Here Are The Products In Category - "
	response.write objSP.ReturnVal("@CatName")
	response.write "</H1>"

	set objXML = Server.CreateObject("MSXML2.DOMDocument")
	set objXSL = Server.CreateObject("MSXML2.DOMDocument")

	objXML.async = false
	objXSL.async = false

	
	objXML.LoadXML objSP.XML				'load the XML from our object into the XML DOMDocument
				
	objXSL.Load Server.MapPath("products.xsl")		'load the XSL file into a DOM Document
	
	response.write objXML.TransformNode(objXSL)		'write the transformed XML to the page

	set objXML = Nothing
	set objXSL = Nothing
else
	'there was an error, display the error message
	response.write objSP.Message

	'you can also get a more technical look at the error message using the ConnErrors property
	response.write objSP.ConnErrors
end if

set objSP = Nothing


%>