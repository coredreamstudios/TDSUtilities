************ SQL2KEXEC/cSPXML Read ME  *********************


OVERVIEW

	Recently I have been working with SQL Server 2000 and it's XML returning features, mostly within an ASP
environment. I found myself writing many different pages that display table-like and hierarchical data. I also 
found myself writing repetitive code in every page, mostly ADO code and the error handling involved with it. 
So I wrote this class to save myself some coding. This class acts as a re-useable object for calling Stored 
Procedures and SQL Queries that return XML in SQL Server 2000. 

	With a few lines of code, you can return xml from SQL Server to your application. You can pass multiple input 
parameters to the stored procedure as well as output parameters for returning singular data. It also lets you
easily access @RETURN_VALUE parameters from your Stored Procedures so that business rules can be applied within
the stored procedure. You can call multiple stored procedures and Queries with while sharing one connection. We
can do this by using the setting the ADO Connection's CursorLocation to adUseClient. 

	I have included a Visual Basic Project with this class, the project can be used to return straight XML 
to the client application.  The project can retrieve XML from SQL Server via a Stored Procedure or an SQL 
statement with the 'FOR XML [AUTO|RAW|EXPLICIT]' statement at the end.  

	I have also included an ASP page (CallSP.asp) and an XSL file (Products.xsl) to show an example of using the 
class from within ASP. The ASP page calls a stored procedure that has an input parameter and an output parameter.
After the XML is returned to the page, we apply an XSL style sheet to transform it into straight HTML on the
server using the MSXML2.DOMDocument object.  The XML could also be sent to the client (IE only) and the XSL applied 
on the client side to free up Server Resources.

	This example uses the Northwind Database that ships with  SQL Server 2000. I also sent along a short sql 
script (SQL_Setup.sql) to create several stored procedure examples. The proper permissions must be set in SQL to 
view or execute the stored procedures from the Visual Basic Project as well as from an ASP Page. 

	I currently have this class compiled in an Active X DLL that supports COM+ transactions. I use it from many 
of my ASP pages in combination with XSL to create quick data driven pages. I could also be added to any 
Visual Basic Project that needs to work with XML from SQL server 2000.

	If anyone has any suggestions on optimizing it, improved error handling or anything else,
I would love to hear their idea's. I can be reached at sullivan_josh@hotmail.com.


Some of the notable Properties/Methods to look at include the following:

CallSPXML	True/False
	- optional parameters are the stored procedure name and the root node to wrap around the xml
	- this calls a stored procedure that will return XML from SQL Server 2000. Input & output parameters 
	  may be used. The parameters have to be added before calling this function, using the AddParam
	  method.
	- the XML is returned in an ADO Stream object using the ADO Command's 'Output Stream' property
	- if UseReturnParameter is set to true (default), this function will fail if the stored procedure 
	  @RETURN_VALUE is anything but 0.

CallQueryXML	True/False
	- optional parameter is the SQL text that you want to run
	- this wraps the given SQL text within special XML tags that can be passed to SQL Server 2000 through
	  the ADO Command's CommandStream property.
	- the XML is returned in an ADO Stream object using the ADO Command's 'Output Stream' property

XML	String
	- this is the XML returned from the most recent query or stored procedure called.

ParamCount	Long
	- this must be set in order to use parameters with the stored procedures.

AddParam
	- adds a parameter to the current stored procedures parameter list. Several different Enumerated constants
	  are used but they can be passed as straight integers.
	- all parameters (output parameters included), must be added before the stored procedure is called

UseReturnParameter	True/False	(Default = True)
	- when calling stored procedures, the object can check the @RETURN_VALUE parameter for your stored
	  procedure (the call will fail if the return parameter is not zero).  If you have no need to use 
	  return parameters, set this to false so that an extra parameter need not be added to the Command Object

ReturnVal	Variant
	- when passed a return parameter name, it will return the parameter from the stored procedure command
	  object.

ReturnParamterValue	Long
	- if @RETURN_VALUE parameter is used, it will return the @RETURN_VALUE of the last stored procedure called

RootNode	String
	- this is the root node to wrap around the xml that is returned by SQL Server 2000. It can be set using
	  'RootName' or '<RootName>'. It will automatically make the closing tag for the root node to avoid error.
	  It can also be set to zero length string for no root node.





NOTES
	- all tables & stored procedures used will need the permissions set in SQL Server 2000
		- usually is something like SERVERNAME\IUSR_SERVERNAME
	- you need to have ADO 2.6 installed to use the xml features of SQL Server 2000