<% 
'*************************************************************************************
'*  1. Module Name				: BA
'*  2. Function Name			: XMLHTTP Request
'*  3. Program ID				: GetPuniSvrConn
'*  4. Program Name				: GetPuniSvrConn
'*  5. Modified date(First)		: 2003.01.1
'*  6. Modified date(Last)		: 
'*  7. Modifier (First)			: KO giyeon
'*  8. Modifier (Last)			: 
'*  9. Comment					: 
'*  10.Version Label			: PS.2003.01.1
'*  11.Compatible PuniSvrConn   : 2,0,0,0 (PS.2003.01.1)
'*************************************************************************************
on error resume next
err.Clear

dim obj 
dim str
Dim ExecuteSql 
Dim Lang
Dim iPath 
iPath = Request.ServerVariables("PATH_INFO")
iPath = Split(iPath,"/") 
Lang = UCase(iPath(UBound(iPath) - 2))


set obj = CreateObject("PuniSvrConn.CuniSvrConn")

strC = Request("Cmd")
strF = Request("Flag")
strCompany = Request("Company")
strClientID = Request("ClientID")
strConnString = Request("ConnString")
strSQL = Request("SQL")

if (strC ="R") then
	If (strF = "S") Then	
		ExecuteSvrDLL = obj.WriteLoginStatus(strCompany, strClientID)
	End If
	If (strF = "E") Then	
		ExecuteSvrDLL = obj.DeleteLoginStatus(strCompany, strClientID)
	end if 	
	Response.Write ExecuteSvrDLL	
else
	ExecuteSql = obj.XMLHTTPConnectDB(strConnString, strF, strSQL,strC,Lang)			
	Response.ContentType = "text/xml"	
	Response.Write ExecuteSql	
end if
if err.number <> 0 then 
	Response.Write err.Description
end if

set obj = nothing
%>
