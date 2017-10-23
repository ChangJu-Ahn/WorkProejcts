<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->


<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          :  
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment
'* 12. Common Coding Guide  : this mark(¢Ð) means that "Do not change" 
'*                            this mark(¢Á) Means that "may  change"
'*                            this mark(¡Ù) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next
Call HideStatusWnd 

Dim PY2G101
					
Set PY2G101 = Server.CreateObject("PY2G101.cCisNewItemReq")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
   
Call PY2G101.Y_MAINT_NEW_ITEM_REQ_SVR(gStrGlobalCollection, "DELETE", Request("txtReqNo") )
	    
If CheckSYSTEMError(Err,True) = True Then
	Set PY2G101 = Nothing
	Response.End
End If		    
			              
Set PY2G101 = Nothing   
                            
%>
<Script Language=vbscript>
	Call parent.DbDeleteOk()
</Script>
