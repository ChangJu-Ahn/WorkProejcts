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
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q4113MB3
'*  4. Program Name         : 수입검사불합격통지 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next
Call HideStatusWnd 

Dim PQIG340
					
Set PQIG340 = Server.CreateObject("PQIG340.cQMtRejReportSimple")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG340.Q_MAINT_REJECT_REPORT_SIMPLE_SVR(gStrGlobalCollection, _
									"DELETE", _
									Request("txtPlantcd"), _
									Request("txtInspReqNo"), _
									1)
	    
If CheckSYSTEMError(Err,True) = True Then
	Set PQIG340 = Nothing
	Response.End
End If		    
			              
Set PQIG340 = Nothing   
                            
%>
<Script Language=vbscript>
	Call parent.DbDeleteOk()
</Script>
