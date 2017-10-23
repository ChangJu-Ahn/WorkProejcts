<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(부서정보등록)
'*  3. Program ID           : B2406ma1
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Complus List         : 
'                             
'*  7. Modified date(First) : 2005/10/19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
    On Error Resume Next																	'☜: Protect system from crashing
    Err.Clear																				'☜: Clear Error status

    Call LoadBasisGlobalInf()
    Call HideStatusWnd             

    Dim PB6G061
    Dim strSpread
    Dim strOrgId
    
    strOrgId = Request("txtOrgId")
    strSpread = Request("txtSpread")     
	
	Set PB6G061 = server.CreateObject("PB6G061.cBControlHorgMas")

    If CheckSYSTEMError(Err,True) = True Then
        Response.End  
    End If	
    
	Call PB6G061.B_MAKE_SEQ_ALL_ZERO(gStrGlobalCollection,strOrgId,strSpread)
	
    If CheckSYSTEMError(Err,True) = True Then
        Set PB6G061 = Nothing
        Response.End  
    End If	

	Set  PB6G061 = Nothing

%>

<Script Language="VBScript">
    With Parent

       .DBSaveOk
	End With       
</Script>	
