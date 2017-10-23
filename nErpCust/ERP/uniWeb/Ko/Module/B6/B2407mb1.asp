<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(horg_abs 부서개편개요)
'*  3. Program ID           : B2403mb1.asp
'*  4. Program Name         : B2403mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B24031ControlHorgAbs
'                             +B24038ListHorgAbs
'*  7. Modified date(First) : 2000/10/25
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
Option Explicit		
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
    Dim strSpread

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call LoadBasisGlobalInf()

    Call HideStatusWnd                                                               '☜: Hide Processing message
    
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)       
    strSpread         = Request("txtSpread")
	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
    End Select

Sub SubBizQueryMulti()
 	On Error Resume Next
 	
	Dim PB6G063	
	Dim import_horg_mas_orgid                                                               '☜: Protect system from crashing
                 
    import_horg_mas_orgid = request("txtOrgid") 

    Set PB6G063 = server.CreateObject ("PB6G063.cBListHorgMasExtra")

    If CheckSYSTEMError(Err,True) = True Then
        Response.End  
    End If	
	
    lgstrData = PB6G063.B_READ_HORG_MAS_EXTRA(gStrGlobalCollection,import_horg_mas_orgid)	
    
    If CheckSYSTEMError(Err,True) = True Then
		Set PB6G063 = Nothing
    End If
	
	Set PB6G061 = Nothing
End Sub

Sub SubBizSaveMulti()
	On Error Resume Next
    Dim PB6G063

    Set PB6G063 = server.CreateObject("PB6G063.cBControlHorgMasExtra")

    If CheckSYSTEMError(Err,True) = True Then
        Response.End  
    End If
    
	Call PB6G063.B_CONTROL_HORG_MAS_EXTRA(gStrGlobalCollection,strSpread)
	
    If CheckSYSTEMError(Err,True) = True Then
        Set PB6G063 = Nothing
        Response.End  
    End If	
   
	Set  PB6G061 = Nothing
	
End Sub
%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"       
             With Parent
                .ggoSpread.Source  = .frm1.vspdData                
                .ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"
                .DBQueryOk
	         End with       
       Case "<%=UID_M0002%>"      
             Parent.DBSaveOk                 
    End Select    
    
       
</Script>	
