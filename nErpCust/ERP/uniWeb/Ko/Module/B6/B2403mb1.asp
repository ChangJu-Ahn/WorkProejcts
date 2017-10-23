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
	
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%

    Dim LngMaxRow		
	Dim LngRow
	Dim GroupCount
	Dim lgstrData
	Dim strdata	
	Dim strSpread
	
	const B457_EG1_E1_orgid = 0
	const B457_EG1_E1_orgnm = 1
	const B457_EG1_E1_orgdt = 2
	const B457_EG1_E1_remarks = 3
	const B457_EG1_E1_currentyn = 4
	const C_Temp = 5
	    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call LoadBasisGlobalInf()

    Call HideStatusWnd                                                               '☜: Hide Processing message
    
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    strSpread         = Request("txtSpread")  
            
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)			                                                         '☜: Save,Update
             Call SubBizSaveMulti()
    End Select

Sub SubBizQueryMulti()	

	On Error Resume Next 
	Dim PB6G051		
	Dim import_org_cd
		

    Err.Clear
  
    import_org_cd = Request("txtOrgid") 

    Set PB6G051 = server.CreateObject ("PB6G051.cBListHorgAbs")
    If CheckSYSTEMError(Err,True) = True Then
        set PB6G051 = nothing
        Response.End  
    End If	
	on error goto 0
    
    On Error Resume Next    
    lgstrData = PB6G051.B_READ_HORG_ABS(gStrGlobalCollection,import_org_cd)
    If CheckSYSTEMError(Err,True) = True Then
        set PB6G051 = nothing
        Response.End  
    End If	
	on error goto 0
     
    set PB6G051 = nothing

End Sub

%>
<Script Language=vbscript>

	Dim LngLastRow      
	Dim LngMaxRow       
	Dim LngRow          
	Dim strTemp
	Dim strData
	
	With parent		
	
<%      
	  GroupCount = Ubound(lgstrData)
	  For LngRow = 0 To GroupCount	     
	    
%>  
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(lgstrData(LngRow,B457_EG1_E1_orgid)))%>"	
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(lgstrData(LngRow,B457_EG1_E1_orgnm)))%>"	
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(lgstrData(LngRow,B457_EG1_E1_orgdt))%>"		
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(lgstrData(LngRow,B457_EG1_E1_remarks)))%>"
<%		
        If lgstrData(LngRow,B457_EG1_E1_currentyn) = "Y" Then
%>
		strData = strData & Chr(11) & "1" 
<%
        Else
%>
		strData = strData & Chr(11) & "0" 
<%
        End If
%>
        strData = strData & Chr(11) & "<%=ConvSPChars(Trim(lgstrData(LngRow,C_Temp)))%>"
		strData = strData & Chr(11) & <%=LngRow%>  + 1               'LngMaxRow + <%=LngRow%>
		strData = strData & Chr(11) & Chr(12)
		
<%      
    Next    
%>    
	End With
</Script>	
<%  

Sub SubBizSaveMulti()
    
    On Error Resume Next 
    Dim PB6G051 
            
    Set  PB6G051 = server.CreateObject("PB6G051.cBControlHorgAbs")  
    If CheckSYSTEMError(Err,True) = True Then
        set PB6G051 = nothing
        Response.End  
    End If	
	on error goto 0
                                                             
    On Error Resume Next  
    call PB6G051.B_CONTROL_HORG_ABS(gStrGlobalCollection,strSpread)
    If CheckSYSTEMError(Err,True) = True Then
        set PB6G051 = nothing
        Response.End  
    End If	
	on error goto 0

    set PB6G051 = nothing

End Sub
%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          
            With Parent
                .ggoSpread.Source     = .frm1.vspdData               
                .ggoSpread.SSShowData  strData 
                .DBQueryOk
	         End with
         
       Case "<%=UID_M0002%>"                                                        '☜ : Save         
             Parent.DBSaveOk
    End Select    
    
       
</Script>	
