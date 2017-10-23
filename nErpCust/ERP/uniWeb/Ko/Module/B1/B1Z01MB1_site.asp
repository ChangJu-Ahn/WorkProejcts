<%@LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
           
<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf

    Dim PB1Z001_KO346
    
    Dim txtQueryCd
    Dim txtQueryNm
    Dim txtDept_cd
    Dim cboRole_type
    Dim txtSelect
    Dim txtFrom
    Dim txtWhere
    Dim txtEtc
    Dim txtRemark
    Dim RtnQueryID
    Dim iCommandSent
	Dim E_PrevNext_Code
    
    
    Call HideStatusWnd        
    
    lgOpModeCRUD = UCASE(TRIM(Request("txtMode")))
    iCommandSent = Request("txtCommand")
    
    Response.Write lgOpModeCRUD
 
    
    txtQueryCd				= Request("txtQueryCd")
    gDepart					= Request("gDepart")    
    txtQueryNm				= UCASE(Request.Form("txtQueryNm"))
    txtDept_cd				= Request.Form("txtDept_cd")
    cboRole_type			= Request.Form("cboRole_type")
    txtSelect				= Request.Form("txtSelect")
    txtFrom					= Request.Form("txtFrom")
    txtWhere				= Request.Form("txtWhere")
    txtEtc					= Request.Form("txtEtc")
    txtRemark				= Request.Form("txtRemark")
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

    lgErrorStatus = "NO"
  
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

	IF lgOpModeCRUD = "UID_M0001" THEN 'CREATE
		Call SubBizSave()
	ELSEIF lgOpModeCRUD = "UID_M0002" THEN 'UPDATE
		Call SubUpdateSave()
	ELSEIF lgOpModeCRUD = "UID_M0003" THEN 'SELECT
		Call SubBizQuery()
	ELSEIF lgOpModeCRUD = "UID_M0004" THEN 'DELETE
		Call SubDeleteBiz()
	END IF

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection


Sub SubDeleteBiz()

    Dim lgStrSQL1
    Dim lgStrSQL2
    Dim intRetVal
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    
	' 해당 Business Object 생성 
	Set PB1Z001_KO346 = Server.CreateObject("PB1Z001_KO346.clsQueryCommand")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If
	
	Call PB1Z001_KO346.B_DELETE_QUERY_COMMAND(gStrGlobalCollection, _
										txtQueryCd)
                              
	If CheckSYSTEMError(Err,True) = True Then
		Set PB1Z001_KO346 = Nothing
		Response.End
	End If
	
	
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Select Data
'============================================================================================================
Sub SubBizQuery()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    ' 해당 Business Object 생성 
	Set PB1Z001_KO346 = Server.CreateObject("PB1Z001_KO346.clsQueryCommand")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If
	
	Call PB1Z001_KO346.B_LIST_QUERY_COMMAND(gStrGlobalCollection, _
				 						  txtQueryCd, _
				 						  gDepart, _
				 						  iCommandSent, _
				 						  E_PrevNext_Code, _
										  EG1_group_export)
										  
	If CheckSYSTEMError(Err,True) = True Then
		Set PB1Z001_KO346 = Nothing
		Response.End
	End If
	
	If Not isEmpty(E_PrevNext_Code) Then
		If Trim(E_PrevNext_Code(0)) = "900011" Or Trim(E_PrevNext_Code(0)) = "900012" Then
			Call DisplayMsgBox(E_PrevNext_Code(0), VbOKOnly, "", "", I_MKSCRIPT)
		End If
	End If
		

%>
<Script Language=vbscript>

With parent																	'☜: 화면 처리 ASP 를 지칭함

    .Frm1.txtQueryCd.Value			= "<%=ConvSPChars(EG1_group_export(0,0))%>"                   'Set condition area
    .Frm1.txtQueryNm.Value			= "<%=ConvSPChars(EG1_group_export(0,1))%>" 
    .Frm1.txtDept_cd.Value			= "<%=ConvSPChars(EG1_group_export(0,2))%>"
    .Frm1.txtDept_nm.Value			= "<%=ConvSPChars(EG1_group_export(0,3))%>"
    .Frm1.cboRole_type.Value		= "<%=ConvSPChars(EG1_group_export(0,4))%>"
            
            
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
	
	ReDim TmpBuffer(0)	

	strData = ""
	strData = strData & Chr(11) & "<%=ConvSPChars(replace(EG1_group_export(0,5),Chr(13) &Chr(10),chr(7)))%>" 
	strData = strData & Chr(11) & "<%=ConvSPChars(replace(EG1_group_export(0,6),Chr(13) &Chr(10),chr(7)))%>"
	strData = strData & Chr(11) & "<%=ConvSPChars(replace(EG1_group_export(0,7),Chr(13) &Chr(10),chr(7)))%>" 
	strData = strData & Chr(11) & "<%=ConvSPChars(replace(EG1_group_export(0,8),Chr(13) &Chr(10),chr(7)))%>" 
	strData = strData & Chr(11) & "<%=ConvSPChars(replace(EG1_group_export(0,9),Chr(13) &Chr(10),chr(7)))%>"  
	strData = strData & Chr(11) & LngMaxRow
	strData = strData & Chr(11) & Chr(12)
			
	TmpBuffer(0) = strData

	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .frm1.vspdData1
	.ggoSpread.SSShowDataByClip iTotalStr
		
	.DbQueryOk1
	
End With	

            
</Script>

<%	

	
End Sub


'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    Dim lgStrSQL1
    Dim lgStrSQL2
    Dim intRetVal
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    ' 해당 Business Object 생성 
	Set PB1Z001_KO346 = Server.CreateObject("PB1Z001_KO346.clsQueryCommand")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If
	
'	Call DisplayMsgBox("x", vbInformation, gStrGlobalCollection, "FASDFADS1111", I_MKSCRIPT)
	
	
	RtnQueryID =  PB1Z001_KO346.B_CREATE_QUERY_COMMAND(gStrGlobalCollection, _
										txtQueryCd, _
										txtQueryNm, _
										txtDept_cd, _
										cboRole_type, _
										txtSelect, _
										txtFrom, _
										txtWhere, _
										txtEtc, _
										txtRemark, _
										iExportDisposition)

	'Call DisplayMsgBox("x", vbInformation, RtnQueryID & "===", "FASDFADS1111", I_MKSCRIPT)                              
	
	If CheckSYSTEMError(Err,True) = True Then
		Set PB1Z001_KO346 = Nothing
		Response.End
	End If
	
End Sub
	    
'============================================================================================================
' Name : SubUpdateSave
' Desc : Query Data from Db
'============================================================================================================
Sub SubUpdateSave()

    Dim lgStrSQL1
    Dim lgStrSQL2
    Dim intRetVal
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    
	' 해당 Business Object 생성 
	Set PB1Z001_KO346 = Server.CreateObject("PB1Z001_KO346.clsQueryCommand")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If
	
	
	Call PB1Z001_KO346.B_UPDATE_QUERY_COMMAND(gStrGlobalCollection, _
										txtQueryCd, _
										txtQueryNm, _
										txtDept_cd, _
										cboRole_type, _
										txtSelect, _
										txtFrom, _
										txtWhere, _
										txtEtc, _										
										txtRemark, _
										iExportDisposition)
										
	RtnQueryID =     txtQueryCd                

	If CheckSYSTEMError(Err,True) = True Then
		Set PB1Z001_KO346 = Nothing
		Response.End
	End If

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode, pCode)
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode, pConn, pRs, pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "U"
			If CheckSYSTEMError(pErr,True) = True Then
			   Call DisplayMsgBox("122918", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
			   ObjectContext.SetAbort
			   Call SetErrorStatus
			Else
			   If CheckSQLError(pConn,True) = True Then
			      Call DisplayMsgBox("122918", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
			      ObjectContext.SetAbort
			      Call SetErrorStatus
			   End If
			End If
    End Select
End Sub

IF lgOpModeCRUD = "UID_M0004" THEN
%>
	<Script Language="VBScript">
	    If Trim("<%=lgErrorStatus%>") = "NO" Then
	       Parent.FncNew
	    End If
	</Script>	
<%
ELSEIF lgOpModeCRUD <> "UID_M0003" THEN
%>
	<Script Language="VBScript">
	    If Trim("<%=lgErrorStatus%>") = "NO" Then
		   Parent.frm1.txtQueryCd.Value  = "<%=ConvSPChars(RtnQueryID)%>"
	       Parent.DBSaveOk
	    End If
	</Script>	
<%
END IF
%>


