<% Option Explicit %>
<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 회계관리 
'*  3. Program ID           : a5130ma1
'*  4. Program Name         : 전표관리번호일괄채번 
'*  5. Program Desc         : 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2006/11/15
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================

Response.Buffer = True												'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->

<%													

On Error Resume Next
Err.Clear 

Call LoadBasisGlobalInf() 
Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""  
    lgOpModeCRUD      = Request("txtMode")												'☜: Read Operation Mode (CRUD)   

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             'Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizSave
' Desc : 
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next
    Err.Clear    
	
	Const A331_EG1_yyyymm = 0
	Const A331_EG1_status = 1
	Const A331_EG1_status_nm = 2
	Const A331_EG1_auto_no_type = 3
	Const A331_EG1_date_info = 4
	Const A331_EG1_auto_no = 5
	Const A331_EG1_working_dt = 6
	Const A331_EG1_working_id = 7
	
	Dim iPAGG150
	Dim iStrData
	Dim iYear
	Dim iExportData
    Dim iLngRow

    Set iPAGG150 = Server.CreateObject("PAGG150.cALkupGLMgmtNoHisSvr")

    If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If

	iYear = Trim(Request("txtYear"))
    
	iExportData = iPAGG150.A_LKUP_GL_MGMT_NO_HISTORY_SVR(gStrGlobalCollection, iYear)
	
	If CheckSYSTEMError(Err, True) = True Then					
		Set iPAGG150 = Nothing        
		Exit Sub
    End If    

    Set iPAGG150 = Nothing

    iStrData = ""	
	For iLngRow = 0 To UBound(iExportData, 1) 
		iStrData = iStrData & Chr(11) & ConvSPChars(iExportData(iLngRow, A331_EG1_yyyymm))
		iStrData = iStrData & Chr(11) & ConvSPChars(iExportData(iLngRow, A331_EG1_status))
		iStrData = iStrData & Chr(11) & ConvSPChars(iExportData(iLngRow, A331_EG1_status_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(iExportData(iLngRow, A331_EG1_auto_no_type))
		iStrData = iStrData & Chr(11) & ConvSPChars(Trim(iExportData(iLngRow, A331_EG1_date_info)))	
		iStrData = iStrData & Chr(11) & ConvSPChars(iExportData(iLngRow, A331_EG1_auto_no))
		iStrData = iStrData & Chr(11) & UNIDateClientFormat(iExportData(iLngRow, A331_EG1_working_dt))
		iStrData = iStrData & Chr(11) & ConvSPChars(Trim(iExportData(iLngRow, A331_EG1_working_id)))
		iStrData = iStrData & Chr(11) & iLngRow + 1
		iStrData = iStrData & Chr(11) & Chr(12)		
	Next

	Response.Write "<Script Language=vbscript>						" & vbcr
	Response.Write " With Parent									" & vbCr
	Response.Write " 	.ggoSpread.Source = .frm1.vspdData	        " & vbCr
	Response.Write "	.ggoSpread.SSShowData """ & iStrData   & """" & vbCr
	Response.Write " 	.DbQueryOk								    " & vbCr
	Response.Write " End With										" & vbCr
	Response.Write "</Script>										" & vbCr
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next	
	Err.Clear

    Dim iPAGG150		
    Dim I1_auto_no_type
    Dim txtSpread 
    Dim iErrorPosition

    Set iPAGG150 = Server.CreateObject("PAGG150.cAExecGLMgmtNoSvr")

    If CheckSYSTEMError(Err, True) = True Then					
		Response.End 
    End If    
	
	I1_auto_no_type = "JM"
	txtSpread = Trim(Request("txtSpread"))
	
    Call iPAGG150.A_EXEC_GL_MGMT_NO_SVR(gStrGlobalCollection, I1_auto_no_type , txtSpread)						

    If CheckSYSTEMError(Err, True) = True Then					
		Set iPAGG150 = Nothing
		Response.End 
    End If    

    Set iPAGG150 = Nothing

	Response.Write " <Script Language=vbscript>	" & vbCr
	Response.Write " With parent				" & vbCr
    Response.Write "	.DbSaveOk  				" & vbCr    
    Response.Write " End With					" & vbCr
    Response.Write " </Script>					" & vbCr
End Sub    
%>

