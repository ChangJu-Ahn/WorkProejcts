<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4231mb1
'*  4. Program Name         : 이자율변경등록 
'*  5. Program Desc         : Register of Loan Change
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002-04-02
'*  8. Modified date(Last)  : 2002-07-12
'*  9. Modifier (First)     : Oh, Soo Min
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")   

Call HideStatusWnd
On Error Resume Next

Dim txtLoanNo
Dim lgOpModeCRUD
Dim lgErrorStatus,lgErrorPos

lgErrorStatus	= "NO"
lgErrorPos		= ""                                                           '☜: Set to space
lgOpModeCRUD	= Request("txtMode")					'☜ : 현재 상태를 받음 
txtLoanNo		= Trim(Request("txtLoanNo"))

Const C_SHEETMAXROWS_D = 100

Select Case lgOpModeCRUD
     Case CStr(UID_M0001)                                                         '☜: Query
          Call SubBizQuery()

     Case CStr(UID_M0002)       
          Call SubBizSave()

     Case CStr(UID_M0003)                                                         '☜: Delete
          Call SubBizDelete()
End Select
   
Response.End 


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	Dim PAFG455LIST	
	Dim E1_f_ln_info, EG1_export_group	
	Dim iLngRow,iLngCol
	Dim iIntLoopCount
	Dim iStrData
	Dim iStrPrevKey
	Dim iIntMaxRows
	Dim iIntQueryCount

	Const EA_f_ln_info_loan_no1 = 0
    Const EA_f_ln_info_loan_nm1 = 1
    Const EA_f_ln_info_loan_fg1 = 2
    Const EA_f_ln_info_loan_dt1 = 3
    Const EA_f_ln_info_doc_cur1 = 4
    Const EA_f_ln_info_loan_amt1 = 5
    Const EA_f_ln_info_due_dt1 = 6
    Const EA_f_ln_info_loan_int_rate1 = 7

    Const EA_f_ln_his_seq2 = 0
    Const EA_f_ln_his_chg_dt2 = 1
    Const EA_f_ln_his_due_dt2 = 2
    Const EA_f_ln_his_int_rate2 = 3
    Const EA_f_ln_his_chg_type2 = 4
    Const EA_f_ln_his_chg_amt2 = 5
    Const EA_f_ln_his_gl_no2 = 6
    Const EA_f_ln_his_temp_gl_no2 = 7
    Const EA_f_ln_his_his_desc2 = 8

    iStrPrevKey		= Trim(Request("lgStrPrevKey"))        
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")

    If Len(Trim(iIntQueryCount))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)
       End If
    Else   
       iIntQueryCount = 0
    End If

	Set PAFG455LIST = Server.CreateObject("PAFG455.bFListLnHis")

	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If

	Call PAFG455LIST.F_LIST_LN_HIS_SVR(gStrGloBalCollection, iStrPrevKey, C_SHEETMAXROWS_D, txtLoanNo, E1_f_ln_info, EG1_export_group)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG455LIST = nothing		
		Exit Sub
    End If
    Set PAFG455LIST	 = nothing

	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write " With parent.frm1		" & vbCr
	Response.Write " .txtLoanNo.value  = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_loan_no1)) & """			" & vbCr
	Response.Write " .txtLoanNm.value  = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_loan_nm1)) & """			" & vbCr
	Response.Write " .txtLoanDt.value  = """ & UNIDateClientFormat(E1_f_ln_info(EA_f_ln_info_loan_dt1)) & """	" & vbCr
	Response.Write " .txtDueDt.value   = """ & UNIDateClientFormat(E1_f_ln_info(EA_f_ln_info_due_dt1)) & """		" & vbCr
	Response.Write " .txtLoanAmt.text = """ & UNINumClientFormat(E1_f_ln_info(EA_f_ln_info_loan_amt1), ggAmtOfMoney.DecPoint, 0)    & """" & vbCr

       'Response.Write " .txtIntRate.text = """ & UNINumClientFormat(E1_f_ln_info(EA_f_ln_info_loan_int_rate1), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write " .txtIntRate.text = """ & UNINumClientFormat(E1_f_ln_info(EA_f_ln_info_loan_int_rate1), "6", 0)  & """	" & vbCr

	Response.Write " .htxtLoanNo.value = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_loan_no1)) & """			" & vbCr
    Response.Write " End with				" & vbcr
    Response.Write "Parent.DbQueryOk		" & vbcr
    Response.Write "</Script>               " & vbCr

	iStrData = ""

	If IsEmpty(EG1_export_group) = False Then
		For iLngRow = 0 To UBound(EG1_export_group, 1)
			iIntLoopCount = iIntLoopCount + 1
			If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
					iStrData = iStrData & Chr(11) & Trim(EG1_export_group(iLngRow, EA_f_ln_his_seq2))
					iStrData = iStrData & Chr(11) & UNIDateClientFormat(Trim(EG1_export_group(iLngRow, EA_f_ln_his_chg_dt2)))
					iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, EA_f_ln_his_int_rate2), ggExchRate.DecPoint, 0)
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, EA_f_ln_his_his_desc2)))
					iStrData = iStrData & Chr(11) & Cstr(iIntMaxRows + iLngRow + 1) 
					iStrData = iStrData & Chr(11) & Chr(12)
			Else
				iStrPrevKey = EG1_export_group(UBound(EG1_export_group, 1), EA_f_ln_his_chg_dt2)
				iIntQueryCount = iIntQueryCount + 1
				Exit For
			End If
		Next
		
		If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
			iStrPrevKey = ""
		    iIntQueryCount = ""
		End If

	End If

	Response.Write " <Script Language=vbscript>								 " & vbCr
	Response.Write " With parent											 " & vbCr
    Response.Write "	.ggoSpread.Source		= .frm1.vspdData			 " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData	  """ & iStrData		& """" & vbCr
    Response.Write "	.lgPageNo				= """ & iIntQueryCount	& """" & vbCr
    Response.Write "	.lgStrPrevKey			= """ & iStrPrevKey		& """" & vbCr
    Response.Write "	.DbQueryOk											 " & vbCr
    Response.Write "End With												 " & vbCr
    Response.Write "</Script>												 " & vbCr 
    
    
End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status
    
	Dim PAFG455CU
	Dim iarrData
	Dim I1_f_ln_info
	Dim txtSpread
	
	iarrData = Split(request("txtSpread"), gRowSep)
	txtLoanNo = Request("htxtLoanNo")

    Const A659_L3_seq = 0
    Const A659_L3_chg_dt = 1
    Const A659_L3_int_rate = 2
    Const A659_L3_chg_type = 3
    Const A659_L3_his_desc = 4

    Set PAFG455CU = server.CreateObject ("PAFG455.cFMntFLnHisSvr")   

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
     
	Call PAFG455CU.F_MAINT_F_LN_HIS_SVR(gStrGlobalCollection, iarrData, txtLoanNo)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG455CU = nothing
		Exit Sub	
    End If
	 
    Set PAFG455CU = nothing
    
	Response.Write "<Script Language=vbscript>		" & vbCr
	Response.Write " parent.DbSaveOk()				" & vbCr
    Response.Write "</Script>						" & vbCr  

End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status    
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub


%>
