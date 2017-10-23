<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Basic Info.
'*  3. Program ID           : A2109MB1
'*  4. Program Name         : 신용카드등록 
'*  5. Program Desc         : Register of Credit Card
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2002/11/26
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Chang Joo, Kim
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->

<%
	Call LoadBasisGlobalInf()

    Dim lgOpModeCRUD
    On Error Resume Next
    Err.Clear

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
          '  Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
          '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iPABG045
    Dim import_b_credit_card
    Dim export_b_credit_card
    Dim export_group
    Dim iStrPrevKey
    Dim iIntMaxRows
    Dim iIntQueryCount
	Dim itxtCredit_No

    ReDim export_b_credit_card(1)

    Dim iStrData
    Dim iIntLoopCount
    Dim iLngRow,iLngCol

    Const C_SHEETMAXROWS_D  = 100

	Const credit_no		= 0
	Const credit_nm		= 1
	Const credit_eng_nm = 2
	CONST EX_CARD_CO_CD = 3
	CONST EX_CARD_CO_NM = 4
	Const credit_type	= 5
	Const cost_cd		= 6
	Const cost_nm		= 7
	Const rgst_no		= 8
	Const expire_dt		= 9
	Const sttl_dt		= 10
	Const use_user_id	= 11
	Const bank_cd		= 12
	Const bank_nm		= 13
	Const bank_acct_no	= 14

    Const A033_E1_credit_no = 0
    Const A033_E1_credit_nm = 1

    On Error Resume Next
    Err.Clear

	iStrPrevKey		= Request("lgStrPrevKey")
	itxtCredit_No	= Request("txtCredit_No")
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")

    If Trim(iStrPrevKey) = "" Then
		import_b_credit_card	= itxtCredit_No
	Else
		import_b_credit_card	= iStrPrevKey
    End If

	Set iPABG045 = Server.CreateObject("PABG045.cBListCreditCardSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If    

	Call iPABG045.B_LIST_CREDIT_CARD_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, import_b_credit_card, itxtCredit_No, export_b_credit_card, export_group)

    If CheckSYSTEMError(Err, True) = True Then
       Set iPABG045 = Nothing
       Exit Sub
    End If

    Set iPABG045 = Nothing

	iStrData = ""

	iIntLoopCount = 0

	For iLngRow = 0 To UBound(export_group, 1)
		iIntLoopCount = iIntLoopCount + 1

	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, credit_no)))			'1
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, credit_nm)))			'2
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, credit_eng_nm)))			'2
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, EX_CARD_CO_CD)))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, EX_CARD_CO_NM)))
				iStrData = iStrData & Chr(11) & Trim(export_group(iLngRow, credit_type))						'3
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, cost_cd)))			'4
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, cost_nm)))			'5
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, rgst_no)))			'6
				iStrData = iStrData & Chr(11) & UNIDateClientFormat(Trim(export_group(iLngRow, expire_dt)))	'8
				iStrData = iStrData & Chr(11) & Trim(export_group(iLngRow, sttl_dt))						'9
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, use_user_id)))			'10
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, bank_cd)))		'11
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, bank_nm)))		'12
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, bank_acct_no))) 		'13
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & iIntMaxRows + iIntLoopCount                     '11
			    iStrData = iStrData & Chr(11) & Chr(12)
	    Else
			iStrPrevKey = export_group(UBound(export_group, 1), 0)
			iIntQueryCount = iIntQueryCount + 1
			Exit For

		End If
	Next

	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = "" 
	End If

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
		Response.Write "	.frm1.txtCredit_Nm.value  = """ & ConvSPChars(export_b_credit_card(A033_E1_credit_nm))    & """" & vbCr
		Response.Write "	.frm1.htxtCredit_No.value = """ & ConvSPChars(export_b_credit_card(A033_E1_credit_no))    & """" & vbCr
		Response.Write "	.lgPageNo = """ & iIntQueryCount    & """" & vbCr
		Response.Write "	.lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)    & """" & vbCr
		Response.Write "	.DbQueryOk " & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
   
Sub SubBizSaveMulti()
    On Error Resume Next
    Err.Clear
    Dim iPABG045

'	Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
'	Dim	lGrpCnt																		'☜: Group Count
	Dim LngMaxRow
	Dim LngRow
'	Dim strStatus
    Dim iErrorPosition

    Set iPABG045 = Server.CreateObject("PABG045.cBMngCreditCardSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set iPABG045 = Nothing
       Exit Sub
    End If

    LngMaxRow  = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수    


    Call iPABG045.B_MANAGE_CREDIT_CARD_SVR(gStrGloBalCollection, Request("txtSpread"), iErrorPosition)
   	If CheckSYSTEMError2(Err, True,iErrorPosition & "행","","","","") = True Then
       Set iPABG045 = Nothing
       Exit Sub
    End If

    Set iPABG045 = Nothing

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub

'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next 
    Err.Clear
End Sub

'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next 
    Err.Clear
End Sub

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next 
    Err.Clear
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next 
    Err.Clear
End Sub

'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()
End Sub

'============================================================================================================
Sub SetErrorStatus()
End Sub
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next 
    Err.Clear
End Sub
%>
