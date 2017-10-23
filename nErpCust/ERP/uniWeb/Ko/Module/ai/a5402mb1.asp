<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<%
    'On Error Resume Next                                                   '☜: Protect prorgram from crashing
    Err.Clear                                                              '☜: Clear Error status
    
%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- Include file="../../inc/adovbs.inc" -->

<%

    Dim lgStrPrevKey

	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB") 	
	Call HideStatusWnd
    lgErrorStatus  = ""
    lgKeyStream    = Split(Request("txtKeyStream"),gColSep) 
    lgStrPrevKey   = Trim(Request("lgStrPrevKey"))                   '☜: Next Key

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	Dim strGlNo, strGlSeq
	Const C_GL_NO = 1
	Const C_GL_SEQ = 2
	
    Dim FromDate, ToDate
    Dim strAccountCd, strAccountNm, strOpenAcctFg, strCurrency
    Dim strFrdate, FrYear, FrMonth, FrDay
    Dim strTodate, ToYear, ToMonth, ToDay
	
	FromDate      = UniConvDate(Request("txtDateFr"))
	ToDate        = UniConvDate(Request("txtDateTo"))
	strAccountCd  = UCase(Trim(Request("txtAccountCd")))
	strOpenAcctFg = UCase(Trim(Request("txtOpenAcctFg")))
	strCurrency   = UCase(Trim(Request("txtCurrency")))
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
	Select Case CStr(Request("txtMode"))
		Case CStr(UID_M0001)                                                         '☜: Query
			Call SubBizQuery()
			Call SubBizQueryMulti()
		Case CStr(UID_M0002)                                                         '☜: Save,Update
		    Call SubBizSaveMulti()
		Case CStr(UID_M0003)                                                         '☜: Delete

			Call ExtractDateFrom(FromDate,	gServerDateFormat,gServerDateType,	FrYear, FrMonth, FrDay)
			Call ExtractDateFrom(ToDate,	gServerDateFormat,gServerDateType,	ToYear,  ToMonth, ToDay)
			
			strFrdate = FrYear & FrMonth & FrDay
			strTodate = ToYear & ToMonth & ToDay

			If strOpenAcctFg = "Y" Then
				Call execClosingCancelAll()
			Else
				Call execClosingAll()
			End If
	End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
    Dim lgStrSQL
    
	'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear  
                                                                          '☜: Clear Error status
	strAccountNm	= ""                                                                         

    If lgStrPrevKey = "" Then

		If strAccountCd <> "" Then
			lgStrSQL = " Select Acct_Nm from A_Acct Where Acct_cd = " & Filtervar(strAccountCd, "''", "S")

			If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                   'R(Read) X(CursorType) X(LockType) 
				If isNull(lgObjRs("Acct_Nm")) = False Then
					strAccountNm = Trim(lgObjRs("Acct_Nm"))
				End If
			End if
			Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
			Call SubCloseRs(lgObjRs)  
		End If		    

		Response.Write  " <Script Language=vbscript> " & vbCr
		Response.Write  " With Parent                " & vbCr
		Response.Write  "    .frm1.txtAccountNm.value =  """ & strAccountNm    & """" & vbCr
		Response.Write  " End With                   " & vbCr        
		Response.Write  " </Script>                  " & vbCr
    End If
    
End Sub	


'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iSelCount
    Dim lgStrSQL
    Dim lgstrData
    Dim lgLngMaxRow
    Dim iDx
    
    Const C_SHEETMAXROWS_D  = 30                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 

    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------


	lgStrSQL = ""
	If strOpenAcctFg = "Y" Then
		lgStrSQL = lgStrSQL & " SELECT " & FilterVar("0", "''", "S") & "  CHOICE, A.GL_NO, A.GL_SEQ, A.GL_DT, A.OPEN_DOC_AMT AMT, A.OPEN_AMT LOC_AMT, A.GL_DESC " 
		lgStrSQL = lgStrSQL & " FROM A_OPEN_ACCT A "
		lgStrSQL = lgStrSQL & " WHERE BASIS_FG = " & FilterVar("Y", "''", "S") & "  "
	Else
		lgStrSQL = lgStrSQL & " SELECT " & FilterVar("0", "''", "S") & "  CHOICE, A.GL_NO, A.ITEM_SEQ GL_SEQ, A.GL_DT, A.ITEM_AMT AMT, A.ITEM_LOC_AMT LOC_AMT,A.ITEM_DESC GL_DESC "
		lgStrSQL = lgStrSQL & " FROM A_GL_ITEM A, A_ACCT B "
		lgStrSQL = lgStrSQL & " WHERE  B.ACCT_CD = A.ACCT_CD "
		lgStrSQL = lgStrSQL & " AND B.BAL_FG = A.DR_CR_FG "
		lgStrSQL = lgStrSQL & " AND A.GL_NO + CONVERT(CHAR(8), ITEM_SEQ) "
		lgStrSQL = lgStrSQL & " NOT IN (SELECT GL_NO + CONVERT(CHAR(8), GL_SEQ) FROM A_OPEN_ACCT) "
	End If
	lgStrSQL = lgStrSQL & " AND A.ACCT_CD = " & Filtervar(strAccountCd, "''", "S")
	lgStrSQL = lgStrSQL & " AND A.GL_DT BETWEEN " & Filtervar(FromDate, null, "S") & " AND " & Filtervar(ToDate, null, "S")
	lgStrSQL = lgStrSQL & " AND A.DOC_CUR = " & Filtervar(strCurrency, "''", "S")

    If 	FncOpenRs("U",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '☜ : No data is found. 
        lgErrorStatus     = "YES"
        Exit Sub
    Else 
       
		lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)

		If UniCInt(lgStrPrevKey,0) > 0 Then
			lgObjRs.Move     = C_SHEETMAXROWS_D * UniCInt(lgStrPrevKey,0)
		End If

		lgstrData = ""
		iDx = 1
		Do While Not lgObjRs.EOF
		
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CHOICE"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_NO"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_SEQ"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("GL_DT"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_DESC"))
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("AMT"),ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("LOC_AMT"),ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
			lgObjRs.MoveNext

			iDx =  iDx + 1
			If iDx > C_SHEETMAXROWS_D Then
				lgStrPrevKey = UniCInt(lgStrPrevKey,0) + 1
				Exit Do
			End If   
		Loop 

	End If
    
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If iDx <= C_SHEETMAXROWS_D Then
		lgStrPrevKey = ""
    End If   

    If CheckSYSTEMError(Err,True) = True Then
		ObjectContext.SetAbort
		Exit Sub
    End If   

	If lgErrorStatus  = "" Then
		Response.Write  " <Script Language=vbscript>                                  " & vbCr
		Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData       " & vbCr
		Response.Write  "    Parent.frm1.vspdData.Redraw = False "						& vbCr  
		Response.Write  "    Parent.ggoSpread.SSShowData """ & lgstrData & """ ,""F"""	& vbCr
		Response.Write  "    Parent.lgStrPrevKey    = """ & lgStrPrevKey & """"			& vbCr
		Response.Write  "    Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData," & lgLngMaxRow + 1 & "," & lgLngMaxRow + iDx - 1 & ",parent.frm1.txtCurrency.value,parent.C_OPENAMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    Parent.DBQueryOk   " & vbCr
		Response.Write  "    Parent.frm1.vspdData.Redraw = True "						& vbCr  
		Response.Write  " </Script>             " & vbCr
	End If
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim iDx
    Dim arrRowVal
    Dim arrColVal
    Dim strOpenAcctFg
    
    strOpenAcctFg = Request("txtOpenAcctFg")
	arrRowVal = Split(Request("txtSpread"), gRowSep)

	For iDx = 1 To Ubound(arrRowVal)
		arrColVal = Split(arrRowVal(iDx-1), gColSep)
		strGlNo = arrColVal(C_GL_NO)
		strGlSeq = arrColVal(C_GL_SEQ)

  		If strOpenAcctFg = "Y" then
			Call execClosingCancel()
		Else
			Call execClosing()
		End If
	Next
    
End Sub    


 '============================================================================================================
' Name : execClosing
' Desc : Query Data from Db
'============================================================================================================
Sub execClosing()
	Dim IntRetCD
	Dim lgErrorPos
    Dim CALLSPNAME
	''On Error Resume Next   
	Err.Clear

		CALLSPNAME =  "usp_a_create_do_Basic_open_acct"
		
		Call SubCreateCommandObject(lgObjComm)
		With lgObjComm
			.CommandText = CALLSPNAME			'CALLSPNAME
			.CommandType = adCmdStoredProc
			.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE"	,adInteger,adParamReturnValue)
			.Parameters.Append lgObjComm.CreateParameter("@i_gl_no"		,adVarWChar,adParamInput,Len(Trim(strGlNo)), strGlNo)
			.Parameters.Append lgObjComm.CreateParameter("@i_gl_seq"	,adInteger,adParamInput, Len(Trim(strGlSeq)), strGlSeq)
			.Parameters.Append lgObjComm.CreateParameter("@usr_id"		,adVarWChar,adParamInput,Len(Trim(gUsrID)), gUsrID)
			.Parameters.Append lgObjComm.CreateParameter("@msg_cd"		,adVarWChar,adParamOutput, 6)
			.Execute ,, adExecuteNoRecords
		End With
		If  Err.number = 0 Then
			IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
			If  IntRetCD = 1 then
				%><Script Language=vbscript>parent.DbexeOk</Script><%
			end if
		Else 
			lgErrorStatus     = "YES"                                                         '☜: Set error status
			Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
			Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
		End if

		Call SubCloseCommandObject(lgObjComm)

		If lgErrorStatus    = "YES" Then
			lgErrorPos = lgErrorPos & CALLSPNAME & gColSep
		End If
		
End Sub

 '============================================================================================================
' Name : execClosingCancel
' Desc : Query Data from Db
'============================================================================================================
Sub execClosingCancel()
	Dim IntRetCD
    Dim CALLSPNAME 
    Dim lgErrorPos	
    'On Error Resume Next                                                               '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
'	call SubBizQuery()  
    CALLSPNAME =  "usp_a_delete_do_Basic_open_acct"

		Call SubCreateCommandObject(lgObjComm)

		With lgObjComm
			.CommandText = CALLSPNAME			'CALLSPNAME
			.CommandType = adCmdStoredProc
			.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE"	,adInteger,adParamReturnValue)
			.Parameters.Append lgObjComm.CreateParameter("@i_gl_no"		,adVarWChar,adParamInput,Len(Trim(strGlNo)), strGlNo)
			.Parameters.Append lgObjComm.CreateParameter("@i_gl_seq"	,adInteger,adParamInput, Len(Trim(strGlSeq)), strGlSeq)
			.Parameters.Append lgObjComm.CreateParameter("@usr_id"		,adVarWChar,adParamInput,Len(Trim(gUsrID)), gUsrID)
			.Parameters.Append lgObjComm.CreateParameter("@msg_cd"		,adVarWChar,adParamOutput, 6 )
			
			.Execute ,, adExecuteNoRecords
		End With

		If  Err.number = 0 Then
			IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
			If  IntRetCD = 1 then
				%><Script Language=vbscript>parent.DbexeOk</Script><%
			End if
		Else 
		
			lgErrorStatus     = "YES"                                                         '☜: Set error status
			Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
			Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
		End if

		Call SubCloseCommandObject(lgObjComm)

		If lgErrorStatus    = "YES" Then
			lgErrorPos = lgErrorPos & CALLSPNAME & gColSep
		End If
End Sub


 '============================================================================================================
' Name : execClosingAll
' Desc : 
'============================================================================================================
Sub execClosingAll()
	Dim IntRetCD
	Dim lgErrorPos
    Dim CALLSPNAME
	On Error Resume Next   
	Err.Clear

    CALLSPNAME =  "usp_a_create_Basic_open_acct"
           
		Call SubCreateCommandObject(lgObjComm)
		With lgObjComm
			.CommandText = CALLSPNAME			'CALLSPNAME
			.CommandType = adCmdStoredProc
			.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE"	,adInteger,adParamReturnValue)
			.Parameters.Append lgObjComm.CreateParameter("@i_f_dt"		,adVarWChar,adParamInput,8,strFrdate)
			.Parameters.Append lgObjComm.CreateParameter("@i_t_dt"		,adVarWChar,adParamInput,8, strTodate)
			.Parameters.Append lgObjComm.CreateParameter("@i_acct_cd"	,adVarWChar,adParamInput,20,strAccountCd)
			.Parameters.Append lgObjComm.CreateParameter("@i_doc_cur"	,adVarWChar,adParamInput,3, strCurrency)
			.Parameters.Append lgObjComm.CreateParameter("@usr_id"		,adVarWChar,adParamInput,13, gUsrID)
			.Parameters.Append lgObjComm.CreateParameter("@msg_cd"		,adVarWChar,adParamOutput, 6)
			.Execute ,, adExecuteNoRecords
		End With
		If  Err.number = 0 Then
			IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
			If  IntRetCD = 1 then
				%><Script Language=vbscript>parent.DbexeOk</Script><%
			end if
		Else 
			lgErrorStatus     = "YES"                                                         '☜: Set error status
			Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
			Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
		End if

		Call SubCloseCommandObject(lgObjComm)

		If lgErrorStatus    = "YES" Then
			lgErrorPos = lgErrorPos & CALLSPNAME & gColSep
		End If
		

End Sub

 '============================================================================================================
' Name : execClosingCancelAll
' Desc : 
'============================================================================================================
Sub execClosingCancelAll()
	Dim IntRetCD
    Dim CALLSPNAME 
    Dim lgErrorPos	
    On Error Resume Next                                                               '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	call SubBizQuery()  
    CALLSPNAME =  "usp_a_delete_Basic_open_acct"

		Call SubCreateCommandObject(lgObjComm)

		With lgObjComm
			.CommandText = CALLSPNAME			'CALLSPNAME
			.CommandType = adCmdStoredProc
			.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE"	,adInteger,adParamReturnValue)
			.Parameters.Append lgObjComm.CreateParameter("@i_f_dt"		,adVarWChar,adParamInput,8,strFrdate)
			.Parameters.Append lgObjComm.CreateParameter("@i_t_dt"		,adVarWChar,adParamInput,8, strTodate)
			.Parameters.Append lgObjComm.CreateParameter("@i_acct_cd"	,adVarWChar,adParamInput,20,strAccountCd)
			.Parameters.Append lgObjComm.CreateParameter("@i_doc_cur"	,adVarWChar,adParamInput,3, strCurrency)
			.Parameters.Append lgObjComm.CreateParameter("@usr_id"		,adVarWChar,adParamInput,13, gUsrID)
			.Parameters.Append lgObjComm.CreateParameter("@msg_cd"		,adVarWChar,adParamOutput, 6 )
			
			.Execute ,, adExecuteNoRecords
		End With

		If  Err.number = 0 Then
			IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
			If  IntRetCD = 1 then
				%><Script Language=vbscript>parent.DbexeOk</Script><%
			End if
		Else 
		
			lgErrorStatus     = "YES"                                                         '☜: Set error status
			Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
			Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
		End if

		Call SubCloseCommandObject(lgObjComm)

		If lgErrorStatus    = "YES" Then
			lgErrorPos = lgErrorPos & CALLSPNAME & gColSep
		End If
	

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
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
%>


