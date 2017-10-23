<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
	Response.Expires = -1                               '☜ : will expire the response immediately
	Response.Buffer = True                              '☜ : The server does not send output to the client until all of the ASP 
														'     scripts on the current page have been processed
	On Error Resume Next                                '☜: Protect system from crashing
	Err.Clear                                           '☜: Clear Error status
	Dim lgErrorStatus, lgErrorPos, lgObjConn, lgObjRs
	Dim lgOpModeCRUD, iStrData
	Dim nParamIndex
	Dim nFieldIndex

	Call LoadBasisGlobalInf()
	Call HideStatusWnd
	'---------------------------------------Common-----------------------------------------------------------
	lgErrorStatus   = "NO"
	lgErrorPos      = ""                                '☜: Set to space
	lgOpModeCRUD    = Request("txtMode")                '☜: Read Operation Mode (CRUD)

	Call SubOpenDB(lgObjConn)                           '☜: Make a DB Connection

	Dim iKey1
	Dim lgCARD_DD
	Dim lgStrSQL
	On Error Resume Next                                '☜: Protect system from crashing
	Err.Clear                                           '☜: Clear Error status

	'----------------------------------------------------------------------------------------------
	' b_bdc_master 정보를 읽어 화면에 표시한다.
	lgStrSQL = "SELECT process_name, use_flag, join_method, tran_flag, start_row, run_time " & _
			   "FROM   b_bdc_master " & _
			   "WHERE  process_id = '" & Trim(Request("txtProcID")) & "'"

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus()
		Call SubCloseRs(lgObjRs)
		Call SubCloseRs(lgObjConn)
		Response.End
	Else 
		Response.Write "<Script Language=vbscript>"                     & vbCr
		Response.Write "With Parent "                                   & vbCr
		Response.Write "    .frm1.txtProcNm.value = """ & lgObjRs("process_name") & """"  & vbCr
		Response.Write "    .frm1.txtNProcNm.value = """ & lgObjRs("process_name") & """"  & vbCr
		If Trim(lgObjRs("use_flag")) = "Y" Then
			Response.Write "    .frm1.Radio1.checked = true" & vbCr
			Response.Write "    .frm1.hUseFlag.value = ""Y"" " & vbCr
		Else
			Response.Write "    .frm1.Radio2.checked = true " & vbCr
			Response.Write "    .frm1.hUseFlag.value = ""N"" " & vbCr
		End If

		If Trim(lgObjRs("join_method")) = "N" Then
			Response.Write "    .frm1.Radio3.checked = true" & vbCr
			Response.Write "    .frm1.hJoinMethod.value = ""N"" " & vbCr
		Else
			Response.Write "    .frm1.Radio4.checked = true " & vbCr
			Response.Write "    .frm1.hJoinMethod.value = ""S"" " & vbCr
		End If

		If Trim(lgObjRs("tran_flag")) = "Y" Then
			Response.Write "    .frm1.Radio6.checked = true " & vbCr
			Response.Write "    .frm1.hTranFlag.value = ""Y"" " & vbCr
		Else
			Response.Write "    .frm1.Radio7.checked = true " & vbCr
			Response.Write "    .frm1.hTranFlag.value = ""N"" " & vbCr
		End If

		Response.Write "    .frm1.txtStartRow.value = """ & lgObjRs("start_row") & """"  & vbCr
		Response.Write "    .frm1.txtRunTime.value = """ & lgObjRs("run_time") & """"  & vbCr
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"                                      & vbCr
		Call SubCloseRs(lgObjRs)
	End If

	'----------------------------------------------------------------------------------------------
	' b_bdc_com 테이블 정보를 읽어들여 배열에 저장한다.
	lgStrSQL = "SELECT process_id, process_seq, com_name, method_name, proc_desc " & _
			   "FROM   b_bdc_com " & _
			   "WHERE  process_id = '" & Trim(Request("txtProcID")) & "' " & _
			   "ORDER BY process_seq"

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
		'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		'Call SetErrorStatus()
		'Call SubCloseRs(lgObjRs)
		'Call SubCloseRs(lgObjConn)
		'Response.End
	Else
		iStrData = ""
		Do while Not (lgObjRs.EOF Or lgObjRs.BOF)
            iStrData = iStrData & Chr(11) & lgObjRs("com_name")
            iStrData = iStrData & Chr(11) & lgObjRs("method_name")
            iStrData = iStrData & Chr(11) & lgObjRs("proc_desc")
            iStrData = iStrData & Chr(11) & Chr(12)
			lgObjRs.MoveNext
		Loop
	
		Response.Write "<Script Language=vbscript>"                     & vbCr
		Response.Write "With Parent "                                   & vbCr
		Response.Write "    .ggoSpread.Source = .frm1.vspdData "        & vbCr
		Response.Write "    .ggoSpread.SSShowData """ & iStrData & """" & vbCr
		Response.Write "    .lgStrPrevKey = """ & iStrNextkey & """"    & vbCr
		Response.Write "    .frm1.vspdData.ReDraw = True "              & vbCr
		Response.Write "    .DbQueryOk  "                               & vbCr
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"                                      & vbCr
		Call SubCloseRs(lgObjRs)
	End If

	'----------------------------------------------------------------------------------------------
	' b_bdc_param 테이블 정보를 읽어들여 배열에 저장한다.
	lgStrSQL = "SELECT process_seq, param_seq, param_name, type, optional " & _
			   "FROM   b_bdc_param " & _
			   "WHERE  process_id = '" & Trim(Request("txtProcID")) & "' " & _
			   "ORDER BY process_seq, param_seq"

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
		'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		'Call SetErrorStatus()
		'Call SubCloseRs(lgObjRs)
		'Call SubCloseRs(lgObjConn)
		'Response.End
	Else
		Response.Write "<Script Language=vbscript>"                     & vbCr
		Response.Write "With Parent "                                   & vbCr
		iStrData = ""
		nParamIndex = CInt(lgObjRs("process_seq"))
		Do while Not (lgObjRs.EOF Or lgObjRs.BOF)
			If nParamIndex <> CInt(lgObjRs("process_seq")) Then
				Response.Write "	.arrParams(" & nParamIndex-1 & ") = """ & iStrData & """" & vbCr
				nParamIndex = CInt(lgObjRs("process_seq"))
				iStrData = ""
			End If
            iStrData = iStrData & Chr(11) & lgObjRs("param_name")
            iStrData = iStrData & Chr(11) & lgObjRs("type")
            iStrData = iStrData & Chr(11) & lgObjRs("optional")
            iStrData = iStrData & Chr(11) & Chr(12)
			lgObjRs.MoveNext
		Loop
		Response.Write "	.arrParams(" & nParamIndex-1 & ") = """ & iStrData & """" & vbCr
		Response.Write "    .ggoSpread.Source = .frm1.vspdData1 "       & vbCr
		Response.Write "    .ggoSpread.SSShowData .arrParams(0) "		& vbCr
		Response.Write "    .lgStrPrevKey = """ & iStrNextkey & """"    & vbCr
		Response.Write "    .frm1.vspdData1.ReDraw = True "             & vbCr
		Response.Write "    .DbQueryOk  "                               & vbCr
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"                                      & vbCr
		Call SubCloseRs(lgObjRs)
	End If

	'----------------------------------------------------------------------------------------------
	' b_bdc_param_field 테이블 정보를 읽어들여 배열에 저장한다.
	lgStrSQL = "SELECT a.process_seq, a.param_seq, a.seq, a.attch_char,	 " & _
			   "       b.field_id, b.sheet_no, b.field_seq, b.field_name " & _
			   "FROM   b_bdc_param_field a,								 " & _
			   "	   b_bdc_field b									 " & _
			   "WHERE  a.process_id='" & Trim(Request("txtProcID")) & "' " & _
			   "  AND  b.process_id=a.process_id						 " & _
			   "  AND  b.field_id = a.field_id							 " & _
			   "ORDER BY a.process_seq, a.param_seq, a.seq				 "

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
	'	Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
	'	Call SetErrorStatus()
	'	Call SubCloseRs(lgObjRs)
	'	Call SubCloseRs(lgObjConn)
	'	Response.End
	Else
		iStrData = ""
		Response.Write "<Script Language=vbscript>"                     & vbCr
		Response.Write "With Parent "                                   & vbCr
		nParamIndex = CInt(lgObjRs("process_seq"))
		nFieldIndex = CInt(lgObjRs("param_seq"))
		Response.Write "	.arrJoins(0,0) = """""						& vbCr
		Do while Not (lgObjRs.EOF Or lgObjRs.BOF)
			If nParamIndex <> CInt(lgObjRs("process_seq")) _
			Or nFieldIndex <> CInt(lgObjRs("param_seq")) Then
				Response.Write "	.arrJoins(" & nParamIndex-1 & "," & nFieldIndex-1 & ") = """ & iStrData & """" & vbCr
				nParamIndex = CInt(lgObjRs("process_seq"))
				nFieldIndex = CInt(lgObjRs("param_seq"))
				iStrData = ""
			End If
            iStrData = iStrData & Chr(11) & lgObjRs("field_id")
            iStrData = iStrData & Chr(11) & lgObjRs("sheet_no")
            iStrData = iStrData & Chr(11) & lgObjRs("field_seq")
            iStrData = iStrData & Chr(11) & lgObjRs("field_name")
            iStrData = iStrData & Chr(11) & lgObjRs("attch_char")
            iStrData = iStrData & Chr(11) & Chr(12)
			lgObjRs.MoveNext
		Loop

		Response.Write "	.arrJoins(" & nParamIndex-1 & "," & nFieldIndex-1 & ") = """ & iStrData & """" & vbCr
		Response.Write "    .ggoSpread.Source = .frm1.vspdData2 "       & vbCr
		Response.Write "    .ggoSpread.SSShowData .arrJoins(0,0)"		& vbCr
		Response.Write "    .lgStrPrevKey = """ & iStrNextkey & """"    & vbCr
		Response.Write "    .frm1.vspdData1.ReDraw = True "             & vbCr
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"                                      & vbCr
		Call SubCloseRs(lgObjRs)
	End If

	'----------------------------------------------------------------------------------------------
	' b_bdc_field 테이블 정보를 읽어들여 배열에 저장한다.
	lgStrSQL = "SELECT field_id, sheet_no, field_seq, field_name, type,	" & _
			   "	   option_flag, parent_field						" & _
			   "FROM   b_bdc_field										" & _
			   "WHERE  process_id='" & Trim(Request("txtProcID")) & "'	" & _
			   "ORDER BY field_id										"

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
		'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		'Call SetErrorStatus()
		'Call SubCloseRs(lgObjRs)
		'Call SubCloseRs(lgObjConn)
		'Response.End
	Else
		iStrData = ""
		Do while Not (lgObjRs.EOF Or lgObjRs.BOF)
            iStrData = iStrData & Chr(11) & lgObjRs("sheet_no")
            iStrData = iStrData & Chr(11) & lgObjRs("field_seq")
            iStrData = iStrData & Chr(11) & lgObjRs("field_name")
            iStrData = iStrData & Chr(11) & lgObjRs("type")
            iStrData = iStrData & Chr(11) & lgObjRs("option_flag")
            iStrData = iStrData & Chr(11) & lgObjRs("parent_field")
            iStrData = iStrData & Chr(11) & Chr(12)
			lgObjRs.MoveNext
		Loop
	
		Response.Write "<Script Language=vbscript>"                     & vbCr
		Response.Write "With Parent "                                   & vbCr
		Response.Write "    .ggoSpread.Source = .frm1.vspdData3 "       & vbCr
		Response.Write "    .ggoSpread.SSShowData """ & iStrData & """" & vbCr
		Response.Write "    .lgStrPrevKey = """ & iStrNextkey & """"    & vbCr
		Response.Write "    .frm1.vspdData1.ReDraw = True "             & vbCr
		Response.Write "    .DbQueryOk  "                               & vbCr
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"                                      & vbCr
		Call SubCloseRs(lgObjRs)
	End If

	Call SubCloseRs(lgObjConn)
	Response.End
%>
