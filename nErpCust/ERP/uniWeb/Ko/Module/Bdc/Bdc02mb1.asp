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
'	On Error Resume Next                                '☜: Protect system from crashing
'	Err.Clear                                           '☜: Clear Error status
	Dim lgErrorStatus, lgErrorPos, lgObjConn, lgObjRs
	Dim lgOpModeCRUD, iStrData
	Dim nSeqIndex
	Dim nFieldIndex
	Dim i, szTemp

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

	'----------------------------------------------------------------------------------------------
	' b_bdc_param_field 테이블 정보를 읽어들여 배열에 저장한다.
	lgStrSQL = " SELECT sheet_no, field_id, field_name 	 " & vbCrLf & _
			   " FROM   b_bdc_field								 " & vbCrLf & _
			   " WHERE  process_id= " & FilterVar(Request("txtProcID"), "''", "S") & vbCrLf &  _
			   " ORDER BY sheet_no, field_id				 "

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus()
		Call SubCloseRs(lgObjRs)
		Call SubCloseRs(lgObjConn)
		Response.End
	Else
		iStrData = ""
		Response.Write "<Script Language=vbscript>"                     & vbCr
		Response.Write "With Parent "                                   & vbCr
		Response.Write "    .ggoSpread.Source = .frm1.vspdData3"       & vbCr
		Response.Write "    .ggoSpread.SSShowData """ 
		Do while Not (lgObjRs.EOF Or lgObjRs.BOF)
            Response.Write Chr(11) & lgObjRs("sheet_no")
            Response.Write Chr(11) & lgObjRs("field_id")
            Response.Write Chr(11) & lgObjRs("field_name")
            Response.Write Chr(11) & Chr(12)
			lgObjRs.MoveNext
		Loop
	    Response.Write """" & vbCr
		Response.Write "	Call .SetSpreadColor(3, -1, -1) " & vbCr
		Response.Write "    .frm1.vspdData3.ReDraw = True "             & vbCr
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"                                      & vbCr
		Call SubCloseRs(lgObjRs)
	End If


	'----------------------------------------------------------------------------------------------
	' Query b_bdc_sql
	lgStrSQL = "SELECT sql_name, ret_value, statement " & _
			   "FROM   b_bdc_sql " & _
			   "WHERE  process_id = " & Filtervar(Request("txtProcID"), "''", "S") & _
			   "ORDER BY seq "

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = True Then
		Response.Write "<Script Language=vbscript>"                     & vbCr
		Response.Write "With Parent "                                   & vbCr
		Response.Write "    .ggoSpread.Source = .frm1.vspdData1 "        & vbCr
		Response.Write "    .ggoSpread.SSShowData """ 
		iStrData = ""
		i = 0
		Do while Not (lgObjRs.EOF Or lgObjRs.BOF)
            Response.Write Chr(11) & lgObjRs("sql_name")
            Response.Write Chr(11) & lgObjRs("ret_value")
            Response.Write Chr(11) & Chr(12)
            szTemp = lgObjRs("statement")
            iStrData = iStrData & ".arrQuery(" & i & ",0) = Replace(Replace(""" & Replace(Replace(szTemp, vbCr, chr(11)), vbLf, chr(12)) & """, Chr(11), vbCr), Chr(12), vbLf)" & vbCrLf
            i = i + 1
			lgObjRs.MoveNext
		Loop
	    
	    Response.Write """" & vbCr
		Response.Write "    .frm1.vspdData1.ReDraw = True "              & vbCr
		Response.Write iStrData
		Response.Write "	.frm1.txtSQLStatements.value = .arrQuery(0, 0) " & vbCrLf
		Response.Write "    .nSpreadIndex1 = 1  "                       & vbCr
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"                                      & vbCr
		Call SubCloseRs(lgObjRs)
	End If

	'----------------------------------------------------------------------------------------------
	' b_bdc_param 테이블 정보를 읽어들여 배열에 저장한다.
	lgStrSQL = " SELECT a.seq, a.param_seq,  " & _
			   "		a.name, a.param_type, a.length, " & _
			   "		a.field_id a, b.field_name      " & _
			   " FROM   b_bdc_sql_param a, b_bdc_field b" & _
			   " WHERE  a.process_id = " & FilterVar(Request("txtProcID"), "''", "S")  & _
			   " AND	a.process_id = b.process_id " & _
			   " AND	a.field_id = b.field_id " & _
			   " ORDER BY a.seq, a.param_seq "

    
	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = True Then
		Response.Write "<Script Language=vbscript>"                     & vbCr
		Response.Write "With Parent "                                   & vbCr
		iStrData = ""
		nSeqIndex = CInt(lgObjRs("seq"))
		
		Do while Not (lgObjRs.EOF Or lgObjRs.BOF)
			If nSeqIndex <> CInt(lgObjRs("seq")) Then
				Response.Write "	.arrQuery(" & nSeqIndex-1 & ", 1) = """ & iStrData & """" & vbCr
				nSeqIndex = CInt(lgObjRs("seq"))
				iStrData = ""
			End If
			
            iStrData = iStrData & lgObjRs(1) & Chr(11) & _
                                lgObjRs(0) & Chr(11) & _
                                lgObjRs(2) & Chr(11) & _
                                lgObjRs(3) & Chr(11) & _
                                lgObjRs(4) & Chr(11) & _
                                lgObjRs(5) & Chr(11) & _
                                lgObjRs(6) & Chr(11) & _
                                Chr(12)
			lgObjRs.MoveNext
		Loop
		
		Response.Write "	.arrQuery(" & nSeqIndex-1 & ", 1) = """ & iStrData & """" & vbCr
		Response.Write "    .ggoSpread.Source = .frm1.vspdData2 "       & vbCr
		Response.Write "    .ggoSpread.SSShowData .arrQuery(0,1) "		& vbCr
		Response.Write "    .frm1.vspdData2.ReDraw = True "             & vbCr
		Response.Write "    .nSpreadIndex2 = 1  "                       & vbCr
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"                                      & vbCr
		Call SubCloseRs(lgObjRs)
	End If

	'----------------------------------------------------------------------------------------------
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "    Parent.DbQueryOk "      & vbCr
	Response.Write "</Script>"                  & vbCr
    
	Call SubCloseRs(lgObjConn)
	Response.End
%>
