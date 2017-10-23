<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
    Dim strSQL
    Dim iLngRow, lgStrPrevKey
    
    On Error Resume Next
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
    Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
    
    Const C1_SHEETMAXROWS_D  = 100
    
    lgLngMaxRow = Trim(Request("txtMaxRows"))
    
    If Trim(Request("lgStrPrevKey")) = "" Then
		lgStrPrevKey = ""
    Else
		lgStrPrevKey = Trim(Request("lgStrPrevKey"))
	End If
	
	strSQL = ""
    strSQL = strSQL & " SELECT TOP " & C1_SHEETMAXROWS_D + 1
    strSQL = strSQL & "    bp_cd"
    strSQL = strSQL & "  , bp_nm"
    strSQL = strSQL & "  , bp_type"
    strSQL = strSQL & "  , usage_flag"
    strSQL = strSQL & "  , create_type"
    strSQL = strSQL & "  , send_dt"
    strSQL = strSQL & "  , mes_receive_flag"
    strSQL = strSQL & "  , err_desc"
    strSQL = strSQL & "  , mes_receive_dt"
	strSQL = strSQL & "	FROM"
	strSQL = strSQL & "  t_if_snd_biz_part_ko119 (nolock) "
	strSQL = strSQL & " WHERE"
	strSQL = strSQL & "      bp_cd   BETWEEN (CASE '" & Request("txtBp_cdFrom") & "' WHEN '' THEN ''           ELSE '" & Request("txtBp_cdFrom") & "'     END)"
	strSQL = strSQL & "		         AND     (CASE '" & Request("txtBp_cdTo")   & "' WHEN '' THEN 'ZZZZ'       ELSE '" & Request("txtBp_cdTo")   & "'     END)"
	strSQL = strSQL & "	 AND send_dt BETWEEN (CASE '" & Request("txtFrDT")      & "' WHEN '' THEN '1900-01-01' ELSE '" & Request("txtFrDT")      & "'     END)"
	strSQL = strSQL & "	             AND     (CASE '" & Request("txtToDT")      & "' WHEN '' THEN '9999-12-31' ELSE '" & Request("txtToDT") & " 23:59:59' END)"
	strSQL = strSQL & "	 AND bp_type LIKE    (CASE '" & Request("txtRadioType") & "' WHEN '' THEN '%'          ELSE '" & Request("txtRadioType") & "'     END)"
	strSQL = strSQL & "	 AND mes_receive_flag LIKE (CASE '" & Request("txtRadioFlag") & "' WHEN '' THEN '%' ELSE '" & Request("txtRadioFlag") & "' END)"
	If Trim(lgStrPrevKey) <> "" Then
		strSQL = strSQL & "  AND bp_cd >= '" & lgStrPrevKey & "'"
	End If
	strSQL = strSQL & " ORDER BY"
	strSQL = strSQL & "  bp_cd"
	
    Call SubOpenDB(lgObjConn)

    If FncOpenRs("R",lgObjConn,lgObjRs,strSQL,"X","X") = False Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
    Else
		Response.Write "<Script Language=""VBScript"">"         & vbCrLf
		Response.Write "With Parent"                            & vbCrLf
		Response.Write "    .ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write "    .ggoSpread.SSShowDataByClip """
		
		iLngRow = 0
		lgStrPrevKey = ""
		Do While Not lgObjRs.EOF
			If iLngRow < C1_SHEETMAXROWS_D Then
				Response.Write Chr(11) & lgObjRs("bp_cd")
				Response.Write Chr(11) & lgObjRs("bp_nm")
				Response.Write Chr(11) & lgObjRs("bp_type")
				Response.Write Chr(11) & lgObjRs("usage_flag")
				Response.Write Chr(11) & lgObjRs("create_type")
				Response.Write Chr(11) & lgObjRs("send_dt")
				Response.Write Chr(11) & lgObjRs("mes_receive_flag")
				Response.Write Chr(11) & lgObjRs("err_desc")
				Response.Write Chr(11) & lgObjRs("mes_receive_dt")
				Response.Write Chr(11) & Chr(12)
			Else
				lgStrPrevKey = Trim(lgObjRs("bp_cd"))
			End If
			
			iLngRow = iLngRow + 1
			
			lgObjRs.MoveNext
		Loop

		Response.Write """"                                               & vbCrLf
		Response.Write "	.lgStrPrevKey = """ & lgStrPrevKey & """    " & vbCrLf
		Response.Write "    .DBQueryOk "                                  & vbCrLf
		Response.Write "End with"                                         & vbCrLf
		Response.Write "</Script>"                                        & vbCrLf

    End If
    
	Call SubCloseRs(lgObjRs)
	Call SubCloseDB(lgObjConn)    
	
    Response.End 
%>
