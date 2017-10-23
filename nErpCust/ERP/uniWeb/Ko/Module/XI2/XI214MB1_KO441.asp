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
    Dim iLngRow, lgStrPrevKey, lgArrPrevKey
    On Error Resume Next
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
    Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
    
    Const C1_SHEETMAXROWS_D  = 100
    
    Const C_item_cd    = 0
    Const C_bp_cd      = 1
    Const C_bp_item_cd = 2
    
    lgLngMaxRow = Trim(Request("txtMaxRows"))
    
    If Trim(Request("lgStrPrevKey")) = "" Then
		lgStrPrevKey = ""
    Else
		lgStrPrevKey = Trim(Request("lgStrPrevKey"))
		lgArrPrevKey = Split(lgStrPrevKey, Chr(11))
	End If

    strSQL = ""
    strSQL = strSQL & " SELECT TOP " & C1_SHEETMAXROWS_D + 1
    strSQL = strSQL & "    a.item_cd"
    strSQL = strSQL & "  , b.item_nm"
    strSQL = strSQL & "  , a.bp_cd"
    strSQL = strSQL & "  , c.bp_nm"
    strSQL = strSQL & "  , a.bp_item_cd"
    strSQL = strSQL & "  , a.bp_item_nm"
    strSQL = strSQL & "  , a.bp_item_spec"
    strSQL = strSQL & "  , a.bp_item_prt_spec"
    strSQL = strSQL & "  , a.create_type"
    strSQL = strSQL & "  , a.send_dt"
    strSQL = strSQL & "  , a.mes_receive_flag"
    strSQL = strSQL & "  , a.err_desc"
    strSQL = strSQL & "  , a.mes_receive_dt"
	strSQL = strSQL & " FROM"
	strSQL = strSQL & "  t_if_snd_bp_item_ko441 a   (nolock) "
	strSQL = strSQL & "  INNER JOIN b_item b        (nolock)  ON b.item_cd = a.item_cd"
	strSQL = strSQL & "  INNER JOIN b_biz_partner c (nolock)  ON c.bp_cd = a.bp_cd"
	strSQL = strSQL & " WHERE"
	strSQL = strSQL & "      a.bp_cd   LIKE    (CASE '" & Request("txtBp_cd")  & "' WHEN '' THEN '%'          ELSE '" & Request("txtBp_cd") & "'         END)"
	strSQL = strSQL & "  AND a.send_dt BETWEEN (CASE '" & Request("txtFrDT")   & "' WHEN '' THEN '1900-01-01' ELSE '" & Request("txtFrDT") & "'          END)"
	strSQL = strSQL & "                AND     (CASE '" & Request("txtToDT")   & "' WHEN '' THEN '9999-12-31' ELSE '" & Request("txtToDT") & " 23:59:59' END)"
	strSQL = strSQL & "  AND a.item_cd LIKE    (CASE '" & Request("txtItemCD") & "' WHEN '' THEN '%'          ELSE '" & Request("txtItemCD") & "'        END)"
	strSQL = strSQL & "  AND a.mes_receive_flag LIKE (CASE '" & Request("txtRadioFlag") & "' WHEN '' THEN '%' ELSE '" & Request("txtRadioFlag") & "' END)"
	If Trim(lgStrPrevKey) <> "" Then
		strSQL = strSQL & "  AND (a.item_cd > '" & lgArrPrevKey(C_item_cd) & "'"
		strSQL = strSQL & "  OR   (a.item_cd = '" & lgArrPrevKey(C_item_cd) & "'"
		strSQL = strSQL & "  AND   (a.bp_cd > '" & lgArrPrevKey(C_bp_cd) & "'"
		strSQL = strSQL & "  OR     (a.bp_cd = '" & lgArrPrevKey(C_bp_cd) & "'"
		strSQL = strSQL & "  AND     a.bp_item_cd >= '" & lgArrPrevKey(C_bp_item_cd) & "'))))"
	End If
	strSQL = strSQL & " ORDER BY"
	strSQL = strSQL & "    a.item_cd"
	strSQL = strSQL & "  , a.bp_cd"
	strSQL = strSQL & "  , a.bp_item_cd"

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
				Response.Write Chr(11) & lgObjRs("item_cd")
				Response.Write Chr(11) & Replace(lgObjRs("item_nm"), """", """""")
				Response.Write Chr(11) & lgObjRs("bp_cd")
				Response.Write Chr(11) & lgObjRs("bp_nm")
				Response.Write Chr(11) & lgObjRs("bp_item_cd")
				Response.Write Chr(11) & Replace(lgObjRs("bp_item_nm"), """", """""")
				Response.Write Chr(11) & Replace(lgObjRs("bp_item_spec"), """", """""")
				Response.Write Chr(11) & Replace(lgObjRs("bp_item_prt_spec"), """", """""")
				Response.Write Chr(11) & lgObjRs("create_type")
				Response.Write Chr(11) & lgObjRs("send_dt")
				Response.Write Chr(11) & lgObjRs("mes_receive_flag")
				Response.Write Chr(11) & lgObjRs("err_desc")
				Response.Write Chr(11) & lgObjRs("mes_receive_dt")
				Response.Write Chr(11) & Chr(12)
			Else
				lgStrPrevKey = lgStrPrevKey & Trim(lgObjRs("item_cd"))    & Chr(11)
				lgStrPrevKey = lgStrPrevKey & Trim(lgObjRs("bp_cd"))      & Chr(11)
				lgStrPrevKey = lgStrPrevKey & Trim(lgObjRs("bp_item_cd")) & Chr(11)
			End If
			
			iLngRow = iLngRow + 1
				
			lgObjRs.MoveNext
		Loop
		Response.Write """"                                               & vbCrLf
		Response.Write "	.lgStrPrevKey = """ & lgStrPrevKey & """    " & vbCrLf
		Response.Write "    .DBQueryOk  "                                 & vbCrLf
		Response.Write "End with"                                         & vbCrLf
		Response.Write "</Script>"                                        & vbCrLf
	End If

    Call SubCloseRs(lgObjRs)
    Call SubCloseDB(lgObjConn)
    Response.End
    
'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
%>

