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
    Call LoadInfTB19029B( "Q", "*", "NOCOOKIE", "MB")
    Call LoadBNumericFormatB("Q", "*", "NOCOOKIE", "MB")
	
	Const C1_SHEETMAXROWS_D  = 100
    
    Const C_prnt_plant_cd  = 0
    Const C_prnt_item_cd   = 1
    Const C_child_item_seq = 2
    
    lgLngMaxRow = Trim(Request("txtMaxRows"))
    
    If Trim(Request("lgStrPrevKey")) = "" Then
		lgStrPrevKey = ""
    Else
		lgStrPrevKey = Trim(Request("lgStrPrevKey"))
		lgArrPrevKey = Split(lgStrPrevKey, Chr(11))
	End If
	
    strSQL = ""
    strSQL = strSQL & " SELECT TOP " & C1_SHEETMAXROWS_D + 1
	strSQL = strSQL & "    a.prnt_item_cd"
	strSQL = strSQL & "  , isnull(b.item_nm,'') AS prnt_item_nm"
	strSQL = strSQL & "  , a.child_item_seq"
	strSQL = strSQL & "  , a.child_item_cd"
	strSQL = strSQL & "  , isnull(c.item_nm,'') AS child_item_nm"
	strSQL = strSQL & "  , a.child_item_qty"
	strSQL = strSQL & "  , a.child_item_unit"
	strSQL = strSQL & "  , a.lot_flg"
	strSQL = strSQL & "  , a.valid_from_dt"
	strSQL = strSQL & "  , a.valid_to_dt"
	strSQL = strSQL & "  , a.create_type"
	strSQL = strSQL & "  , a.send_dt"
	strSQL = strSQL & "  , a.mes_receive_flag"
	strSQL = strSQL & "  , a.err_desc"
	strSQL = strSQL & "  , a.mes_receive_dt"
	strSQL = strSQL & " FROM"
	strSQL = strSQL & "  t_if_snd_bom_ko119 a     (nolock) "
	strSQL = strSQL & "  INNER JOIN b_item b      (nolock) ON b.item_cd = a.prnt_item_cd"
	strSQL = strSQL & "  LEFT OUTER JOIN b_item c (nolock) ON c.item_cd = a.child_item_cd"
	strSQL = strSQL & " WHERE"
	strSQL = strSQL & "      a.prnt_plant_cd LIKE    (CASE '" & Request("txtPlantCD")   & "' WHEN '' THEN '%'          ELSE '" & Request("txtPlantCD") & "'       END)"
	strSQL = strSQL & "  AND a.prnt_item_cd  LIKE    (CASE '" & Request("txtPItemCD")   & "' WHEN '' THEN '%'          ELSE '" & Request("txtPItemCD") & "'       END)"
	strSQL = strSQL & "  AND a.child_item_cd LIKE    (CASE '" & Request("txtCItemCD")   & "' WHEN '' THEN '%'          ELSE '" & Request("txtCItemCD") & "'       END)"
	strSQL = strSQL & "  AND a.send_dt       BETWEEN (CASE '" & Request("txtFrDT")      & "' WHEN '' THEN '1900-01-01' ELSE '" & Request("txtFrDT") & "'          END)"
	strSQL = strSQL & "                      AND     (CASE '" & Request("txtToDT")      & "' WHEN '' THEN '9999-12-31' ELSE '" & Request("txtToDT") & " 23:59:59' END)"
	strSQL = strSQL & "  AND a.mes_receive_flag LIKE (CASE '" & Request("txtRadioFlag") & "' WHEN '' THEN '%'          ELSE '" & Request("txtRadioFlag") & "'     END)"
	If Trim(lgStrPrevKey) <> "" Then
		strSQL = strSQL & "  AND (a.prnt_plant_cd > '" & lgArrPrevKey(C_prnt_plant_cd) & "'"
		strSQL = strSQL & "  OR   (a.prnt_plant_cd = '" & lgArrPrevKey(C_prnt_plant_cd) & "'"
		strSQL = strSQL & "  AND   (a.prnt_item_cd > '" & lgArrPrevKey(C_prnt_item_cd) & "'"
		strSQL = strSQL & "  OR     (a.prnt_item_cd = '" & lgArrPrevKey(C_prnt_item_cd) & "'"
		strSQL = strSQL & "  AND     a.child_item_seq >= '" & lgArrPrevKey(C_child_item_seq) & "'))))"
	End If
	strSQL = strSQL & " ORDER BY"
	strSQL = strSQL & "    a.prnt_plant_cd"
	strSQL = strSQL & "  , a.prnt_item_cd"
	strSQL = strSQL & "  , a.child_item_seq"

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
				Response.Write Chr(11) & lgObjRs("prnt_item_cd")
				Response.Write Chr(11) & Replace(lgObjRs("prnt_item_nm"), """", """""")
				Response.Write Chr(11) & lgObjRs("child_item_seq")
				Response.Write Chr(11) & lgObjRs("child_item_cd")
				Response.Write Chr(11) & Replace(lgObjRs("child_item_nm"), """", """""")
				Response.Write Chr(11) & UNINumClientFormat(lgObjRs("child_item_qty"), ggQty.DecPoint, 0)
				Response.Write Chr(11) & lgObjRs("child_item_unit")
				Response.Write Chr(11) & lgObjRs("lot_Flg")
				Response.Write Chr(11) & UNIDateClientFormat(lgObjRs("valid_from_dt"))
				Response.Write Chr(11) & UNIDateClientFormat(lgObjRs("valid_to_dt"))
				Response.Write Chr(11) & lgObjRs("create_type")
				Response.Write Chr(11) & lgObjRs("send_dt")
				Response.Write Chr(11) & lgObjRs("mes_receive_flag")
				Response.Write Chr(11) & lgObjRs("err_desc")
				Response.Write Chr(11) & lgObjRs("mes_receive_dt")

				Response.Write Chr(11) & Chr(12)
			Else
				lgStrPrevKey = lgStrPrevKey & Trim(Request("txtPlantCD"))     & Chr(11)
				lgStrPrevKey = lgStrPrevKey & Trim(lgObjRs("prnt_item_cd"))   & Chr(11)
				lgStrPrevKey = lgStrPrevKey & Trim(lgObjRs("child_item_seq")) & Chr(11)
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

