<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Module Name          : 구매 
'*  2. Function Name        : 매입일괄등록 
'*  3. Program ID           : M5121BB1
'*  4. Program Name         :
'*  5. Program Desc         : 매입일괄등록 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2005/08/30
'*  8. Modified date(Last)  : 2005/09/08
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Shim Hae Young
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
' =======================================================================================================

%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%


	Call LoadBasisGlobalInf()

	Call loadInfTB19029B("I", "*","NOCOOKIE","BB")

    Call HideStatusWnd                                            '☜: Hide Processing message
    Call SubCreateCommandObject(lgObjComm)
    Call SubOpenDB(lgObjConn)
    Call SubBizQuery()
    Call SubCloseDB(lgObjConn)

    Call SubBizBatch()
    Call SubCloseCommandObject(lgObjComm)

'============================================================================================================
' Name : SubBizBatch
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim intRetCD
    Dim iObjRs
    Dim iArrBillNo		' 추가된 매출채권번호 
    Dim iStrArFlag		' 확정여부 
	Dim iStrWorkType	' 작업유형 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    IntRetCD = 0


    Set iObjRs = Server.CreateObject("ADODB.Recordset")

    With lgObjComm
		.CommandTimeout = 0
		.CommandText = "dbo.usp_m_iv_batch"
        .CommandType = adCmdStoredProc

	    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger,adParamReturnValue)
		.Parameters.Append .CreateParameter("@from_dt", adDBTimeStamp,adParamInput,,UNIConvDate(Request("txtConFromDt")))
	    .Parameters.Append .CreateParameter("@to_dt", adDBTimeStamp,adParamInput,,UNIConvDate(Request("txtConToDt")))

		If Trim(Request("txtConMovType")) <> "" Then
		    .Parameters.Append .CreateParameter("@io_type", adVarXChar,adParamInput,5,Replace(Request("txtConMovType"), "'", "''"))
		Else
		    .Parameters.Append .CreateParameter("@io_type", adVarXChar,adParamInput,5,"%")
		End If

		If Trim(Request("txtConSppl")) <> "" Then
		    .Parameters.Append .CreateParameter("@bp_cd", adVarXChar,adParamInput,10,Replace(Request("txtConSppl"), "'", "''"))
		Else
		    .Parameters.Append .CreateParameter("@bp_cd", adVarXChar,adParamInput,10,"%")
		End If

		If Trim(Request("txtConPurGrp")) <> "" Then
		    .Parameters.Append .CreateParameter("@pur_grp_cond", adVarXChar,adParamInput,4,Replace(Request("txtConPurGrp"), "'", "''"))
		Else
		    .Parameters.Append .CreateParameter("@pur_grp_cond", adVarXChar,adParamInput,4,"%")
		End If

		.Parameters.Append .CreateParameter("@iv_dt", adDBTimeStamp,adParamInput,,UNIConvDate(Request("txtIvDt")))
	    .Parameters.Append .CreateParameter("@pur_grp", adVarXChar,adParamInput,4,Replace(Request("txtPurGrp"), "'", "''"))

	    .Parameters.Append .CreateParameter("@iv_type", adVarXChar,adParamInput,5,Replace(Request("txtIvType"), "'", "''"))
	    .Parameters.Append .CreateParameter("@vat_type", adVarXChar,adParamInput,5,Replace(Request("txtVAT"), "'", "''"))

	    .Parameters.Append .CreateParameter("@vat_rt", adDouble,adParamInput,,UniConvNum(Request("txtVatRt"),0))


	    .Parameters.Append .CreateParameter("@vat_inc_flg", adXChar,adParamInput,4,Replace(Request("txtVatFlg"), "'", "''"))
	    .Parameters.Append .CreateParameter("@pay_meth", adXChar,adParamInput,4,Replace(Request("txtPayMeth"), "'", "''"))
		.Parameters.Append .CreateParameter("@user_id", adVarXChar,adParamInput,13,Replace(Request("txtUserId"), "'", "''"))
		.Parameters.Append .CreateParameter("@issue_dt_fg", adVarXChar,adParamInput,1,Replace(Request("txtIssueDTFg"), "'", "''"))

       	Set iObjRs = .Execute
    End With

    If CheckSYSTEMError(Err,True) = True Then
       IntRetCD = -1
       Exit Sub
    End If

    IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

    If CDbl(intRetCD) = 0 Then
		iArrBillNo = iObjRs.GetRows

		iObjRs.Close
    	Set iObjRs = Nothing

		Call DisplayMsgBox("204262", vbOKOnly, iArrBillNo(0, 0) & "~" & iArrBillNo(1, 0) & " (" & iArrBillNo(2, 0) & ")", "", I_MKSCRIPT)
    Else
       Call DisplayMsgBox(IntRetCd, vbInformation, "", "", I_MKSCRIPT)
       If Not(iObjRs Is Nothing) then
			Set iObjRs = Nothing
       End If
    End If

	Response.Write  " <Script Language=vbscript> " & vbCr
	Response.Write  "  Call Parent.frm1.txtConFromDt.focus  " & vbCr
	Response.Write  " </Script>                  " & vbCr
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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

	Dim iDx

	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear


	'-------------------
	'입출고형태 코드체크 
	'-------------------
	If Trim(Request("txtConMovType")) <> "" Then
    	lgStrSQL = ""
    	Call SubMakeSQLStatements("0", Trim(Request("txtConMovType")), "", "")           '☜ : Make sql statements


    	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists

    		Call DisplayMsgBox("171900", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtConMovTypeNm.Value  = """"" & vbCrLf   'Set condition area
    		Response.Write "parent.Frm1.txtConMovType.focus" & vbCrLf   'Set condition area
    		Response.Write "</Script>" & vbcRLf
    		Response.End
    	Else
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtConMovTypeNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
    		Response.Write "</Script>" & vbcRLf
    	End If


    	Call SubCloseRs(lgObjRs)
    End If

	'-------------------
	'공급처 코드체크 
	'-------------------
	If Trim(Request("txtConSppl")) <> "" Then
    	lgStrSQL = ""
    	Call SubMakeSQLStatements("1", Trim(Request("txtConSppl")), "", "")           '☜ : Make sql statements


    	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists

    		Call DisplayMsgBox("229927", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtConSpplNm.Value  = """"" & vbCrLf   'Set condition area
    		Response.Write "parent.Frm1.txtConSppl.focus" & vbCrLf   'Set condition area
    		Response.Write "</Script>" & vbcRLf
    		Response.End
    	Else
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtConSpplNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
    		Response.Write "</Script>" & vbcRLf
    	End If


    	Call SubCloseRs(lgObjRs)
    End If

	'-------------------
	'구매그룹 코드체크 
	'-------------------
	If Trim(Request("txtConPurGrp")) <> "" Then
    	lgStrSQL = ""
    	Call SubMakeSQLStatements("2", Trim(Request("txtConPurGrp")), "", "")           '☜ : Make sql statements


    	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists

    		Call DisplayMsgBox("125100", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtConPurGrpNm.Value  = """"" & vbCrLf   'Set condition area
    		Response.Write "parent.Frm1.txtConPurGrp.focus" & vbCrLf   'Set condition area
    		Response.Write "</Script>" & vbcRLf
    		Response.End
    	Else
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtConPurGrpNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
    		Response.Write "</Script>" & vbcRLf
    	End If


    	Call SubCloseRs(lgObjRs)
    End If


	'-------------------
	'매입형태 코드체크 
	'-------------------
	If Trim(Request("txtIvType")) <> "" Then
    	lgStrSQL = ""
    	Call SubMakeSQLStatements("3", Trim(Request("txtIvType")), "", "")           '☜ : Make sql statements


    	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists

    		Call DisplayMsgBox("171800", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtIvTypeNm.Value  = """"" & vbCrLf   'Set condition area
    		Response.Write "parent.Frm1.txtIvType.focus" & vbCrLf   'Set condition area
    		Response.Write "</Script>" & vbcRLf
    		Response.End
    	Else
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtIvTypeNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
    		Response.Write "</Script>" & vbcRLf
    	End If


    	Call SubCloseRs(lgObjRs)
    End If


	'-------------------
	'구매그룹 코드체크 
	'-------------------
	If Trim(Request("txtPurGrp")) <> "" Then
    	lgStrSQL = ""
    	Call SubMakeSQLStatements("4", Trim(Request("txtPurGrp")), "", "")           '☜ : Make sql statements


    	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists

    		Call DisplayMsgBox("125100", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtPurGrpNm.Value  = """"" & vbCrLf   'Set condition area
    		Response.Write "parent.Frm1.txtPurGrp.focus" & vbCrLf   'Set condition area
    		Response.Write "</Script>" & vbcRLf
    		Response.End
    	Else
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtPurGrpNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
    		Response.Write "</Script>" & vbcRLf
    	End If


    	Call SubCloseRs(lgObjRs)
    End If


	'-------------------
	'VAT 코드체크 
	'-------------------
	If Trim(Request("txtVAT")) <> "" Then
    	lgStrSQL = ""
    	Call SubMakeSQLStatements("5", Trim(Request("txtVAT")), "", "")           '☜ : Make sql statements


    	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists

    		Call DisplayMsgBox("175122", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtVATNm.Value  = """"" & vbCrLf   'Set condition area
    		Response.Write "parent.Frm1.txtVAT.focus" & vbCrLf   'Set condition area
    		Response.Write "</Script>" & vbcRLf
    		Response.End
    	Else
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtVATNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
    		Response.Write "</Script>" & vbcRLf
    	End If


    	Call SubCloseRs(lgObjRs)
    End If

	'-------------------
	'결제방법 코드체크 
	'-------------------
	If Trim(Request("txtPayMeth")) <> "" Then
    	lgStrSQL = ""
    	Call SubMakeSQLStatements("6", Trim(Request("txtPayMeth")), "", "")           '☜ : Make sql statements


    	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists

    		Call DisplayMsgBox("200054", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtPayMethNm.Value  = """"" & vbCrLf   'Set condition area
    		Response.Write "parent.Frm1.txtPayMeth.focus" & vbCrLf   'Set condition area
    		Response.Write "</Script>" & vbcRLf
    		Response.End
    	Else
    		Response.Write "<Script Language = VBScript>" & vbCrLf
    		Response.Write "parent.Frm1.txtPayMethNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
    		Response.Write "</Script>" & vbcRLf
    	End If


    	Call SubCloseRs(lgObjRs)
    End If

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


	Dim iSelCount

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Select Case pDataType

		Case "0"
			lgStrSQL = "SELECT IO_Type_Cd,IO_Type_NM  FROM M_Mvmt_type WHERE USAGE_FLG=" & FilterVar("Y", "''", "S") & " AND  IO_Type_Cd =" & FilterVar(pCode, "''", "S")

		Case "1"
			lgStrSQL = "SELECT BP_CD,BP_NM FROM B_Biz_Partner  WHERE Bp_Type in (" & FilterVar("S", "''", "S") & "," & FilterVar("CS", "''", "S") & ") AND usage_flag=" & FilterVar("Y", "''", "S") & "  AND in_out_flag = " & FilterVar("O", "''", "S") & " AND   BP_CD=" & FilterVar(pCode, "''", "S")

		Case "2"
			lgStrSQL = "SELECT PUR_GRP,PUR_GRP_NM  FROM B_Pur_Grp WHERE USAGE_FLG=" & FilterVar("Y", "''", "S") & " AND  PUR_GRP =" & FilterVar(pCode, "''", "S")

		Case "3"
			lgStrSQL = "SELECT IV_TYPE_CD,IV_TYPE_NM FROM M_IV_TYPE  WHERE import_flg=" & FilterVar("N", "''", "S") & " AND  IV_TYPE_CD=" & FilterVar(pCode, "''", "S")

		Case "4"
			lgStrSQL = "SELECT PUR_GRP,PUR_GRP_NM  FROM B_Pur_Grp WHERE USAGE_FLG=" & FilterVar("Y", "''", "S") & " AND  PUR_GRP =" & FilterVar(pCode, "''", "S")

		Case "5"
			lgStrSQL = "SELECT b_minor.MINOR_CD,b_minor.MINOR_NM,b_configuration.REFERENCE  FROM B_MINOR,b_configuration WHERE b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd "
			lgStrSQL = lgStrSQL & "and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1" & " AND  b_minor.MINOR_CD=" & FilterVar(pCode, "''", "S")

		Case "6"
			lgStrSQL = "SELECT b_minor.MINOR_CD,b_minor.MINOR_NM,b_configuration.REFERENCE  FROM B_Minor,b_configuration WHERE b_minor.Major_Cd=" & FilterVar("B9004", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"
			lgStrSQL = lgStrSQL & "AND  b_minor.MINOR_CD=" & FilterVar(pCode, "''", "S")
	End Select


	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>


