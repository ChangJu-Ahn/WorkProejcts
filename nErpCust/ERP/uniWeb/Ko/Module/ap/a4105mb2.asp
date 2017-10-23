<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next
Err.Clear 
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strDueDt															
Dim strDueFg
Dim strDocCur																
Dim strPayBpCd	
Dim strPaymType																
DIm tempApno
Dim strCond
Dim strCond2
Dim strCond3
DIm strApno
Dim strBatchAllcNo
Dim strAllcDt
Dim StrMaxRows
Dim StrFlgMode

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")   
    Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB") 

    Call HideStatusWnd() 
	Call SplitRequest()
	
   	lgErrorStatus = ""
    lgIntFlgMode = StrFlgMode

	Call MakeSQLWhere()
	Call SubOpenDB(lgObjConn)  
	Call SubBizQuerySingle()
    Call SubCloseDB(lgObjConn) 

'============================================================================================================
' Name : SubBizQuerySingle
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuerySingle()
	Dim lgStrSQL
	Dim lgLngMaxRow
	Dim iDx

	On Error Resume Next                                                                 '☜: Protect system from crashing
	Err.Clear                                                                            '☜: Clear Error status

    If Cstr(lgIntFlgMode) = Cstr(OPMD_UMODE) Then
		lgStrSQL = lgStrSQL & "SELECT * FROM ("
		lgStrSQL = lgStrSQL & " SELECT " & FilterVar("1", "''", "S") & "  is_checked, A.pay_bp_cd, B.bp_nm, A.ap_no, A.ap_dt, A.ap_due_dt, A.doc_cur, A.ap_amt,A.cls_amt, A.bal_amt, "
		lgStrSQL = lgStrSQL & " CASE WHEN A.ap_due_dt > " & FilterVar(strAllcDt ,null,"S") & " THEN " & FilterVar("IN", "''", "S") & "  ELSE " & FilterVar("OVER", "''", "S") & "  END over_due_fg, A.acct_cd, A.ap_desc "
		lgStrSQL = lgStrSQL & " FROM ufn_A_ClsApByBatchPayment(" & FilterVar(strBatchAllcNo ,null,"S") & ") A"
		lgStrSQL = lgStrSQL & "	LEFT JOIN b_biz_partner B ON B.bp_cd = A.pay_bp_cd"
		lgStrSQL = lgStrSQL & " WHERE A.ap_no IN (" & strApno & ") "
		lgStrSQL = lgStrSQL & " UNION ALL"
		lgStrSQL = lgStrSQL & " SELECT " & FilterVar("0", "''", "S") & "  is_checked, A.pay_bp_cd, B.bp_nm, A.ap_no, A.ap_dt, A.ap_due_dt, A.doc_cur, A.ap_amt, " & FilterVar("0", "''", "S") & "  cls_amt, A.cls_amt bal_amt,"
		lgStrSQL = lgStrSQL & " CASE WHEN A.ap_due_dt >" & FilterVar(strAllcDt ,null,"S") & " THEN " & FilterVar("IN", "''", "S") & "  ELSE " & FilterVar("OVER", "''", "S") & "  END over_due_fg, A.acct_cd, A.ap_desc"
		lgStrSQL = lgStrSQL & " FROM ufn_A_ClsApByBatchPayment(" & FilterVar(strBatchAllcNo ,null,"S") & ") A"
		lgStrSQL = lgStrSQL & "	LEFT JOIN b_biz_partner B ON B.bp_cd = A.pay_bp_cd"
		lgStrSQL = lgStrSQL & " WHERE A.ap_no NOT IN (" & strApno & ")) TMP "
		lgStrSQL = lgStrSQL & strcond3
		lgStrSQL = lgStrSQL & " ORDER BY ap_due_dt asc , ap_no asc "
	Else
		lgStrSQL = lgStrSQL & "SELECT * FROM ("
		lgStrSQL = lgStrSQL & " SELECT " & FilterVar("1", "''", "S") & "  is_checked, A.PAY_BP_CD, B.bp_nm, A.ap_no, A.ap_dt, A.ap_due_dt, A.doc_cur, A.ap_amt, A.bal_amt cls_amt, A.bal_amt,"
		lgStrSQL = lgStrSQL & " CASE WHEN A.ap_due_dt > " & FilterVar(strAllcDt ,null,"S") & " THEN " & FilterVar("IN", "''", "S") & "  ELSE " & FilterVar("OVER", "''", "S") & "  END over_due_fg, A.acct_cd, A.ap_desc"
		lgStrSQL = lgStrSQL & " FROM a_open_ap A LEFT JOIN b_biz_partner B ON B.bp_cd = A.pay_bp_cd"
		lgStrSQL = lgStrSQL & " WHERE A.ap_amt >0  AND A.conf_fg=" & FilterVar("C", "''", "S") & "  and A.gl_no<>''" & strcond 
		lgStrSQL = lgStrSQL & " UNION ALL"
		lgStrSQL = lgStrSQL & " SELECT " & FilterVar("0", "''", "S") & "  is_checked, A.PAY_BP_CD, B.bp_nm, A.ap_no, A.ap_dt, A.ap_due_dt, A.doc_cur, A.ap_amt, " & FilterVar("0", "''", "S") & "  cls_amt, a.bal_amt,"
		lgStrSQL = lgStrSQL & " CASE WHEN A.ap_due_dt > " & FilterVar(strAllcDt ,null,"S") & " THEN " & FilterVar("IN", "''", "S") & "  ELSE " & FilterVar("OVER", "''", "S") & "  END over_due_fg, A.acct_cd, A.ap_desc"
		lgStrSQL = lgStrSQL & " FROM a_open_ap A LEFT JOIN b_biz_partner B ON B.bp_cd = A.pay_bp_cd"
		lgStrSQL = lgStrSQL & " WHERE A.ap_sts = " & FilterVar("O", "''", "S") & "  AND A.ap_amt >0 AND A.conf_fg=" & FilterVar("C", "''", "S") & "  and A.gl_no<>''" & strcond2 & ") AP"
		lgStrSQL = lgStrSQL & " ORDER BY AP.ap_due_dt asc , ap_no asc "
	End If	

'	Response.Write lgStrSQL
'	Response.End
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
		lgStrPrevKey = ""
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
		lgErrorStatus = "YES"
		Exit Sub 
	Else
		iDx = 1
		lgstrData = ""
		lgLngMaxRow = StrMaxRows														'☜: Read Operation Mode (CRUD)

		Do While Not lgObjRs.EOF
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IS_CHECKED"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_BP_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("AP_NO"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("AP_DT"))          
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("AP_DUE_DT"))   
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOC_CUR"))
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("AP_AMT"),ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BAL_AMT"),ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("CLS_AMT"),ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OVER_DUE_FG"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("AP_DESC"))
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
			
			lgObjRs.MoveNext
			iDx =  iDx + 1
   		Loop 
	End If

	Call SubCloseRs(lgObjRs)															'☜: Release RecordSSet

	If CheckSYSTEMError(Err,True) = True Then
		ObjectContext.SetAbort
		Exit Sub
	End If   
End Sub       
   
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  MakeSQLWhere()
	Dim arrVal
	Dim ii
	Dim temp
	strApno=""

	arrVal=Split(tempApno, vbTab) 
	
	If UBound(arrVal) > 1 Then	
		For ii = 1 To UBound(arrVal) - 1
			If ii<>UBound(arrVal)-1 Then
				strApno= strApno & FilterVar(arrVal(ii),null,"S") & ","
			Else
				strApno= strApno & FilterVar(arrVal(ii),null,"S") 
			End If
		Next
	Else
		strApno="''"
	End If

	If strDueFg Then
		strCond = " AND A.ap_due_dt = " & FilterVar(strDueDt , "''", "S")
	Else
		strCond = " AND A.ap_due_dt <=" & FilterVar(strDueDt , "''", "S")
	End If
	
	strCond = strCond & " AND A.doc_cur = " & FilterVar(strDocCur , "''", "S")
	strCond = strCond & " And A.pay_bp_cd= "	& FilterVar(strPayBpCd , "''", "S")
	If strPaymType <> "" Then strCond = strCond & " And A.paym_type = " & FilterVar(strPaymType , "''", "S")	
	
	
	strcond2= strcond
	strcond = strCond & "AND A.ap_no IN (" & strApno & ")"
	strcond2= strcond2 & "AND A.ap_no NOt IN (" & strApno & ")"
	If strPaymType <> "" Then strCond2 = strCond2 & " And A.paym_type = " & FilterVar(strPaymType , "''", "S")
	

	strCond3 = strCond3 & " WHERE doc_cur = " & FilterVar(strDocCur , "''", "S")
	strCond3 = strCond3 & " And pay_bp_cd= "	& FilterVar(strPayBpCd , "''", "S")
End Sub

Sub SplitRequest()
	Dim arrRequest,arrtemp0,arrTemp,arrTemp1,arrTemp2

	arrTemp0 = Trim(Request.Form("hstrVal"))

	arrTemp = Split(arrTemp0,"?")
	arrTemp1 = split(arrtemp(1),"&")

	Redim arrRequest(9)
	For ii= 0 to UBound(arrtemp1,1) 
		arrTemp2 = Split(arrTemp1(ii),"=")
		arrRequest(ii) = arrTemp2(1)
	Next			

	StrFlgMode	    = Trim(arrRequest(0))    
	strDueDt		= UNIConvDate(arrRequest(2))
	strDueFg		= Trim(arrRequest(3))    
	strPayBpCd		= Trim(arrRequest(1))   
	strDocCur		= Trim(arrRequest(4))  
	tempApno		= Trim(arrRequest(5))
	strBatchAllcNo  = Trim(arrRequest(6))
	strAllcDt		= UNIConvDate(arrRequest(7)) 
	strPaymType     = Trim(arrRequest(8))	
	strmaxRows      = Trim(arrRequest(9))
End Sub

%>

<Script Language=vbscript>
 If "<%=lgErrorStatus%>" = "" Then    
       'Set condition data to hidden area
       Parent.ggoSpread.Source  = Parent.frm1.vspdData2
       Parent.ggoSpread.SSShowData "<%=lgstrData%>"            '☜ : Display data
       Parent.DbQueryOk2
    End If  
</Script>	


