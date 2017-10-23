<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->


<% 
Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I", "*","NOCOOKIE","QB")


On Error Resume Next
Err.clear

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag                              '☜ : DBAgent Parameter 선언Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim rs0, rs1, rs2, rs3
Dim lgPageNo                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgstrData, lgstrData2
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim txtNDrAmt1, txtNCrAmt1
Dim strWhere
Const C_SHEETMAXROWS_D  = 100
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

    Call HideStatusWnd 

    lgPageNo			= Request("lgPageNo")                               '☜ : Next key flag
    lgSelectList		= Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT		= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList			= Request("lgTailList")                                 '☜ : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
	Dim intLoopCnt
	On Error Resume Next
	Err.clear

    Dim  ColCnt
    Dim  iStr

	intLoopCnt = rs0.recordcount

	If cint(intLoopCnt) <> 0 Then
		lgstrData = ""
		txtNDrAmt1 = 0
		Do while Not (rs0.EOF Or rs0.BOF)
		    iStr = ""
			For ColCnt = 0 To UBound(lgSelectListDT) - 1 
		        iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			Next
			lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
			rs0.MoveNext
		Loop
		If NOT(rs2.EOF) And NOT(rs2.BOF) Then
			txtNDrAmt1 = rs2(0)
		End If

	End If

	rs0.Close : 	Set rs0 = Nothing 
	rs2.Close : 	Set rs2 = Nothing 

	intLoopCnt = rs1.recordcount

	If cint(intLoopCnt) <> 0 Then
		lgstrData2 = ""
		txtNCrAmt1 = 0
		Do while Not (rs1.EOF Or rs1.BOF)
		    iStr = ""
			For ColCnt = 0 To UBound(lgSelectListDT) - 1 
		        iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs1(ColCnt))
			Next
			lgstrData2      = lgstrData2      & iStr & Chr(11) & Chr(12)
			rs1.MoveNext
		Loop
		If NOT(rs3.EOF) And NOT(rs3.BOF) Then
			txtNCrAmt1 = rs3(0)
		End If
	End If

	rs1.Close :		Set rs1 = Nothing 
	rs3.Close :		Set rs3 = Nothing 

End Sub
'==========================================================================================
' Set DB Agent arg
'==========================================================================================
Sub FixUNISQLData()
	On Error Resume Next
	Err.clear

    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
	
	Dim txtBizAreaCd, txtIssuedDt, txtIssuedDt2, txtAcctCd, txtFromGlDt, txtToGlDt
	Dim strSelect
	
	txtBizAreaCd = UCase(Trim(Request("txtBizAreaCd")))
	txtIssuedDt  = Request("txtIssuedDt")
	txtIssuedDt2 = Request("txtIssuedDt2")
	txtAcctCd    = UCase(Trim(Request("txtAcctCd")))
	txtFromGlDt  = UniConvDate(Request("txtFromGlDt"))
	txtToGlDt    = UniConvDate(Request("txtToGlDt"))
	
	strWhere = ""
	If txtBizAreaCd <> "" Then	strWhere = strWhere & "	AND B.BIZ_AREA_CD = " & Filtervar(txtBizAreaCd, "''", "S")
	If txtIssuedDt	<> "" Then	strWhere = strWhere & "	AND B.ISSUED_DT >= " & Filtervar(UniConvDate(txtIssuedDt), null, "S")
	If txtIssuedDt2 <> "" Then	strWhere = strWhere & "	AND B.ISSUED_DT <= " & Filtervar(UniConvDate(txtIssuedDt2), null, "S")
	If txtAcctCd	<> "" Then	strWhere = strWhere & "	AND A.ACCT_CD = " & Filtervar(txtAcctCd, "''", "S")
	If txtFromGlDt	<> "" Then	strWhere = strWhere & "	AND B.GL_DT >= " & Filtervar(UniConvDate(txtFromGlDt), null, "S")
	If txtToGlDt	<> "" Then	strWhere = strWhere & "	AND B.GL_DT <= " & Filtervar(UniConvDate(txtToGlDt), null, "S")

	strSelect = ""
	'strSelect = strSelect & "SELECT DISTINCT C.BATCH_NO,C.ITEM_SEQ, (CASE C.REVERSE_FG WHEN " & FilterVar("Y", "''", "S") & "  THEN CASE T3.REVERSE_FG WHEN " & FilterVar("M", "''", "S") & "  THEN C.ITEM_LOC_AMT*-1 WHEN " & FilterVar("R", "''", "S") & "  THEN C.ITEM_LOC_AMT ELSE 0 END ELSE C.ITEM_LOC_AMT END) ITEM_LOC_AMT " & vbcr
	strSelect = strSelect & "SELECT DISTINCT FilterVar(C.BATCH_NO, "''", "S"), C.ITEM_SEQ, (CASE C.REVERSE_FG WHEN " & FilterVar("Y", "''", "S") & "  THEN CASE T3.REVERSE_FG WHEN " & FilterVar("M", "''", "S") & "  THEN C.ITEM_LOC_AMT*-1 WHEN " & FilterVar("R", "''", "S") & "  THEN C.ITEM_LOC_AMT ELSE 0 END ELSE C.ITEM_LOC_AMT END) ITEM_LOC_AMT " & vbcr
	
	'strSelect = strSelect & ",T1.TRANS_TYPE,T3.TRANS_NM,C.JNL_CD,E.JNL_NM,ISNULL(C.EVENT_CD,'') EVENT_CD ,ISNULL(F.JNL_NM,'') EVENT_NM " & vbcr
	strSelect = strSelect & ", FilterVar(T1.TRANS_TYPE, "''", "S"), FilterVar(T3.TRANS_NM, "''", "S"), FilterVar(C.JNL_CD, "''", "S"),FilterVar(E.JNL_NM, "''", "S"), FilterVar(ISNULL(C.EVENT_CD,''), "''", "S") EVENT_CD , FilterVar(ISNULL(F.JNL_NM,''), "''", "S") EVENT_NM " & vbcr
	
	strSelect = strSelect & "FROM A_GL_ITEM A  " & vbcr
	strSelect = strSelect & "INNER JOIN A_GL B ON (A.GL_NO = B.GL_NO) " & vbcr

    UNISqlId(0) = "A5144MA107"	'조회 
    UNISqlId(1) = "A5144MA107"	'조회 
    UNISqlId(2) = "A5144MA108"	'조회 
    UNISqlId(3) = "A5144MA108"	'조회 

	Redim UNIValue(3,12)
	
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
	UNIValue(0,1) = strSelect
    UNIValue(0,2) = "DR"
    UNIValue(0,3) = strWhere
    UNIValue(0,4) = strSelect
    UNIValue(0,5) = "DR"
    UNIValue(0,6) = strWhere
    UNIValue(0,7) = strSelect
    UNIValue(0,8) = "DR"
    UNIValue(0,9) = strWhere
    UNIValue(0,10) = "DR"
    UNIValue(0,11) = strWhere

    UNIValue(1,0) = lgSelectList                                          '☜: Select list
	UNIValue(1,1) = strSelect
    UNIValue(1,2) = "CR"
    UNIValue(1,3) = strWhere
    UNIValue(1,4) = strSelect
    UNIValue(1,5) = "CR"
    UNIValue(1,6) = strWhere
    UNIValue(1,7) = strSelect
    UNIValue(1,8) = "CR"
    UNIValue(1,9) = strWhere
    UNIValue(1,10) = "CR"
    UNIValue(1,11) = strWhere

	UNIValue(2,0) = strSelect
    UNIValue(2,1) = "DR"
    UNIValue(2,2) = strWhere
    UNIValue(2,3) = strSelect
    UNIValue(2,4) = "DR"
    UNIValue(2,5) = strWhere
    UNIValue(2,6) = strSelect
    UNIValue(2,7) = "DR"
    UNIValue(2,8) = strWhere
    UNIValue(2,9) = "DR"
    UNIValue(2,10) = strWhere

	UNIValue(3,0) = strSelect
    UNIValue(3,1) = "CR"
    UNIValue(3,2) = strWhere
    UNIValue(3,3) = strSelect
    UNIValue(3,4) = "CR"
    UNIValue(3,5) = strWhere
    UNIValue(3,6) = strSelect
    UNIValue(3,7) = "CR"
    UNIValue(3,8) = strWhere
    UNIValue(3,9) = "CR"
    UNIValue(3,10) = strWhere

    UNIValue(0,12) = UCase(Trim(lgTailList))
    UNIValue(1,12) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	On Error Resume Next
	Err.clear

    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If

    If (rs0.EOF And rs0.BOF) AND (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		rs1.Close:		Set rs1 = Nothing
		rs2.Close:		Set rs2 = Nothing
		rs3.Close:		Set rs3 = Nothing
		Set lgADF = Nothing
		Exit Sub
	Else
		Call  MakeSpreadSheetData()
    End If						
		    
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing    	    

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()


End Sub


'==========================================================================================

%>

<Script Language=vbscript>
	With parent
		.ggoSpread.Source = .frm1.vspdData1 
		.ggoSpread.SSShowData "<%=lgstrData%>"
		.frm1.txtNDrAmt1.text = "<%=UNINumClientFormat(txtNDrAmt1, ggAmtOfMoney.DecPoint, 0)%>"                            '☜: Display data 
		.ggoSpread.Source = .frm1.vspdData2 
		.ggoSpread.SSShowData "<%=lgstrData2%>"                            '☜: Display data 
		.frm1.txtNCrAmt1.text = "<%=UNINumClientFormat(txtNCrAmt1, ggAmtOfMoney.DecPoint, 0)%>"
		.DbQueryOk3
	End with
</Script>
