
<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4214ma1
'*  4. Program Name         : 차입처별 차입금 현황조회 
'*  5. Program Desc         : Query of Loan Balance
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002.04.25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : An, do hyun
'* 10. Modifier (Last)      : 2003.05.19
'* 11. Comment              : 
'=======================================================================================================
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("Q", "A","NOCOOKIE","QB")
%>
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5     '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Const C_SHEETMAXROWS_D = 100

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim	txtLoanType, txtLoanTypeNm, txtLoanPlcNm
Dim cboConfFg, cboApSts, cboLoanFg, txtDocCur, txtLoanPlcFg, txtLoanPlcCd
Dim	txtBaseFromDt, txtBaseToDt
Dim nextDtFr, nextDtTo
Dim txtPreTotAmt, txtPlanTotAmt, txtLoanTotAmt, txtNextTotAmt, txtPayTotAmt, txtBalTotAmt
Dim txtPreTotLocAmt, txtPlanTotLocAmt, txtLoanTotLocAmt, txtNextTotLocAmt, txtPayTotLocAmt, txtBalTotLocAmt
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1
Dim  iLoopCount
Dim  LngMaxRow

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL


'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

    Call HideStatusWnd 

	' 1024 byte 초과로 post 방식으로 전환 
	
	lgPageNo       = UNICInt(Trim(Request("hlgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	    
	lgMaxCount     = C_SHEETMAXROWS_D                           '☜ : 한번에 가져올수 있는 데이타 건수 
	lgSelectList   = Request("hlgSelectList")                               '☜ : select 대상목록 
	lgSelectListDT = Split(Request("hlgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgTailList     = Request("hlgTailList")                                 '☜ : Orderby value
	lgDataExist    = "No"

	LngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

	txtLoanType		= Trim(Request("hLoanType"))
	txtBaseFromDt	= UNIConvDate(Request("hBaseFromDt"))
	txtBaseToDt		= UNIConvDate(Request("hBaseToDt"))
	cboConfFg		= Trim(Request("hConfFg"))
	cboApSts		= Trim(Request("hApSts"))
	cboLoanFg		= Trim(Request("hLoanFg"))
	txtDocCur		= UCase(Trim(Request("hDocCur")))
	txtLoanPlcFg	= Trim(Request("hLoanPlcFg"))
	txtLoanPlcCd	= UCase(Trim(Request("hLoanPlcCd")))
	strBizAreaCd	= Trim(UCase(Request("htxtBizAreaCd")))            '사업장From
	strBizAreaCd1	= Trim(UCase(Request("htxtBizAreaCd1")))            '사업장To

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("hlgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("hlgInternalCd"))
	lgSubInternalCd		= Trim(Request("hlgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("hlgAuthUsrID"))
	
    Call FixUNISQLData()
	Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()

    Dim  RecordCnt
    Dim  ColCnt
    Dim  iRowStr

    lgstrData = ""

    lgDataExist    = "Yes"

    If CDbl(lgPageNo) > 0 Then
       rs0.Move     = CDbl(lgMaxCount) * CDbl(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
       
    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
				
        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(5)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Dim strWhere
	Dim strSUM, strJOIN

    Redim UNIValue(5,19)

    UNISqlId(0) = "F4214MA101"
	If txtLoanPlcFg = "BK" Then
		UNISQLID(1) = "ABANKNM"
	Else
		UNISQLID(1) = "ABPNM"
	End If
	UNISQLID(2) = "AMINORNM"
    UNISqlId(3) = "F4214MA102"
    UNISqlId(4) = "A_GETBIZ"
    UNISqlId(5) = "A_GETBIZ"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	nextDtFr = UNIDateAdd("D",1,txtBaseToDt,gServerDateFormat)
	nextDtTo = UNIDateAdd("M",1,nextDtFr,gServerDateFormat)
	nextDtTo = UNIDateAdd("D",-1,nextDtTo,gServerDateFormat)

	strWhere = ""
	If cboConfFg	= "C" Then	strWhere = strWhere & " and LN.conf_fg IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("E", "''", "S") & " ) " 
	If cboConfFg	= "U" Then	strWhere = strWhere & " and LN.conf_fg   =  " & FilterVar(cboConfFg , "''", "S") & " " 
	If txtLoanPlcFg <> "" Then strWhere = strWhere & " AND LN.loan_plc_type = " & FilterVar(txtLoanPlcFg ,"''"	,"S")
	If cboApSts <> "" Then strWhere = strWhere & " AND LN.rdp_cls_fg = " & FilterVar(cboApSts ,"''"	,"S")
	If cboLoanFg <> "" Then strWhere = strWhere & " AND LN.loan_fg = " & FilterVar(cboLoanFg ,"''"	,"S")
	If txtDocCur <> "" Then strWhere = strWhere & " AND LN.doc_cur = " & FilterVar(txtDocCur ,"''"	,"S")
	If txtLoanPlcCd <> "" Then
		If txtLoanPlcFg = "BK" Then
			strWhere = strWhere & " and LN.Loan_bank_cd = " & FilterVar(txtLoanPlcCd ,"''"	,"S")
		Else
			strWhere = strWhere & " and LN.bp_cd = " & FilterVar(txtLoanPlcCd ,"''"	,"S")
		End If
	End If
	If txtLoanType <> "" Then strWhere = strWhere & " AND LN.loan_type = " & FilterVar(txtLoanType ,"''"	,"S")
	
	if strBizAreaCd <> "" then
		strWhere = strWhere & " AND LN.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere = strWhere & " AND LN.BIZ_AREA_CD >= " & FilterVar("", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere = strWhere & " AND LN.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere = strWhere & " AND LN.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND LN.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND LN.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND LN.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND LN.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If

	strWhere	= strWhere	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL


	strJOIN = "LEFT JOIN b_bank BK ON LN.loan_bank_cd = BK.bank_cd "
	strJOIN = strJOIN & "LEFT JOIN B_BIZ_PARTNER BP ON LN.bp_cd = BP.bp_cd "
	strJOIN = strJOIN & "LEFT JOIN b_minor MI ON LN.loan_type = MI.minor_cd AND MI.major_cd = " & FilterVar("F1000", "''", "S") & "  "
	
	strSUM = ""
	strSUM = strSUM & " SUM(ISNULL(PL1.PLAN_AMT, 0))-SUM(ISNULL(PM1.PAY_AMT, 0)) PreTotAmt "
	strSUM = strSUM & " ,SUM(ISNULL(PL2.PLAN_AMT, 0)) PlanTotAmt "
	strSUM = strSUM & " ,SUM(ISNULL(PM2.PAY_AMT, 0)) + SUM(ISNULL(LN1.bas_amt, 0)) PayTotAmt "
	strSUM = strSUM & " ,SUM(ISNULL(LN1.LOAN_AMT, 0)) LoanTotAmt "
	strSUM = strSUM & " ,SUM(ISNULL(LN1.LOAN_AMT, 0))-(SUM(ISNULL(PM2.PAY_AMT, 0)) + SUM(ISNULL(LN1.bas_amt, 0))) + (SUM(ISNULL(PL1.PLAN_AMT, 0))-SUM(ISNULL(PM1.PAY_AMT, 0))) - SUM(ISNULL(HI.CHG_AMT,0)) BalTotAmt "
	strSUM = strSUM & " ,SUM(ISNULL(PL3.PLAN_AMT, 0)) NextTotAmt "
	strSUM = strSUM & " ,SUM(ISNULL(PL1.PLAN_LOC, 0))-SUM(ISNULL(PM1.PAY_LOC, 0)) PreTotLocAmt "
	strSUM = strSUM & " ,SUM(ISNULL(PL2.PLAN_LOC, 0)) PlanTotLocAmt "
	strSUM = strSUM & " ,SUM(ISNULL(PM2.PAY_LOC, 0)) + SUM(ISNULL(LN1.bas_loc, 0)) PayTotLocAmt "
	strSUM = strSUM & " ,SUM(ISNULL(LN1.LOAN_LOC, 0)) LoanTotLocAmt "
	strSUM = strSUM & " ,SUM(ISNULL(LN1.LOAN_LOC, 0))-(SUM(ISNULL(PM2.PAY_LOC, 0)) + SUM(ISNULL(LN1.bas_loc, 0))) + (SUM(ISNULL(PL1.PLAN_LOC, 0))-SUM(ISNULL(PM1.PAY_LOC, 0))) - SUM(ISNULL(HI.CHG_LOC_AMT,0))  BalTotLocAmt"
	strSUM = strSUM & " ,SUM(ISNULL(PL3.PLAN_LOC, 0))  NextTotLocAmt "

	UNIValue(0,1) = strJOIN
    UNIValue(0,2) = FilterVar(txtBaseFromDt ,""       ,"S") 
    UNIValue(0,3) = FilterVar(txtBaseFromDt ,""       ,"S") 
    UNIValue(0,4) = FilterVar(txtBaseFromDt ,""       ,"S") 
    UNIValue(0,5) = FilterVar(txtBaseToDt ,""       ,"S") 
    UNIValue(0,6) = FilterVar(txtBaseFromDt ,""       ,"S") 
    UNIValue(0,7) = FilterVar(txtBaseToDt ,""       ,"S") 
    UNIValue(0,8) = FilterVar(txtBaseFromDt ,""       ,"S") 
    UNIValue(0,9) = FilterVar(txtBaseToDt ,""       ,"S") 
    UNIValue(0,10) = FilterVar(txtBaseFromDt ,""       ,"S") 
    UNIValue(0,11) = FilterVar(txtBaseToDt ,""       ,"S") 
    UNIValue(0,12) = FilterVar(txtBaseFromDt ,""       ,"S") 
    UNIValue(0,13) = FilterVar(txtBaseToDt ,""       ,"S") 
    UNIValue(0,14) = FilterVar(txtBaseToDt ,""       ,"S") 
    UNIValue(0,15) = FilterVar(nextDtTo ,""       ,"S") 
    UNIValue(0,16) = FilterVar(txtBaseFromDt ,""       ,"S") 
    UNIValue(0,17) = FilterVar(txtBaseToDt ,""       ,"S") 
    UNIValue(0,18) = strWhere
    
	If txtLoanPlcFg = "BK" Then
		UNIValue(1,0) = FilterVar(txtLoanPlcCd ,""       ,"S")
	Else
		UNIValue(1,0) = FilterVar(txtLoanPlcCd ,"''"       ,"S")
	End If
	UNIValue(2,0) = FilterVar("F1000" ,""       ,"S")
	UNIValue(2,1) = FilterVar(txtLoanType ,""       ,"S")
    
    UNIValue(3,0) = strSUM
    UNIValue(3,1) = strJOIN
    UNIValue(3,2) = FilterVar(txtBaseFromDt, "''", "S") 
    UNIValue(3,3) = FilterVar(txtBaseFromDt, "''", "S") 
    UNIValue(3,4) = FilterVar(txtBaseFromDt, "''", "S") 
    UNIValue(3,5) = FilterVar(txtBaseToDt, "''", "S") 
    UNIValue(3,6) = FilterVar(txtBaseFromDt, "''", "S") 
    UNIValue(3,7) = FilterVar(txtBaseToDt, "''", "S") 
    UNIValue(3,8) = FilterVar(txtBaseFromDt, "''", "S") 
    UNIValue(3,9) = FilterVar(txtBaseToDt, "''", "S") 
    UNIValue(3,10) = FilterVar(txtBaseFromDt, "''", "S") 
    UNIValue(3,11) = FilterVar(txtBaseToDt, "''", "S") 
	
	'2003/12/12 Oh Soo Min 수정    
    UNIValue(3,12) = FilterVar(txtBaseToDt, "''", "S") 
    UNIValue(3,13) = FilterVar(nextDtTo, "''", "S") 
    UNIValue(3,14) = FilterVar(txtBaseToDt, "''", "S") 
    UNIValue(3,15) = FilterVar(txtBaseFromDt, "''", "S") 
    UNIValue(3,16) = FilterVar(txtBaseToDt, "''", "S") 
    UNIValue(3,17) = strWhere
    
	UNIValue(4,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(5,0)  = FilterVar(strBizAreaCd1, "''", "S")

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    'rs1
	If txtLoanPlcCd <> "" Then
	    If Not (rs1.EOF OR rs1.BOF) Then
			txtLoanPlcNm = Trim(rs1(1))
		Else
			txtLoanPlcNm = ""
			If txtLoanPlcFg = "BK" Then
				Call DisplayMsgBox("800123", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
			ElseIf txtLoanPlcFg = "BP" Then
				Call DisplayMsgBox("126100", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
			End If			
	        rs1.Close
		    Set rs1 = Nothing 
			Exit sub
		End IF
		rs1.Close
		Set rs1 = Nothing
	End If
	
    'rs2
	If txtLoanType <> "" Then
	    If Not (rs2.EOF OR rs2.BOF) Then
			txtLoanTypeNm = Trim(rs2(1))
		Else
			txtLoanTypeNm = ""
			Call DisplayMsgBox("140936", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs2.Close
		    Set rs2 = Nothing 
			Exit sub
		End IF
		rs2.Close
		Set rs2 = Nothing
	End If

    'rs4
    If Not( rs4.EOF OR rs4.BOF) Then
   		strBizAreaCd = Trim(rs4(0))
		strBizAreaNm = Trim(rs4(1))
	Else
		strBizAreaCd = ""
		strBizAreaNm = ""
		
    End IF
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBizAreaCd.value = "<%=strBizAreaCd%>"
			.txtBizAreaNm.value = "<%=strBizAreaNm%>"
		End With
		</Script>
<%   
    rs4.Close
    Set rs4 = Nothing
    
    ' rs5
    If Not( rs5.EOF OR rs5.BOF) Then
   		strBizAreaCd1 = Trim(rs5(0))
		strBizAreaNm1 = Trim(rs5(1))
	Else
		strBizAreaCd1 = ""
		strBizAreaNm1 = ""
		
    End IF
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBizAreaCd1.value = "<%=strBizAreaCd1%>"
			.txtBizAreaNm1.value = "<%=strBizAreaNm1%>"
		End With
		</Script>
<%
    rs5.Close
    Set rs5 = Nothing


    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("140900", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()

	    'rs3
		If Not (rs3.EOF OR rs3.BOF) Then
			txtPreTotAmt = Trim(rs3(0))
			txtPlanTotAmt = Trim(rs3(1))
			txtPayTotAmt = Trim(rs3(2))
			txtLoanTotAmt = Trim(rs3(3))
			txtBalTotAmt = Trim(rs3(4))
			txtNextTotAmt = Trim(rs3(5))
			txtPreTotLocAmt = Trim(rs3(6))			'Local Amount
			txtPlanTotLocAmt = Trim(rs3(7))
			txtPayTotLocAmt = Trim(rs3(8))
			txtLoanTotLocAmt = Trim(rs3(9))
			txtBalTotLocAmt = Trim(rs3(10))
			txtNextTotLocAmt = Trim(rs3(11))

		Else
			txtPreTotAmt = ""
			txtPreTotLocAmt = ""
			txtPlanTotAmt = ""
			txtPlanTotLocAmt = ""
			txtLoanTotAmt = ""
			txtLoanTotLocAmt = ""
			txtNextTotAmt = ""
			txtNextTotLocAmt = ""
			txtPayTotAmt = ""
			txtPayTotLocAmt = ""
			txtBalTotAmt = ""
			txtBalTotLocAmt = ""
		End IF

		rs3.Close
		Set rs3 = Nothing
    End If
    
End Sub

%>
<Script Language=vbscript>
With Parent
	If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.hConfFg.value		= "<%=ConvSPChars(cboConfFg)%>" 
			.frm1.hApSts.value		= "<%=ConvSPChars(cboApSts)%>" 
			.frm1.hLoanFg.value		= "<%=ConvSPChars(cboLoanFg)%>" 
			.frm1.hDocCur.value		= "<%=ConvSPChars(txtDocCur)%>" 
			.frm1.hLoanPlcFg.value	= "<%=ConvSPChars(txtLoanPlcFg)%>" 
			.frm1.hLoanPlcCd.value	= "<%=ConvSPChars(txtLoanPlcCd)%>" 
			.Frm1.hLoanType.Value		= "<%=ConvSPChars(txtLoanType)%>"                  'For Next Search
			.Frm1.hBaseFromDt.Value	= "<%=txtBaseFromDt%>"                  'For Next Search
			.Frm1.hBaseToDt.Value		= "<%=txtBaseFromDt%>"                  'For Next Search
			.Frm1.htxtBizAreaCd.value = Trim(.Frm1.txtBizAreaCd.value)
			.Frm1.htxtBizAreaCd1.value = Trim(.Frm1.txtBizAreaCd1.value)
       End If
'		If "<%=txtDocCur%>"   <> "" Then
'			.frm1.txtPreTotAmt.Text		= "<%=UNINumClientFormat(txtPreTotAmt,ggAmtOfMoney.DecPoint, 0)%>"			'rs3 값 
'			.frm1.txtPlanTotAmt.Text	= "<%=UNINumClientFormat(txtPlanTotAmt,ggAmtOfMoney.DecPoint, 0)%>"
'			.frm1.txtLoanTotAmt.Text	= "<%=UNINumClientFormat(txtLoanTotAmt,ggAmtOfMoney.DecPoint, 0)%>"
'			.frm1.txtNextTotAmt.Text	= "<%=UNINumClientFormat(txtNextTotAmt,ggAmtOfMoney.DecPoint, 0)%>"
'			.frm1.txtPayTotAmt.Text		= "<%=UNINumClientFormat(txtPayTotAmt,ggAmtOfMoney.DecPoint, 0)%>"
'			.frm1.txtBalTotAmt.Text		= "<%=UNINumClientFormat(txtBalTotAmt,ggAmtOfMoney.DecPoint, 0)%>"
'		Else
			.frm1.txtPreTotAmt.Text		= "<%=UNINumClientFormat(txtPreTotAmt,2, 0)%>"			'rs3 값 
			.frm1.txtPlanTotAmt.Text	= "<%=UNINumClientFormat(txtPlanTotAmt,2, 0)%>"
			.frm1.txtLoanTotAmt.Text	= "<%=UNINumClientFormat(txtLoanTotAmt,2, 0)%>"
			.frm1.txtNextTotAmt.Text	= "<%=UNINumClientFormat(txtNextTotAmt,2, 0)%>"
			.frm1.txtPayTotAmt.Text		= "<%=UNINumClientFormat(txtPayTotAmt,2, 0)%>"
			.frm1.txtBalTotAmt.Text		= "<%=UNINumClientFormat(txtBalTotAmt,2, 0)%>"
'		End If

		.frm1.txtPreTotAmt.Text		= "<%=UNINumClientFormat(txtPreTotAmt,2, 0)%>"			'rs3 값 
		.frm1.txtPlanTotAmt.Text	= "<%=UNINumClientFormat(txtPlanTotAmt,2, 0)%>"
		.frm1.txtLoanTotAmt.Text	= "<%=UNINumClientFormat(txtLoanTotAmt,2, 0)%>"
		.frm1.txtNextTotAmt.Text	= "<%=UNINumClientFormat(txtNextTotAmt,2, 0)%>"
		.frm1.txtPayTotAmt.Text		= "<%=UNINumClientFormat(txtPayTotAmt,2, 0)%>"
		.frm1.txtBalTotAmt.Text		= "<%=UNINumClientFormat(txtBalTotAmt,2, 0)%>"
		.frm1.txtPreTotLocAmt.Text	= "<%=UNINumClientFormat(txtPreTotLocAmt,ggAmtOfMoney.DecPoint, 0)%>"			'rs3 값 
		.frm1.txtPlanTotLocAmt.Text	= "<%=UNINumClientFormat(txtPlanTotLocAmt,ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtLoanTotLocAmt.Text	= "<%=UNINumClientFormat(txtLoanTotLocAmt,ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtNextTotLocAmt.Text	= "<%=UNINumClientFormat(txtNextTotLocAmt,ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtPayTotLocAmt.Text	= "<%=UNINumClientFormat(txtPayTotLocAmt,ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtBalTotLocAmt.Text	= "<%=UNINumClientFormat(txtBalTotLocAmt,ggAmtOfMoney.DecPoint, 0)%>"

       .ggoSpread.Source  = .frm1.vspdData
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '☜ : Display data
		.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",2),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",3),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",4),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",5),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",6),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",7),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",1),parent.GetKeyPos("A",8),   "A" ,"I","X","X")
		.frm1.vspdData.Redraw = True
    End If

	.DbQueryOk()
	.frm1.txtLoanTypeNm.value = "<%=txtLoanTypeNm%>"			'rs1 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기 
	.frm1.txtLoanPlcNm.value = "<%=txtLoanPlcNm%>"			'rs2 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기 
	
End With
</Script>

