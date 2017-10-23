<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'********************************************************************************************************
'*  1. Module Name          : Sales																		*
'*  2. Function Name        : 판매경비관리																*
'*  3. Program ID           : S6111MA1																	*
'*  4. Program Name         : 판매경비등록																*
'*  5. Program Desc         :																			*
'*  6. Comproxy List        : PS9G111.dll, PS9G118.dll, PB0C003.dll, PB0C004.dll
'*  7. Modified date(First) : 2000/04/26																*
'*  8. Modified date(Last)  : 2001/12/19																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : Cho Sung Hyun																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/26 : 화면 design												*
'*							  2. 2000/09/22 : 4th Coding Start											*
'*							  3. 2001/12/19 : Date 표준적용												*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

<%'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================%>
Const BIZ_PGM_ID = "s6111mb1_lko441.asp"												'☆: Head Query 비지니스 로직 ASP명 

Const ConXchRate = "ConXchRate"
Const ConVatType = "ConVatType"

'☆: Spread Sheet의 Column별 상수 
Dim C_PostingFlg 			'확정여부 
Dim C_ChargeNo 				'경비관리번호 
Dim C_ChargeType 			'경비항목 
Dim C_ChargeTypePop 		'경비항목팝업 
Dim C_ChargeTypeNm			'경비항목명 
Dim C_BpCd 					'거래처 
Dim C_BpPop 				'거래처팝업 
Dim C_BpNm					'거래처명 
'2008-04-21 6:07오후 :: hanc
Dim C_ext1_Cd 					'거래처 
Dim C_ext1_Pop 				'거래처팝업 
Dim C_ext1_Nm					'거래처명 
Dim C_ChargeDt				'발생일 
Dim C_VatType 				'계산서종류 
Dim C_VatTypePop 			'계산서종류팝업 
Dim C_VatTypeNm				'계산서종류명 
Dim C_Curr					'화폐 
Dim C_CurrPop				'화폐팝업 
Dim C_ChargeDocAmt			'발생금액 
Dim C_XchCalop				'환율연산자 
Dim C_XchRate				'환율 
Dim C_ChargeLocAmt			'발생자국금액 
Dim C_VatRate				'VAT율 
Dim C_VatDocAmt 			'VAT금액 
Dim C_VatLocAmt 			'VAT자국금액 
Dim C_TaxBizArea 			'세금신고사업장 
Dim C_TaxBizAreaPop 		'세금신고사업장팝업 
Dim C_TaxBizAreaNm			'세금신고사업장명 
Dim C_PayDueDt				'지급만기일 
Dim C_PayType				'지급유형 
Dim C_PayTypePop			'지급유형팝업 
Dim C_PayTypeNm 			'지급유형명 
Dim C_PayDocAmt 			'지급액 
Dim C_PayLocAmt 			'지급자국액 
Dim C_CheckNo				'어음번호 
Dim C_CheckNoPop			'어음번호팝업 
Dim C_BankCd 				'출금은행코드 
Dim C_BankPop				'출금은행팝업 
Dim C_BankNm 				'출금은행명 
Dim C_BankAcct 				'출금계좌 
Dim C_BankAcctPop 			'출금계좌팝업 
Dim C_RefRemark				'기타참조사항 

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim	lgBlnVatChangedFlag		' 부가세유형의 변경여부 
Dim IsOpenPop				' Popup

'========================================================================================================
Sub initSpreadPosVariables()  
	C_PostingFlg = 1		'확정여부 
	C_ChargeNo = 2			'경비관리번호 
	C_ChargeType = 3		'경비항목 
	C_ChargeTypePop = 4		'경비항목팝업 
	C_ChargeTypeNm = 5		'경비항목명 
	C_BpCd = 6				'거래처 
	C_BpPop = 7				'거래처팝업 
	C_BpNm = 8				'거래처명 
	C_ChargeDt = 9			'발생일 
	C_VatType = 10			'계산서종류 
	C_VatTypePop = 11		'계산서종류팝업 
	C_VatTypeNm = 12		'계산서종류명 
	C_Curr = 13				'화폐 
	C_CurrPop = 14			'화폐팝업 
	C_ChargeDocAmt = 15		'발생금액 
	C_XchCalop = 16			'환율연산자 
	C_XchRate = 17			'환율 
	C_ChargeLocAmt = 18		'발생자국금액 
	C_VatRate = 19			'VAT율 
	C_VatDocAmt = 20		'VAT금액 
	C_VatLocAmt = 21		'VAT자국금액 
	C_TaxBizArea = 22		'세금신고사업장 
	C_TaxBizAreaPop = 23	'세금신고사업장팝업 
	C_TaxBizAreaNm = 24		'세금신고사업장명 
	C_PayDueDt = 25			'지급만기일 
	C_PayType = 26			'지급유형 
	C_PayTypePop = 27		'지급유형팝업 
	C_PayTypeNm = 28		'지급유형명 
	C_PayDocAmt = 29		'지급액 
	C_PayLocAmt = 30		'지급자국액 
	C_CheckNo = 31			'어음번호 
	C_CheckNoPop = 32		'어음번호팝업 
	C_BankCd = 33			'출금은행코드 
	C_BankPop = 34			'출금은행팝업 
	C_BankNm = 35			'출금은행명 
	C_BankAcct = 36			'출금계좌 
	C_BankAcctPop = 37		'출금계좌팝업 
'2008-04-21 5:59오후 :: hanc
	C_ext1_Cd = 38				'세금계산서거래처 
	C_ext1_Pop = 39				'세금계산서거래처팝업 
	C_ext1_Nm = 40				'세금계산서거래처명 
	C_RefRemark = 41		'기타참조사항 

End Sub

'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0  

End Sub

'========================================================================================================
Sub SetDefaultVal()
	frm1.txtConProcessStepCd.focus
	lgBlnFlgChgValue = False
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	
	With frm1.vspdData

	    .MaxCols = C_ext1_Nm + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	    .Col = .MaxCols															'☜: 공통콘트롤 사용 Hidden Column
	    .ColHidden = True

	    .MaxRows = 0
	    ggoSpread.Source = frm1.vspdData

		.ReDraw = false
		
       ggoSpread.Spreadinit "V20030301",,parent.gAllowDragDropSpread    

       Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCheck C_PostingFlg, "", 5,,,True
	    ggoSpread.SSSetEdit C_ChargeNo, "경비관리번호",18,,,18,2
	    ggoSpread.SSSetEdit C_ChargeType, "경비항목",15,,,20,2
	    ggoSpread.SSSetButton C_ChargeTypePop
	    ggoSpread.SSSetEdit C_ChargeTypeNm, "경비항목명", 15
	    ggoSpread.SSSetEdit C_BpCd, "거래처", 15,,,10,2
	    ggoSpread.SSSetButton C_BpPop
	    ggoSpread.SSSetEdit C_BpNm, "거래처명", 15
		ggoSpread.SSSetDate C_ChargeDt, "발생일",10,2,parent.gDateFormat
	    ggoSpread.SSSetEdit C_VatType, "VAT유형",13,,,5,2
	    ggoSpread.SSSetButton C_VatTypePop
	    ggoSpread.SSSetEdit C_VatTypeNm, "VAT유형명",20
	    ggoSpread.SSSetEdit C_Curr, "화폐",10,,,3,2
	    ggoSpread.SSSetButton C_CurrPop
	    ggoSpread.SSSetFloat C_ChargeDocAmt,"발생금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit C_XchCalop, "환율연산자",15
	    ggoSpread.SSSetFloat C_XchRate,"환율",15,parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat C_ChargeLocAmt,"발생자국금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat C_VatRate,"VAT율" ,15, parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat C_VatDocAmt, "VAT금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat C_VatLocAmt, "VAT자국금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetDate C_PayDueDt, "지급만기일",15,2,parent.gDateFormat
	    ggoSpread.SSSetEdit C_TaxBizArea, "세금신고사업장",18,,,10,2
	    ggoSpread.SSSetButton C_TaxBizAreaPop
	    ggoSpread.SSSetEdit C_TaxBizAreaNm, "세금신고사업장명",25
	    ggoSpread.SSSetEdit C_PayType, "지급유형",10,,,5,2
	    ggoSpread.SSSetButton C_PayTypePop
	    ggoSpread.SSSetEdit C_PayTypeNm, "지급유형명",15
	    ggoSpread.SSSetFloat C_PayDocAmt, "지급금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat C_PayLocAmt, "지급자국금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit C_CheckNo, "어음번호", 20,,,30,2
		ggoSpread.SSSetButton C_CheckNoPop
		ggoSpread.SSSetEdit C_BankCd, "출금은행", 10,,,10,2
		ggoSpread.SSSetButton C_BankPop
		ggoSpread.SSSetEdit C_BankNm, "출금은행명", 20
	    ggoSpread.SSSetEdit C_BankAcct, "출금계좌", 30,,,30,2
	    ggoSpread.SSSetButton C_BankAcctPop
		ggoSpread.SSSetEdit C_RefRemark, "기타참조사항", 50
		'2008-04-21 6:00오후 :: hanc
	    ggoSpread.SSSetEdit C_ext1_Cd, "세금계산서거래처", 15,,,10,2
	    ggoSpread.SSSetButton C_ext1_Pop
	    ggoSpread.SSSetEdit C_ext1_Nm, "세금계산서거래처명", 15
		
		Call ggoSpread.MakePairsColumn(C_ChargeType,C_ChargeTypePop)
		Call ggoSpread.MakePairsColumn(C_BpCd,C_BpPop)
		Call ggoSpread.MakePairsColumn(C_VatType,C_VatTypePop)
		Call ggoSpread.MakePairsColumn(C_Curr,C_CurrPop)
		Call ggoSpread.MakePairsColumn(C_TaxBizArea,C_TaxBizAreaPop)
		Call ggoSpread.MakePairsColumn(C_PayType,C_PayTypePop)
		Call ggoSpread.MakePairsColumn(C_CheckNo,C_CheckNoPop)
		Call ggoSpread.MakePairsColumn(C_BankCd,C_BankPop)
		Call ggoSpread.MakePairsColumn(C_BankAcct,C_BankAcctPop)
		Call ggoSpread.MakePairsColumn(C_ext1_Cd,C_ext1_Pop)
       
		.ReDraw = true
   
    End With
    
End Sub

'========================================================================================================
Sub SetSpreadLock()
End Sub

'========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	Dim i 
	    
    With frm1
    
    .vspdData.ReDraw = False

    ggoSpread.SSSetRequired	C_ChargeType, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ChargeTypeNm, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_BpCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BpNm, pvStartRow, pvEndRow
    '2008-04-21 6:01오후 :: hanc
    ggoSpread.SSSetRequired C_ext1_Cd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ext1_Nm, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_ChargeDt, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_VatTypeNm, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_Curr, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_ChargeDocAmt, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_VatRate, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_PayTypeNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BankNm, pvStartRow, pvEndRow

    ggoSpread.SSSetProtected C_XchCalop, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_PayDueDt, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_TaxBizArea, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_TaxBizAreaNm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_TaxBizAreaPop, pvStartRow, pvEndRow

	'--- 화페단위 
	.vspdData.Col = C_Curr 		
	If UCase(Trim(.vspdData.Text)) = UCase(parent.gCurrency) Then
		ggoSpread.SSSetProtected C_XchRate, pvStartRow, pvEndRow
'		ggoSpread.SSSetProtected C_ChargeLocAmt, pvStartRow, pvEndRow
'		ggoSpread.SSSetProtected C_VatLocAmt, pvStartRow, pvEndRow
		
'		ggoSpread.SSSetProtected C_PayLocAmt, pvStartRow, pvEndRow
	Else
		ggoSpread.SSSetRequired	C_XchRate, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_ChargeLocAmt, pvStartRow, pvEndRow
	End If

	For i = pvStartRow to pvEndRow
		Call ChangePayType(i)
		Call VATTypeEditColor(i)
	Next

    .vspdData.ReDraw = True
    
    End With

End Sub

'========================================================================================================
Function OpenHdrSalesCharge(Byval strCode,Byval strName,Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	On Error Resume Next

	If strCode.readOnly = True Then Exit Function

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	Case 1,3	<% '진행구분 %>
		arrParam(1) = "B_Minor"								<%' TABLE 명칭 %>
		arrParam(2) = Trim(strCode.value)					<%' Code Condition%>
		arrParam(3) = Trim(strName.value)					<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD=" & FilterVar("S9014", "''", "S") & ""					<%' Where Condition%>
		arrParam(5) = "진행구분"							<%' TextBox 명칭 %>

		arrField(0) = "Minor_CD"							<%' Field명(0)%>
		arrField(1) = "Minor_NM"							<%' Field명(1)%>

		arrHeader(0) = "진행구분"							<%' Header명(0)%>
		arrHeader(1) = "진행구분명"						<%' Header명(1)%>

	Case 2,4	<% '영업그룹 %>
		arrParam(1) = "B_SALES_GRP"							<%' TABLE 명칭 %>
		arrParam(2) = Trim(strCode.value)					<%' Code Condition%>
		arrParam(3) = Trim(strName.value)					<%' Name Cindition%>
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						<%' Where Condition%>
		arrParam(5) = "영업그룹"							<%' TextBox 명칭 %>

	    arrField(0) = "SALES_GRP"							<%' Field명(0)%>
	    arrField(1) = "SALES_GRP_NM"						<%' Field명(1)%>
		
	    arrHeader(0) = "영업그룹"							<%' Header명(0)%>
	    arrHeader(1) = "영업그룹명"						<%' Header명(1)%>

	End Select

	arrParam(0) = arrParam(5)								<%' 팝업 명칭 %>
	arrParam(3) = ""

	strCode.focus

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If Err.number <> 0 Then 
			Err.Clear 
			Exit Function
		End If
	Else
		Call SetHdrSalesCharge(arrRet,iWhere,strCode,strName)
	End If	
	
End Function

'========================================================================================================
Function OpenProcessStep(strProcessStep,strCode,strName,Byval iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim ProcessStep,strPopUp, strprogid
	Dim arrVal

	Select Case iWhere
	Case 2
		If frm1.txtBasNo.readOnly = True Then Exit Function
	End Select

'	On Error Resume Next

	If Trim(strProcessStep.value) = "" Then
		Call DisplayMsgBox("206150", "X", "X", "X")
'		Msgbox "진행구분코드를 먼저 입력하세요.",VbInformation,"상태 알림 메세지"
		strProcessStep.focus
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	strCode.focus

	Select Case UCase(Trim(strProcessStep.value))
	Case "SN","EB","ED","EL","EN","EO","EA","EM"
		Select Case UCase(Trim(strProcessStep.value))
		Case "SN"         <% '수주 %>
			strPopUp = "../s3/s3111pa1.asp"
			strprogid = "s3111pa1"
		Case "EB"         <% '수출 선적 %>
			strPopUp = "../s8/s5211pa1.asp"
			strprogid = "s5211pa1"
		Case "ED"         <% '수출 통관 %>
			strPopUp = "../s6/s4211pa1.asp"
			strprogid = "s4211pa1"
		Case "EL"         <% '수출 L/C %>
			strPopUp = "../s4/s3211pa1.asp"
			strprogid = "s3211pa1"
		Case "EN"         <% '수출 NEGO %>
			strPopUp = "../sa/s7111pa1.asp"
			strprogid = "s7111pa1"
		Case "EO"         <% '수출 Local L/C %>
			strPopUp = "../s4/s3211pa2.asp"
			strprogid = "s3211pa2"
		Case "EA"         <% '수출 L/C Amend %>
			strPopUp = "../s4/s3221pa1.asp"
			strprogid = "s3221pa1"
		Case "EM"         <% '수출 Local L/C Amend %>
			strPopUp = "../s4/s3221pa2.asp"
			strprogid = "s3221pa2"
		End Select

		iCalledAspName = AskPRAspName(strprogid)
		
		if Trim(iCalledAspName) = "" then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, strprogid, "x")
			gblnWinEvent = False
			exit Function
		end if

		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrVal), _
				"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Case "BN"         <% '매출 %>
	
		arrVal = ""

		strprogid = "s5111pa1"

		iCalledAspName = AskPRAspName(strprogid)
		
		if Trim(iCalledAspName) = "" then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, strprogid, "x")
			gblnWinEvent = False
			exit Function
		end if
		arrRet = window.showModalDialog(iCalledAspName & "?txtExceptFlag=A", Array(window.parent, arrVal), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Case "DN"         <% '출하 %>
		strPopUp = "../s5/s4111pa1.asp"
		strprogid = "s4111pa1"
		iCalledAspName = AskPRAspName(strprogid)
		
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, strprogid, "x")
			gblnWinEvent = False
			Exit Function
		End if

		arrRet = window.showModalDialog(iCalledAspName & "?txtExceptFlag=A", Array(window.parent), _
				"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Case Else
		IsOpenPop = False
		Call DisplayMsgBox("206153", "X", "X", "X")
		Exit Function

	End Select

	IsOpenPop = False

	Select Case UCase(Trim(strProcessStep.value))
	Case "SN","DN","BN","EA","EM"
		If arrRet = "" Then
			If Err.Number <> 0 Then
				Err.Clear 
			End If
			Exit Function
		Else
			Call SetProcessStep(arrRet,iWhere,strProcessStep,strCode,strName)
		End If
	Case Else
		If arrRet(0) = "" Then
			If Err.Number <> 0 Then
				Err.Clear 
			End If
			Exit Function
		Else
			Call SetProcessStep(arrRet,iWhere,strProcessStep,strCode,strName)
		End If
	End Select	
	
End Function

'========================================================================================================
Function OpenSalesCharge(Byval strCode, Byval iWhere, Byval GridRow)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol,TempCd

	OpenSalesCharge	= False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	Case 1	'경비항목 

		arrParam(1) = "B_TRADE_CHARGE CHR,A_JNL_ITEM JNL"	<%' TABLE 명칭 %>
		arrParam(2) = Trim(strCode)							<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "CHR.CHARGE_CD = JNL.JNL_CD AND CHR.MODULE_TYPE = " & FilterVar("S", "''", "S") & "  AND " _
						& "JNL.JNL_TYPE = " & FilterVar("EC", "''", "S") & ""				<%' Where Condition%>
		arrParam(5) = "경비항목"						<%' TextBox 명칭 %>

		arrField(0) = "CHR.CHARGE_CD"						<%' Field명(0)%>
		arrField(1) = "JNL.JNL_NM"							<%' Field명(1)%>

		arrHeader(0) = "경비항목"						<%' Header명(0)%>
		arrHeader(1) = "경비항목명"						<%' Header명(1)%>

	Case 2	'거래처 
		arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>	    
		arrParam(2) = Trim(strCode)							<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "bp_type IN ( " & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ", " & FilterVar("S", "''", "S") & " ) AND usage_flag = " & FilterVar("Y", "''", "S") & " "	<%' Where Condition%>
		arrParam(5) = "거래처"							<%' TextBox 명칭 %>

		arrField(0) = "BP_CD"								<%' Field명(0)%>
		arrField(1) = "BP_NM"								<%' Field명(1)%>
		
	    arrHeader(0) = "거래처"							<%' Header명(0)%>
	    arrHeader(1) = "거래처명"						<%' Header명(1)%>

	Case 21	''2008-04-21 6:06오후 :: hanc거래처 
		arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>	    
		arrParam(2) = Trim(strCode)							<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "bp_type IN ( " & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ", " & FilterVar("S", "''", "S") & " ) AND usage_flag = " & FilterVar("Y", "''", "S") & " "	<%' Where Condition%>
		arrParam(5) = "세금계산서거래처"							<%' TextBox 명칭 %>

		arrField(0) = "BP_CD"								<%' Field명(0)%>
		arrField(1) = "BP_NM"								<%' Field명(1)%>
		
	    arrHeader(0) = "거래처"							<%' Header명(0)%>
	    arrHeader(1) = "거래처명"						<%' Header명(1)%>

	Case 3	'VAT유형 
		arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config"<%' TABLE 명칭 %>
		arrParam(2) = Trim(strCode)							<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD" _
						& " And Config.MINOR_CD = Minor.MINOR_CD" _
						& " And Config.SEQ_NO = 1"			<%' Where Condition%>	
		arrParam(5) = "VAT유형"							<%' TextBox 명칭 %>

	    arrField(0) = "Minor.MINOR_CD"						<%' Field명(0)%>
	    arrField(1) = "Minor.MINOR_NM"						<%' Field명(1)%>
	    arrField(2) = "Config.REFERENCE"					<%' Field명(2)%>

		arrHeader(0) = "VAT유형"						<%' Header명(0)%>
		arrHeader(1) = "VAT유형명"						<%' Header명(1)%>
		arrHeader(2) = "VAT율"							

	Case 4	'화폐 
		arrParam(1) = "B_CURRENCY"							<%' TABLE 명칭 %>
		arrParam(2) = strCode								<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "화폐"							<%' TextBox 명칭 %>
		
	    arrField(0) = "CURRENCY"							<%' Field명(0)%>
	    arrField(1) = "CURRENCY_DESC"						<%' Field명(1)%>
	    
	    arrHeader(0) = "화폐"							<%' Header명(0)%>
	    arrHeader(1) = "화폐명"							<%' Header명(1)%>

	Case 5	'지급유형 
		arrParam(0) = "지급유형"								<%' 팝업 명칭 %>
		arrParam(1) = "B_CONFIGURATION Config, B_MINOR Minor"		<%' TABLE 명칭 %>
		arrParam(2) = Trim(strCode)									<%' Code Condition%>
		arrParam(3) = ""											<%' Name Cindition%>
		arrParam(4) = "Config.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " AND Config.SEQ_NO = " & FilterVar("1", "''", "S") & "  " _
					& "AND Config.MINOR_CD = Minor.MINOR_CD AND Config.MAJOR_CD = Minor.MAJOR_CD " _
					& "AND Config.MINOR_CD <> " & FilterVar("PP", "''", "S") & " " _
					& "AND Config.REFERENCE IN(" & FilterVar("RP", "''", "S") & "," & FilterVar("P", "''", "S") & " )"			<%' Where Condition%>
		arrParam(5) = "지급유형"								<%' TextBox 명칭 %>

		arrField(0) = "Config.MINOR_CD"								<%' Field명(0)%>
		arrField(1) = "Minor.MINOR_NM"								<%' Field명(1)%>

	    arrHeader(0) = "지급유형"								<%' Header명(0)%>
	    arrHeader(1) = "지급유형명"								<%' Header명(1)%>

	Case 6	'출금계좌 
		arrParam(1) = "B_BANK BK, F_DPST DP"	
		arrParam(2) = Trim(strCode)		
		arrParam(3) = ""				

		frm1.vspdData.Col = C_BankCd
		If Trim(frm1.vspdData.Text) = "" Then
			Call DisplayMsgBox("205152", "X", "출금은행", "X")
			frm1.vspdData.Action = 0
			IsOpenPop = False
			Exit Function
		End If
		
		arrParam(4) = "BK.BANK_CD=DP.BANK_CD And BK.BANK_CD = " _
			& FilterVar(Trim(frm1.vspdData.Text), "''", "S")
		arrParam(5) = "출금계좌"						

		arrField(0) = "DP.BANK_ACCT_NO"		
		arrField(1) = "BK.BANK_CD"	
		arrField(2) = "BK.BANK_NM"				

		arrHeader(0) = "출금계좌"			
		arrHeader(1) = "출금은행"			
		arrHeader(2) = "출금은행명"			
		
	Case 7	'출금은행코드 
		arrParam(1) = "B_BANK BK, F_DPST DP"	
		arrParam(2) = Trim(strCode)			
		arrParam(3) = ""					
		arrParam(4) = "BK.BANK_CD=DP.BANK_CD" 
		arrParam(5) = "출금은행"		

		arrField(0) = "BK.BANK_CD"	
		arrField(1) = "BK.BANK_NM"		
		arrField(2) = "DP.BANK_ACCT_NO"			

		arrHeader(0) = "출금은행"						
		arrHeader(1) = "출금은행명"						
		arrHeader(2) = "출금계좌"						

	Case 8	'어음번호 

		Dim strBpCd, strChargeDt, strChargeLocAmt, strVatLocAmt, strTotAmt, iDblPayAmt
		
		frm1.vspdData.Row = GridRow

		<% '거래처 %>
		frm1.vspdData.Col = C_BpCd
		strBpCd = Trim(frm1.vspdData.Text)
		If Len(strBpCd) = 0 Then 
			MsgBox "거래처를 입력하세요", vbInformation, "<%=gLogoName%>"
			frm1.vspdData.Action = 0
			IsOpenPop = False
			Exit Function
		End If

		<% '발생일 %>
		frm1.vspdData.Col = C_ChargeDt
		strChargeDt = UNIConvDate(Trim(frm1.vspdData.Text))
		If Len(strChargeDt) = 0 Then 
			MsgBox "발생일를 입력하세요", vbInformation, "<%=gLogoName%>"
			frm1.vspdData.Action = 0
			IsOpenPop = False
			Exit Function
		End If

		<% '자국금액 %>
		frm1.vspdData.Col = C_ChargeLocAmt
		strChargeLocAmt = UNICDbl(Trim(frm1.vspdData.Text))
		If Len(strChargeLocAmt) = 0 Then 
			MsgBox "자국금액을 입력하세요", vbInformation, "<%=gLogoName%>"
			frm1.vspdData.Action = 0
			IsOpenPop = False
			Exit Function
		End If

		<% 'VAT자국금액 %>
		frm1.vspdData.Col = C_VatLocAmt
		strVatLocAmt = UNICDbl(Trim(frm1.vspdData.Text))
		If Len(strVatLocAmt) = 0 Then 
			MsgBox "VAT자국금액을 입력하세요", vbInformation, "<%=gLogoName%>"
			frm1.vspdData.Action = 0
			IsOpenPop = False
			Exit Function
		End If
        
		strTotAmt =strChargeLocAmt + strVatLocAmt

		'지급금액 
		frm1.vspdData.Col = C_PayLocAmt
		iDblPayAmt = UNICDbl(Trim(frm1.vspdData.Text))

		If UNICDbl(iDblPayAmt) = 0 Then
			Call DisplayMsgBox("173133",  "X", "지급금액", "X")
			frm1.vspdData.Action = 0
			IsOpenPop = False
			Exit Function
		Else
			If strTotAmt <> iDblPayAmt Then
				Call DisplayMsgBox("206137", "X", "X", "X")
				frm1.vspdData.Action = 0
				IsOpenPop = False
				Exit Function
			End If
		End If

        '--- 화페단위 
		frm1.vspdData.Col = C_Curr 	
'		strTotAmt =UNIConvNumPCToCompanyByCurrency(strChargeLocAmt + strVatLocAmt,trim(frm1.vspdData.text),parent.ggAmtOfMoneyNo, "X" , "X")


		arrParam(1) = "F_NOTE"								<%' TABLE 명칭 %>
		arrParam(2) = Trim(strCode)							<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "NOTE_FG = " & FilterVar("D3", "''", "S") & " AND NOTE_STS = " & FilterVar("BG", "''", "S") & " AND BP_CD =  " & FilterVar(strBpCd , "''", "S") & "" _
					& " AND (ISSUE_DT <=  " & FilterVar(strChargeDt , "''", "S") & " And DUE_DT >=  " & FilterVar(strChargeDt , "''", "S") & ")" _
					& " AND NOTE_AMT = (" & strTotAmt & ")"			<%' Where Condition%>
										
		arrParam(5) = "어음번호"						<%' TextBox 명칭 %>

		arrField(0) = "NOTE_NO"								<%' Field명(0)%>
		arrField(1) = "NOTE_FG"								<%' Field명(5)%>
		arrField(2) = "NOTE_STS"							<%' Field명(6)%>

		arrHeader(0) = "어음번호"						<%' Header명(0)%>
		arrHeader(1) = "어음구분"						<%' Header명(1)%>
		arrHeader(2) = "어음상태"						<%' Header명(2)%>

	Case 9	'세금신고사업장 

		' 2002-09-30 : 세금 신고 사업장 Table변경 
		arrParam(1) = "B_TAX_BIZ_AREA"						<%' TABLE 명칭 %>
		arrParam(2) = Trim(strCode)							<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>		
		arrParam(5) = "세금신고사업장"					<%' TextBox 명칭 %>
		
		arrField(0) = "TAX_BIZ_AREA_CD"						<%' Field명(0)%>
		arrField(1) = "TAX_BIZ_AREA_NM"						<%' Field명(1)%>

		
		arrHeader(0) = "사업장"							<%' Header명(0)%>
		arrHeader(1) = "사업장명"						<%' Header명(1)%>

	End Select

	arrParam(0) = arrParam(5)							<%' 팝업 명칭 %>

	Select Case iWhere
	Case 6, 7
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Case 8
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=650px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Case Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	End Select
	

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSalesCharge(arrRet, iWhere, GridRow)
		OpenSalesCharge	= True
	End If	
	
	
End Function

<%'======================================   GetTaxBizAreaForSalesGrp()  =====================================
'	Description : 영업그룹이 변경될 경우 관련된 세금신고 사업장 정보를 Fetch한다.
'==================================================================================================== %>
Function GetTaxBizAreaForSalesGrp()
	Dim iIntRowIndex
	Dim iStrBpCd, iStrPreBpCd, iStrPreTaxBizArea
	Dim iBlnFetchFlag
	
	iBlnFetchFlag = False
	
	With frm1
		For iIntRowIndex = 1 To .vspdData.MaxRows
			.vspdData.Row = iIntRowIndex
			
			.vspddata.Col = C_VatType
			If Trim(.vspddata.text) <> "" Then
				.vspdData.Col = C_BpCd : iStrBpCd = .vspdData.Text
				If iBlnFetchFlag Then
					If iStrBpCd = iStrPreBpCd Then
						'거래처가 이전Row의 거래처와 동일한 경우 
						.vspddata.Col = C_TaxBizArea : .vspddata.text = iStrPreTaxBizArea
					Else
						Call GetTaxBizArea("BA", iIntRowIndex)
						iStrPreBpCd = iStrBpCd
						.vspddata.Col = C_TaxBizArea : iStrPreTaxBizArea = .vspddata.text
					End If
				Else
					Call GetTaxBizArea("BA", iIntRowIndex)
					iStrPreBpCd = iStrBpCd
					.vspddata.Col = C_TaxBizArea : iStrPreTaxBizArea = .vspddata.text
					iBlnFetchFlag = True
				End If
			End If
		Next
	End With
End Function

<%'======================================   GetTaxBizArea()  =====================================
'	Description : 세금신고 사업장 정보를 Fetch한다.
'==================================================================================================== %>
Function GetTaxBizArea(Byval pvStrFlag, ByVal pvIntRow)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrBpCd, iStrSalesGrp, iStrTaxBizArea
	Dim iStrRs
	Dim iArrTaxBizArea

	GetTaxBizArea = False

	With frm1
		' 부가세 유형이 등록되지 않은 경우에는 세금신고사업장을 Fetch하지 않는다 
		.vspddata.Row = pvIntRow
		.vspddata.Col = C_VatType
		If .vspddata.text = "" Then
			GetTaxBizArea = False
			Exit Function
		End If
	
		<%'세금신고 사업장 Edting시 유효값 Check 및 사업장 명 Fetch %>	
		If pvStrFlag = "NM" Then
			.vspddata.Col = C_TaxBizArea
			iStrTaxBizArea = .vspdData.Text
		Else
			.vspdData.Col = C_BpCd
			iStrBpCd = .vspdData.Text
			iStrSalesGrp = frm1.txtSalesGrp.value
			<%'발행처와 영업 그룹이 모두 등록되어 있는 경우 종합코드에 설정된 rule을 따른다 %>
			If Len(iStrBpCd) > 0 And Len(iStrSalesGrp) > 0 Then
				pvStrFlag = "*"
			ElseIf pvStrFlag = "VT" Then
				If Len(iStrBpCd) > 0 Then
					pvStrFlag = "BP"
				ElseIf Len(iStrSalesGrp) > 0 Then
					pvStrFlag = "BA"
				Else
					Exit Function			
				End If
			End If
		End if
	End With
	
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetTaxBizArea ( " & FilterVar(iStrBpCd, "''", "S") & ",  " & FilterVar(iStrSalesGrp, "''", "S") & ",  " & FilterVar(iStrTaxBizArea, "''", "S") & ",  " & FilterVar(pvStrFlag, "''", "S") & ") "
	iStrWhereList = ""
	
	Err.Clear    
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		If iStrRs <> "" Then
			iArrTaxBizArea = Split(iStrRs, Chr(11))
			With frm1
				.vspdData.Col = C_TaxBizArea
				.vspdDAta.Text = iArrTaxBizArea(1)
				.vspdData.Col = C_TaxBizAreaNm
				.vspdDAta.Text = iArrTaxBizArea(2)
			End With
			GetTaxBizArea = True
		End If		
	Else
		If Err.number <> 0 Then Err.Clear 

		' 세금 신고 사업장을 Editing한 경우 
		If pvStrFlag = "NM" Then
			If Not OpenSalesCharge(iStrTaxBizArea, 9, pvIntRow) Then
				With frm1
					.vspdData.Col = C_TaxBizArea
					.vspdDAta.Text = ""
					.vspdData.Col = C_TaxBizAreaNm
					.vspdDAta.Text = ""
				End With
			Else
				GetTaxBizArea = True
			End if
		End if
		
	End if	

End Function

'========================================================================================================
Function SetHdrSalesCharge(Byval arrRet,Byval iWhere,strCode,strName)

	strCode.value = arrRet(0)
	strName.value = arrRet(1) 

	Select Case iWhere
	Case 3
		lgBlnFlgChgValue = True
	Case 4
		Call GetTaxBizAreaForSalesGrp
		lgBlnFlgChgValue = True
	End Select

End Function

'========================================================================================================
Function SetProcessStep(Byval arrRet,Byval iWhere,strProcessStep,strCode,strName)

	Select Case iWhere
	Case 1
		Select Case UCase(Trim(strProcessStep.value))
		Case "EB","ED","EL","EN","EO"	<% '수출 선적,수출 통관,수출 L/C,수출 NEGO,수출 Local L/C %>
			frm1.txtConBasNo.value = arrRet(0)
		Case "SN","DN","BN","EA","EM"	<% '수주,출하,매출,L/C Amend, Local L/C Amend %>
			frm1.txtConBasNo.value = arrRet
		End Select

	Case 2
		Select Case UCase(Trim(strProcessStep.value))
		Case "EB","ED","EL","EO"	<% '수출 선적,수출 통관,수출 L/C,수출 NEGO,수출 Local L/C %>
			strCode.value = arrRet(0)
			frm1.txtBasDocNo.value = ""
			If UBound(arrRet) > 0 Then frm1.txtBasDocNo.value = arrRet(1)			
		Case "SN","DN","BN","EA","EM"	<% '수주,출하,매출,L/C Amend, Local L/C Amend %>
			strCode.value = arrRet
		Case "EN"
			strCode.value = arrRet(0)
		End Select

		lgBlnFlgChgValue = True

	End Select
		
End Function

'========================================================================================================
Function SetSalesCharge(Byval arrRet,ByVal iWhere,ByVal GridRow)

	With frm1

		ggoSpread.Source = frm1.vspdData

		Select Case iWhere
		Case 1	'경비항목 
			.vspdData.Col = C_ChargeType
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_ChargeTypeNm
			.vspdData.Text = arrRet(1)

		Case 2	'거래처 
		    
			.vspdData.Col = C_BpCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_BpNm
			.vspdData.Text = arrRet(1)
			Call GetTaxBizArea("BP", GridRow)
			
            ' Check Ryu
            Call vspdData_Change(C_BpCd ,Cint(GridRow))

			lgBlnFlgChgValue = True
			Exit Function

		Case 21	''2008-04-21 6:02오후 :: hanc 세금계산서거래처 
		    
			.vspdData.Col = C_ext1_Cd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_ext1_Nm
			.vspdData.Text = arrRet(1)
			

			lgBlnFlgChgValue = True
			Exit Function

		Case 3	'계산서종류 
			.vspdData.Col = C_VatType
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_VatTypeNm
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_VatRate
			.vspdData.Text = arrRet(2)

			Call GetTaxBizArea("VT", GridRow)
			Call vspdData_Change(C_ChargeDocAmt ,Cint(GridRow))
			Call vspdData_Change(C_ChargeLocAmt ,Cint(GridRow))
			Call VATTypeEditColor(GridRow)

			lgBlnFlgChgValue = True
			Exit Function
			
		Case 4	'화폐 
			.vspdData.Col = C_Curr
			.vspdData.Text = arrRet(0)
			Call vspdData_Change(C_Curr,frm1.vspdData.ActiveRow)
		Case 5	'지급유형 
			.vspdData.Col = C_PayType
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_PayTypeNm
			.vspdData.Text = arrRet(1)
			Call ChangePayType(GridRow)
			
            ' Check Ryu
            Call vspdData_Change(C_PayType ,Cint(GridRow))
            
			lgBlnFlgChgValue = True
			Exit Function

		Case 6	'출금계좌 
			.vspdData.Col = C_BankAcct
			.vspdData.Text = arrRet(0)
			
		Case 7	'출금은행코드 
			.vspdData.Col = C_BankCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_BankNm
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_BankAcct
			.vspdData.Text = arrRet(2)

		Case 8	'어음번호 
			.vspdData.Col = C_CheckNo
			.vspdData.Text = arrRet(0)

		Case 9	'세금신고사업장 
			.vspdData.Col = C_TaxBizArea
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_TaxBizAreaNm
			.vspdData.Text = arrRet(1)
			
			'ggoSpread.UpdateRow GridRow 'v2.5에서 사용된방법 
            ' Check Ryu
            Call vspdData_Change(C_TaxBizArea ,Cint(GridRow))			

			lgBlnFlgChgValue = True
			Exit Function
		End Select
	
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		<% ' 변경이 읽어났다고 알려줌 %>

	End With

	lgBlnFlgChgValue = True
	
End Function

<%
'========================================================================================
' Function Desc : DbQuery중 Header가 성공적으로 조회된 경우 
'========================================================================================
%>
Function HdrQueryOk()														<%'☆: 조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												<%'⊙: Indicates that current mode is Update mode%>
  
	Call ggoOper.SetReqAttr(frm1.txtProcessStepCd, "Q")
	Call ggoOper.SetReqAttr(frm1.txtSalesGrp, "Q")

	If Not IsNull(Trim(frm1.txtBasNo.value)) Then
		Call ggoOper.SetReqAttr(frm1.txtBasNo, "Q")
	End If

End Function


<% '===================================   ChangeVatType()  ======================================
'	Description : 계산서종류 변동 
'==================================================================================================== %>
Sub ChangeVatType(ByVal pvIntRow)

	Dim strVal

	lgBlnVatChangedFlag = False

	frm1.vspdData.Row = pvIntRow
	frm1.vspdData.Col = C_VatType
	IF Trim(frm1.vspdData.Text) <> "" Then
		Call LayerShowHide(1)
		
		strVal = BIZ_PGM_ID & "?txtMode=" & ConVatType									<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtSpread=" & "1" & parent.gColSep & "B9001" & parent.gColSep & Trim(frm1.vspdData.Text) & parent.gColSep & pvIntRow & parent.gRowSep
		strVal = strVal & "&txtUpdtUserId=" & parent.gUsrID									<%'☆: 조회 조건 데이타 %>
		
   		Call RunMyBizASP(MyBizASP, strVal)												<%'☜: 비지니스 ASP 를 가동 %>
   		'세금신고사업장 Default 값 Fetch
		Call GetTaxBizArea("VT", pvIntRow)
		If Not lgBlnVatChangedFlag Then Call VATTypeEditColor(pvIntRow)
	Else
		Call VATTypeEditColor(pvIntRow)
	End If 
End Sub

<% '===================================   ChangeVatTypeOk()  ======================================
'	Description : 계산서종류 변동 
'==================================================================================================== %>
Function ChangeVatTypeOk(GRow)
	'-- VAT율 
	frm1.vspdData.Row = Cint(GRow)
	frm1.vspdData.Col = C_VatRate
	frm1.vspdData.Text = UNIFormatNumber(UNICDbl(frm1.txtSpread.value),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
	Call vspdData_Change(C_ChargeDocAmt ,Cint(GRow))
	Call vspdData_Change(C_ChargeLocAmt ,Cint(GRow))
	Call VATTypeEditColor(GRow)
	lgBlnVatChangedFlag = TRUE
End Function

<% '===================================   ChangeXchRate()  ======================================
'	Description : 환율변동 
'==================================================================================================== %>
Sub ChangeXchRate(strCSFalg)
	If strCSFalg = "Client" Then
		Call ClientXchRateCalcu
		Exit Sub
	End If
	
	'-------------------- 서버단 잠시 Stop --------------------	
	Dim strVal

	Call LayerShowHide(1)

	strVal = BIZ_PGM_ID & "?txtMode=" & ConXchRate										<%'☜: 비지니스 처리 ASP의 상태 %>
	strVal = strVal & "&txtSpread=" & Trim(frm1.txtSpread.value)
	strVal = strVal & "&txtUpdtUserId=" & parent.gUsrID									<%'☆: 조회 조건 데이타 %>
 
	Call RunMyBizASP(MyBizASP, strVal)												<%'☜: 비지니스 ASP 를 가동 %>

End Sub

<% '===================================   ChangeXchRateOk()  ======================================
'	Description : 환율변동 
'==================================================================================================== %>
Function ChangeXchRateOk(GRow)

	frm1.vspdData.Row = Cint(GRow)
	frm1.vspdData.Col = C_ChargeLocAmt
	frm1.vspdData.Text = UNIFormatNumber(UNICDbl(frm1.txtSpread.value), ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)

	Call vspdData_Change(C_XchRate,Cint(GRow))

End Function


<% '===================================   ClientXchRateCalcu()  =====================================
'	Description : Client Side에서 환율 계산 
'==================================================================================================== %>
Sub ClientXchRateCalcu()

	Dim arrXchVal, arrXchTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    arrXchTemp = Split(frm1.txtSpread.value, parent.gRowSep)
    
	arrXchVal = Split(arrXchTemp(0), parent.gColSep)

	frm1.vspdData.Row = Trim(arrXchVal(5))
	frm1.vspdData.Col = C_XchCalop
	
	Select Case Trim(frm1.vspdData.Text)

	    Case "+"
	    	frm1.vspdData.Col = C_ChargeLocAmt
	    	frm1.vspdData.Text = UNIFormatNumber(UNICDbl(Trim(arrXchVal(3))) + UNICDbl(Trim(arrXchVal(4))),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	    Case "-"
	    	frm1.vspdData.Col = C_ChargeLocAmt
	    	frm1.vspdData.Text = UNIFormatNumber(UNICDbl(Trim(arrXchVal(3))) - UNICDbl(Trim(arrXchVal(4))),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	    Case "*"
	    	frm1.vspdData.Col = C_ChargeLocAmt
	    	frm1.vspdData.Text = UNIFormatNumber(UNICDbl(Trim(arrXchVal(3))) * UNICDbl(Trim(arrXchVal(4))),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	    Case "/"
	    	frm1.vspdData.Col = C_ChargeLocAmt
	    	frm1.vspdData.Text = UNIFormatNumber(UNICDbl(Trim(arrXchVal(3))) * UNICDbl(Trim(arrXchVal(4))),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	    Case Else
	    	.vspdData.Col = C_ChargeLocAmt
	    	.vspdData.Text = UNIFormatNumber(UNICDbl(Trim(arrXchVal(3))),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	End Select

	Call vspdData_Change(C_ChargeLocAmt ,Cint(Trim(arrXchVal(5))))
	Call vspdData_Change(C_PayDocAmt ,Cint(Trim(arrXchVal(5))))

End Sub


<% '=====================================  ChangePayType() ===========================================
'	Description : 지급유형의 변화에 따른 Edit 처리 
'==================================================================================================== %>
Function ChangePayType(Byval GridRow)

	With frm1

		.vspdData.Row = GridRow
		.vspdData.Col = C_PayType

		Select Case UCase(Trim(.vspdData.Text))
		Case "NP"		<% '지급어음 %>
			ggoSpread.SpreadUnLock C_CheckNo,GridRow,C_CheckNoPop,GridRow	<% '어음번호 %>

			If GetSetupMod(parent.gSetupMod, "A") = "Y" Then
				ggoSpread.SSSetRequired C_CheckNo, GridRow, GridRow
			Else
				ggoSpread.SSSetProtected C_CheckNoPop, GridRow, GridRow
			End If

			ggoSpread.SpreadLock C_BankCd,GridRow,C_BankAcctPop,GridRow
			.vspdData.Col = C_BankCd : .vspdData.Text = ""
			.vspdData.Col = C_BankNm : .vspdData.Text = ""
			.vspdData.Col = C_BankAcct : .vspdData.Text = ""

		Case "DP"		<% '예적금 %>
			ggoSpread.SpreadUnLock C_BankCd,GridRow,C_BankPop,GridRow		<% '출금은행 %>
			ggoSpread.SpreadUnLock C_BankAcct,GridRow,C_BankAcctPop,GridRow	<% '출금계좌 %>
			ggoSpread.SSSetRequired C_BankCd, GridRow, GridRow
			ggoSpread.SSSetRequired C_BankAcct, GridRow, GridRow

			ggoSpread.SpreadLock C_CheckNo,GridRow,C_CheckNoPop,GridRow
			.vspdData.Col = C_CheckNo : .vspdData.Text = ""

		Case Else
			ggoSpread.SpreadLock C_CheckNo,GridRow,C_CheckNoPop,GridRow
			ggoSpread.SpreadLock C_BankCd,GridRow,C_BankAcctPop,GridRow

			.vspdData.Col = C_CheckNo : .vspdData.Text = ""
			.vspdData.Col = C_BankCd : .vspdData.Text = ""
			.vspdData.Col = C_BankNm : .vspdData.Text = ""
			.vspdData.Col = C_BankAcct : .vspdData.Text = ""

		End Select

		.vspdData.Col = C_PayType
		<% '지급유형이 입력된 경우 : 지급금액, 지급자국금액이 필수입력항목 %>
		Select Case Len(Trim(.vspdData.Text))
		Case 0
			ggoSpread.SpreadUnLock C_PayDocAmt,GridRow,C_PayLocAmt,GridRow
		Case Else
			ggoSpread.SSSetRequired C_PayDocAmt, GridRow, GridRow	'지급액 
			ggoSpread.SSSetRequired C_PayLocAmt, GridRow, GridRow	'지급자국액 
		End Select

		'--- 화페단위 
		.vspdData.Col = C_Curr
		If UCase(Trim(.vspdData.Text)) = UCase(parent.gCurrency) Then ggoSpread.SSSetProtected C_PayLocAmt, GridRow, GridRow	'지급자국액 

	End With

End Function


<% '=====================================  VATTypeEditColor() ========================================
'	Description : 세금계산서종류가 입력되는 경우 필수입력처리 
'==================================================================================================== %>
Function VATTypeEditColor(Byval GridRow)

	With frm1
		'-- 계산서종류	有 -> Requried
		'--				無 -> Protected
		.vspdData.Col = C_VatType 
		.vspdData.Row = GridRow

		<% '세금계산서종류가 입력된 경우 : 세금신고사업장은 필수입력항목 %>
		Select Case Len(Trim(.vspdData.Text))
		Case 0
			ggoSpread.SpreadLock C_TaxBizArea, GridRow,C_TaxBizAreaPop,GridRow
			'--- 세금신고사업장 
			.vspdData.Col = C_TaxBizArea	:	.vspdData.Text = ""
			.vspdData.Col = C_TaxBizAreaNm	:	.vspdData.Text = ""

			ggoSpread.SpreadLock C_VatDocAmt, GridRow,C_VatLocAmt,GridRow

			'--- VAT금액 
			.vspdData.Col = C_VatDocAmt	:	.vspdData.Text = 0
			'--- VAT자국금액 
			.vspdData.Col = C_VatLocAmt	:	.vspdData.Text = 0
			'--- VAT율 
			.vspdData.Col = C_VatRate	:	.vspdData.Text = 0

		Case Else
			ggoSpread.SpreadUnLock C_TaxBizArea, GridRow,C_TaxBizAreaPop,GridRow
			ggoSpread.SSSetRequired	C_TaxBizArea, GridRow, GridRow
			ggoSpread.SpreadUnLock C_VatDocAmt, GridRow,C_VatLocAmt,GridRow
			ggoSpread.SSSetRequired	C_VatDocAmt, GridRow, GridRow
			ggoSpread.SSSetRequired	C_VatLocAmt, GridRow, GridRow
		End Select

		'--- 화페단위 
		.vspdData.Col = C_Curr
'		If UCase(Trim(.vspdData.Text)) = UCase(parent.gCurrency) Then ggoSpread.SSSetProtected C_VatLocAmt, GridRow, GridRow	'VAT자국액 

	End With

End Function


<% '=====================================  VATRateEditColor() ========================================
'	Description : VAT율이 0보다 큰 경우 필수입력처리 
'==================================================================================================== %>
Function VATRateEditColor(Byval GridRow)

	With frm1

		'-- VAT Rate >  0 -> VAT금액,VAT자국금액 필수입력 
		'--			 <= 0 -> VAT금액,VAT자국금액 Optional
		.vspdData.Col = C_VatRate
		.vspdData.Row = GridRow

		Select Case UNICDbl(Trim(.vspdData.Text))
		Case 0
			ggoSpread.SpreadUnLock C_VatDocAmt, GridRow,C_VatLocAmt,GridRow

			'--- VAT금액 
			.vspdData.Col = C_VatDocAmt	:	.vspdData.Text = 0
			'--- VAT자국금액 
			.vspdData.Col = C_VatLocAmt	:	.vspdData.Text = 0

		Case Else
			ggoSpread.SSSetRequired	C_VatDocAmt, GridRow, GridRow
			ggoSpread.SSSetRequired	C_VatLocAmt, GridRow, GridRow
		End Select

		'--- 화페단위 
		.vspdData.Col = C_Curr
'		If UCase(Trim(.vspdData.Text)) = UCase(parent.gCurrency) Then ggoSpread.SSSetProtected C_VatLocAmt, GridRow, GridRow	'VAT자국액 

	End With

End Function

<% '===================================   SetQuerySpreadColor()  ======================================
'	Description : 조회후 그리드 Color
'==================================================================================================== %>
Sub SetQuerySpreadColor(ByVal lRow)
	
    Dim SCol
    Dim SRow
    
    With frm1

    .vspdData.ReDraw = False

		For SRow = UNICDbl(frm1.txtPrevMaxRow.value) + 1 to .vspdData.MaxRows
			.vspdData.Row = SRow
			.vspdData.Col = C_PostingFlg
			Select Case UNICDbl(.vspdData.text)
			Case 0
				ggoSpread.SSSetRequired	C_ChargeType, SRow, SRow
				ggoSpread.SSSetProtected C_ChargeTypeNm, SRow, SRow
				ggoSpread.SSSetProtected C_ChargeNo, SRow, SRow
				ggoSpread.SSSetRequired C_BpCd, SRow, SRow
				ggoSpread.SSSetProtected C_BpNm, SRow, SRow
				'2008-04-21 6:05오후 :: hanc
				ggoSpread.SSSetRequired C_ext1_Cd, SRow, SRow
				ggoSpread.SSSetProtected C_ext1_Nm, SRow, SRow
				ggoSpread.SSSetRequired C_ChargeDt, SRow, SRow
				ggoSpread.SSSetProtected C_VatTypeNm, SRow, SRow
				ggoSpread.SSSetRequired C_Curr, SRow, SRow
				ggoSpread.SSSetRequired C_ChargeDocAmt, SRow, SRow

				'--- 화페단위 
				.vspdData.Col = C_Curr 		
				If UCase(Trim(.vspdData.Text)) = UCase(parent.gCurrency) Then
					ggoSpread.SSSetProtected C_XchRate, SRow, SRow
'					ggoSpread.SSSetProtected C_ChargeLocAmt, SRow, SRow
'					ggoSpread.SSSetProtected C_VatLocAmt, SRow, SRow
'					ggoSpread.SSSetProtected C_PayLocAmt, SRow, SRow
				Else
					ggoSpread.SSSetRequired	C_XchRate, SRow, SRow
					ggoSpread.SSSetRequired	C_ChargeLocAmt, SRow, SRow
				End If

'				ggoSpread.SSSetProtected C_ChargeLocAmt, SRow, SRow
				ggoSpread.SSSetProtected C_VatRate, SRow, SRow
				ggoSpread.SSSetProtected C_PayTypeNm, SRow, SRow
				ggoSpread.SSSetProtected C_BankNm, SRow, SRow

				ggoSpread.SSSetProtected C_XchCalop, SRow, SRow
				ggoSpread.SSSetRequired C_PayDueDt, SRow, SRow

				'--- 계산서종류에 값이 있는경우 
				ggoSpread.SSSetProtected C_TaxBizAreaNm, SRow, SRow
				.vspdData.Col = C_VatType 		
				If Len(Trim(.vspdData.Text)) Then
					ggoSpread.SSSetRequired C_TaxBizArea, SRow, SRow
				Else
					ggoSpread.SSSetProtected C_TaxBizArea, SRow, SRow
				End If
				
				<% 'SMJ 지금액이 입력된경우 지급유형을 필수 항목으로 처리 %>
				.vspdData.Col = C_PayDocAmt
				If .vspdData.Text > 0 Then
					ggoSpread.SSSetRequired C_PayType, SRow, SRow	'지급유형 필수 
				Else
					ggoSpread.SpreadUnLock C_PayType, SRow, SRow	'지급유형 
				End If					
					
				Call ChangePayType(SRow)
				Call VATTypeEditColor(SRow)
			Case 1
				For SCol = C_PostingFlg + 1 To C_RefRemark
					ggoSpread.SSSetProtected SCol, SRow, SRow
				Next
			End Select
		Next

    .vspdData.ReDraw = True
    
    End With

End Sub

<% '==========================================   CookiePage()  ======================================
'	Description : JUMP시 Load화면으로 조건부로 Value
'==================================================================================================== %>
Function CookiePage(Byval Kubun)

	Const CookieSplit = 4877						<%'Cookie Split String : CookiePage Function Use%>
	Dim strTemp, arrVal

	If Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then Exit Function

		frm1.txtConProcessStepCd.value =  arrVal(0) 
		frm1.txtConSalesGrp.value =  arrVal(1) 
		frm1.txtConSalesGrpNm.value =  arrVal(2) 
		frm1.txtConBasNo.value =  arrVal(3) 

		frm1.txtProcessStepCd.value =  arrVal(0) 
		frm1.txtSalesGrp.value =  arrVal(1) 
		frm1.txtSalesGrpNm.value =  arrVal(2) 
		frm1.txtBasNo.value =  arrVal(3) 
		
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		WriteCookie CookieSplit , ""

		Call FncQuery()
		
	End IF

End Function

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_PostingFlg	= iCurColumnPos(1)
			C_ChargeNo		= iCurColumnPos(2)
			C_ChargeType	= iCurColumnPos(3)
			C_ChargeTypePop = iCurColumnPos(4)
			C_ChargeTypeNm	= iCurColumnPos(5)
			C_BpCd			= iCurColumnPos(6)
			C_BpPop			= iCurColumnPos(7)
			C_BpNm			= iCurColumnPos(8)
			C_ext1_Cd			= iCurColumnPos(9)
			C_ext1_Pop			= iCurColumnPos(10)
			C_ext1_Nm			= iCurColumnPos(11)
			C_ChargeDt		= iCurColumnPos(12)
			C_VatType		= iCurColumnPos(13)
			C_VatTypePop	= iCurColumnPos(14)
			C_VatTypeNm		= iCurColumnPos(15)
			C_Curr			= iCurColumnPos(16)
			C_CurrPop		= iCurColumnPos(17)
			C_ChargeDocAmt	= iCurColumnPos(18)
			C_XchCalop		= iCurColumnPos(19)
			C_XchRate		= iCurColumnPos(20)
			C_ChargeLocAmt	= iCurColumnPos(21)
			C_VatRate		= iCurColumnPos(22)
			C_VatDocAmt		= iCurColumnPos(23)
			C_VatLocAmt		= iCurColumnPos(24)
			C_TaxBizArea	= iCurColumnPos(25)
			C_TaxBizAreaPop = iCurColumnPos(26)
			C_TaxBizAreaNm	= iCurColumnPos(27)
			C_PayDueDt		= iCurColumnPos(28)
			C_PayType		= iCurColumnPos(29)
			C_PayTypePop	= iCurColumnPos(30)
			C_PayTypeNm		= iCurColumnPos(31)
			C_PayDocAmt		= iCurColumnPos(32)
			C_PayLocAmt		= iCurColumnPos(33)
			C_CheckNo		= iCurColumnPos(34)
			C_CheckNoPop	= iCurColumnPos(35)
			C_BankCd		= iCurColumnPos(36)
			C_BankPop		= iCurColumnPos(37)
			C_BankNm		= iCurColumnPos(38)
			C_BankAcct		= iCurColumnPos(39)
			C_BankAcctPop	= iCurColumnPos(40)
			C_RefRemark		= iCurColumnPos(41)

    End Select    
End Sub

'========================================================================================================= 
Sub Form_Load()

	Call InitVariables														'⊙: Initializes local global variables
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet
	Call SetDefaultVal	
	'폴더/조회/입력 
	'/삭제/저장/한줄In
	'/한줄Out/취소/이전 
	'/다음/복사/엑셀 
	'/인쇄/찾기 
    Call SetToolbar("11101111001011")										'⊙: 버튼 툴바 제어 
	Call CookiePage(0)

End Sub
'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================================= 
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
   
		If Row > 0 Then   
			Select Case Col
   
			Case C_ChargeTypePop					'--경비항목 
			    .Col = C_ChargeType
			    .Row = Row
			    Call OpenSalesCharge(.Text, 1, Row)
			Case C_BpPop							'--거래처 
			    .Col = C_BpCd
			    .Row = Row
			    Call OpenSalesCharge(.Text, 2, Row)
			Case C_ext1_Pop							'--'2008-04-21 6:06오후 :: hanc거래처 
			    .Col = C_ext1_Cd
			    .Row = Row
			    Call OpenSalesCharge(.Text, 21, Row)
			Case C_VatTypePop						'--계산서종류 
			    .Col = C_VatType
			    .Row = Row
			    Call OpenSalesCharge(.Text, 3, Row)
			Case C_CurrPop							'--화폐 
			    .Col = C_Curr
			    .Row = Row
			    Call OpenSalesCharge(.Text, 4, Row)
			Case C_PayTypePop						'--지급유형 
			    .Col = C_PayType
			    .Row = Row
			    Call OpenSalesCharge(.Text, 5, Row)
			Case C_BankAcctPop						'--출금계좌 
			    .Col = C_BankAcct
			    .Row = Row
			    Call OpenSalesCharge(.Text, 6, Row)
			Case C_BankPop							'--출금은행코드 
			    .Col = C_BankCd
			    .Row = Row
			    Call OpenSalesCharge(.Text, 7, Row)
			Case C_CheckNoPop						'--어음번호 
			    .Col = C_CheckNo
			    .Row = Row
			    Call OpenSalesCharge(.Text, 8, Row)
			Case C_TaxBizAreaPop					'--세금신고사업장 
			    .Col = C_TaxBizArea
			    .Row = Row
			    Call OpenSalesCharge(.Text, 9, Row)
   			End Select

   			Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")  
   		End If		    

	End With

End Sub

<%
'==========================================================================================
'   Event Desc : 영업그룹 변경시 Default 세금신고사업장 Fetch
'==========================================================================================
%>
Sub txtSalesGrp_OnChange()
	If Trim(frm1.txtSalesGrp.value) = "" Then
		frm1.txtSalesGrpNm.value = ""
	Else
		<%'영업그룹과 관련된 세금신고사업장을 Fetch한다. %>
		Call GetTaxBizAreaForSalesGrp
	End if
End Sub

'========================================================================================================= 
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData
End Sub

'========================================================================================================= 
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================= 
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================================================================================= 
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================= 
Sub vspdData_Change(ByVal Col , ByVal Row)
    
	Dim strChgVal,strDocAmt,strVatDocAmt,strVatLocAmt
	Dim CurrColumn
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	If Row < 0 Then Exit Sub

	lgBlnFlgChgValue = True

	Select Case Col						    				

	    Case C_PayType
	    	Call ChangePayType(Row)

	    Case C_BpCd
	    	'관련 세금신고사업장 Fetch
	    	Call GetTaxBizArea("BP",Row)

	    Case C_XchRate, C_Curr, C_ChargeDt
	    
	    	CurrColumn = Col

			Select Case CurrColumn

		        Case C_Curr

					Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Curr,C_ChargeDocAmt,"A","X","X")
					Call FixDecimalPlaceByCurrency2(frm1.vspdData,Row,Parent.gCurrency,C_ChargeLocAmt,"A","X","X")
					Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Curr,C_VatDocAmt,"A","X","X")
					Call FixDecimalPlaceByCurrency2(frm1.vspdData,Row,Parent.gCurrency,C_VatLocAmt,"A","X","X")	
					Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Curr,C_PayDocAmt,"A","X","X")
					Call FixDecimalPlaceByCurrency2(frm1.vspdData,Row,Parent.gCurrency,C_PayLocAmt,"A","X","X")	
					
					Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_Curr,C_ChargeDocAmt,"A","I","X","X")
					Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,Row,Row,Parent.gCurrency,C_ChargeLocAmt,"A","I","X","X")
					Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_Curr,C_VatDocAmt,"A","I","X","X")
					Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,Row,Row,Parent.gCurrency,C_VatLocAmt,"A" ,"I","X","X")         
					Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_Curr,C_PayDocAmt,"A","I","X","X")
					Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,Row,Row,Parent.gCurrency,C_PayLocAmt,"A" ,"I","X","X")         
		            	
				Case C_ChargeDt
					If Trim(frm1.vspdData.text) = "" Or Trim(frm1.vspdData.text) = parent.gCurrency Then
						Exit Sub
					End If	

		    End Select

		    With frm1
				.vspdData.Row = Row
				.vspdData.Col = 0

				Select Case .vspdData.Text

				    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

						strChgVal=""
						strDocAmt=""

						'--- 화페단위 
				        .vspdData.Col = C_Curr 		

				        If Trim(.vspdData.Text) = "" Then
							'MsgBox "화폐를 입력하세요", vbInformation, "<%=gLogoName%>"
							.vspdData.Action = 0
							Exit Sub
						ElseIf UCase(Trim(.vspdData.Text)) = UCase(parent.gCurrency) Then

					        '--- 발생금액 
					        .vspdData.Col = C_ChargeDocAmt 
					        strDocAmt =UNICDbl(Trim(.vspdData.Text))

					        '--- 자국금액 
					        .vspdData.Col = C_ChargeLocAmt 
					        .vspdData.Text = strDocAmt
					        		
					        ggoSpread.SpreadLock C_XchRate, Row,C_XchRate,Row
					 '      ggoSpread.SpreadLock C_ChargeLocAmt, Row,C_ChargeLocAmt,Row
					 '      ggoSpread.SpreadLock C_VatLocAmt, Row,C_VatLocAmt,Row
					 '      ggoSpread.SpreadLock C_PayLocAmt, Row,C_PayLocAmt,Row

					        '--- 환율 
					        .vspdData.Col = C_XchRate 
					        .vspdData.Text = 1

					        '--- 환율연산자 
					        .vspdData.Col = C_XchCalop
					        .vspdData.Text = "*"

						    Call vspdData_Change(C_ChargeLocAmt ,Row)
						    Call vspdData_Change(C_PayDocAmt ,Row)
						    Exit Sub
				        Else
							ggoSpread.SpreadUnLock C_XchRate, Row,C_XchRate,Row
							ggoSpread.SpreadUnLock C_ChargeLocAmt, Row,C_ChargeLocAmt,Row
							ggoSpread.SpreadUnLock C_PayLocAmt, Row,C_PayLocAmt,Row							
														
							ggoSpread.SSSetRequired	C_XchRate, Row, Row
							ggoSpread.SSSetRequired	C_ChargeLocAmt, Row, Row
							
							.vspdData.Col = C_VatType 
							If Trim(.vspdData.Text) <> "" Then
								ggoSpread.SpreadUnLock C_VatLocAmt, Row,C_VatLocAmt,Row
								ggoSpread.SSSetRequired	C_VatLocAmt, Row, Row
							End If

				        End If

						If CurrColumn = C_XchRate Then 
					        '--- 발생금액 
					        .vspdData.Col = C_ChargeDocAmt 
					        If Trim(.vspdData.Text) = 0 Then
								Exit Sub
					        End If
					     End If

				        '--- 발생일 
				        .vspdData.Col = C_ChargeDt 
				        If Trim(.vspdData.Text) = "" Then
							Exit Sub
				        End If

				        '--- 발생일 
				        .vspdData.Col = C_ChargeDt 
				        strChgVal = strChgVal & Trim(.vspdData.Text) & parent.gColSep
						'--- 화페단위 
				        .vspdData.Col = C_Curr 		
				        strChgVal = strChgVal & Trim(.vspdData.Text) & parent.gColSep
						'--- 화페단위(자국)
				        strChgVal = strChgVal & parent.gCurrency & parent.gColSep
				        '--- 발생금액 
				        .vspdData.Col = C_ChargeDocAmt 
				        strChgVal = strChgVal & Trim(.vspdData.Text) & parent.gColSep
						'--- 환율 
				        .vspdData.Col = C_XchRate 		
				        strChgVal = strChgVal & Trim(.vspdData.Text) & parent.gColSep
						'--- 현 Row No
				        strChgVal = strChgVal & Row & parent.gRowSep

				End Select

			End With

    		frm1.txtSpread.value = strChgVal

		    Select Case CurrColumn
		            
		        Case C_Curr
	
		        	Call ChangeXchRate("Server")
		        	Call ChangePayType(Row)
		        	Call VATTypeEditColor(Row)

		        Case C_XchRate
		
		        	frm1.vspdData.Col = C_XchCalop
		        	If Trim(frm1.vspdData.Text) = "" Then
		        		Call ChangeXchRate("Server")
		        	Else
		        		Call ChangeXchRate("Client")		
		        	End If

				Case C_ChargeDt				
		        	Call ChangeXchRate("Server")

		        	
		    End Select


	    Case C_ChargeDocAmt
	
		    With frm1
				.vspdData.Row = Row
				.vspdData.Col = 0

				Select Case .vspdData.Text
		
				    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

						strChgVal=""
						strDocAmt=""
						strVatDocAmt=""

						'--- 발생금액 
						.vspdData.Col = C_ChargeDocAmt 
						strDocAmt = UNICDbl(Trim(.vspdData.Text))
						If strDocAmt <> 0 Then
							.vspdData.Col = C_VatRate
							strChgVal =UNICDbl(Trim(.vspdData.Text))
							
							 '--- 화페단위 
							.vspdData.Col = C_Curr 
							strVatDocAmt = UNIConvNumPCToCompanyByCurrency((strDocAmt * strChgVal) / 100,Trim(frm1.vspdData.text),parent.ggAmtOfMoneyNo, "X" , "X") 		
							
							'--- VAT금액 
							.vspdData.Col = C_VatDocAmt
							.vspdData.Text = strVatDocAmt
							
							
							'--- 자국금액 계산 
							'Call vspdData_Change(C_Curr,Row)
							Call vspdData_Change(C_XchRate,Row)
						End If
				End Select
			End With

			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Curr,C_ChargeDocAmt,"A","X","X")	

	    Case C_ChargeLocAmt
		    With frm1
				.vspdData.Row = Row
				.vspdData.Col = 0

				Select Case .vspdData.Text
				    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

						strChgVal=""
						strDocAmt=""
						strVatLocAmt=""

						'--- 발생금액 
						.vspdData.Col = C_ChargeLocAmt 
						strDocAmt = UNICDbl(Trim(.vspdData.Text))
						If strDocAmt <> 0 Then
							.vspdData.Col = C_VatRate
							strChgVal = UNICDbl(Trim(.vspdData.Text))

							'--- 화페단위 
							.vspdData.Col = C_Curr 
							strVatLocAmt = UNIConvNumPCToCompanyByCurrency(strDocAmt * (strChgVal/100),Trim(frm1.vspdData.text),parent.ggAmtOfMoneyNo, "X" , "X") 		
							.vspdData.Col = C_VatLocAmt
							.vspdData.Text = strVatLocAmt
							
							'--- 자국금액 계산 
							'==Call vspdData_Change(C_Curr,Row)
						End If
				End Select
			End With

			Call FixDecimalPlaceByCurrency2(frm1.vspdData,Row,Parent.gCurrency,C_ChargeLocAmt,"A","X","X")			

	    Case C_VatType

	    	With frm1
	    		.vspdData.Row = Row
	    		.vspdData.Col = 0
	    		If .vspdData.Text <> ggoSpread.DeleteFlag Then	Call ChangeVatType(Row)
	    	End With

	    Case C_PayDocAmt
	    	    With frm1

	    			Dim strPayDocAmt
	    			Dim strXchRate

	    			.vspdData.Row = Row
	    			
	    			'--화폐 
	    			strXchRate = ""
	    			.vspdData.Col = C_XchRate	:	strXchRate = Trim(.vspdData.Text)
	    			'--지급금액 
	    			strPayDocAmt = ""
	    			.vspdData.Col = C_PayDocAmt	:	strPayDocAmt = Trim(.vspdData.Text)

					<% 'SMJ 지금액이 입력된경우 지급유형을 필수 항목으로 처리 %>
					If strPayDocAmt > 0 Then
						ggoSpread.SSSetRequired C_PayType, Row, Row	'지급유형 필수 
					Else
						ggoSpread.SpreadUnLock C_PayType, Row, Row	'지급유형 
					End If
									
	    			'--- 화페단위 
	    			.vspdData.Col = C_Curr 		
	    			If Trim(.vspdData.Text) = "" Then
	    				MsgBox "화폐를 입력하세요", vbInformation, "<%=gLogoName%>"
	    				.vspdData.Action = 0
	    				Exit Sub
	    			ElseIf UCase(Trim(.vspdData.Text)) = UCase(parent.gCurrency) Then
	    				.vspdData.Col = C_PayLocAmt	:	.vspdData.Text = strPayDocAmt
	    				Exit Sub
	    			End If

	    			.vspdData.Col = 0
	    			Select Case .vspdData.Text
	    			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

	    				.vspdData.Col = C_XchCalop
	
	    				Select Case Trim(.vspdData.Text)
	    				Case "+"
	    					.vspdData.Col = C_PayLocAmt
	    					.vspdData.Text = UNIFormatNumber(UNICDbl(strPayDocAmt) + UNICDbl(strXchRate),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	    				Case "-"
	    					.vspdData.Col = C_PayLocAmt
	    					.vspdData.Text = UNIFormatNumber(UNICDbl(strPayDocAmt) - UNICDbl(strXchRate),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	    				Case "*"
	    					.vspdData.Col = C_PayLocAmt
	    					.vspdData.Text = UNIFormatNumber(UNICDbl(strPayDocAmt) * UNICDbl(strXchRate),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	    				Case "/"
	    					.vspdData.Col = C_PayLocAmt
	    					.vspdData.Text = UNIFormatNumber(UNICDbl(strPayDocAmt) / UNICDbl(strXchRate),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	    				Case Else
	    					.vspdData.Col = C_PayLocAmt
	    					.vspdData.Text =UNIFormatNumber(UNICDbl(strPayDocAmt),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	    				End Select

	    			End Select
	    		End With

			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Curr,C_PayDocAmt,"A","X","X")	

	    Case C_TaxBizArea
	    	' 세금신고사업장명 Fetch
	    	With frm1
	    		.vspdData.Row = Row
	    		.vspdData.Col = C_TaxBizArea
	    		
	    		If Trim(.vspdData.text) = "" Then
	    			.vspdData.Col = C_TaxBizAreaNm : .vspdData.text = ""
	    		Else
	    			Call GetTaxBizArea("NM",Row)
	    		End If
	    	End With

	    Case C_VatDocAmt

			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Curr,C_VatDocAmt,"A","X","X")	

	    Case C_VatLocAmt

			Call FixDecimalPlaceByCurrency2(frm1.vspdData,Row,Parent.gCurrency,C_VatLocAmt,"A","X","X")	
			
	    Case C_PayLocAmt

			Call FixDecimalPlaceByCurrency2(frm1.vspdData,Row,Parent.gCurrency,C_PayLocAmt,"A","X","X")	
	
	End Select
    
End Sub

'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_ChargeDocAmt
            Call EditModeCheck(frm1.vspdData, Row, C_Curr, C_ChargeDocAmt, "A" ,"I", Mode, "X", "X")
        Case C_ChargeLocAmt
            Call EditModeCheck2(frm1.vspdData, Row, Parent.gCurrency, C_ChargeLocAmt, "A" ,"I", Mode, "X", "X")
        Case C_VatDocAmt
            Call EditModeCheck(frm1.vspdData, Row, C_Curr, C_VatDocAmt, "A" ,"I", Mode, "X", "X")        
        Case C_VatLocAmt
            Call EditModeCheck2(frm1.vspdData, Row, Parent.gCurrency, C_VatLocAmt, "A" ,"I", Mode, "X", "X")
        Case C_PayDocAmt
            Call EditModeCheck(frm1.vspdData, Row, C_Curr, C_PayDocAmt, "A" ,"I", Mode, "X", "X")
        Case C_PayLocAmt
            Call EditModeCheck2(frm1.vspdData, Row, Parent.gCurrency, C_PayLocAmt, "A" ,"I", Mode, "X", "X")
    End Select
End Sub

'========================================================================================================= 
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		Call DisableToolBar(Parent.TBC_QUERY)
		Call DbQuery
    End if    
End Sub


'========================================================================================================= 
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

<%    '-----------------------
    'Check previous data area
    '----------------------- %>
    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")										<%'⊙: Clear Contents  Field%>
    Call InitVariables															<%'⊙: Initializes local global variables%>

<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

<%  '-----------------------
    'Query function call area
    '----------------------- %>
	Call SetToolbar("11101111001011")
	Call ggoOper.LockField(Document, "N")

    Call DbQuery																<%'☜: Query db data%>

    FncQuery = True																<%'⊙: Processing is OK%>
        
End Function

'========================================================================================================= 
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          <%'⊙: Processing is NG%>
    
<%  '-----------------------
    'Check previous data area
    '-----------------------%>
    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
<%  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------%>
    Call ggoOper.ClearField(Document, "A")                                      <%'⊙: Clear Condition,contents Field%>

    Call ggoOper.LockField(Document, "N")                                       <%'⊙: Lock  Suitable  Field%>
    Call InitVariables															<%'⊙: Initializes local global variables%>

	'폴더/조회/입력 
	'/삭제/저장/한줄In
	'/한줄Out/취소/이전 
	'/다음/복사/엑셀 
	'/인쇄 
    Call SetToolbar("11101111001011")										'⊙: 버튼 툴바 제어 
    Call SetDefaultVal
    
    Set gActiveElement = document.ActiveElement 
    FncNew = True																<%'⊙: Processing is OK%>

End Function

'========================================================================================================= 
Function FncDelete() 
    
    Exit Function
    Err.Clear                                                               '☜: Protect system from crashing    
    
    FncDelete = False														<%'⊙: Processing is NG%>
    
<%  '-----------------------
    'Precheck area
    '-----------------------%>
    If lgIntFlgMode <> parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002", "X", "X", "X")
        'Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then                                                '☜: Delete db data
       Exit Function                                                        '☜:
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition,Contents Field
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

'========================================================================================================= 
Function FncSave() 
    Dim IntRetCD 
	Dim iDblChargeLocAmt
	Dim iDblVATLocAmt
	    
    FncSave = False                                                         <%'⊙: Processing is NG%>
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	frm1.txtConProcessStepCd.focus
	
<%  '-----------------------
    'Precheck area
    '-----------------------%>
    If lgBlnFlgChgValue = False Or ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        'Call MsgBox("No data changed!!", vbInformation)
        Exit Function
    End If
    
<%  '-----------------------
    'Check content area
    '-----------------------%>
	<%'⊙: If MULTI/SINGLEMULTI %>
    If Not chkField(Document, "2") Then		<% '⊙: Check contents area %>
			Exit Function
		End If

	If Not ggoSpread.SSDefaultCheck Then		<% '⊙: Check contents area %>
		Exit Function
	End If

    ggoSpread.Source = frm1.vspdData

	With frm1
		.vspdData.Col = 0
		If .vspdData.Text = ggoSpread.InsertFlag Or .vspdData.Text = ggoSpread.UpdateFlag Then
			.vspdData.Col = C_PayType
			If .vspdData.text = "NP" Then
				.vspdData.Col = C_ChargeLocAmt	:	iDblChargeLocAmt = Trim(.vspdData.Text)
				.vspdData.Col = C_VatLocAmt		:	iDblVATLocAmt	 = Trim(.vspdData.Text)

				.vspdData.Col = C_PayLocAmt
				If UNICDbl(.vspdData.text) <> UNICDbl(iDblChargeLocAmt) + UNICDbl(iDblVATLocAmt) Then
					Call DisplayMsgBox("206137", "X", "X", "X")
					Exit Function
				End If
			End If
		End If
	End With
	
<%  '-----------------------
    'Save function call area
    '-----------------------%>
    CAll DbSave				                                                <%'☜: Save db data%>
    
    FncSave = True                                                          <%'⊙: Processing is OK%>
    
End Function

'========================================================================================================= 
Function FncCopy() 
	Dim IntRetCD
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error stat
    '----------  Coding part  -------------------------------------------------------------
    FncCopy = False                                                               '☜: Processing is NG
    
    If frm1.vspdData.Maxrows < 1 Then Exit Function
    
	With frm1
		.vspdData.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		
		.vspdData.Row = frm1.vspdData.ActiveRow
		.vspdData.Col = C_PostingFlg
		.vspdData.text = 0
		.vspdData.Col = C_ChargeNo
		.vspdData.text = ""

		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_Curr,C_ChargeDocAmt,"A" ,"I","X","X")         
		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,Parent.gCurrency,C_ChargeLocAmt,"A" ,"I","X","X")         
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_Curr,C_VatDocAmt,"A" ,"I","X","X")         
		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,Parent.gCurrency,C_VatLocAmt,"A" ,"I","X","X")         
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_Curr,C_PayDocAmt,"A" ,"I","X","X")         
		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,Parent.gCurrency,C_PayLocAmt,"A" ,"I","X","X")         

		SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
		
		.vspdData.ReDraw = True
	End With
	
	If Err.number = 0 Then
       FncCopy = True                                                            '☜: Processing is OK
    End If
    
    lgBlnFlgChgValue = True

End Function

'========================================================================================================= 
Function FncCancel() 
	If frm1.vspdData.Maxrows < 1 Then Exit Function
    frm1.vspdData.Redraw = False    	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_Curr,C_ChargeDocAmt,"A" ,"I","X","X")         
	Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,Parent.gCurrency,C_ChargeLocAmt,"A" ,"I","X","X")         
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_Curr,C_VatDocAmt,"A" ,"I","X","X")         
	Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,Parent.gCurrency,C_VatLocAmt,"A" ,"I","X","X")         
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,C_Curr,C_PayDocAmt,"A" ,"I","X","X")         
	Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,Parent.gCurrency,C_PayLocAmt,"A" ,"I","X","X")         
    frm1.vspdData.Redraw = True     
End Function

'========================================================================================================= 
Function FncInsertRow(pvRowCnt) 
	Dim IntRetCD
    Dim imRow,i
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False 
    
    If IsNumeric(Trim(pvRowCnt)) then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End If
	    
	With frm1
		.vspdData.ReDraw = False
		
		.vspdData.focus
		ggoSpread.Source = .vspdData

		ggoSpread.InsertRow, imRow

		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
    
		lgBlnFlgChgValue = True
    <% '----------  Coding part  -------------------------------------------------------------%>   

		For i = .vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1
		.vspdData.Row = i			

		.vspdData.Col = C_ChargeDt
		.vspdData.Text = "<%=EndDate%>"
		.vspdData.Col = C_ChargeDocAmt
		.vspdData.Text = 0

		'.vspdData.Col = C_Curr
		'.vspdData.Text = parent.gCurrency
		.vspdData.Col = C_XchRate
		.vspdData.Text = 0
		.vspdData.Col = C_XchCalop
		.vspdData.Text = "*"
		.vspdData.Col = C_ChargeLocAmt
		.vspdData.Text = 0
		.vspdData.Col = C_VatRate
		.vspdData.Text = 0
		.vspdData.Col = C_VatDocAmt
		.vspdData.Text = 0

		'.vspdData.Col = C_PayDueDt
		'.vspdData.Text = "<%=EndDate%>"
		.vspdData.Col = C_PayDocAmt
		.vspdData.Text = 0
		.vspdData.Col = C_PayLocAmt
		.vspdData.Text = 0

		Next

		.vspdData.ReDraw = True
    End With
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If  
    
    Set gActiveElement = document.ActiveElement 

End Function

'========================================================================================================= 
Function FncDeleteRow() 

    If frm1.vspdData.Maxrows < 1 Then Exit Function

    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1  

    .vspdData.focus
    ggoSpread.Source = .vspdData 
    
    <% '----------  Coding part  -------------------------------------------------------------%>   
	lDelRows = ggoSpread.DeleteRow
	
    lgBlnFlgChgValue = True
    
    End With
    
    Set gActiveElement = document.ActiveElement
   
End Function

'========================================================================================================= 
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================================= 
Function FncPrev() 
    On Error Resume Next                                                    <%'☜: Protect system from crashing%>

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
        'Call MsgBox("조회한후에 지원됩니다.", vbInformation)
        Exit Function
    ElseIf lgPrevNo = "" Then
		Call DisplayMsgBox("900011", "X", "X", "X")  '☜ 바뀐부분 
		'Call MsgBox("이전 데이타가 없습니다..", vbInformation)
    End If

End Function

'========================================================================================================= 
Function FncNext() 
    On Error Resume Next                                                    <%'☜: Protect system from crashing%>

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
        'Call MsgBox("조회한후에 지원됩니다.", vbInformation)
        Exit Function
    ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")  '☜ 바뀐부분 
		'Call MsgBox("이전 데이타가 없습니다..", vbInformation)
    End If

End Function

'========================================================================================================= 
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'========================================================================================================= 
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLEMULTI, False)
End Function

'========================================================================================================= 
Sub FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub


'========================================================================================================= 
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================= 
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    
	Call ggoSpread.ReOrderingSpreadData()
	Call SetQuerySpreadColor(1)

    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, -1, -1 ,C_Curr,C_ChargeDocAmt,"A","I","X","X")
    Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData, -1, -1 ,parent.gCurrency,C_ChargeLocAmt,"A","I","X","X")
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, -1, -1 ,C_Curr,C_VatDocAmt,"A","I","X","X")
    Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData, -1, -1 ,parent.gCurrency,C_VatLocAmt,"A","I","X","X")
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, -1, -1 ,C_Curr,C_PayDocAmt,"A","I","X","X")
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData, -1, -1 ,C_Curr,C_PayLocAmt,"A","I","X","X")


End Sub

'========================================================================================================= 
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'========================================================================================================= 
Function DbDelete() 
    On Error Resume Next                                                    <%'☜: Protect system from crashing%>
End Function

'========================================================================================================= 
Function DbDeleteOk()														<%'☆: 삭제 성공후 실행 로직 %>
    On Error Resume Next                                                    <%'☜: Protect system from crashing%>
End Function

'========================================================================================================= 
Function DbQuery() 

    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
    DbQuery = False                                                         <%'⊙: Processing is NG%>

	Call LayerShowHide(1)
	    
    Dim strVal

    If lgIntFlgMode = parent.OPMD_UMODE Then    
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtProcessStepCd=" & Trim(frm1.txtHConProcessStepCd.value)		<%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtBasNo=" & Trim(frm1.txtHConBasNo.value)
		strVal = strVal & "&txtSalesGrp=" & Trim(frm1.txtHConSalesGrp.value)		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtProcessStepCd=" & Trim(frm1.txtConProcessStepCd.value)	<%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtBasNo=" & Trim(frm1.txtConBasNo.value)
		strVal = strVal & "&txtSalesGrp=" & Trim(frm1.txtConSalesGrp.value)		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End If	

	frm1.txtPrevMaxRow.value = Trim(frm1.vspdData.MaxRows)

	Call RunMyBizASP(MyBizASP, strVal)												<%'☜: 비지니스 ASP 를 가동 %>
	
    DbQuery = True																	<%'⊙: Processing is NG%>

End Function

'========================================================================================================= 
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												<%'⊙: Indicates that current mode is Update mode%>
  
    Call ggoOper.LockField(Document, "Q")									<%'⊙: This function lock the suitable field%>
    Call SetToolbar("11101111001111")					   
	Call SetQuerySpreadColor(1)
	lgBlnFlgChgValue = False

    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus		
    Else
       frm1.txtConProcessStepCd.focus
    End If     

End Function

'========================================================================================
Function DbSave() 

    Err.Clear																<%'☜: Protect system from crashing%>
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	
    DbSave = False                                                          '⊙: Processing is NG
    
	Call LayerShowHide(1)

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 0    
		strVal = ""
		strDel = ""
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag							'☜: 신규 
					strVal = strVal & "C" & parent.gColSep	& lRow & parent.gColSep'☜: C=Create
		        Case ggoSpread.UpdateFlag							'☜: 수정 
					strVal = strVal & "U" & parent.gColSep	& lRow & parent.gColSep'☜: U=Update
				Case ggoSpread.DeleteFlag							'☜: 삭제 
					strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep'☜: D=Delete
		            '--- 경비관리번호 
		            .vspdData.Col = C_ChargeNo
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_PostingFlg 		            
		            If Trim(.vspdData.Text) = "1" Then
						strDel = strDel & "Y" & parent.gColSep
		            Else
						strDel = strDel & "N" & parent.gColSep
		            End If

					'--- 경비항목 
		            .vspdData.Col = C_ChargeType 		
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
					'--- 거래처 
		            .vspdData.Col = C_BpCd
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
					'--- 발생일 
		            .vspdData.Col = C_ChargeDt 		
		            strDel = strDel & Trim(UNIConvDate(.vspdData.Text)) & parent.gColSep
					'--- 계산서종류 
		            .vspdData.Col = C_VatType 		
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
					'--- 화폐	
		            .vspdData.Col = C_Curr 		
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
					'--- 발생금액 
		            .vspdData.Col = C_ChargeDocAmt 		
		            strDel = strDel & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- 환율 
		            .vspdData.Col = C_XchRate 		
		            strDel = strDel & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- 자국금액 
		            .vspdData.Col = C_ChargeLocAmt 		
		            strDel = strDel & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- VAT율 
		            .vspdData.Col = C_VatRate 		
		            strDel = strDel & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- VAT발생금액 
		            .vspdData.Col = C_VatDocAmt 		
		            strDel = strDel & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- VAT자국금액 
		            .vspdData.Col = C_VatLocAmt 		
		            strDel = strDel & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- 지급유형 
		            .vspdData.Col = C_PayType 		
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            '--- 어음번호 
		            .vspdData.Col = C_CheckNo 		
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
					'--- 물대포함 
					strDel = strDel & "C" & parent.gColSep
					'--- 출금계좌 
		            .vspdData.Col = C_BankAcct 		
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep		       
					'--- 출금은행코드 
		            .vspdData.Col = C_BankCd 		            
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
					'--- 기타참조사항 
		            .vspdData.Col = C_RefRemark 		
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep

					'--- 지급만기일 
		            .vspdData.Col = C_PayDueDt 		
		            strDel = strDel & Trim(UNIConvDate(.vspdData.Text)) & parent.gColSep
					'--- 세금신고사업장 
		            .vspdData.Col = C_TaxBizArea 		
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
					'--- 지급액 
		            .vspdData.Col = C_PayDocAmt
		            strDel = strDel & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- 지급자국액 
		            .vspdData.Col = C_PayLocAmt
		            strDel = strDel & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
		            
		            '--- 환율연산자 
		            .vspdData.Col = C_XchCalop
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep					
		            
		            lGrpCnt = lGrpCnt + 1 
			End Select

			Select Case .vspdData.Text
				Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

					'--- 확정여부 
		            .vspdData.Col = C_PostingFlg 		            
		            If Trim(.vspdData.Text) = "1" Then
						strVal = strVal & "Y" & parent.gColSep
		            Else
						strVal = strVal & "N" & parent.gColSep
		            End If
		            '--- 경비관리번호 
		            .vspdData.Col = C_ChargeNo
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					'--- 경비항목 
		            .vspdData.Col = C_ChargeType 		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					'--- 거래처 
		            .vspdData.Col = C_BpCd
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            '--- 발생일 
		            .vspdData.Col = C_ChargeDt 		
		            strVal = strVal & Trim(UNIConvDate(.vspdData.Text)) & parent.gColSep
					'--- 계산서종류 
		            .vspdData.Col = C_VatType 		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					'--- 화폐	
		            .vspdData.Col = C_Curr 		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					'--- 발생금액 
		            .vspdData.Col = C_ChargeDocAmt 		
		            strVal = strVal & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- 환율 
		            .vspdData.Col = C_XchRate 		
		            strVal = strVal & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- 자국금액 
		            .vspdData.Col = C_ChargeLocAmt 		
		            strVal = strVal & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- VAT율 
		            .vspdData.Col = C_VatRate 		
		            strVal = strVal & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- VAT발생금액 
		            .vspdData.Col = C_VatDocAmt 		
		            strVal = strVal & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- VAT자국금액 
		            .vspdData.Col = C_VatLocAmt 		
		            strVal = strVal & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- 지급유형 
		            .vspdData.Col = C_PayType 		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            '--- 어음번호 
		            .vspdData.Col = C_CheckNo 		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					'--- 물대포함 
					strVal = strVal & "C" & parent.gColSep
					'--- 출금계좌 
		            .vspdData.Col = C_BankAcct 		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		       
					'--- 출금은행코드 
		            .vspdData.Col = C_BankCd 		            
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					'--- 기타참조사항 
		            .vspdData.Col = C_RefRemark 		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

					'--- 지급만기일 
		            .vspdData.Col = C_PayDueDt 		
		            strVal = strVal & Trim(UNIConvDate(.vspdData.Text)) & parent.gColSep
					'--- 세금신고사업장 
		            .vspdData.Col = C_TaxBizArea 		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					'--- 지급액 
		            .vspdData.Col = C_PayDocAmt
		            strVal = strVal & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
					'--- 지급자국액 
		            .vspdData.Col = C_PayLocAmt
		            strVal = strVal & Trim(UNIConvNum(.vspdData.Text,0)) & parent.gColSep
		            
		            '--- 환율연산자 
		            .vspdData.Col = C_XchCalop
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

					'--- ext1_qty 
		            strVal = strVal & "0" & parent.gColSep
		            '--- ext2_qty
		            strVal = strVal & "0" & parent.gColSep
		            '--- ext3_qty
		            strVal = strVal & "0" & parent.gColSep
		            '--- ext1_amt
		            strVal = strVal & "0" & parent.gColSep
		            '--- ext2_amt
		            strVal = strVal & "0" & parent.gColSep
		            '--- ext3_amt
		            strVal = strVal & "0" & parent.gColSep
		            '--- ext1_cd
		            .vspdData.Col = C_ext1_Cd
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            '--- ext2_cd
		            .vspdData.Col = C_ext1_nm
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            '--- ext3_cd
		            strVal = strVal & "" & parent.gRowSep		            
		            
		            lGrpCnt = lGrpCnt + 1 
		    End Select       
		Next
	
		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
Function DbSaveOk()															<%'☆: 저장 성공후 실행 로직 %>

    Call ggoOper.ClearField(Document, "2")
    Call InitVariables    
    Call FncQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>판매경비등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>진행구분</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConProcessStepCd" ALT="진행구분" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcessStep" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenHdrSalesCharge frm1.txtConProcessStepCd,frm1.txtConProcessStepNm,1">&nbsp;<INPUT NAME="txtConProcessStepNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConSalesGrp" ALT="영업그룹" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenHdrSalesCharge frm1.txtConSalesGrp,frm1.txtConSalesGrpNm,2">&nbsp;<INPUT NAME="txtConSalesGrpNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>발생근거번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConBasNo" TYPE=TEXT SIZE=20 ALT="발생근거번호" MAXLENGTH=18 TAG="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBasNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProcessStep frm1.txtConProcessStepCd,frm1.txtConBasNo,frm1.txtConBasNo,1"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>진행구분</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtProcessStepCd" ALT="진행구분" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcessStep" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenHdrSalesCharge frm1.txtProcessStepCd,frm1.txtProcessStepNm,3">&nbsp;<INPUT NAME="txtProcessStepNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>영업그룹</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrp" ALT="영업그룹" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenHdrSalesCharge frm1.txtSalesGrp,frm1.txtSalesGrpNm,4">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발생근거번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBasNo" TYPE=TEXT SIZE=20 ALT="발생근거번호" MAXLENGTH=18 TAG="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBasNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProcessStep frm1.txtProcessStepCd,frm1.txtBasNo,frm1.txtBasDocNo,2"></TD>
								<TD CLASS=TD5 NOWRAP>발생문서번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBasDocNo" SIZE=35 MAXLENGTH=35 TAG="24"></TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>> <IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtHConProcessStepCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHConBasNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHConSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHAcctFlag" tag="24">

<INPUT TYPE=HIDDEN NAME="txtPrevMaxRow" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
