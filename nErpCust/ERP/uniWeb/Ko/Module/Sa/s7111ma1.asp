<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : S7111MA1    																*
'*  4. Program Name         : NEGO 등록																	*
'*  5. Program Desc         : NEGO 등록    																*
'*  6. Comproxy List        : PSAG111.dll, PSAG119.dll               									*
'*  7. Modified date(First) : 2000/05/08																*
'*  8. Modified date(Last)  : 2000/05/08																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : An Chang Hwan																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/05/08 : 화면 design												*
'*							  2. 2000/05/08 : Coding Start												*
'******************************************************************************************************** 
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBS">
Option Explicit					<% '☜: indicates that All variables must be declared in advance %>

'============================================  1.2.1 Global 상수 선언  ==================================
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

	Const BIZ_PGM_ID = "s7111mb1.asp"						'☆: 비지니스 로직 ASP명 
	Const BIZ_PGM_BASDATAQUERY_ID = "s7111mb2.asp"			'☆: 비지니스 로직 ASP명 : 매출채권참조 
	Const BIZ_PGM_POSTING_ID = "s7111mb3.asp"				'☆: 비지니스 로직 ASP명 : 확정 
	Const EXPORT_CHARGE_ENTRY_ID = "s6111ma1"				'☆: 이동할 ASP명 
	Const TAB1 = 1
	Const TAB2 = 2

	'------ Minor Code PopUp을 위한 Major Code정의 ------ 
	Const gstrSubTypeMajor = "S9072"
	Const gstrCollectTypeMajor = "B9004"
	Const gstrNegoTypeMajor = "S9071"

'============================================  1.2.2 Global 변수 선언  ==================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
	
	Dim gSelframeFlg								'현재 TAB의 위치를 나타내는 Flag 
	Dim gblnWinEvent								'~~~ ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
													'	 PopUp Window가 사용중인지 여부를 나타내는 variable 
	Dim lgBlnFlawChgFlg	
	Dim gtxtChargeType
	Dim glsTab
	Dim gTabMaxCnt

'========================================================================================================
	Function InitVariables()
		lgIntFlgMode = parent.OPMD_CMODE						<%'⊙: Indicates that current mode is Create mode%>
		lgBlnFlgChgValue = False								<%'⊙: Indicates that no value changed%>
		lgBlnFlawChgFlg = False
		lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
		
		<% '------ Coding part ------ %>
		gblnWinEvent = False
	End Function
	
'========================================================================================================
	Sub SetDefaultVal()
		frm1.txtNegoDt.text = EndDate
		frm1.txtPayExpiryDt.text = EndDate
		frm1.txtNegoReqDt.text = EndDate
		frm1.txtNegoDocAmt.text = UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
		frm1.txtNegoLocAmt.text = UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
		frm1.txtXchRate.text = UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
		frm1.txtLocCurrency.value = parent.gCurrency
		frm1.btnPosting.disabled = True
		lgBlnFlgChgValue = False
	End Sub
	
'========================================================================================================
	Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
		<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
	End Sub	

'========================================================================================================
	Function ClickTab1()
		If gSelframeFlg = TAB1 Then Exit Function
		
		Call changeTabs(TAB1)
		frm1.txtNEGONo.focus
		
		gSelframeFlg = TAB1
	End Function

	Function ClickTab2()
		If gSelframeFlg = TAB2 Then Exit Function
		
		Call changeTabs(TAB2)
		frm1.txtNEGONo.focus
		
		gSelframeFlg = TAB2
	End Function

'===NEGO관리번호=============================================================================================
	Function OpenNEGONoPop()
		Dim iCalledAspName
		Dim strRet
		
		If gblnWinEvent = True Or UCase(frm1.txtNEGONo.className) = "PROTECTED" Then Exit Function
		
		iCalledAspName = AskPRAspName("S7111PA1")
		
		if Trim(iCalledAspName) = "" then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S7111PA1", "x")
			gblnWinEvent = False
			exit Function
		end if

		gblnWinEvent = True
		
		strRet = window.showModalDialog(iCalledAspName, array(window.parent), _
				"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False
		
		If strRet(0) = "" Then
			frm1.txtNEGONo.focus
			Exit Function
		Else
			Call SetNEGONo(strRet)
			frm1.txtNEGONo.focus
		End If
	End Function

'===매출채권참조=============================================================================================
	Function OpenBillRef()
		Dim iCalledAspName
		Dim strRet

		If lgIntFlgMode = parent.OPMD_UMODE Then 
			Call DisplayMsgBox("200005", "x", "x", "x")
			Exit function
		End If 
		iCalledAspName = AskPRAspName("S5111RA1")
		
		if Trim(iCalledAspName) = "" then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S5111RA1", "x")
			exit Function
		end if
		
		strRet = window.showModalDialog(iCalledAspName, array(window.parent), _
				"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
		If strRet(0) = "" Then
			Exit Function
		Else
			Call SetBillRef(strRet)
		End If
	End Function

'========================================================================================================
	Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strMajorCd)
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = strPopPos								<%' 팝업 명칭 %>
		arrParam(1) = "B_Minor"								<%' TABLE 명칭 %>
		arrParam(2) = Trim(strMinorCD)						<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""		<%' Where Condition%>
		arrParam(5) = strPopPos								<%' TextBox 명칭 %>

		arrField(0) = "Minor_CD"							<%' Field명(0)%>
		arrField(1) = "Minor_NM"							<%' Field명(1)%>

		arrHeader(0) = strPopPos							<%' Header명(0)%>
		arrHeader(1) = strPopPos  & "명"				<%' Header명(1)%>
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetMinorCd(strMajorCd, arrRet)
		End If
	End Function

'========================================================================================================
	Function OpenCollectType()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = "입금유형"						<%' 팝업 명칭 %>
		arrParam(1) = "B_Configuration A, (Select * From B_Minor Where Major_cd = " & FilterVar("A1006", "''", "S") & ") B"								<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtCollectType.value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "a.major_cd = " & FilterVar("B9004", "''", "S") & " and a.reference = b.minor_cd and a.minor_cd =  " & FilterVar(frm1.txtPayTerms.value , "''", "S") & ""					<%' Where Condition%>
		arrParam(5) = "입금유형"						<%' TextBox 명칭 %>

		arrField(0) = "A.Reference"							<%' Field명(0)%>
		arrField(1) = "B.minor_nm"							<%' Field명(1)%>

		arrHeader(0) = "입금유형"						<%' Header명(0)%>
		arrHeader(1) = "입금유형명"						<%' Header명(1)%>

		frm1.txtCollectType.focus 
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetCollectType(arrRet)
		End If
	End Function

'========================================================================================================
	Function OpenSalesGroup()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = "영업그룹"						<%' 팝업 명칭 %>
		arrParam(1) = "B_SALES_GRP"							<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtSalesGroup.value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						<%' Where Condition%>
		arrParam(5) = "영업그룹"						<%' TextBox 명칭 %>

		arrField(0) = "SALES_GRP"							<%' Field명(0)%>
		arrField(1) = "SALES_GRP_NM"						<%' Field명(1)%>

		arrHeader(0) = "영업그룹"						<%' Header명(0)%>
		arrHeader(1) = "영업그룹명"						<%' Header명(1)%>

		frm1.txtSalesGroup.focus 
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetSalesGroup(arrRet)
		End If
	End Function

'========================================================================================================
	Function OpenBankPop()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = "은행"							<%' 팝업 명칭 %>
		arrParam(1) = "B_BANK"								<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtIncomeBank.value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "은행"							<%' TextBox 명칭 %>

		arrField(0) = "BANK_CD"								<%' Field명(0)%>
		arrField(1) = "BANK_NM"								<%' Field명(1)%>

		arrHeader(0) = "은행"							<%' Header명(0)%>
		arrHeader(1) = "은행명"							<%' Header명(1)%>

		frm1.txtIncomeBank.focus 
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetBank(arrRet)
		End If
	End Function
'========================================================================================================
	Function OpenNegoBankPop()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = "은행"							<%' 팝업 명칭 %>
		arrParam(1) = "B_BANK"								<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtIncomeBank.value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "은행"							<%' TextBox 명칭 %>

		arrField(0) = "BANK_CD"								<%' Field명(0)%>
		arrField(1) = "BANK_NM"								<%' Field명(1)%>

		arrHeader(0) = "은행"							<%' Header명(0)%>
		arrHeader(1) = "은행명"							<%' Header명(1)%>

		frm1.txtIncomeBank.focus 
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetNegoBank(arrRet)
		End If
	End Function

'========================================================================================================
	Function OpenBankAcct()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)
		Dim IncomeBank
		
		If gblnWinEvent = True Then Exit Function
		
		If Len(frm1.txtIncomeBank.value) Then
			IncomeBank =  frm1.txtIncomeBank.value
		Else
			Call DisplayMsgBox("205152", "x", frm1.txtIncomeBank.Alt, "x")
			frm1.txtIncomeBank.focus 
			Exit Function
		End If
		
		gblnWinEvent = True

		arrParam(0) = "입금계좌"																<%' 팝업 명칭 %>
		arrParam(1) = "B_BANK_ACCT, B_BANK"															<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtAccountNo.value)													<%' Code Condition%>
		arrParam(3) = ""																			<%' Name Cindition%>
		arrParam(4) = "B_BANK_ACCT.BANK_CD=B_BANK.BANK_CD And B_BANK.BANK_CD =  " & FilterVar(IncomeBank , "''", "S") & ""		<%' Where Condition%>
		arrParam(5) = "입금계좌"																<%' TextBox 명칭 %>

		arrField(0) = "B_BANK_ACCT.BANK_ACCT_NO"													<%' Field명(0)%>
		arrField(1) = "ED10" & parent.gColSep & "B_BANK_ACCT.BANK_CD"															<%' Field명(1)%>
		arrField(2) = "ED20" & parent.gColSep & "B_BANK.BANK_NM"																<%' Field명(2)%>

		arrHeader(0) = "입금계좌"																<%' Header명(0)%>
		arrHeader(1) = "입금은행"																<%' Header명(1)%>
		arrHeader(2) = "입금은행명"																<%' Header명(2)%>

		frm1.txtAccountNo.focus 
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetBankAcct(arrRet)
		End If
	End Function

'========================================================================================================
	Function SetNEGONo(strRet)
		frm1.txtNEGONo.value = strRet(0)
	End Function
'========================================================================================================
	Function SetCollectType(strRet)
		frm1.txtCollectType.value = strRet(0)
		frm1.txtCollectTypeNm.value = strRet(1)
		
		lgBlnFlgChgValue = True
	End Function
'========================================================================================================
	Function SetBillRef(strRet)
		Call ggoOper.ClearField(Document, "2")											<% '⊙: Clear Contents  Field %>
		Call SetRadio()
		Call SetDefaultVal
		
		frm1.txtHBillNo.value = strRet(0)
		frm1.txtHBLDocNo.value = strRet(1)
		frm1.txtHLCNo.value = strRet(2)
		frm1.txtHBLFlag.value = strRet(3)
		
		Dim strVal
		Call LayerShowHide(1)		
		
		strVal = BIZ_PGM_BASDATAQUERY_ID & "?txtBillNo=" & Trim(frm1.txtHBillNo.value)		<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtBLNo=" & Trim(frm1.txtHBLDocNo.value)
		strVal = strVal & "&txtLCNo=" & Trim(frm1.txtHLCNo.value)						<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtBLFlag=" & Trim(frm1.txtHBLFlag.value)

		Call RunMyBizASP(MyBizASP, strVal)												<%'☜: 비지니스 ASP 를 가동 %>
	
		
		frm1.btnPosting.disabled = True
		
		
		lgBlnFlgChgValue = True
	End Function

'========================================================================================================
	Function SetMinorCd(strMajorCd, arrRet)
		Select Case strMajorCd
			Case gstrSubTypeMajor
				frm1.txtSubType.Value = arrRet(0)

			Case gstrCollectTypeMajor
				frm1.txtCollectType.Value = arrRet(0)
				frm1.txtCollectTypeNm.Value = arrRet(1)
			
			Case gstrNegoTypeMajor
				frm1.txtNegoType.value = arrRet(0)
				frm1.txtNegoTypeNm.value = arrRet(1)
			Case Else
		End Select

		lgBlnFlgChgValue = True
	End Function

'========================================================================================================
	Function SetSalesGroup(arrRet)
		frm1.txtSalesGroup.value = arrRet(0)
		frm1.txtSalesGroupNm.value = arrRet(1)

		lgBlnFlgChgValue = True
	End Function

'========================================================================================================
	Function SetBank(arrRet)
		frm1.txtIncomeBank.Value = arrRet(0)
		frm1.txtIncomeBankNm.Value = arrRet(1)

		lgBlnFlgChgValue = True
	End Function
'========================================================================================================
	Function SetNegoBank(arrRet)
		frm1.txtNegoBank.Value = arrRet(0)
		frm1.txtNegoBankNm.Value = arrRet(1)

		lgBlnFlgChgValue = True
	End Function
'========================================================================================================
	Function SetBankAcct(arrRet)
		frm1.txtAccountNo.Value = arrRet(0)
		lgBlnFlgChgValue = True
	End Function

'========================================================================================================
	Function SetRadio()
		Dim blnOldFlag

		blnOldFlag = lgBlnFlgChgValue

		frm1.rdoFlawExist2.checked = True
		frm1.rdoPostingflg2.checked = True

		lgBlnFlgChgValue = blnOldFlag
	End Function

'========================================================================================================
	Function LoadExportCharge()
		Dim strDtlOpenParam

		WriteCookie "txtChargeType", "EN"
		WriteCookie "txtBasNo", UCase(Trim(frm1.txtNEGONo.value))

		strDtlOpenParam = EXPORT_CHARGE_ENTRY_ID

		PgmJump(EXPORT_CHARGE_ENTRY_ID)
	End Function

'========================================================================================================
	Function PostNego()
		If Trim(frm1.txtHNEGONo.value) = "" Then
			Call DisplayMsgBox("900002", "x", "x", "x")	<% '⊙: "Will you destory previous data" %>
			'Call MsgBox("조회를 선행하십시오.", parent.VB_INFORMATION)
			Exit Function
		End If

		Dim strVal

		Call LayerShowHide(1)

		strVal = BIZ_PGM_POSTING_ID & "?txtNEGONo=" & Trim(frm1.txtNEGONo.value)		<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtgChangeOrgId=" & parent.gChangeOrgId
		strVal = strVal & "&txtUpdtUserId=" & parent.gUsrID									<%'☆: 조회 조건 데이타 %>

		Call RunMyBizASP(MyBizASP, strVal)												<%'☜: 비지니스 ASP 를 가동 %>
	End Function
	
'========================================================================================================
	Function PostingOk()
		Dim blnOldFlag

		blnOldFlag = lgBlnFlgChgValue

		frm1.rdoPostingflg2.checked = True

		lgBlnFlgChgValue = blnOldFlag
		
		Call FncQuery()
	End Function

<%
'========================================================================================================
'=	Event Desc : 수주참조나 L/C 참조가 되었을때 해당 textbox 를 Protect 시킨다.							=
'========================================================================================================
%>
	Sub ProtectFlawRelTag()
		With frm1
			Call ggoOper.SetReqAttr(.txtCollectType, "Q")
			Call ggoOper.SetReqAttr(.txtPayDt, "Q")
			Call ggoOper.SetReqAttr(.txtAccountNo, "Q")
			Call ggoOper.SetReqAttr(.txtIncomeBank, "Q")
		End With
	End Sub	
<%
'========================================================================================================
'=	Event Desc : 출고참조시 수주참조나 L/C 참조로인해 Protect 되어있는 textbox 를 Release 한다.			=
'========================================================================================================
%>
	Sub ReleaseFlawRelTag()
		With frm1
			Call ggoOper.SetReqAttr(.txtCollectType, "N")
			Call ggoOper.SetReqAttr(.txtPayDt, "N")
			Call ggoOper.SetReqAttr(.txtAccountNo, "D")
			Call ggoOper.SetReqAttr(.txtIncomeBank, "D")
		End With			
	End Sub

'========================================================================================================
	Function CookiePage(ByVal Kubun)

		On Error Resume Next

		Const CookieSplit = 4877						<%'Cookie Split String : CookiePage Function Use%>
		Dim strTemp, arrVal

		Select Case Kubun
		
		Case 1
			WriteCookie CookieSplit , frm1.txtBLNo.value
		Case 0
			strTemp = ReadCookie(CookieSplit)

			If strTemp = "" then Exit Function
				
			frm1.txtNegoNo.value =  strTemp
			
			If Err.number <> 0 Then
				Err.Clear
				WriteCookie CookieSplit , ""
				Exit Function 
			End If
			
			Call FncQuery()
						
			WriteCookie CookieSplit , ""
		Case 2	
			WriteCookie CookieSplit , "EN" & parent.gRowSep & frm1.txtSalesGroup.value & parent.gRowSep & frm1.txtSalesGroupNm.value & parent.gRowSep & frm1.txtNEGONo.value 
		
		End Select 		
	End Function
	

<%
'==========================================================================================
'   Event Desc : All Document Body Protected
'==========================================================================================
%>
	Function ProtectBody()

	    On Error Resume Next
	    
		Dim elmCnt, strTagName

		For elmCnt = 1 to frm1.length - 1
			If Left(frm1.elements(elmCnt).getAttribute("tag"),1) = "2" Then
				Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "Q")
			End If

			If Err.number <> 0 Then	Err.Clear
		Next
		

	End Function

<%
'==========================================================================================
'   Event Desc : All Document Body Release
'==========================================================================================
%>
	Function ReleaseBody()
	    On Error Resume Next
	    
		Dim elmCnt, strTagName

		For elmCnt = 1 to frm1.length - 1
			If Left(frm1.elements(elmCnt).getAttribute("tag"),1) = "2" Then
				Select Case Left(frm1.elements(elmCnt).getAttribute("tag"),2)
				Case "21"
					Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "D")
				Case "22","23"
					Call ggoOper.SetReqAttr(frm1.elements(elmCnt), "N")
				End Select
			End If

			If Err.number <> 0 Then	Err.Clear
		Next
	End Function

'============================================================================================================
	Function ProtectXchRate()
		If frm1.txtCurrency.value = parent.gCurrency Then
			Call ggoOper.SetReqAttr(frm1.txtXchRate, "Q")
		End If	
	End Function

'====================================================================================================
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
	Sub CurFormatNumericOCX()
		With frm1

			'NegoL 금액 
			ggoOper.FormatFieldByObjectOfCur .txtNegoDocAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
			'근거금액 
			ggoOper.FormatFieldByObjectOfCur .txtBaseDocAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
			 '환율 
			ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec

		End With
	End Sub

'========================================================================================================
	Sub Form_Load()
		Call LoadInfTB19029
		Call AppendNumberPlace("6","3","2")												<% '⊙: Load table , B_numeric_format %>
		Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
		Call ggoOper.LockField(Document, "N")											<% '⊙: Lock  Suitable  Field %>
		Call SetDefaultVal
		Call SetToolbar("11100000000011")												<% '⊙: 버튼 툴바 제어 %>
		Call InitVariables
		Call changeTabs(TAB1)		

		frm1.txtNEGONo.focus
		Set gActiveElement = document.activeElement 
		
		Call rdoFlawExist2_OnClick
		Call CookiePage(0)
        glsTab = "Y"
        gTabMaxCnt = 2
        
	End Sub
	
'========================================================================================================
	Sub Form_QueryUnload(Cancel, UnloadMode)
	End Sub
	
'========================================================================================================
	Sub btnNEGONoOnClick()
		Call OpenNEGONoPop()
	End Sub

'========================================================================================================
	Sub btnSubTypeOnClick()
		If frm1.txtSubType.readOnly <> True Then
			frm1.txtSubType.focus
			Call OpenMinorCd(frm1.txtSubType.value, "", "환차손익", gstrSubTypeMajor)
		End If
	End Sub
'========================================================================================================
	Sub btnNegoTypeOnClick()
		If frm1.txtNegoType.readOnly <> True Then
			frm1.txtNegoType.focus 
			Call OpenMinorCd(frm1.txtNegoType.value, frm1.txtNegoTypeNm.value , "NEGO유형", gstrNegoTypeMajor)
		End If
	End Sub

'========================================================================================================
	Sub btnCollectTypeOnClick()
		If frm1.txtCollectType.readOnly <> True Then
			Call OpenCollectType()
		End If
	End Sub

'========================================================================================================
	Sub btnNegoBankOnClick()
		If frm1.txtNegoBank.readOnly <> True Then
			Call OpenNegoBankPop()
		End If
	End Sub

'========================================================================================================
	Sub btnSalesGroupOnClick()
		If frm1.txtSalesGroup.readOnly <> True Then
			Call OpenSalesGroup()
		End If
	End Sub

'========================================================================================================
	Sub btnAccountNoOnClick()
		If frm1.txtAccountNo.readOnly <> True Then
			Call OpenBankAcct()
		End If
	End Sub

'========================================================================================================
	Sub btnIncomeBankOnClick()
		If frm1.txtIncomeBank.readOnly <> True Then
			Call OpenBankPop()
		End If
	End Sub

'========================================================================================================
	Sub btnPosting_OnClick()
		If frm1.btnPosting.disabled <> True Then
			Dim IntRetCD

			IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")

			If IntRetCD = vbNo Then
				Exit Sub
			End If

			Call PostNego()
		End If
	End Sub

'========================================================================================================
	Sub rdoFlawExist1_OnClick()
		If frm1.rdoPostingflg2.checked = True Then    '확정에 따른 모든필드를 Protect 시키고 하서 OnClick Event 가 발생하기때문에 Check 를 해줌 
			Call ReleaseFlawRelTag()
		End If
	End Sub
'========================================================================================================
	Sub rdoFlawExist2_OnClick()
		frm1.txtPayDt.text = ""
		frm1.txtCollectType.value = ""
		frm1.txtCollectTypeNm.value = ""		
		frm1.txtAccountNo.value = ""
		frm1.txtIncomeBank.value = ""
		frm1.txtIncomeBankNm.value = ""
		Call ProtectFlawRelTag()
	End Sub
'========================================================================================================
	Sub txtNegoDt_Change()
		lgBlnFlgChgValue = True
	End Sub

'========================================================================================================
	Sub txtXchCommRate_Change()
		lgBlnFlgChgValue = True
	End Sub

'========================================================================================================
	Sub txtPayExpiryDt_Change()
		lgBlnFlgChgValue = True
	End Sub
'========================================================================================================
	Sub txtNegoReqDt_Change()
		lgBlnFlgChgValue = True
	End Sub
'========================================================================================================
	Sub txtNegoLocAmt_Change()
		lgBlnFlgChgValue = True
	End Sub
'========================================================================================================
	Sub txtPayDt_Change()
		lgBlnFlgChgValue = True
	End Sub
'========================================================================================================
	Sub txtCollectType_Change()
		lgBlnFlgChgValue = True
	End Sub

'=======================================================================================================%>
	Sub txtNegoDt_DblClick(Button)
	    If Button = 1 Then
	        frm1.txtNegoDt.Action = 7 
			Call SetFocusToDocument("M")	
			Frm1.txtNegoDt.Focus
	    End If
	End Sub

	Sub txtPayExpiryDt_DblClick(Button)
	    If Button = 1 Then
	        frm1.txtPayExpiryDt.Action = 7
	    	Call SetFocusToDocument("M")	
			Frm1.txtPayExpiryDt.Focus
	    End If
	End Sub	

	Sub txtPayDt_DblClick(Button)
	    If Button = 1 Then
	        frm1.txtPayDt.Action = 7
	    	Call SetFocusToDocument("M")	
			Frm1.txtPayDt.Focus
	    End If
	End Sub	

	Sub txtNegoReqDt_DblClick(Button)
	    If Button = 1 Then
	        frm1.txtNegoReqDt.Action = 7
	    	Call SetFocusToDocument("M")	
			Frm1.txtNegoReqDt.Focus
	    End If
	End Sub	

'========================================================================================================
	Sub txtXchRate_Change()
		If frm1.txtCurrency.value = parent.gCurrency Then
			frm1.txtXchRate.text = 1
			frm1.txtNegoLocAmt.text =  UNIFormatNumber(UNICDbl(frm1.txtNegoDocAmt.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
		Else
			If frm1.txtExchRateOp.value = "*" Then												<%'☜: 비지니스 ASP 를 가동 %>
			   frm1.txtNegoLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtNegoDocAmt.text) * UNICDbl(frm1.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
			ElseIf frm1.txtExchRateOp.value = "/" Then
				If frm1.txtXchRate.text <> 0 Then
					frm1.txtNegoLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtNegoDocAmt.text) / UNICDbl(frm1.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
				Else
					frm1.txtNegoLocAmt.text = 0
				End If
			End If
		End If
		lgBlnFlgChgValue = True
	End Sub
'========================================================================================================
	Sub txtNegoDocAmt_Change()
		
		With frm1
			If frm1.txtExchRateOp.value = "*" Then												<%'☜: 비지니스 ASP 를 가동 %>
				frm1.txtNegoLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtNegoDocAmt.text) * UNICDbl(frm1.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
			ElseIf frm1.txtExchRateOp.value = "/" Then
				If frm1.txtXchRate.text <> 0 Then
					frm1.txtNegoLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtNegoDocAmt.text) / UNICDbl(frm1.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
				Else
					frm1.txtNegoLocAmt.text = 0
				End If
			End If								
		End With
		lgBlnFlgChgValue = True
	End Sub
'========================================================================================================
Function JumpChgCheck()

	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then Exit Function
	End If

	Call CookiePage(2)
	Call PgmJump(EXPORT_CHARGE_ENTRY_ID)

End Function

'========================================================================================================
	Function FncQuery()
		Dim IntRetCD

		FncQuery = False													<% '⊙: Processing is NG %>

		Err.Clear															<% '☜: Protect system from crashing %>

		<% '------ Check previous data area ------ %>
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")			<% '⊙: "Will you destory previous data" %>
'			IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "2")								<% '⊙: Clear Contents  Field %>

		<% '------ Check condition area ------ %>
		If Not chkField(Document, "1") Then									<% '⊙: This function check indispensable field %>
			Exit Function
		End If

		<% '------ Query function call area ------ %>
		Call DbQuery()														<% '☜: Query db data %>

		FncQuery = True														<% '⊙: Processing is OK %>
	End Function
	
'========================================================================================================
	Function FncNew()
		Dim IntRetCD 

		FncNew = False                                                          <%'⊙: Processing is NG%>

		<% '------ Check previous data area ------ %>
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "x", "x")
'			IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)

			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		<% '------ Erase condition area ------ %>
		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "A")									<%'⊙: Clear Condition,Contents Field%>
		Call ggoOper.LockField(Document, "N")									<%'⊙: Lock  Suitable  Field%>
		Call SetDefaultVal
		Call SetRadio()
		Call InitVariables														<%'⊙: Initializes local global variables%>
		Call SetToolbar("11100000000011")										<% '⊙: 버튼 툴바 제어 %>
		Call ReleaseBody()
		Call rdoFlawExist2_OnClick

		frm1.txtNEGONo.focus
		Set gActiveElement = document.activeElement 

		FncNew = True															<%'⊙: Processing is OK%>
	End Function
	
'========================================================================================================
	Function FncDelete()
		Dim IntRetCD

		FncDelete = False												<% '⊙: Processing is NG %>
		
		<% '------ Precheck area ------ %>
		If lgIntFlgMode <> parent.OPMD_UMODE Then								<% 'Check if there is retrived data %>
			Call DisplayMsgBox("900002", "x", "x", "x")
'			Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
			Exit Function
		End If

		IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "x", "x")

		If IntRetCD = vbNo Then
			Exit Function
		End If

		<% '------ Delete function call area ------ %>
		Call DbDelete													<% '☜: Delete db data %>

		FncDelete = True												<% '⊙: Processing is OK %>
	End Function

'========================================================================================================
	Function FncSave()
		Dim IntRetCD
		
		FncSave = False													<% '⊙: Processing is NG %>
		
		Err.Clear														<% '☜: Protect system from crashing %>
		
		<% '------ Precheck area ------ %>
		If lgBlnFlgChgValue = False Then								<% 'Check if there is retrived data %>
		    IntRetCD = DisplayMsgBox("900001", "x", "x", "x")					<% '⊙: No data changed!! %>
'		    Call MsgBox("No data changed!!", vbInformation)
		    Exit Function
		End If
		
		<% '------ Check contents area ------ %>
		If Not chkField(Document, "2") Then						<% '⊙: Check contents area %>
		    If gPageNo > 0 Then
		        gSelframeFlg = gPageNo
		    End If
			Exit Function
		End If
		
		If Len(Trim(frm1.txtExpireDt.Text)) And Len(Trim(frm1.txtNegoDt.Text)) Then
			If UniConvDateToYYYYMMDD(frm1.txtNegoDt.Text, parent.gDateFormat, "-") > UniConvDateToYYYYMMDD(frm1.txtExpireDt.Text, parent.gDateFormat, "-") Then
				Call DisplayMsgBox("970023", "x", frm1.txtExpireDt.Alt, frm1.txtNegoDt.Alt)
				'MsgBox "pObjToDt(은)는 pObjFromDt보다 크거나 같아야 합니다.", vbExclamation, "uniERP(Warning)"
				Call ClickTab1()
				frm1.txtNegoDt.Focus
				Set gActiveElement = document.activeElement 
				Exit Function
			End If
		End If
		
		If UNICDbl(frm1.txtNegoDocAmt.text) <= 0 Then
			Call DisplayMsgBox("970022", "x", "NEGO금액","0")
			Call ClickTab1()			
			frm1.txtNegoDocAmt.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
		
		If UNICDbl(frm1.txtXchRate.text) <= 0 Then
			Call DisplayMsgBox("970022", "x", "환율","0")
			Call ClickTab1()			
			frm1.txtXchRate.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
		
		<% '------ Save function call area ------ %>
		Call DbSave														<% '☜: Save db data %>
		
		FncSave = True													<% '⊙: Processing is OK %>
	End Function

'========================================================================================================
	Function FncCopy()
		Dim IntRetCD

		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")			<%'⊙: "Will you destory previous data"%>
'			IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)

			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		lgIntFlgMode = parent.OPMD_CMODE													<%'⊙: Indicates that current mode is Crate mode%>

		<% '------ 조건부 필드를 삭제한다. ------ %>
		Call ggoOper.ClearField(Document, "1")										<%'⊙: Clear Condition Field%>
		Call ggoOper.LockField(Document, "N")										<%'⊙: This function lock the suitable field%>
	End Function

'========================================================================================================
	Function FncPrint()
		Call parent.FncPrint()
	End Function

'========================================================================================================
	Function FncPrev() 
		Dim strVal
    
		If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
		    Call DisplayMsgBox("900002", "x", "x", "x")  '☜ 바뀐부분 
		    'Call MsgBox("조회한후에 지원됩니다.", vbInformation)
		    Exit Function
		End If

		Call LayerShowHide(1)

		frm1.txtPrevNext.value = "PREV"

		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtNEGONo=" & Trim(frm1.txtNEGONo1.value)			<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		<%'☆: 조회 조건 데이타 %>
		     
		Call RunMyBizASP(MyBizASP, strVal)
	End Function

'========================================================================================================
	Function FncNext() 
	    Dim strVal
	    
	    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
	        Call DisplayMsgBox("900002", "x", "x", "x")  '☜ 바뀐부분 
	        'Call MsgBox("조회한후에 지원됩니다.", vbInformation)
	        Exit Function
	    End If

		Call LayerShowHide(1)

		frm1.txtPrevNext.value = "NEXT"

	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							<%'☜: 비지니스 처리 ASP의 상태 %>
	    strVal = strVal & "&txtNEGONo=" & Trim(frm1.txtNEGONo1.value)				<%'☆: 조회 조건 데이타 %>
	    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		<%'☆: 조회 조건 데이타 %>
	         
		Call RunMyBizASP(MyBizASP, strVal)
	End Function

'========================================================================================================
	Function FncExcel() 
		Call parent.FncExport(parent.C_SINGLE)
	End Function

'========================================================================================================
	Function FncFind() 
		Call parent.FncFind(parent.C_SINGLE, True)
	End Function

'========================================================================================================
	Function FncExit()
		Dim IntRetCD

		FncExit = False

		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")			<%'⊙: "Will you destory previous data"%>

'			IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		FncExit = True
	End Function

'========================================================================================================
	Function DbQuery()
		Err.Clear															<%'☜: Protect system from crashing%>

		DbQuery = False														<%'⊙: Processing is NG%>

		Dim strVal

		Call LayerShowHide(1)		
	
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtNEGONo=" & Trim(frm1.txtNEGONo.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&empty=empty"

		Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>
	
		DbQuery = True														<%'⊙: Processing is NG%>
	End Function

'========================================================================================================
	Function DbSave()
		Err.Clear															<%'☜: Protect system from crashing%>

		DbSave = False														<%'⊙: Processing is NG%>
		
		Dim strVal

		Call LayerShowHide(1)

		With frm1
			.txtMode.value = parent.UID_M0002										<%'☜: 비지니스 처리 ASP 의 상태 %>
			.txtFlgMode.value = lgIntFlgMode
			.txtUpdtUserId.value = parent.gUsrID
			.txtInsrtUserId.value = parent.gUsrID

			Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		End With

		DbSave = True														<%'⊙: Processing is NG%>
	End Function
	
'========================================================================================================
	Function DbDelete()
		Err.Clear															<%'☜: Protect system from crashing%>

		DbDelete = False													<%'⊙: Processing is NG%>

		If frm1.rdoPostingflg1.checked = True Then
			Call DisplayMsgBox("207124", "x", "x", "x")
			Exit Function
		End If

		Dim strVal

		Call LayerShowHide(1)

		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003					<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtNEGONo=" & Trim(frm1.txtNEGONo1.value)		<%'☜: 삭제 조건 데이타 %>
		strVal = strVal & "&empty=empty"

		Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>

		DbDelete = True														<%'⊙: Processing is NG%>
	End Function

'========================================================================================================
	Function DbQueryOk()													<% '☆: 조회 성공후 실행로직 %>
		<% '------ Reset variables area ------ %>
		lgIntFlgMode = parent.OPMD_UMODE											<% '⊙: Indicates that current mode is Update mode %>
		lgBlnFlgChgValue = False

		Call ggoOper.LockField(Document, "Q")								<% '⊙: This function lock the suitable field %>
		Call SetToolbar("111110001101111")
		frm1.txtLocCurrency.value = parent.gCurrency		
		
		' 입금유형이 존재하는 경우에만 Enable
'		If frm1.rdoFlawExist1.checked Then
			frm1.btnPosting.disabled = False
'		ELSE
'			frm1.btnPosting.disabled = True
'		END IF
		
		If frm1.rdoPostingflg1.checked = True Then
			Call ProtectBody()
		ElseIf frm1.rdoPostingflg2.checked = True Then
			Call ReleaseBody() 
		End If	

		Call ProtectXchRate()			

		lgBlnFlgChgValue = False
		frm1.txtNEGONo.focus
	
	End Function

'========================================================================================================
	Function BillQueryOk()													<% '☆: 조회 성공후 실행로직 %>
		Call SetToolbar("111010000000111")
		Call txtNegoDocAmt_Change()
		Call ProtectXchRate()
	End Function
	
'========================================================================================================
	Function DbSaveOk()														<%'☆: 저장 성공후 실행 로직 %>
		Call InitVariables
		Call FncQuery()
	End Function
	
'========================================================================================================
	Function DbDeleteOk()													<%'☆: 삭제 성공후 실행 로직 %>
		lgBlnFlgChgValue = False
		Call FncNew()
	End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
		<TABLE <%=LR_SPACE_TYPE_00%>>
			<TR>
				<TD <%=HEIGHT_TYPE_00%>>&nbsp;<% ' 상위 여백 %></TD>
			</TR>
			<TR HEIGHT=23>
				<TD WIDTH=100%>
					<TABLE <%=LR_SPACE_TYPE_10%>>
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
									<TR>
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>NEGO정보</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>NEGO기타</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD WIDTH=* align=right><A href="vbscript:OpenBillRef">매출채권참조</A></TD>
							<TD WIDTH=10>&nbsp;</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR HEIGHT=*>
				<TD WIDTH=100% CLASS="Tab11">
					<!-- 첫번째 탭 내용 
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">   --->
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
						</TR>
						<TR>
							<TD HEIGHT=20 WIDTH=100%>
								<FIELDSET CLASS="CLSFLD">
									<TABLE <%=LR_SPACE_TYPE_40%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>NEGO 관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNEGONo" SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="NEGO관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNEGONo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call btnNEGONoOnClick()"></TD>
											<TD CLASS=TDT NOWRAP></TD>
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
							<TD>
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">	
									<!--<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0> -->
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>NEGO관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtNegoNo1" TYPE=TEXT SIZE=20 MAXLENGTH=18 TAG="25XXXU" ALT="NEGO관리번호"></TD>
											<TD CLASS=TD5 NOWRAP>NEGO유형</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNegoType" SIZE=10 MAXLENGTH=5 TAG="22XXXU" ALT="NEGO유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNegoType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnNegoTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtNegoTypeNm" SIZE=20 TAG="24"></TD>
								    	</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>매입번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtNegoDocNo" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="22XXXU" ALT="매입번호"></TD>
											<TD CLASS=TD5 NOWRAP>매입은행</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNegoBank" SIZE=10 MAXLENGTH=10 TAG="22XXXU" ALT="매입은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNegoBank" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnNegoBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtNegoBankNm" SIZE=20 TAG="24"></TD>
										</TR>	
										<TR>	
											<TD CLASS=TD5 NOWRAP>NEGO일</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtNegoDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="NEGO일"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>확정여부</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" TAG="24X" VALUE="Y" ID="rdoPostingflg1"><LABEL FOR="rdoPostingflg1">확정</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" VALUE="N" TAG="24X" CHECKED ID="rdoPostingflg2"><LABEL FOR="rdoPostingflg2">미확정</LABEL></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>NEGO금액</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtNegoDocAmt" CLASS=FPDS140 tag="22X2Z" ALT="NEGO금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>													
														<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXZU" ALT="화폐"></TD>
													</TR>
												</TABLE>
											</TD>				
											<TD CLASS=TD5 NOWRAP>문자금액</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNegoAmtTxt" MAXLENGTH=50 SIZE=35 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>환율</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchRate" CLASS=FPDS140 tag="22X5Z" ALT="환율" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS=TD5 NOWRAP>NEGO자국금액</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtNegoLocAmt" CLASS=FPDS140 tag="22X2Z" ALT="NEGO자국금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>													
														<TD>
															&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU">
														</TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>영업그룹</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="22XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnSalesGroupOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14"></TD>
											<TD CLASS=TD5 NOWRAP>지급기일</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtPayExpiryDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="지급기일"></OBJECT>');</SCRIPT></TD>
										</TR>	
										<TR>
											<TD CLASS=TD5 NOWRAP>매입의뢰일</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtNegoReqDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="매입의뢰일"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>입금유무</TD>
											<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFlawExist" TAG="21X" VALUE="Y" ID="rdoFlawExist1">
												<LABEL FOR="rdoFlawExist1">Y</LABEL>
												&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFlawExist" VALUE="N" TAG="21X" CHECKED ID="rdoFlawExist2">
												<LABEL FOR="rdoFlawExist2">N</LABEL>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>입금일</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtPayDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="입금일"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>입금유형</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCollectType" SIZE=10 MAXLENGTH=4 TAG="23XXXU" ALT="입금유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCollectType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnCollectTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtCollectTypeNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>					
											<TD CLASS=TD5 NOWRAP>입금은행</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncomeBank" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="입금은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIncomeBank" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnIncomeBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtIncomeBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>입금계좌번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAccountNo" SIZE=32 MAXLENGTH=30 TAG="21XXXU" ALT="입금계좌"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAccountNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnAccountNoOnClick()"></TD>
										</TR>
										<TR>	
											<TD CLASS=TD5 NOWRAP>발행지역</TD>
											<TD CLASS=TD6 COLSPAN=3><INPUT TYPE=TEXT NAME="txtNegoPubZone" ALT="발행지역" MAXLENGTH=50 SIZE=86 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>비고</TD>
											<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemarks1" ALT="비고" TYPE=TEXT MAXLENGTH=120 SIZE=86 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemarks2" ALT="비고" TYPE=TEXT MAXLENGTH=120 SIZE=86 TAG="21X"></TD>
										</TR>			
										<%Call SubFillRemBodyTD5656(6)%>
									</TABLE>
								</DIV>
								<!-- 두번째 탭 내용 -->
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>연 환가요율</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchCommRate" CLASS=FPDS140 tag="21X5Z" ALT="연 환가요율" Title="FPDOUBLESINGLE">&nbsp;(%)</OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS=TD5 NOWRAP>통지번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAdvNo" ALT="통지번호" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="21XXXU"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>매출관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillNo" TYPE=TEXT SIZE=20 MAXLENGTH=18 TAG="24XXXU"></TD>
											<TD CLASS=TD5 NOWRAP>B/L번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBLNo" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="24XXXU"></TD>
										</TR>
										<TR>	
											<TD CLASS=TD5 NOWRAP>L/C관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCNo" TYPE=TEXT SIZE=20 MAXLENGTH=18 TAG="24XXXU"></TD>
											<TD CLASS=TD5 NOWRAP>L/C번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>근거금액</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtBaseDocAmt" CLASS=FPDS140 tag="24X2Z" ALT="근거금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>													
														<TD>
															&nbsp;<INPUT TYPE=TEXT NAME="txtBaseCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXZU">
														</TD>
													</TR>
												</TABLE>
											</TD>
											<TD CLASS=TD5 NOWRAP>최종선적일</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtLatestShipDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="최종선적일"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>				
											<TD CLASS=TD5 NOWRAP>L/C개설일</TD>						
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtOpenDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="L/C개설일"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>유효일</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtExpireDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="유효일"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>개설은행</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" TAG="24XXXU" ALT="개설은행">&nbsp;<INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>가격조건</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoterms" SIZE=20 MAXLENGTH=5 MAXLENGTH=3 TAG="24XXXU" ALT="가격조건"></TD>
										</TR>
										<TR>	
											<TD CLASS=TD5 NOWRAP>결제방법</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="결제방법">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" MAXLENGTH=50 SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>결제기간</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayDur" ALT="결제기간" STYLE="TEXT-ALIGN: right" TYPE=TEXT MAXLENGTH=3 SIZE=5 TAG="24X7" ALT="결제기간">&nbsp;일</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>수입자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수입자">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>수출자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수출자">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>대행자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="대행자">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>제조자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="제조자">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
										</TR>
									<%Call SubFillRemBodyTD5656(6)%>
									</TABLE>
								</DIV>
							</TD>
						</TR>
					</TABLE>				
				</TD>
			</TR>
			<TR HEIGHT=20>
				<TD WIDTH=100%>
					<TABLE <%=LR_SPACE_TYPE_30%>>
						<TD ><BUTTON NAME="btnPosting" CLASS="CLSMBTN">확정</BUTTON></TD>
						<TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck()">판매경비등록</A></TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX = -1></IFRAME></TD>
			</TR>
		</TABLE>
		<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtBankCd" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtHBillNo" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtHBLDocNo" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtHLCNo" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtHBLFlag" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtHNEGONo" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24" TABINDEX = -1> 
		<INPUT TYPE=HIDDEN NAME="txtExchRateOp" tag="24" TABINDEX = -1>
	</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
