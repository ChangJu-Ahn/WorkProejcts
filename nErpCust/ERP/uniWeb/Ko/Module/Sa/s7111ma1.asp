<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : S7111MA1    																*
'*  4. Program Name         : NEGO ���																	*
'*  5. Program Desc         : NEGO ���    																*
'*  6. Comproxy List        : PSAG111.dll, PSAG119.dll               									*
'*  7. Modified date(First) : 2000/05/08																*
'*  8. Modified date(Last)  : 2000/05/08																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : An Chang Hwan																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/05/08 : ȭ�� design												*
'*							  2. 2000/05/08 : Coding Start												*
'******************************************************************************************************** 
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBS">
Option Explicit					<% '��: indicates that All variables must be declared in advance %>

'============================================  1.2.1 Global ��� ����  ==================================
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

	Const BIZ_PGM_ID = "s7111mb1.asp"						'��: �����Ͻ� ���� ASP�� 
	Const BIZ_PGM_BASDATAQUERY_ID = "s7111mb2.asp"			'��: �����Ͻ� ���� ASP�� : ����ä������ 
	Const BIZ_PGM_POSTING_ID = "s7111mb3.asp"				'��: �����Ͻ� ���� ASP�� : Ȯ�� 
	Const EXPORT_CHARGE_ENTRY_ID = "s6111ma1"				'��: �̵��� ASP�� 
	Const TAB1 = 1
	Const TAB2 = 2

	'------ Minor Code PopUp�� ���� Major Code���� ------ 
	Const gstrSubTypeMajor = "S9072"
	Const gstrCollectTypeMajor = "B9004"
	Const gstrNegoTypeMajor = "S9071"

'============================================  1.2.2 Global ���� ����  ==================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
	
	Dim gSelframeFlg								'���� TAB�� ��ġ�� ��Ÿ���� Flag 
	Dim gblnWinEvent								'~~~ ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
													'	 PopUp Window�� ��������� ���θ� ��Ÿ���� variable 
	Dim lgBlnFlawChgFlg	
	Dim gtxtChargeType
	Dim glsTab
	Dim gTabMaxCnt

'========================================================================================================
	Function InitVariables()
		lgIntFlgMode = parent.OPMD_CMODE						<%'��: Indicates that current mode is Create mode%>
		lgBlnFlgChgValue = False								<%'��: Indicates that no value changed%>
		lgBlnFlawChgFlg = False
		lgIntGrpCount = 0										<%'��: Initializes Group View Size%>
		
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

'===NEGO������ȣ=============================================================================================
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

'===����ä������=============================================================================================
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

		arrParam(0) = strPopPos								<%' �˾� ��Ī %>
		arrParam(1) = "B_Minor"								<%' TABLE ��Ī %>
		arrParam(2) = Trim(strMinorCD)						<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""		<%' Where Condition%>
		arrParam(5) = strPopPos								<%' TextBox ��Ī %>

		arrField(0) = "Minor_CD"							<%' Field��(0)%>
		arrField(1) = "Minor_NM"							<%' Field��(1)%>

		arrHeader(0) = strPopPos							<%' Header��(0)%>
		arrHeader(1) = strPopPos  & "��"				<%' Header��(1)%>
		
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

		arrParam(0) = "�Ա�����"						<%' �˾� ��Ī %>
		arrParam(1) = "B_Configuration A, (Select * From B_Minor Where Major_cd = " & FilterVar("A1006", "''", "S") & ") B"								<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtCollectType.value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "a.major_cd = " & FilterVar("B9004", "''", "S") & " and a.reference = b.minor_cd and a.minor_cd =  " & FilterVar(frm1.txtPayTerms.value , "''", "S") & ""					<%' Where Condition%>
		arrParam(5) = "�Ա�����"						<%' TextBox ��Ī %>

		arrField(0) = "A.Reference"							<%' Field��(0)%>
		arrField(1) = "B.minor_nm"							<%' Field��(1)%>

		arrHeader(0) = "�Ա�����"						<%' Header��(0)%>
		arrHeader(1) = "�Ա�������"						<%' Header��(1)%>

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

		arrParam(0) = "�����׷�"						<%' �˾� ��Ī %>
		arrParam(1) = "B_SALES_GRP"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtSalesGroup.value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						<%' Where Condition%>
		arrParam(5) = "�����׷�"						<%' TextBox ��Ī %>

		arrField(0) = "SALES_GRP"							<%' Field��(0)%>
		arrField(1) = "SALES_GRP_NM"						<%' Field��(1)%>

		arrHeader(0) = "�����׷�"						<%' Header��(0)%>
		arrHeader(1) = "�����׷��"						<%' Header��(1)%>

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

		arrParam(0) = "����"							<%' �˾� ��Ī %>
		arrParam(1) = "B_BANK"								<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtIncomeBank.value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "����"							<%' TextBox ��Ī %>

		arrField(0) = "BANK_CD"								<%' Field��(0)%>
		arrField(1) = "BANK_NM"								<%' Field��(1)%>

		arrHeader(0) = "����"							<%' Header��(0)%>
		arrHeader(1) = "�����"							<%' Header��(1)%>

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

		arrParam(0) = "����"							<%' �˾� ��Ī %>
		arrParam(1) = "B_BANK"								<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtIncomeBank.value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "����"							<%' TextBox ��Ī %>

		arrField(0) = "BANK_CD"								<%' Field��(0)%>
		arrField(1) = "BANK_NM"								<%' Field��(1)%>

		arrHeader(0) = "����"							<%' Header��(0)%>
		arrHeader(1) = "�����"							<%' Header��(1)%>

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

		arrParam(0) = "�Աݰ���"																<%' �˾� ��Ī %>
		arrParam(1) = "B_BANK_ACCT, B_BANK"															<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtAccountNo.value)													<%' Code Condition%>
		arrParam(3) = ""																			<%' Name Cindition%>
		arrParam(4) = "B_BANK_ACCT.BANK_CD=B_BANK.BANK_CD And B_BANK.BANK_CD =  " & FilterVar(IncomeBank , "''", "S") & ""		<%' Where Condition%>
		arrParam(5) = "�Աݰ���"																<%' TextBox ��Ī %>

		arrField(0) = "B_BANK_ACCT.BANK_ACCT_NO"													<%' Field��(0)%>
		arrField(1) = "ED10" & parent.gColSep & "B_BANK_ACCT.BANK_CD"															<%' Field��(1)%>
		arrField(2) = "ED20" & parent.gColSep & "B_BANK.BANK_NM"																<%' Field��(2)%>

		arrHeader(0) = "�Աݰ���"																<%' Header��(0)%>
		arrHeader(1) = "�Ա�����"																<%' Header��(1)%>
		arrHeader(2) = "�Ա������"																<%' Header��(2)%>

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
		Call ggoOper.ClearField(Document, "2")											<% '��: Clear Contents  Field %>
		Call SetRadio()
		Call SetDefaultVal
		
		frm1.txtHBillNo.value = strRet(0)
		frm1.txtHBLDocNo.value = strRet(1)
		frm1.txtHLCNo.value = strRet(2)
		frm1.txtHBLFlag.value = strRet(3)
		
		Dim strVal
		Call LayerShowHide(1)		
		
		strVal = BIZ_PGM_BASDATAQUERY_ID & "?txtBillNo=" & Trim(frm1.txtHBillNo.value)		<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtBLNo=" & Trim(frm1.txtHBLDocNo.value)
		strVal = strVal & "&txtLCNo=" & Trim(frm1.txtHLCNo.value)						<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtBLFlag=" & Trim(frm1.txtHBLFlag.value)

		Call RunMyBizASP(MyBizASP, strVal)												<%'��: �����Ͻ� ASP �� ���� %>
	
		
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
			Call DisplayMsgBox("900002", "x", "x", "x")	<% '��: "Will you destory previous data" %>
			'Call MsgBox("��ȸ�� �����Ͻʽÿ�.", parent.VB_INFORMATION)
			Exit Function
		End If

		Dim strVal

		Call LayerShowHide(1)

		strVal = BIZ_PGM_POSTING_ID & "?txtNEGONo=" & Trim(frm1.txtNEGONo.value)		<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtgChangeOrgId=" & parent.gChangeOrgId
		strVal = strVal & "&txtUpdtUserId=" & parent.gUsrID									<%'��: ��ȸ ���� ����Ÿ %>

		Call RunMyBizASP(MyBizASP, strVal)												<%'��: �����Ͻ� ASP �� ���� %>
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
'=	Event Desc : ���������� L/C ������ �Ǿ����� �ش� textbox �� Protect ��Ų��.							=
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
'=	Event Desc : ��������� ���������� L/C ���������� Protect �Ǿ��ִ� textbox �� Release �Ѵ�.			=
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
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
	Sub CurFormatNumericOCX()
		With frm1

			'NegoL �ݾ� 
			ggoOper.FormatFieldByObjectOfCur .txtNegoDocAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
			'�ٰűݾ� 
			ggoOper.FormatFieldByObjectOfCur .txtBaseDocAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
			 'ȯ�� 
			ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec

		End With
	End Sub

'========================================================================================================
	Sub Form_Load()
		Call LoadInfTB19029
		Call AppendNumberPlace("6","3","2")												<% '��: Load table , B_numeric_format %>
		Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
		Call ggoOper.LockField(Document, "N")											<% '��: Lock  Suitable  Field %>
		Call SetDefaultVal
		Call SetToolbar("11100000000011")												<% '��: ��ư ���� ���� %>
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
			Call OpenMinorCd(frm1.txtSubType.value, "", "ȯ������", gstrSubTypeMajor)
		End If
	End Sub
'========================================================================================================
	Sub btnNegoTypeOnClick()
		If frm1.txtNegoType.readOnly <> True Then
			frm1.txtNegoType.focus 
			Call OpenMinorCd(frm1.txtNegoType.value, frm1.txtNegoTypeNm.value , "NEGO����", gstrNegoTypeMajor)
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
		If frm1.rdoPostingflg2.checked = True Then    'Ȯ���� ���� ����ʵ带 Protect ��Ű�� �ϼ� OnClick Event �� �߻��ϱ⶧���� Check �� ���� 
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
			If frm1.txtExchRateOp.value = "*" Then												<%'��: �����Ͻ� ASP �� ���� %>
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
			If frm1.txtExchRateOp.value = "*" Then												<%'��: �����Ͻ� ASP �� ���� %>
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
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then Exit Function
	End If

	Call CookiePage(2)
	Call PgmJump(EXPORT_CHARGE_ENTRY_ID)

End Function

'========================================================================================================
	Function FncQuery()
		Dim IntRetCD

		FncQuery = False													<% '��: Processing is NG %>

		Err.Clear															<% '��: Protect system from crashing %>

		<% '------ Check previous data area ------ %>
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")			<% '��: "Will you destory previous data" %>
'			IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?", vbYesNo)
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "2")								<% '��: Clear Contents  Field %>

		<% '------ Check condition area ------ %>
		If Not chkField(Document, "1") Then									<% '��: This function check indispensable field %>
			Exit Function
		End If

		<% '------ Query function call area ------ %>
		Call DbQuery()														<% '��: Query db data %>

		FncQuery = True														<% '��: Processing is OK %>
	End Function
	
'========================================================================================================
	Function FncNew()
		Dim IntRetCD 

		FncNew = False                                                          <%'��: Processing is NG%>

		<% '------ Check previous data area ------ %>
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "x", "x")
'			IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. �ű��۾��� �Ͻðڽ��ϱ�?", vbYesNo)

			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		<% '------ Erase condition area ------ %>
		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "A")									<%'��: Clear Condition,Contents Field%>
		Call ggoOper.LockField(Document, "N")									<%'��: Lock  Suitable  Field%>
		Call SetDefaultVal
		Call SetRadio()
		Call InitVariables														<%'��: Initializes local global variables%>
		Call SetToolbar("11100000000011")										<% '��: ��ư ���� ���� %>
		Call ReleaseBody()
		Call rdoFlawExist2_OnClick

		frm1.txtNEGONo.focus
		Set gActiveElement = document.activeElement 

		FncNew = True															<%'��: Processing is OK%>
	End Function
	
'========================================================================================================
	Function FncDelete()
		Dim IntRetCD

		FncDelete = False												<% '��: Processing is NG %>
		
		<% '------ Precheck area ------ %>
		If lgIntFlgMode <> parent.OPMD_UMODE Then								<% 'Check if there is retrived data %>
			Call DisplayMsgBox("900002", "x", "x", "x")
'			Call MsgBox("��ȸ���Ŀ� ������ �� �ֽ��ϴ�.", vbInformation)
			Exit Function
		End If

		IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "x", "x")

		If IntRetCD = vbNo Then
			Exit Function
		End If

		<% '------ Delete function call area ------ %>
		Call DbDelete													<% '��: Delete db data %>

		FncDelete = True												<% '��: Processing is OK %>
	End Function

'========================================================================================================
	Function FncSave()
		Dim IntRetCD
		
		FncSave = False													<% '��: Processing is NG %>
		
		Err.Clear														<% '��: Protect system from crashing %>
		
		<% '------ Precheck area ------ %>
		If lgBlnFlgChgValue = False Then								<% 'Check if there is retrived data %>
		    IntRetCD = DisplayMsgBox("900001", "x", "x", "x")					<% '��: No data changed!! %>
'		    Call MsgBox("No data changed!!", vbInformation)
		    Exit Function
		End If
		
		<% '------ Check contents area ------ %>
		If Not chkField(Document, "2") Then						<% '��: Check contents area %>
		    If gPageNo > 0 Then
		        gSelframeFlg = gPageNo
		    End If
			Exit Function
		End If
		
		If Len(Trim(frm1.txtExpireDt.Text)) And Len(Trim(frm1.txtNegoDt.Text)) Then
			If UniConvDateToYYYYMMDD(frm1.txtNegoDt.Text, parent.gDateFormat, "-") > UniConvDateToYYYYMMDD(frm1.txtExpireDt.Text, parent.gDateFormat, "-") Then
				Call DisplayMsgBox("970023", "x", frm1.txtExpireDt.Alt, frm1.txtNegoDt.Alt)
				'MsgBox "pObjToDt(��)�� pObjFromDt���� ũ�ų� ���ƾ� �մϴ�.", vbExclamation, "uniERP(Warning)"
				Call ClickTab1()
				frm1.txtNegoDt.Focus
				Set gActiveElement = document.activeElement 
				Exit Function
			End If
		End If
		
		If UNICDbl(frm1.txtNegoDocAmt.text) <= 0 Then
			Call DisplayMsgBox("970022", "x", "NEGO�ݾ�","0")
			Call ClickTab1()			
			frm1.txtNegoDocAmt.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
		
		If UNICDbl(frm1.txtXchRate.text) <= 0 Then
			Call DisplayMsgBox("970022", "x", "ȯ��","0")
			Call ClickTab1()			
			frm1.txtXchRate.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
		
		<% '------ Save function call area ------ %>
		Call DbSave														<% '��: Save db data %>
		
		FncSave = True													<% '��: Processing is OK %>
	End Function

'========================================================================================================
	Function FncCopy()
		Dim IntRetCD

		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")			<%'��: "Will you destory previous data"%>
'			IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?", vbYesNo)

			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		lgIntFlgMode = parent.OPMD_CMODE													<%'��: Indicates that current mode is Crate mode%>

		<% '------ ���Ǻ� �ʵ带 �����Ѵ�. ------ %>
		Call ggoOper.ClearField(Document, "1")										<%'��: Clear Condition Field%>
		Call ggoOper.LockField(Document, "N")										<%'��: This function lock the suitable field%>
	End Function

'========================================================================================================
	Function FncPrint()
		Call parent.FncPrint()
	End Function

'========================================================================================================
	Function FncPrev() 
		Dim strVal
    
		If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
		    Call DisplayMsgBox("900002", "x", "x", "x")  '�� �ٲ�κ� 
		    'Call MsgBox("��ȸ���Ŀ� �����˴ϴ�.", vbInformation)
		    Exit Function
		End If

		Call LayerShowHide(1)

		frm1.txtPrevNext.value = "PREV"

		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtNEGONo=" & Trim(frm1.txtNEGONo1.value)			<%'��: ��ȸ ���� ����Ÿ %>
		strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		<%'��: ��ȸ ���� ����Ÿ %>
		     
		Call RunMyBizASP(MyBizASP, strVal)
	End Function

'========================================================================================================
	Function FncNext() 
	    Dim strVal
	    
	    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
	        Call DisplayMsgBox("900002", "x", "x", "x")  '�� �ٲ�κ� 
	        'Call MsgBox("��ȸ���Ŀ� �����˴ϴ�.", vbInformation)
	        Exit Function
	    End If

		Call LayerShowHide(1)

		frm1.txtPrevNext.value = "NEXT"

	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							<%'��: �����Ͻ� ó�� ASP�� ���� %>
	    strVal = strVal & "&txtNEGONo=" & Trim(frm1.txtNEGONo1.value)				<%'��: ��ȸ ���� ����Ÿ %>
	    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		<%'��: ��ȸ ���� ����Ÿ %>
	         
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
			IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")			<%'��: "Will you destory previous data"%>

'			IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		FncExit = True
	End Function

'========================================================================================================
	Function DbQuery()
		Err.Clear															<%'��: Protect system from crashing%>

		DbQuery = False														<%'��: Processing is NG%>

		Dim strVal

		Call LayerShowHide(1)		
	
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtNEGONo=" & Trim(frm1.txtNEGONo.value)		<%'��: ��ȸ ���� ����Ÿ %>
		strVal = strVal & "&empty=empty"

		Call RunMyBizASP(MyBizASP, strVal)									<%'��: �����Ͻ� ASP �� ���� %>
	
		DbQuery = True														<%'��: Processing is NG%>
	End Function

'========================================================================================================
	Function DbSave()
		Err.Clear															<%'��: Protect system from crashing%>

		DbSave = False														<%'��: Processing is NG%>
		
		Dim strVal

		Call LayerShowHide(1)

		With frm1
			.txtMode.value = parent.UID_M0002										<%'��: �����Ͻ� ó�� ASP �� ���� %>
			.txtFlgMode.value = lgIntFlgMode
			.txtUpdtUserId.value = parent.gUsrID
			.txtInsrtUserId.value = parent.gUsrID

			Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		End With

		DbSave = True														<%'��: Processing is NG%>
	End Function
	
'========================================================================================================
	Function DbDelete()
		Err.Clear															<%'��: Protect system from crashing%>

		DbDelete = False													<%'��: Processing is NG%>

		If frm1.rdoPostingflg1.checked = True Then
			Call DisplayMsgBox("207124", "x", "x", "x")
			Exit Function
		End If

		Dim strVal

		Call LayerShowHide(1)

		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003					<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtNEGONo=" & Trim(frm1.txtNEGONo1.value)		<%'��: ���� ���� ����Ÿ %>
		strVal = strVal & "&empty=empty"

		Call RunMyBizASP(MyBizASP, strVal)									<%'��: �����Ͻ� ASP �� ���� %>

		DbDelete = True														<%'��: Processing is NG%>
	End Function

'========================================================================================================
	Function DbQueryOk()													<% '��: ��ȸ ������ ������� %>
		<% '------ Reset variables area ------ %>
		lgIntFlgMode = parent.OPMD_UMODE											<% '��: Indicates that current mode is Update mode %>
		lgBlnFlgChgValue = False

		Call ggoOper.LockField(Document, "Q")								<% '��: This function lock the suitable field %>
		Call SetToolbar("111110001101111")
		frm1.txtLocCurrency.value = parent.gCurrency		
		
		' �Ա������� �����ϴ� ��쿡�� Enable
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
	Function BillQueryOk()													<% '��: ��ȸ ������ ������� %>
		Call SetToolbar("111010000000111")
		Call txtNegoDocAmt_Change()
		Call ProtectXchRate()
	End Function
	
'========================================================================================================
	Function DbSaveOk()														<%'��: ���� ������ ���� ���� %>
		Call InitVariables
		Call FncQuery()
	End Function
	
'========================================================================================================
	Function DbDeleteOk()													<%'��: ���� ������ ���� ���� %>
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
				<TD <%=HEIGHT_TYPE_00%>>&nbsp;<% ' ���� ���� %></TD>
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
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>NEGO����</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>NEGO��Ÿ</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD WIDTH=* align=right><A href="vbscript:OpenBillRef">����ä������</A></TD>
							<TD WIDTH=10>&nbsp;</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR HEIGHT=*>
				<TD WIDTH=100% CLASS="Tab11">
					<!-- ù��° �� ���� 
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
											<TD CLASS=TD5 NOWRAP>NEGO ������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNEGONo" SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="NEGO������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNEGONo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call btnNEGONoOnClick()"></TD>
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
											<TD CLASS=TD5 NOWRAP>NEGO������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtNegoNo1" TYPE=TEXT SIZE=20 MAXLENGTH=18 TAG="25XXXU" ALT="NEGO������ȣ"></TD>
											<TD CLASS=TD5 NOWRAP>NEGO����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNegoType" SIZE=10 MAXLENGTH=5 TAG="22XXXU" ALT="NEGO����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNegoType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnNegoTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtNegoTypeNm" SIZE=20 TAG="24"></TD>
								    	</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���Թ�ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtNegoDocNo" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="22XXXU" ALT="���Թ�ȣ"></TD>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNegoBank" SIZE=10 MAXLENGTH=10 TAG="22XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNegoBank" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnNegoBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtNegoBankNm" SIZE=20 TAG="24"></TD>
										</TR>	
										<TR>	
											<TD CLASS=TD5 NOWRAP>NEGO��</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtNegoDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="NEGO��"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>Ȯ������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" TAG="24X" VALUE="Y" ID="rdoPostingflg1"><LABEL FOR="rdoPostingflg1">Ȯ��</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" VALUE="N" TAG="24X" CHECKED ID="rdoPostingflg2"><LABEL FOR="rdoPostingflg2">��Ȯ��</LABEL></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>NEGO�ݾ�</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtNegoDocAmt" CLASS=FPDS140 tag="22X2Z" ALT="NEGO�ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>													
														<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXZU" ALT="ȭ��"></TD>
													</TR>
												</TABLE>
											</TD>				
											<TD CLASS=TD5 NOWRAP>���ڱݾ�</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNegoAmtTxt" MAXLENGTH=50 SIZE=35 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>ȯ��</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchRate" CLASS=FPDS140 tag="22X5Z" ALT="ȯ��" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS=TD5 NOWRAP>NEGO�ڱ��ݾ�</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtNegoLocAmt" CLASS=FPDS140 tag="22X2Z" ALT="NEGO�ڱ��ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>													
														<TD>
															&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU">
														</TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�����׷�</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="22XXXU" ALT="�����׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnSalesGroupOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14"></TD>
											<TD CLASS=TD5 NOWRAP>���ޱ���</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtPayExpiryDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="���ޱ���"></OBJECT>');</SCRIPT></TD>
										</TR>	
										<TR>
											<TD CLASS=TD5 NOWRAP>�����Ƿ���</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtNegoReqDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="�����Ƿ���"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>�Ա�����</TD>
											<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFlawExist" TAG="21X" VALUE="Y" ID="rdoFlawExist1">
												<LABEL FOR="rdoFlawExist1">Y</LABEL>
												&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFlawExist" VALUE="N" TAG="21X" CHECKED ID="rdoFlawExist2">
												<LABEL FOR="rdoFlawExist2">N</LABEL>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�Ա���</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtPayDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="�Ա���"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>�Ա�����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCollectType" SIZE=10 MAXLENGTH=4 TAG="23XXXU" ALT="�Ա�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCollectType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnCollectTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtCollectTypeNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>					
											<TD CLASS=TD5 NOWRAP>�Ա�����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncomeBank" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="�Ա�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIncomeBank" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnIncomeBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtIncomeBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>�Աݰ��¹�ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAccountNo" SIZE=32 MAXLENGTH=30 TAG="21XXXU" ALT="�Աݰ���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAccountNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call btnAccountNoOnClick()"></TD>
										</TR>
										<TR>	
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 COLSPAN=3><INPUT TYPE=TEXT NAME="txtNegoPubZone" ALT="��������" MAXLENGTH=50 SIZE=86 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���</TD>
											<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemarks1" ALT="���" TYPE=TEXT MAXLENGTH=120 SIZE=86 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemarks2" ALT="���" TYPE=TEXT MAXLENGTH=120 SIZE=86 TAG="21X"></TD>
										</TR>			
										<%Call SubFillRemBodyTD5656(6)%>
									</TABLE>
								</DIV>
								<!-- �ι�° �� ���� -->
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>�� ȯ������</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchCommRate" CLASS=FPDS140 tag="21X5Z" ALT="�� ȯ������" Title="FPDOUBLESINGLE">&nbsp;(%)</OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS=TD5 NOWRAP>������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAdvNo" ALT="������ȣ" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="21XXXU"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillNo" TYPE=TEXT SIZE=20 MAXLENGTH=18 TAG="24XXXU"></TD>
											<TD CLASS=TD5 NOWRAP>B/L��ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBLNo" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="24XXXU"></TD>
										</TR>
										<TR>	
											<TD CLASS=TD5 NOWRAP>L/C������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCNo" TYPE=TEXT SIZE=20 MAXLENGTH=18 TAG="24XXXU"></TD>
											<TD CLASS=TD5 NOWRAP>L/C��ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�ٰűݾ�</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtBaseDocAmt" CLASS=FPDS140 tag="24X2Z" ALT="�ٰűݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>													
														<TD>
															&nbsp;<INPUT TYPE=TEXT NAME="txtBaseCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXZU">
														</TD>
													</TR>
												</TABLE>
											</TD>
											<TD CLASS=TD5 NOWRAP>����������</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtLatestShipDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="����������"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>				
											<TD CLASS=TD5 NOWRAP>L/C������</TD>						
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtOpenDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="L/C������"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>��ȿ��</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtExpireDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="��ȿ��"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" TAG="24XXXU" ALT="��������">&nbsp;<INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoterms" SIZE=20 MAXLENGTH=5 MAXLENGTH=3 TAG="24XXXU" ALT="��������"></TD>
										</TR>
										<TR>	
											<TD CLASS=TD5 NOWRAP>�������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="�������">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" MAXLENGTH=50 SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>�����Ⱓ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayDur" ALT="�����Ⱓ" STYLE="TEXT-ALIGN: right" TYPE=TEXT MAXLENGTH=3 SIZE=5 TAG="24X7" ALT="�����Ⱓ">&nbsp;��</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
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
						<TD ><BUTTON NAME="btnPosting" CLASS="CLSMBTN">Ȯ��</BUTTON></TD>
						<TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck()">�ǸŰ����</A></TD>
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
