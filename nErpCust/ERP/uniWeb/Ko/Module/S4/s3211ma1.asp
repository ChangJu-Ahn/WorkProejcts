<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : ����																		*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211ma1.asp																*
'*  4. Program Name         : L/C ���																	*
'*  5. Program Desc         : L/C ���																	*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/07/12																*
'*  8. Modified date(Last)  : 2001/08/29																*
'*  9. Modifier (First)     : Kim Hyungsuk 																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : ȭ�� design												*
'*							  2. 2000/07/12 : Coding ReStart											*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                        

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim prDBSYSDate

Dim EndDate ,StartDate

prDBSYSDate = "<%=GetSvrDate%>"

EndDate = parent.UNIConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company
StartDate = parent.UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID				= "s3211mb1.asp"		 
Const BIZ_PGM_SOQRY_ID			= "s3211mb3.asp"	
Const BIZ_PGM_CAL_AMT_ID		= "s3211mb5.asp"
Const LC_DETAIL_ENTRY_ID		= "s3212ma1"			'��: �̵��� ASP�� 
Const LCAMEND_HDR_ENTRY_ID		= "s3221ma1"			'��: �̵��� ASP�� 
Const EXPORT_CHARGE_ENTRY_ID	= "s6111ma1"			'��: �̵��� ASP�� 
    
Const TAB1 = 1
Const TAB2 = 2
Const TAB3 = 3
Const TAB4 = 4
	
'------ Minor Code PopUp�� ���� Major Code���� ------ 
Const gstrLCTypeMajor		= "S9000"				'L/C ���� 
Const gstrTransportMajor	= "B9009"				'��۹�� 
Const gstrFreightMajor		= "S9007"				'�������ҹ�� 	
Const gstrCreditCoreMajor	= "S9003"				'�ſ������ü 
Const gstrLoadPortMajor		= "B9092"				'������ 
Const gstrDischgePortMajor	= "B9092"				'������ 
Const gstrOriginMajor		= "B9094"				'������ 
	
Dim gSelframeFlg					'���� TAB�� ��ġ�� ��Ÿ���� Flag 
Dim gblnWinEvent					

'========================================================================================================
Function InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE								'��: Indicates that current mode is Create mode
	lgBlnFlgChgValue = False								'��: Indicates that no value changed
	lgIntGrpCount = 0										'��: Initializes Group View Size
		
	 '------ Coding part ------ 
	gblnWinEvent = False
End Function
	
'========================================================================================================
Sub SetDefaultVal()
	With frm1
		.txtLocCurrency.value	= parent.gCurrency
		lgBlnFlgChgValue		= False
	End With
End Sub	

'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %> 
<% Call LoadBNumericFormatA("I","*","NOCOOKIE","MA") %>
	
End Sub
		

'========================================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
		
	Call changeTabs(TAB1)
		
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
		
	Call changeTabs(TAB2)
		
	gSelframeFlg = TAB2
End Function
	
Function ClickTab3()
	If gSelframeFlg = TAB3 Then Exit Function
		
	Call changeTabs(TAB3)
		
	gSelframeFlg = TAB3
End Function
	
Function ClickTab4()
	If gSelframeFlg = TAB4 Then Exit Function
		
	Call changeTabs(TAB4)
		
	gSelframeFlg = TAB4
End Function

'========================================================================================================
Function OpenLCNoPop()
	
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
	
	If gblnWinEvent = True Or UCase(frm1.txtLCNo.className) = "PROTECTED" Then 
		Exit Function
	End If
			
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("s3211pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3211pa1", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetLCNo(strRet)
	End If	

End Function

'========================================================================================================
Function OpenSORef()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If lgIntFlgMode = parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "x", "x", "x")
		Exit function
	End If 
				
	If gblnWinEvent = True Then Exit Function
		
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("s3111ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3111ra1", "X")
		gblnWinEvent = False
		Exit Function
	End If		
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
				
	If strRet = "" Then
		Exit Function
	Else
			
		Call SetSORef(strRet)
	End If	
End Function	

'========================================================================================================
Function OpenBankPop(strBankCd, strBankNm, Byval iwhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

		arrParam(0) = "����"							' �˾�  ��Ī 
		arrParam(1) = "B_Bank"								' TABLE ��Ī 
		arrParam(2) = Trim(strBankCd)						'	Code Condition		
		arrParam(3) = ""									' Name Condotion		
		arrParam(4) = ""									' Where Condition
		arrParam(5) = "����"							' TextBox ��Ī 
		
		arrField(0) = "Bank_cd"								' Field��(0)
		arrField(1) = "BANK_NM"							' Field��(1)
	    
		arrHeader(0) = "����"							' Header��(0)
		arrHeader(1) = "�����"							' Header��(1)

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call Setbank(iwhere, arrRet)
	End If	
		
End Function


'========================================================================================================
Function OpenMinorCd(strMinorCD, strMinorNM, strPopNm, strMajorCd)
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
		
		arrParam(0) = strPopNm								' �˾�  ��Ī 
		arrParam(1) = "B_MINOR"								' TABLE ��Ī 
		arrParam(2) = Trim(strMinorCD)						'	Code Condition	
		arrParam(3) = ""									' Name Condotion	
		arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""		' Where Condition
		arrParam(5) = strPopNm								' TextBox ��Ī 
		
		arrField(0) = "Minor_CD"							' Field��(0)
		arrField(1) = "Minor_NM"							' Field��(1)
	    
		arrHeader(0) = strPopNm								' Header��(0)
		arrHeader(1) = strPopNm & "��"					' Header��(1)

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMinorCD(strMajorCd, arrRet)
	End If	
			
End Function

'========================================================================================================
Function OpenPort(strMinorCD, strMinorNM, strPopNm, iwhere)
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
		
		arrParam(0) = strPopNm								' �˾�  ��Ī 
		arrParam(1) = "B_MINOR"								' TABLE ��Ī 
		arrParam(2) = Trim(strMinorCD)						'	Code Condition	
		arrParam(3) = ""									' Name Condotion	
		arrParam(4) = "MAJOR_CD = " & FilterVar("B9092", "''", "S") & ""					' Where Condition
		arrParam(5) = strPopNm								' TextBox ��Ī 
		
		arrField(0) = "Minor_CD"							' Field��(0)
		arrField(1) = "Minor_NM"							' Field��(1)
	    
		arrHeader(0) = strPopNm								' Header��(0)
		arrHeader(1) = strPopNm	& "��"					' Header��(1)

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetOpenPort(iwhere, arrRet)
	End If	
			
End Function

'========================================================================================================
Function OpenBizPartner(strBizPartnerCD, strBizPartnerNM, strPopPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = strPopPos							' �˾� ��Ī 
	arrParam(1) = "B_BIZ_PARTNER"					' TABLE ��Ī 
	arrParam(2) = Trim(strBizPartnerCD)				' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "bp_type IN ( " & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ", " & FilterVar("S", "''", "S") & " ) AND usage_flag = " & FilterVar("Y", "''", "S") & " "	' Where Condition
	arrParam(5) = strPopPos							' TextBox ��Ī 

	arrField(0) = "BP_CD"							' Field��(0)
	arrField(1) = "BP_NM"							' Field��(1)

	arrHeader(0) = strPopPos						' Header��(0)
	arrHeader(1) = strPopPos & "��"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizPartner(strPopPos, arrRet)
	End If
End Function	
	

'========================================================================================================
Function OpenCountry()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "����"							' �˾� ��Ī 
	arrParam(1) = "B_COUNTRY"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtOriginCntry.value)		' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "����"							' TextBox ��Ī 

	arrField(0) = "COUNTRY_CD"							' Field��(0)
	arrField(1) = "COUNTRY_NM"							' Field��(1)

	arrHeader(0) = "����"							' Header��(0)
	arrHeader(1) = "������"							' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCountry(arrRet)
	End If
End Function

'========================================================================================================
Function SetLCNo(strRet)
	frm1.txtLCNo.value = strRet(0)
	frm1.txtLCNo.focus
End Function


'========================================================================================================
Function SetSORef(strRet)
	
	Call ggoOper.ClearField(Document, "A")								 '��: Clear Content  Field 
	Call SetRadio()
	Call InitVariables													 '��: Initializes local global variables 
	Call SetDefaultVal
					
	frm1.txtSONo.value = strRet

	Dim strVal

				
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If

		
	strVal = BIZ_PGM_SOQRY_ID & "?txtSONo=" & Trim(frm1.txtSONo.value)	'��: �����Ͻ� ó�� ASP�� ���� 
		
		
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

	lgBlnFlgChgValue = True
End Function


'========================================================================================================

Function SetBank(Byval iwhere, arrRet)
	select Case UCase(iwhere)
	
		Case "0"
			frm1.txtAdvBank.Value = arrRet(0)
			frm1.txtAdvBankNm.Value = arrRet(1)
			frm1.txtAdvBank.focus
		Case "1"
			frm1.txtOpenBank.value = arrRet(0)
			frm1.txtOpenBankNm.value = arrRet(1)
			frm1.txtOpenBank.focus
		Case "2"
			frm1.txtPayBank.Value = arrRet(0)
			frm1.txtPayBankNm.Value = arrRet(1)
			frm1.txtPayBank.focus
		Case "3"
			frm1.txtRenegoBank.Value = arrRet(0)
			frm1.txtRenegoBankNm.Value = arrRet(1)
			frm1.txtRenegoBank.focus
		Case "4"
			frm1.txtConfirmBank.Value = arrRet(0)
			frm1.txtConfirmBankNm.Value = arrRet(1)	
			frm1.txtConfirmBank.focus
		Case Else
	
	End Select
	
	lgBlnFlgChgValue = True
	
End Function

'========================================================================================================
Function SetMinorCd(strMajorCd,arrRet)
	Select Case strMajorCd
		
		Case gstrLCTypeMajor
			frm1.txtLCType.Value = arrRet(0)
			frm1.txtLCTypeNm.Value = arrRet(1)
			frm1.txtLCType.focus			

		Case gstrTransportMajor
			frm1.txtTransport.Value = arrRet(0)
			frm1.txtTransportNm.Value = arrRet(1)
			frm1.txtTransport.focus		

		Case gstrFreightMajor
			frm1.txtFreight.Value = arrRet(0)
			frm1.txtFreightNm.Value = arrRet(1)
			frm1.txtFreight.focus		

		Case gstrCreditCoreMajor
			frm1.txtCreditCore.Value = arrRet(0)
			frm1.txtCreditCoreNm.Value = arrRet(1)
			frm1.txtCreditCore.focus		
	
		Case gstrOriginMajor
			frm1.txtOrigin.value = arrRet(0)
			frm1.txtOriginNm.value = arrRet(1) 
			frm1.txtOrigin.focus	
		Case Else
		
	End Select

	lgBlnFlgChgValue = True
	
End Function
	

'========================================================================================================
Function SetOpenPort(iwhere, arrRet)
	
	Select Case iwhere				
		Case 0
			frm1.txtLoadingPort.Value = arrRet(0)
			frm1.txtLoadingPortNm.Value = arrRet(1)	
			frm1.txtLoadingPort.focus
		Case 1
			frm1.txtDischgePort.Value = arrRet(0)
			frm1.txtDischgePortNm.Value = arrRet(1)	
			frm1.txtDischgePort.focus
	End Select			
					
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function SetBizPartner(strPopPos, arrRet)
	Select Case UCase(strPopPos)
		Case "����ó"
			frm1.txtNotifyParty.Value = arrRet(0)
			frm1.txtNotifyPartyNm.Value = arrRet(1)
			frm1.txtNotifyParty.focus	
		Case "������"
			frm1.txtAgent.Value = arrRet(0)
			frm1.txtAgentNm.Value = arrRet(1)
			frm1.txtAgent.focus		
		Case "������"
			frm1.txtManufacturer.Value = arrRet(0)
			frm1.txtManufacturerNm.Value = arrRet(1)
			frm1.txtManufacturer.focus		
		Case "��Ź��"
			frm1.txtConsignee.value = arrRet(0)
			frm1.txtConsigneeNm.value = arrRet(1)  	
			frm1.txtConsignee.focus		
		Case Else
	End Select

	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function SetCountry(arrRet)
	frm1.txtOriginCntry.Value = arrRet(0)
	frm1.txtOriginCntryNm.Value = arrRet(1)
	frm1.txtOriginCntry.focus
	
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp, arrVal
	Select Case Kubun
		Case 1
			WriteCookie CookieSplit, frm1.txtLCNo.value
		Case 0
				
			strTemp = ReadCookie(CookieSplit)
				
			If strTemp = "" then Exit Function
					
			frm1.txtLCNo.value =  strTemp
					
			If Err.number <> 0 Then
				Err.Clear
				WriteCookie CookieSplit , ""
				Exit Function 
			End If
				
			Call MainQuery()
							
			WriteCookie CookieSplit , ""
				
		Case 2
			WriteCookie CookieSplit , "EL" & parent.gRowSep & frm1.txtSalesGroup.value & parent.gRowSep & frm1.txtSalesGroupNm.value & parent.gRowSep & frm1.txtLCNo.value 
		
	End Select
	
End Function


'========================================================================================================
Function SetRadio()
	Dim blnOldFlag

	blnOldFlag = lgBlnFlgChgValue

	frm1.rdoPartailShip1.checked = True
	frm1.rdoTranshipment1.checked = True
	frm1.rdoBLAwFlg1.checked = True
	frm1.rdoTransfer1.checked = True
	frm1.rdoChargeCd1.checked = True

	lgBlnFlgChgValue = blnOldFlag
End Function

'========================================================================================================
Function JumpChgCheck(ByVal IWhere)
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then Exit Function
	End If	
		
	Select Case IWhere
	Case 0 
		Call CookiePage(1)
		Call PgmJump(LC_DETAIL_ENTRY_ID)
	Case 1
		Call PgmJump(LCAMEND_HDR_ENTRY_ID)
	Case 2
		Call CookiePage(2)
		Call PgmJump(EXPORT_CHARGE_ENTRY_ID)
	End Select	
End Function


'============================================================================================================
Function ProtectXchRate()
	If frm1.txtCurrency.value = parent.gCurrency Then
		Call ggoOper.SetReqAttr(frm1.txtXchRate, "Q")
	End If	
End Function


'===========================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		'�����ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		    
	End With
End Sub

'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029																 '��: Load table , B_numeric_format 
	Call AppendNumberPlace("6", "3", "0")
	Call AppendNumberPlace("7", "2", "0")
	Call AppendNumberPlace("8", "2", "4")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '��: Lock  Suitable  Field 
		
	Call SetDefaultVal
		
	 '----------  Coding part  ------------------------------------------------------------- 

	Call SetToolBar("11100000000011")										 '��: ��ư ���� ���� 

	Call InitVariables
	Call CookiePage(0)	
	Call changeTabs(TAB1)
    gTabMaxCnt = 4
    gIsTab = "Y"
	frm1.txtLCNo.focus
End Sub

'========================================================================================================
Sub btnLCNoOnClick()
	If frm1.txtLCNo.readOnly <> True Then
		Call OpenLCNoPop()
	End If
End Sub


'========================================================================================================
Sub btnAdvBankOnClick()
	If frm1.txtAdvBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtAdvBank.value,	frm1.txtAdvBankNm.value, "0")
	End If
End Sub


'========================================================================================================
Sub btnOpenBankOnClick()
	If frm1.txtOpenBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtOpenBank.value, frm1.txtOpenBankNm.value, "1")
	End If
End Sub

'========================================================================================================
Sub btnPayBankOnClick()
	If frm1.txtPayBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtPayBank.value, frm1.txtPAyBankNm.value, "2")
	End If
End Sub

'========================================================================================================
Sub btnRenegoBankOnClick()
	If frm1.txtRenegoBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtRenegoBank.value, frm1.txtRenegoBankNm.value, "3")
	End If
End Sub

'========================================================================================================
Sub btnConfirmBankOnClick()
	If frm1.txtConfirmBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtConfirmBank.value, frm1.txtConfirmBankNm.value, "4")
	End If
End Sub


'========================================================================================================
Sub btnNotifyPartyOnClick()
	If frm1.txtNotifyParty.readOnly <> True Then
		Call OpenBizPartner(frm1.txtNotifyParty.value, frm1.txtNotifyPartyNm.value, "����ó")
	End If
End Sub

'========================================================================================================
Sub btnAgentOnClick()
	If frm1.txtAgent.readOnly <> True Then
		Call OpenBizPartner(frm1.txtAgent.value, frm1.txtAgentNm.value, "������")
	End If
End Sub


'========================================================================================================
Sub btnManufacturerOnClick()
	If frm1.txtManufacturer.readOnly <> True Then
		Call OpenBizPartner(frm1.txtManufacturer.value, frm1.txtManufacturerNm.value, "������")
	End If
End Sub


'=======================================================================================================
Sub btnOriginOnClick()
	If frm1.txtOrigin.readOnly <> True Then
		Call OpenMinorCd(frm1.txtOrigin.value, frm1.txtOriginNm.value, "������", gstrOriginMajor)
	End If
End Sub
	

'========================================================================================================
Sub btnOriginCntryOnClick()
	If frm1.txtOriginCntry.readOnly <> True Then
		Call OpenCountry()
	End If
End Sub


'========================================================================================================
Sub btnLCTypeOnClick()
	If frm1.txtLCType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtLCType.value, frm1.txtLCTypeNm.value, "L/C����", gstrLCTypeMajor)
	End If
End Sub

'========================================================================================================
Sub btnLoadingPortOnClick()
	If frm1.txtLoadingPort.readOnly <> True Then
		Call OpenPort(frm1.txtLoadingPort.value, frm1.txtLoadingPortNm.value, "������", 0)
	End If
End Sub

'========================================================================================================
Sub btnDischgePortOnClick()
	If frm1.txtDischgePort.readOnly <> True Then
		Call OpenPort(frm1.txtDischgePort.value, frm1.txtDischgePortNm.value, "������", 1)
	End If
End Sub


'========================================================================================================
Sub btnTransportOnClick()
	If frm1.txtTransport.readOnly <> True Then
		Call OpenMinorCd(frm1.txtTransport.value, frm1.txtTransportNm.value, "��۹��", gstrTransportMajor)
	End If
End Sub


'========================================================================================================
Sub btnFreightOnClick()
	If frm1.txtFreight.readOnly <> True Then
		Call OpenMinorCd(frm1.txtFreight.value, frm1.txtFreightNm.value, "�������ҹ��", gstrFreightMajor)
	End If
End Sub


'========================================================================================================
Sub btnCreditCoreOnClick()
	If frm1.txtCreditCore.readOnly <> True Then
		Call OpenMinorCd(frm1.txtCreditCore.value, frm1.txtCreditCoreNm.value, "�ſ������ü", gstrCreditCoreMajor)
	End If
End Sub
	

'========================================================================================================
Sub btnConsigneeOnClick()
	If frm1.txtCreditCore.readOnly <> True Then
		Call OpenBizPartner(frm1.txtConsignee.value, frm1.txtConsigneeNm.value, "��Ź��")
	End If
End Sub	

'========================================================================================================
Sub rdoPartailShip1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoPartailShip2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub rdoTranshipment1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoTranshipment2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub chkInvCnt_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub chkPackList_OnClick()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub rdoBLAwFlg1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoBLAwFlg2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub chkCertOriginFlg_OnClick()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub chkInsurPolicy_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtAdvDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAdvDt.Action = 7 
        Call SetFocusToDocument("M")
        frm1.txtAdvDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtExpireDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtExpireDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtExpireDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtShipFinDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtShipFinDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtShipFinDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub


'=======================================================================================================
Sub txtOpenDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtOpenDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtOpenDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub


'=======================================================================================================
Sub txtLatestShipDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtLatestShipDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtLatestShipDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtAdvDt_Change()
    lgBlnFlgChgValue = True
End Sub


'=======================================================================================================
Sub txtExpireDt_Change()
    lgBlnFlgChgValue = True
End Sub


'=======================================================================================================
Sub txtShipFinDt_Change()
    lgBlnFlgChgValue = True
End Sub


'=======================================================================================================
Sub txtOpenDt_Change()
    lgBlnFlgChgValue = True
End Sub


'=======================================================================================================
Sub txtLatestShipDt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub rdoTransfer1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoTransfer2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================

Sub rdoChargeCd1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoChargeCd2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================

Sub txtXchRate_Change()
	Err.Clear																			'��: Protect system from crashing
	If frm1.txtCurrency.value = parent.gCurrency Then
		frm1.txtXchRate.text = 1
		frm1.txtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtDocAmt.text) * UNICDbl(frm1.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	Else
		If Len(frm1.txtCurrency.value) Then
			If IsNumeric(frm1.txtXchRate.text) = True And IsNumeric(frm1.txtDocAmt.text) = True Then
				If frm1.txtExchRateOp.value = "*" then
					frm1.txtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtDocAmt.text) * UNICDbl(frm1.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
				ElseIf frm1.txtExchRateOp.value = "/" then
					frm1.txtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtDocAmt.text) / UNICDbl(frm1.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
				End if 												
				lgBlnFlgChgValue = True
			End If
		End If
	End If
End Sub


'========================================================================================================

Sub txtDocAmt_Change()
	With frm1
		
	If .txtCurrency.value = parent.gCurrency Then
		.txtXchRate.text = 1
		.txtLocAmt.text = UNIFormatNumber(UNICDbl(.txtDocAmt.text) * UNICDbl(.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	Else	
		If Len(.txtCurrency.value) Then
			If IsNumeric(.txtXchRate.text) = True And IsNumeric(.txtDocAmt.text) = True Then
				If .txtExchRateOp.value = "*" then
					.txtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtDocAmt.text) * UNICDbl(frm1.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
				ElseIf .txtExchRateOp.value = "/" then
					.txtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtDocAmt.text) / UNICDbl(frm1.txtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
				End If 
				lgBlnFlgChgValue = True
			End If
		End If
	End If
	End With
End Sub


'========================================================================================================
Sub txtLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub txttolerance_Change()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub txtLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub



'========================================================================================================
Sub txtFileDt_Change()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub txtInvCnt_Change()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub txtPackList_Change()
	lgBlnFlgChgValue = True
End Sub






'========================================================================================================
Function FncQuery()

    Dim IntRetCD 
	    
    FncQuery = False                                                        
	    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x") '�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	 '------ Erase contents area ------ 
	Call ggoOper.ClearField(Document, "2")								 '��: Clear Contents  Field 
	Call SetDefaultVal
	Call InitVariables													 '��: Initializes local global variables 

	 '------ Check condition area ------ 
	If Not chkField(Document, "1") Then							 '��: This function check indispensable field 
		Exit Function
	End If

	 '------ Query function call area ------ 
		
	Call DbQuery()														 '��: Query db data 

	FncQuery = True														 '��: Processing is OK 
End Function
	


'========================================================================================================
Function FncNew()
	Dim IntRetCD 

	FncNew = False														'��: Processing is NG

	 '------ Check previous data area ------ 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	 '------ Erase condition area ------ 
	 '------ Erase contents area ------ 
	Call ggoOper.ClearField(Document, "A")								'��: Clear Condition Field	
	Call ggoOper.LockField(Document, "N")								'��: Lock  Suitable  Field
	Call SetDefaultVal
	Call SetRadio()
	Call SetToolBar("11100000000011")										 '��: ��ư ���� ���� 
	Call InitVariables													'��: Initializes local global variables
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
	FncNew = True														'��: Processing is OK
End Function
	

'========================================================================================================
Function FncDelete()
	Dim IntRetCD

	FncDelete = False												 '��: Processing is NG 
		
	 '------ Precheck area ------ 
	If lgIntFlgMode <> parent.OPMD_UMODE Then								 'Check if there is retrived data 
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "x", "x")

	If IntRetCD = vbNo Then
		Exit Function
	End If

	 '------ Delete function call area ------ 
	Call DbDelete													 '��: Delete db data 

	FncDelete = True												 '��: Processing is OK 
End Function

'========================================================================================================
Function FncSave()
	Dim IntRetCD
		
	FncSave = False														 '��: Processing is NG 
		
	Err.Clear															 '��: Protect system from crashing 
		
	frm1.txtLCNo.focus
	Set gActiveElement = document.activeElement 
		
	 '------ Precheck area ------ 
	If lgBlnFlgChgValue = False Then								 'Check if there is retrived data 
	    IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		    					 '��: No data changed!! 
	    Exit Function
	End If
		
	 '------ Check contents area ------ 
	If Not chkField(Document, "2") Then								 '��: Check contents area 
		Call SetToolBar("11101000000111")
	    If gPageNo > 0 Then
	        gSelframeFlg = gPageNo
	    End If
	    Exit Function
	End If 

	If Len(Trim(frm1.txtAdvDt.Text)) And Len(Trim(frm1.txtOpenDt.Text)) Then
		If parent.UniConvDateToYYYYMMDD(frm1.txtOpenDt.Text, parent.gDateFormat, "-") > parent.UniConvDateToYYYYMMDD(frm1.txtAdvDt.Text, parent.gDateFormat, "-") Then
			Call DisplayMsgBox("970023", "x", frm1.txtAdvDt.Alt, frm1.txtOpenDt.Alt)
			Call ClickTab1()
			frm1.txtAdvDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If

	If Len(Trim(frm1.txtLatestShipDt.Text)) And Len(Trim(frm1.txtAdvDt.Text)) Then
		If parent.UniConvDateToYYYYMMDD(frm1.txtAdvDt.Text, parent.gDateFormat, "-") > parent.UniConvDateToYYYYMMDD(frm1.txtLatestShipDt.Text, parent.gDateFormat, "-") Then
			Call DisplayMsgBox("970023", "x", frm1.txtLatestShipDt.Alt, frm1.txtAdvDt.Alt)
			Call ClickTab2()
			frm1.txtLatestShipDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If

	If Len(Trim(frm1.txtExpireDt.Text)) And Len(Trim(frm1.txtLatestShipDt.Text)) Then
		If parent.UniConvDateToYYYYMMDD(frm1.txtLatestShipDt.Text, parent.gDateFormat, "-") > parent.UniConvDateToYYYYMMDD(frm1.txtExpireDt.Text, parent.gDateFormat, "-") Then
			Call DisplayMsgBox("970023", "x", frm1.txtExpireDt.Alt, frm1.txtLatestShipDt.Alt)
			'MsgBox "pObjToDt(��)�� pObjFromDt���� ũ�ų� ���ƾ� �մϴ�.", vbExclamation, "uniERP(Warning)"
			Call ClickTab1()
			frm1.txtExpireDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If
	
	If UNICDbl(frm1.txtDocAmt.text) <= 0 Then
		Call DisplayMsgBox("970022", "x", "�����ݾ�","0")
		Call ClickTab1()			
		frm1.txtDocAmt.focus 
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
		
		
	 '------ Save function call area ------ 
	Call DbSave															 '��: Save db data 
		
	FncSave = True														 '��: Processing is OK 
End Function


'========================================================================================================
Function FncCopy()
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")			'��: "Will you destory previous data"
'			IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	lgIntFlgMode = parent.OPMD_CMODE											'��: Indicates that current mode is Crate mode

	 '------ ���Ǻ� �ʵ带 �����Ѵ�. ------ 
	Call ggoOper.ClearField(Document, "1")								'��: Clear Condition Field
	Call ggoOper.LockField(Document, "N")								'��: This function lock the suitable field
	frm1.txtLCNo1.value = "" 
	lgBlnFlgChgValue = True
End Function


'========================================================================================================
Function FncCancel() 
	On Error Resume Next												'��: Protect system from crashing
End Function


'========================================================================================================
Function FncInsertRow()
	On Error Resume Next												'��: Protect system from crashing
End Function


'========================================================================================================
Function FncDeleteRow()
	On Error Resume Next												'��: Protect system from crashing
End Function

'========================================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================
Function FncPrev() 
    Dim strVal
    Dim IntRetCD
	    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "x", "x", "x")  '�� �ٲ�κ� 
        'Call MsgBox("��ȸ���Ŀ� �����˴ϴ�.", vbInformation)
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x") '�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

					
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If



	frm1.txtPrevNext.value = "PREV"

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo1.value)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		'��: ��ȸ ���� ����Ÿ 
	         
	Call RunMyBizASP(MyBizASP, strVal)
End Function


'========================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD
	    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "x", "x", "x")  '�� �ٲ�κ� 
        'Call MsgBox("��ȸ���Ŀ� �����˴ϴ�.", vbInformation)
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x") '�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

					
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If



	frm1.txtPrevNext.value = "NEXT"

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo1.value)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		'��: ��ȸ ���� ����Ÿ 
	         
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function


'========================================================================================================
Function DbQuery()
	Err.Clear															'��: Protect system from crashing

	DbQuery = False														'��: Processing is NG
							
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If
		
	Dim strVal

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtLcKind=" & "M"
		
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
	DbQuery = True														'��: Processing is NG
End Function
	

'========================================================================================================
Function DbSave()
	Err.Clear															'��: Protect system from crashing

	DbSave = False														'��: Processing is NG

	Dim strVal		
					
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If
		
	If frm1.chkSONoFlg.checked = True Then
		frm1.txtSoNoFlg.value = "Y"
	End If	
		
	With frm1
		.txtMode.value = parent.UID_M0002										'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
	End With

	DbSave = True														'��: Processing is NG
End Function
	

'========================================================================================================
Function DbDelete()
	Err.Clear															'��: Protect system from crashing

	DbDelete = False													'��: Processing is NG

	Dim strVal

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003					'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo1.value)			'��: ���� ���� ����Ÿ 
	strVal = strVal & "&txtSONo=" & Trim(frm1.txtSONo.value)			'��: ���� ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

	DbDelete = True														'��: Processing is NG
End Function
	

'========================================================================================================
Function DbQueryOk()													 '��: ��ȸ ������ ������� 
	 '------ Reset variables area ------ 
	lgIntFlgMode = parent.OPMD_UMODE									'��: Indicates that current mode is Update mode 
	lgBlnFlgChgValue = False
	frm1.txtPrevNext.value = ""
		
	Call ggoOper.LockField(Document, "Q")								 '��: This function lock the suitable field 
	Call SetToolBar("111110001101111")
	Call ProtectXchRate()
	If gSelframeFlg <> TAB1 Then
		Call ClickTab1()
	End If
End Function

'========================================================================================================
Function SOQueryOK()													 '��: ��ȸ ������ ������� 
	Call ProtectXchRate()
	Call txtDocAmt_Change()
	Call SetToolBar("11101000000011")	
End Function

'========================================================================================================
Function DbSaveOk()														'��: ���� ������ ���� ���� 
	Call InitVariables
	Call MainQuery()
End Function
	

'========================================================================================================
Function DbDeleteOk()													'��: ���� ������ ���� ���� 
	lgBlnFlgChgValue = False
	Call MainNew()
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
		<TABLE <%=LR_SPACE_TYPE_00%>>
			<TR>
				<TD <%=HEIGHT_TYPE_00%>></TD>
			</TR>
			
			<TR HEIGHT=23>
				<TD WIDTH=100%>
					<TABLE <%=LR_SPACE_TYPE_10%>>
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD CLASS="CLSMTAB">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
									<TR>
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C �ݾ�����</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTAB">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C ��������</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>	
							<TD CLASS="CLSMTAB">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab3()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���񼭷�</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTAB">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab4()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����ױ�Ÿ</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD WIDTH=* align=right><A href="vbscript:OpenSORef">��������</A></TD>
							<TD WIDTH=10>&nbsp;</TD>
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
											<TD CLASS=TD5 NOWRAP>L/C������ȣ</TD>
											<TD	CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCNo" SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="L/C������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCNo" ALIGN=top TYPE="BUTTON" ONCLICK ="vbscript:btnLCNoOnClick()"></TD>
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
							<TD WIDTH=100% VALIGN=TOP>
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">	
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>L/C������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCNo1"  SIZE=20 MAXLENGTH=18 TAG="25XXXU" ALT="L/C������ȣ"></TD>
											<TD CLASS=TD5 NOWRAP>���ֹ�ȣ</TD>
											<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE=TEXT NAME="txtSONo" SIZE=20 MAXLENGTH=18 TAG="24XXXU" ALT="���ֹ�ȣ">&nbsp;&nbsp;&nbsp;
												<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="25X" VALUE="Y" NAME="chkSONoFlg" ID="chkSONoFlg">
												<LABEL FOR="chkSONoFlg">���ֹ�ȣ����</LABEL>
											</TD>
										</TR>	
										<TR>	
											<TD CLASS=TD5 NOWRAP>L/C��ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="L/C��ȣ" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="22XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdvNo" SIZE=35 MAXLENGTH=35 TAG="21XXXU" ALT="������ȣ"></TD>
										</TR>		
										<TR>
											<TD CLASS=TD5 NOWRAP>L/C����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCType" SIZE=10 MAXLENGTH=5 STYLE="TEXT-ALIGN: left" TAG="22XXXU" ALT="L/C����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnLCTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtLCTypeNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>��ȿ��</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtExpireDt" CLASS=FPDTYYYYMMDD tag="22X" Title="FPDATETIME" ALT="��ȿ��"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdvBank" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" TAG="22XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAdvbank" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnAdvBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtAdvBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtAdvDt" CLASS=FPDTYYYYMMDD tag="22X" Title="FPDATETIME" ALT="������"></OBJECT>');</SCRIPT></TD>
											</TR>									
										<TR>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" TAG="22XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenBank" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnOpenBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>������</TD>						
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtOpenDt" CLASS=FPDTYYYYMMDD tag="22X" Title="FPDATETIME" ALT="������"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�����ݾ�</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtDocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="22X2Z" ALT="�����ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>
														<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="ȭ��"></TD>
													</TR>
												</TABLE>
											</TD>	
											<TD CLASS=TD5 NOWRAP>�����ڱ��ݾ�</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtLocAmt" TABINDEX = "-1" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2Z" ALT="�����ڱ��ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
														<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="�ڱ�ȭ��"></TD>
													</TR>
												</TABLE>
											</TD>	
										</TR>	
										<TR>							
											<TD CLASS=TD5 NOWRAP>ȯ��</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchRate" style="HEIGHT: 20px; WIDTH: 150px" tag="22X5Z" ALT="ȯ��" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�����ݾװ����������</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txttolerance" style="HEIGHT: 20px; WIDTH: 150px" tag="21X8Z" ALT="�����ݾװ����������" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;%
											</TD>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
										</TR>	
										<TR>	
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoTerms" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="��������">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtIncoTermsNm" SIZE=20 TAG="24"></TD>										
											<TD CLASS=TD5 NOWRAP>�����׷�</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="�����׷�">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=5 TAG="24XXU" ALT="�������">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>�����Ⱓ</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtPayDur" style="HEIGHT: 20px; WIDTH: 50px" TAG="24X6" ALT="�����Ⱓ" Title="FPDOUBLESINGLE"><PARAM NAME="MaxValue" VALUE="999"><PARAM NAME="MinValue" VALUE="0"></OBJECT>');</SCRIPT>&nbsp;DAYS</TD>
										</TR>								
										<%Call SubFillRemBodyTD5656(10)%>
									</TABLE>
								</DIV>
								
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>��۹��</TD>
											<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtTransport" SIZE=10 MAXLENGTH=5 TAG="22XXXU" ALT="��۹��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransport" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnTransportOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtTransportNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtLoadingPort" ALT="������" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingPort" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnLoadingPortOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingPortNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtDischgePort" ALT="������" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgePort" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnDischgePortOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtDischgePortNm" SIZE=20 TAG="24"></TD>
										</TR>	
										<TR>
											<TD CLASS=TD5 NOWRAP>������������</TD>	
											<TD CLASS=TD656 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 NAME="txtLatestShipDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="������������"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>ȯ�����</TD>
											<TD CLASS=TD656 COLSPAN=3>
													<INPUT TYPE="RADIO" CLASS="RADIO" VALUE="Y" NAME="rdoTranshipment" TAG="21X" ID="rdoTranshipment1"><LABEL FOR="rdoTranshipment1">Y</LABEL>&nbsp;&nbsp;&nbsp;
													<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTranshipment" TAG="21X" VALUE="N" CHECKED ID="rdoTranshipment2"><LABEL FOR="rdoTranshipment2">N</LABEL>
											</TD>
										</TR>
										<TR>	
											<TD CLASS=TD5 NOWRAP>���Ҽ������</TD>
											<TD CLASS=TD656 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="21X" VALUE="Y" CHECKED ID="rdoPartailShip1"><LABEL FOR="rdoPartailShip1">Y</LABEL>&nbsp;&nbsp;&nbsp;
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="21X" VALUE="N" ID="rdoPartailShip2"><LABEL FOR="rdoPartailShip2">N</LABEL>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�ε����</TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtDeliveryPlce" ALT="�ε����" TYPE=TEXT MAXLENGTH=120 SIZE=35 TAG="21X"></TD>
										</TR>	
										<%Call SubFillRemBodyTD656(15)%>
									</TABLE>
								</DIV>
					
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>�������ñⰣ</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtFileDt" style="HEIGHT: 20px; WIDTH: 50px" tag="21X7" ALT="�������ñⰣ" Title="FPDOUBLESINGLE"><PARAM NAME="MaxValue" VALUE="99"><PARAM NAME="MinValue" VALUE="0"></OBJECT>');</SCRIPT>&nbsp;DAYS
											</TD>
											<TD CLASS=TD5 NOWRAP>�������ñⰣ ����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFileDtTxt" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="21X" ALT="�������ñⰣ ����"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�������</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtInvCnt" style="HEIGHT: 20px; WIDTH: 50px" tag="21X7" ALT="�������" Title="FPDOUBLESINGLE"><PARAM NAME="MaxValue" VALUE="99"><PARAM NAME="MinValue" VALUE="0"></OBJECT>');</SCRIPT>&nbsp;��</TD>
											<TD CLASS=TD5 NOWRAP>�������</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtPackList" style="HEIGHT: 20px; WIDTH: 50px" tag="21X7" ALT="�������" Title="FPDOUBLESINGLE"><PARAM NAME="MaxValue" VALUE="99"><PARAM NAME="MinValue" VALUE="0"></OBJECT>');</SCRIPT>&nbsp;��</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP><INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkCertOriginFlg" ID="chkCertOriginFlg"></TD>
											<TD CLASS=TD6 NOWRAP><LABEL FOR="chkCertOriginFlg">����������</LABEL></TD>
											<TD CLASS=TD5 NOWRAP>B/L����</TD>
											<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoBLAwFlg" TAG="21X" VALUE="Y" CHECKED ID="rdoBLAwFlg1">
												<LABEL FOR="rdoBLAwFlg">BILL OF LADING</LABEL>
												<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoBLAwFlg" TAG="21X" VALUE="N" ID="rdoBLAwFlg2">
												<LABEL FOR="rdoBLAwFlg">AIRWAY BILL</LABEL>
											</TD>
										</TR>	
										<TR>	
											<TD CLASS=TD5 NOWRAP>�������ҹ��</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFreight" SIZE=10 MAXLENGTH=5 TAG="21X" ALT="�������ҿ���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFreight" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnFreightOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtFreightNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>����ó</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNotifyParty" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNotifyParty" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnNotifyPartyOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtNotifyPartyNm" SIZE=20 TAG="24"></TD>	
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��Ź��</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtConsignee"  SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="��Ź��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConsignee" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnConsigneeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtConsigneeNm" SIZE=20 TAG="24"></TD></TD>	
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�κ�����</TD>
											<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtInsurPolicy" ALT="����κ�����" TYPE=TEXT MAXLENGTH=30 SIZE=84 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���񼭷�</TD>
											<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDoc1"  TYPE=TEXT MAXLENGTH=120 SIZE=84 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDoc2"  TYPE=TEXT MAXLENGTH=120 SIZE=84 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDoc3"  TYPE=TEXT MAXLENGTH=120 SIZE=84 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDoc4"  TYPE=TEXT MAXLENGTH=120 SIZE=84 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDoc5"  TYPE=TEXT MAXLENGTH=120 SIZE=84 TAG="21X"></TD>
										</TR>
										<%Call SubFillRemBodyTD5656(8)%>
									</TABLE>
								</DIV>
						
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayBank" TYPE=TEXT SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayBank" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnPayBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtPayBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>RENEGO����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRenegoBank" TYPE=TEXT SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="RENEGO����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRenegoBank" ALIGN=TOP TYPE="BUTTON" ONCLICK ="vbscript:btnRenegoBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtRenegoBankNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>Ȯ������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtConfirmBank" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="Ȯ������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConfirmBank" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnConfirmBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtConfirmBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>�������û���</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBankTxt" SIZE=35 MAXLENGTH=120 TAG="21X" ALT="�������û���"></TD>
										</TR>	
										<TR>
											<TD CLASS=TD5 NOWRAP>�絵��뿩��</TD>
											<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoTransfer" TAG="21X" VALUE="Y" CHECKED ID="rdoTransfer1"><LABEL FOR="rdoTransfer">Y</LABEL>
												<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoTransfer" TAG="21X" VALUE="N" ID="rdoTransfer2"><LABEL FOR="rdoTransfer">N</LABEL>
											</TD>
											<TD CLASS=TD5 NOWRAP>�ſ������ü</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCreditCore" SIZE=10 MAXLENGTH=5 TAG="21XXXU" ALT="�ſ������ü"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCreditCore" ALIGN=TOP TYPE="BUTTON" ONCLICK ="vbscript:btnCreditCoreOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtCreditCoreNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������ �δ���</TD>
											<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoChargeCd" TAG="21X" VALUE="Y" CHECKED ID="rdoChargeCd1"><LABEL FOR="rdoTransfer">��û��</LABEL>
												<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoChargeCd" TAG="21X" VALUE="N" ID="rdoChargeCd2"><LABEL FOR="rdoTransfer">������</LABEL>
											</TD>	
											<TD CLASS=TD5 NOWRAP>������ ����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtChargeTxt" SIZE=35 MAXLENGTH=30 TAG="21X" ALT="������ ��������"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��ݰ��� ����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPaymentTxt" SIZE=35 MAXLENGTH=120 TAG="21X" ALT="��� ��������"></TD>
											<TD CLASS=TD5 NOWRAP>�������� ����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtShipment" SIZE=35 MAXLENGTH=120 TAG="21X" ALT="�������� ��������"></TD>
										</TR>											
										<TR>
											<TD CLASS=TD5 NOWRAP>����������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPreAdvRef" ALT="������ ��������" TYPE=TEXT MAXLENGTH=120 SIZE=35 TAG="21X"></TD>
											<TD CLASS=TD5 NOWRAP>���ȸ��</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTransportComp" ALT="���ȸ��" TYPE=TEXT MAXLENGTH=50 SIZE=35 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOrigin" ALT="������" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrigin" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnOriginOnClick()">&nbsp;<INPUT NAME="txtOriginNm" ALT="��������" TYPE=TEXT MAXLENGTH=30 SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>����������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOriginCntry" ALT="����������" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOriginCntry" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnOriginCntryOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginCntryNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAgent" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnAgentOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnManufacturer" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnManufacturerOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��Ÿ����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRemark" ALT="��Ÿ����" TYPE=TEXT MAXLENGTH=120 SIZE=35 TAG="21X"></TD>
											<TD CLASS=TD5 NOWRAP>AMEND��</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime5 NAME="txtAmendDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="AMEND��"></OBJECT>');</SCRIPT>
											</TD>
										<%Call SubFillRemBodyTD5656(11)%>
										</TR>
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
						<TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck(0)">L/C�������</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(1)">AMEND���</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(2)">�ǸŰ����</A></TD>
						<!--<TD WIDTH=50>&nbsp;</TD>-->
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = -1></IFRAME></TD>
			</TR>
		</TABLE>
		<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtHLCNo" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtHSoNo" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtPrevNext" TAG="24" TABINDEX = -1> 
		<INPUT TYPE=HIDDEN NAME="txtSoNoFlg" TAG="24" TABINDEX = -1> 
		<INPUT TYPE=HIDDEN NAME="txtHOpenDt" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtExchRateOp" TAG="24" TABINDEX = -1> 
	</FORM>
	<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
		<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

</HTML>

