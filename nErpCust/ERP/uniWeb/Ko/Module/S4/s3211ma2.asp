<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211ma2.asp																*
'*  4. Program Name         : Local L/C등록																*
'*  5. Program Desc         : Local L/C등록																*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/10																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/03/22 : Coding Start												*
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

EndDate = UNIConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)    
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID				= "s3211mb2.asp"	
Const BIZ_PGM_SOQRY_ID			= "s3211mb4.asp"	
Const BIZ_PGM_CAL_AMT_ID		= "s3211mb5.asp"
	
Const LC_DETAIL_ENTRY_ID		= "s3212ma2"			
Const LCAMEND_HDR_ENTRY_ID		= "s3221ma2"			
Const EXPORT_CHARGE_ENTRY_ID	= "s6111ma1"		 
	
Const TAB1 = 1
Const TAB2 = 2

 '------ Minor Code PopUp을 위한 Major Code정의 ------ 
Const gstrMLCTypeMajor			= "S9000"				'Master L/C 유형  
Const gstrLCTypeMajor			= "S9001"				'local L/C 유형 
Const gstrPayTermsMajor			= "B9004"

Dim gSelframeFlg					 '현재 TAB의 위치를 나타내는 Flag 
Dim gblnWinEvent					 '~~~ ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 

'========================================================================================================
Function InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE								
	lgBlnFlgChgValue = False								
	lgIntGrpCount = 0										
		
	 '------ Coding part ------ 
	gblnWinEvent = False
End Function
	
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtLocCurrency.value = parent.gCurrency
	lgBlnFlgChgValue = False
End Sub
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %> 
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
	
'========================================================================================================
Function OpenLCNoPop()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Or UCase(frm1.txtLCNo.className) = "PROTECTED" Then Exit Function
		
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("s3211pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3211pa2", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

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
		
	iCalledAspName = AskPRAspName("s3111ra2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3111ra2", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	If strRet = "" Then
		Exit Function
	Else
		Call SetSORef(strRet)
	End If
	
End Function

'========================================================================================================
Function OpenDNRef()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If lgIntFlgMode = parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "x", "x", "x")
		Exit function
	End If 
		
	iCalledAspName = AskPRAspName("s4111ra4")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s4111ra4", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	If strRet = "" Then
		Exit Function
	Else
		Call SetSORef(strRet)
	End If
End Function

'========================================================================================================
Function OpenBankPop(strBankCd, strBankNm, strPopPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "은행"						
	arrParam(1) = "B_BANK"							
	arrParam(2) = Trim(strBankCd)					
	arrParam(3) = ""								
	arrParam(4) = ""								
	arrParam(5) = "은행"						

	arrField(0) = "BANK_CD"							
	arrField(1) = "BANK_NM"							

	arrHeader(0) = "은행"						
	arrHeader(1) = "은행명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBank(strPopPos, arrRet)
	End If
End Function


'========================================================================================================
Function OpenMinorCd(strMinorCD, strMinorNM, strPopNm, strMajorCd)
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
		
		arrParam(0) = strPopNm								' 팝업  명칭 
		arrParam(1) = "B_MINOR"								' TABLE 명칭 
		arrParam(2) = Trim(strMinorCD)						'	Code Condition	
		arrParam(3) = ""									' Name Condotion	
		arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""		' Where Condition
		arrParam(5) = strPopNm								' TextBox 명칭 
		
		arrField(0) = "Minor_CD"							' Field명(0)
		arrField(1) = "Minor_NM"							' Field명(1)
	    
		arrHeader(0) = strPopNm								' Header명(0)
		arrHeader(1) = strPopNm & "명"					' Header명(1)

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
Function OpenPayTerms()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
		
	arrParam(0) = "결제방법"							
	arrParam(1) = "b_minor,b_configuration"					
	arrParam(2) = Trim(frm1.txtPayTerms.Value)				
	arrParam(3) = ""										
	arrParam(4) = "b_minor.MINOR_CD = b_configuration.MINOR_CD AND b_minor.MAJOR_CD = " & FilterVar(gstrPayTermsMajor, "''", "S") & " AND b_configuration.REFERENCE = " & FilterVar("L", "''", "S") & " "
	arrParam(5) = "결제방법"											

	arrField(0) = "b_minor.Minor_CD"											
	arrField(1) = "b_minor.Minor_NM"											

	arrHeader(0) = "결제방법"												
	arrHeader(1) = "결제방법명"												
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPayTerms(arrRet)
	End If
End Function
	

'========================================================================================================
Function SetLCNo(strRet)
	frm1.txtLCNo.value = strRet(0)
	frm1.txtLCNo.focus
End Function


'========================================================================================================
Function SetSORef(strRet)
	Call ggoOper.ClearField(Document, "A")								 
	Call SetRadio()
	Call InitVariables													 
	Call SetDefaultVal

	frm1.txtSONo.value = strRet

	Dim strVal
					
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If

	strVal = BIZ_PGM_SOQRY_ID & "?txtSONo=" & Trim(frm1.txtSONo.value)	

	Call RunMyBizASP(MyBizASP, strVal)									

	lgBlnFlgChgValue = True
End Function


'========================================================================================================
Function SetBank(strPopPos, arrRet)
	Select Case UCase(strPopPos)
		Case "FROMBANK"
			frm1.txtFromBank.Value = arrRet(0)
			frm1.txtFromBankNm.Value = arrRet(1)
			frm1.txtFromBank.focus
				
		Case "OPENBANK"
			frm1.txtOpenBank.Value = arrRet(0)
			frm1.txtOpenBankNm.Value = arrRet(1)
			frm1.txtOpenBank.focus
				
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

		Case gstrMLCTypeMajor
			frm1.txtMLCType.Value = arrRet(0)
			frm1.txtMLCTypeNm.Value = arrRet(1)
			frm1.txtMLCType.focus
			
		Case Else
		
	End Select

	lgBlnFlgChgValue = True
	
End Function


'========================================================================================================
Function SetPayTerms(arrRet)
		
	frm1.txtPayTerms.Value = arrRet(0)
	frm1.txtPayTermsNm.Value = arrRet(1)
	frm1.txtPayTerms.focus
		
	lgBlnFlgChgValue = True
		
End Function

'========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp, arrVal
		
	Select Case Kubun
		
		Case 1
			WriteCookie CookieSplit , frm1.txtLCNo.value

		Case 0

			strTemp = ReadCookie(CookieSplit)
					
			If strTemp = "" Then Exit Function
					
			frm1.txtLCNo.value =  strTemp

			If Err.number <> 0 Then
				Err.Clear
				WriteCookie CookieSplit , ""
				Exit Function 
			Else 
				Call MainQuery()
			End If
							
			WriteCookie CookieSplit , ""
				
		Case 2
			WriteCookie CookieSplit , "EO" & parent.gRowSep & frm1.txtSalesGroup.value & parent.gRowSep & frm1.txtSalesGroupNm.value & parent.gRowSep & frm1.txtLCNo.value
			 			
	End Select
End Function	


'========================================================================================================
Function LoadLCDtl()
	Dim strDtlOpenParam

	WriteCookie "txtLCNo", UCase(Trim(frm1.txtLCNo.value))

	strDtlOpenParam = LC_DETAIL_ENTRY_ID

	document.location.href = GetUserPath & strDtlOpenParam
End Function


'========================================================================================================
Function OpenCookie()
	frm1.txtLCNo.value = ReadCookie("txtLCNo")
	WriteCookie "txtLCNo.value", ""
End Function


'========================================================================================================
Function SetRadio()
	Dim blnOldFlag

	blnOldFlag = lgBlnFlgChgValue

	frm1.rdoPartailShip1.checked = True

	lgBlnFlgChgValue = blnOldFlag
End Function

'=========================================================================== 
Function JumpChgCheck(ByVal IWhere)
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
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
	
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		'개설금액 
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, parent.ggamtofmoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		
		ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		
	End With
End Sub

'========================================================================================================

Sub Form_Load()
	Call LoadInfTB19029																 '⊙: Load table , B_numeric_format 
	Call AppendNumberPlace("6", "3", "0")
	Call AppendNumberPlace("7", "2", "0")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock  Suitable  Field 
		
	Call SetDefaultVal
	Call CookiePage(0)
	Call InitVariables
		
	 '----------  Coding part  ------------------------------------------------------------- 

	Call SetToolBar("11100000000011")											 '⊙: 버튼 툴바 제어 
	Call changeTabs(TAB1)
	frm1.txtLCNo.focus
	Set gActiveElement = document.activeElement 
    gIsTab     = "Y" 
    gTabMaxCnt = 2   
End Sub

'========================================================================================================
Sub btnLCNoOnClick()
	If frm1.txtLCNo.readOnly <> True Then
		Call OpenLCNoPop()
	End If
End Sub


'========================================================================================================
Sub btnFromBankOnClick()
	If frm1.txtFromBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtFromBank.value, frm1.txtFromBankNm.value, "FROMBANK")
	End If
End Sub

'========================================================================================================
Sub btnOpenBankOnClick()
	If frm1.txtOpenBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtOpenBank.value, frm1.txtOpenBankNm.value, "OPENBANK")
	End If
End Sub

'========================================================================================================
Sub btnLCTypeOnClick()
	If frm1.txtLCType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtLCType.value, frm1.txtLCTypeNm.value, "LOCAL L/C유형", gstrLCTypeMajor)
	End If
End Sub
	
'========================================================================================================
Sub btnMLCTypeOnClick()
	If frm1.txtMLCType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtMLCType.value, frm1.txtMLCTypeNm.value, "Master L/C유형", gstrMLCTypeMajor)
	End If
End Sub
'========================================================================================================
Sub btnPayTermsOnClick()
	Call OpenPayTerms()
End Sub

'========================================================================================================
Sub rdoPartailShip1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoPartailShip2_OnPropertyChange()
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
Sub txtExpiryDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtExpiryDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtExpiryDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtMExpiryDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtMExpiryDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtMExpiryDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub txtOpenDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtOpenDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtOpenDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtAmendDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAmendDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtAmendDt.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtMoveDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtMoveDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtMoveDt.Focus
    End If
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
Sub txtXchRate_Change()
	If frm1.txtCurrency.value = parent.gCurrency Then
		frm1.txtXchRate.text = 1
		frm1.txtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtDocAmt.text) * UNICDbl(frm1.txtXchRate.text), ggamtofmoney.DecPoint, -2, 0, ggamtofmoney.RndPolicy, ggamtofmoney.RndUnit)
	Else 	
		If Len(frm1.txtCurrency.value) Then
			If IsNumeric(frm1.txtXchRate.text) = True And IsNumeric(frm1.txtDocAmt.text) = True Then
				If frm1.txtExchRateOp.value = "*" then
					frm1.txtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtDocAmt.text) * UNICDbl(frm1.txtXchRate.text), ggamtofmoney.DecPoint, -2, 0, ggamtofmoney.RndPolicy, ggamtofmoney.RndUnit)
				ElseIf frm1.txtExchRateOp.value = "/" then
					frm1.txtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtDocAmt.text) / UNICDbl(frm1.txtXchRate.text), ggamtofmoney.DecPoint, -2, 0, ggamtofmoney.RndPolicy, ggamtofmoney.RndUnit)
				End If 												
				lgBlnFlgChgValue = True
			End If
		End If
	End If	
End Sub


'========================================================================================================
Sub txtDocAmt_Change()
	If frm1.txtCurrency.value = parent.gCurrency Then
		frm1.txtXchRate.text = 1
		frm1.txtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtDocAmt.text) * UNICDbl(frm1.txtXchRate.text), ggamtofmoney.DecPoint, -2, 0, ggamtofmoney.RndPolicy, ggamtofmoney.RndUnit)
	Else 			
		If IsNumeric(frm1.txtXchRate.text) = True And IsNumeric(frm1.txtDocAmt.text) = True Then
			If frm1.txtExchRateOp.value = "*" then
				frm1.txtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtDocAmt.text) * UNICDbl(frm1.txtXchRate.text), ggamtofmoney.DecPoint, -2, 0, ggamtofmoney.RndPolicy, ggamtofmoney.RndUnit)
			ElseIf frm1.txtExchRateOp.value = "/" then
				frm1.txtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtDocAmt.text) / UNICDbl(frm1.txtXchRate.text), ggamtofmoney.DecPoint, -2, 0, ggamtofmoney.RndPolicy, ggamtofmoney.RndUnit)
			End If 
			lgBlnFlgChgValue = True
		End If
	End If
End Sub


'========================================================================================================
Sub txtAdvDt_Change()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub txtExpiryDt_Change()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub txtMExpiryDt_Change()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub txtOpenDt_Change()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtMoveDt_Change()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
Sub txtAmendDt_Change()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtFileDt_Change()
	lgBlnFlgChgValue = True
End Sub	

'========================================================================================================	
Sub txtPayDur_Change()
	lgBlnFlgChgValue = True
End Sub	
	
'========================================================================================================
Function FncQuery()
	Dim IntRetCD

	FncQuery = False													 '⊙: Processing is NG 

	Err.Clear															  
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")	 '⊙: "Will you destory previous data" 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	 '------ Erase contents area ------ 
	Call ggoOper.ClearField(Document, "2")								 '⊙: Clear Contents  Field 
	Call SetDefaultVal
	Call InitVariables													 '⊙: Initializes local global variables 

	 '------ Check condition area ------ 
	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field 
		Exit Function
	End If

	 '------ Query function call area ------ 
	Call DbQuery()														 '☜: Query db data 

	FncQuery = True														 '⊙: Processing is OK 
End Function
	

'========================================================================================================
Function FncNew()
	Dim IntRetCD 

	FncNew = False														'⊙: Processing is NG

	 '------ Check previous data area ------ 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	Call ggoOper.ClearField(Document, "A")								
	Call ggoOper.LockField(Document, "N")								'⊙: Lock  Suitable  Field
	Call SetDefaultVal
	Call SetRadio()
	Call SetToolBar("11100000000011")										 '⊙: 버튼 툴바 제어 
	Call InitVariables													'⊙: Initializes local global variables
	Call changeTabs(TAB1)
		
	FncNew = True														'⊙: Processing is OK
End Function
	

'========================================================================================================
Function FncDelete()
	Dim IntRetCD

	FncDelete = False												 '⊙: Processing is NG 
		
	 '------ Precheck area ------ 
	If lgIntFlgMode <> parent.OPMD_UMODE Then						'Check if there is retrived data 
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "x", "x")

	If IntRetCD = vbNo Then
		Exit Function
	End If

	 '------ Delete function call area ------ 
	Call DbDelete													 '☜: Delete db data 

	FncDelete = True												 '⊙: Processing is OK 
End Function


'========================================================================================================
Function FncSave()
	Dim IntRetCD
		
	FncSave = False														 '⊙: Processing is NG 
		
	Err.Clear															  
		
	 '------ Precheck area ------ 
	If lgBlnFlgChgValue = False Then								 'Check if there is retrived data 
	    IntRetCD = DisplayMsgBox("900001", "x", "x", "x")					 '⊙: No data changed!! 
'		    Call MsgBox("No data changed!!", vbInformation)
	    Exit Function
	End If
		
	 '------ Check contents area ------ 
	If Not chkField(Document, "2") Then							 '⊙: Check contents area 
	    If gPageNo > 0 Then
	        gSelframeFlg = gPageNo
	    End If
		Exit Function
	End If

	If Len(Trim(frm1.txtAdvDt.Text)) Then
		If UniConvDateToYYYYMMDD(frm1.txtOpenDt.Text, parent.gDateFormat, "-") > UniConvDateToYYYYMMDD(frm1.txtAdvDt.Text, parent.gDateFormat, "-") Then
			Call DisplayMsgBox("970023", "x", frm1.txtAdvDt.Alt, frm1.txtOpenDt.Alt)
			Call changeTabs(TAB1)
			frm1.txtAdvDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If

	If Len(Trim(frm1.txtMoveDt.Text)) Then
		If UniConvDateToYYYYMMDD(frm1.txtAdvDt.Text, parent.gDateFormat, "-") > UniConvDateToYYYYMMDD(frm1.txtMoveDt.Text, parent.gDateFormat, "-") Then
			Call DisplayMsgBox("970023", "x", frm1.txtMoveDt.Alt, frm1.txtAdvDt.Alt)
			Call changeTabs(TAB1)
			frm1.txtMoveDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If

	If Len(Trim(frm1.txtExpiryDt.Text)) Then
		If UniConvDateToYYYYMMDD(frm1.txtMoveDt.Text, parent.gDateFormat, "-") > UniConvDateToYYYYMMDD(frm1.txtExpiryDt.Text, parent.gDateFormat, "-") Then
			Call DisplayMsgBox("970023", "x", frm1.txtExpiryDt.Alt, frm1.txtMoveDt.Alt)
			Call changeTabs(TAB1)
			frm1.txtExpiryDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If

	If Len(Trim(frm1.txtMExpiryDt.Text)) Then
		If UniConvDateToYYYYMMDD(frm1.txtMoveDt.Text, parent.gDateFormat, "-") > UniConvDateToYYYYMMDD(frm1.txtMExpiryDt.Text, parent.gDateFormat, "-") Then
			Call DisplayMsgBox("970023", "x", frm1.txtMExpiryDt.Alt, frm1.txtMoveDt.Alt)
			Call changeTabs(TAB1)
			frm1.txtMExpiryDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If


	If UNICDbl(frm1.txtDocAmt.text) <= 0 Then
		Call DisplayMsgBox("970023", "x", "개설금액","0")
		Call changeTabs(TAB1)
		frm1.txtDocAmt.focus 
		Set gActiveElement = document.activeElement 
		Exit Function  
	End If	
		
	If UNICDbl(frm1.txtXchRate.text) <= 0 Then
		Call DisplayMsgBox("970023", "x", "환율","0")
		Call changeTabs(TAB1)
		frm1.txtXchRate.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If			
		
	 '------ Save function call area ------ 
	Call DbSave															 '☜: Save db data 
		
	FncSave = True														 '⊙: Processing is OK 
End Function


'========================================================================================================
Function FncCopy()
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, , "x", "x")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	lgIntFlgMode = parent.OPMD_CMODE											'⊙: Indicates that current mode is Crate mode

	 '------ 조건부 필드를 삭제한다. ------ 
	Call ggoOper.ClearField(Document, "1")								'⊙: Clear Condition Field
	Call ggoOper.LockField(Document, "N")								'⊙: This function lock the suitable field
End Function


'========================================================================================================
Function FncCancel() 
	On Error Resume Next												
End Function


'========================================================================================================
Function FncInsertRow()
	On Error Resume Next												
End Function

'========================================================================================================
Function FncDeleteRow()
	On Error Resume Next												
End Function


'========================================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================================
Function FncPrev() 
    Dim strVal
	    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002" , "x", "x", "x")  '☜ 바뀐부분 
        Exit Function
    End If
					
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If

	frm1.txtPrevNext.value = "PREV"

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo1.value)				'☆: 조회 조건 데이타 
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		'☆: 조회 조건 데이타 
	         
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
					
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If

	frm1.txtPrevNext.value = "NEXT"

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo1.value)				'☆: 조회 조건 데이타 
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)		'☆: 조회 조건 데이타 
	         
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
Function DbQuery()
	Err.Clear															

	DbQuery = False														'⊙: Processing is NG
		
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If
		
	Dim strVal

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)			'☆: 조회 조건 데이타 
		
	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True														'⊙: Processing is NG
End Function
	

'========================================================================================================
Function DbSave()
	Err.Clear															

	DbSave = False														'⊙: Processing is NG

	If frm1.chkSONoFlg.checked = True Then
		frm1.txtSONoFlg.value = "Y"
	Else
		frm1.txtSONoFlg.value = "N"
	End IF	
	
	Dim strVal
				
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If

		
	With frm1
		.txtMode.value = parent.UID_M0002										'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
		.txtInsrtUserId.value = parent.gUsrID
						
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

	DbSave = True														'⊙: Processing is NG
End Function
	

'========================================================================================================
Function DbDelete()
	Err.Clear															

	DbDelete = False													'⊙: Processing is NG

	Dim strVal

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003					'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo1.value)			'☜: 삭제 조건 데이타 
	strVal = strVal & "&txtSONo=" & Trim(frm1.txtSONo.value)			'☜: 삭제 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 

	DbDelete = True														'⊙: Processing is NG
End Function
	

'========================================================================================================
Function DbQueryOk()													 '☆: 조회 성공후 실행로직 
	 '------ Reset variables area ------ 
	lgIntFlgMode = parent.OPMD_UMODE											 '⊙: Indicates that current mode is Update mode 
	lgBlnFlgChgValue = False

	Call ggoOper.LockField(Document, "Q")								 '⊙: This function lock the suitable field 
		
	Call SetToolBar("111110001101111")
		
	frm1.txtLCNo.focus
				
	'If gSelframeFlg <> TAB1 Then
		Call ClickTab1()
	'End If
End Function

'========================================================================================================
Function SOQueryOK()													 '☆: 조회 성공후 실행로직 
	Call ProtectXchRate()
	Call SetToolBar("11101000000011")	
End Function
	
'========================================================================================================
Function DbSaveOk()														'☆: 저장 성공후 실행 로직 
	Call InitVariables
	Call MainQuery()
End Function
	

'========================================================================================================
Function DbDeleteOk()													'☆: 삭제 성공후 실행 로직 
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
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
									<TR>
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Local L/C 정보</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>구비서류 및 기타</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>					
							<TD WIDTH=* align=right><A href="vbscript:OpenSORef">수주참조</A>&nbsp;|&nbsp;<A href="vbscript:OpenDNRef">출하참조</A></TD>							
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
											<TD CLASS=TD5 NOWRAP>LOCAL L/C관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCNo" SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="LOCAL L/C관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCNo" ALIGN=top TYPE="BUTTON" ONCLICK ="vbscript:btnLCNoOnClick()"></TD>
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
											<TD CLASS=TD5 NOWRAP>LOCAL L/C관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCNo1" ALT="LOCAL LC관리번호" TYPE=TEXT MAXLENGTH=18 SIZE=20 TAG="25XXXU"></TD>
										    <TD CLASS=TD5 NOWRAP>Master L/C관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMLCNo1" ALT="Master L/C관리번호" TYPE=TEXT MAXLENGTH=18 SIZE=20 TAG="25XXXU"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>LOCAL L/C번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LOCAL L/C번호" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="22XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>Master L/C번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMLCDocNo" ALT="Master L/C번호" TYPE=TEXT MAXLENGTH=35 SIZE=20 TAG="21XXXU">&nbsp;-&nbsp;<INPUT NAME="txtMLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>									
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>LOCAL L/C유형</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCType" SIZE=10 MAXLENGTH=5 TAG="22XXXU" ALT="LOCAL L/C유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnLCTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtLCTypeNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>Master L/C유형</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMLCType" SIZE=10 MAXLENGTH=5 TAG="21XXXU" ALT="Master L/C유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnMLCTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtMLCTypeNm" SIZE=20 TAG="24"></TD>		
										</TR>
									    <TR>   	
											<TD CLASS=TD5 NOWRAP>수주번호</TD>
											<TD CLASS=TD6 NOWRAP>
												<INPUT NAME="txtSONo" ALT="수주번호" TYPE=TEXT MAXLENGTH=18 SIZE=20 TAG="24XXXU">
												<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="25X" VALUE="Y" NAME="chkSONoFlg" ID="chkSONoFlg">
												<LABEL FOR="chkSONoFlg">수주번호지정</LABEL>
											</TD>
											<TD CLASS=TD5 NOWRAP>통지번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAdvNo" ALT="통지번호" TYPE=TEXT MAXLENGTH=35 SIZE=30 TAG="21XXXU"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>유효일</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtExpiryDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22X1" ALT="유효일"> </OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS=TD5 NOWRAP>Master L/C유효일</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtMExpiryDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="21X1" ALT="Master L/C유효일"> </OBJECT>');</SCRIPT>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>추심의뢰은행</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFromBank" SIZE=10 MAXLENGTH=10 TAG="22XXXU" ALT="추심의뢰은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromBank" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnFromBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtFromBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>통지일</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtAdvDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22X1" ALT="통지일"> </OBJECT>');</SCRIPT>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>개설은행</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10 MAXLENGTH=10 TAG="22XXXU" ALT="개설은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenBank" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnOpenBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>개설일</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtOpenDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22X1" ALT="개설일"> </OBJECT>');</SCRIPT>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>개설금액</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtDocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="22X2Z" ALT="개설금액"></OBJECT>');</SCRIPT></TD>
														<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐"></TD>
													</TR>
												</TABLE>	
											</TD>
											<TD CLASS=TD5 NOWRAP>개설자국금액</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="22X2Z" ALT="개설자국금액"></OBJECT>');</SCRIPT></TD>
														<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="자국화폐"></TD>
													</TR>
												</TABLE>
											</TD>	
										</TR>
										<TR>	
											<TD CLASS=TD5 NOWRAP>환율</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name=txtXchRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="22X5Z" ALT="환율"> </OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS=TD5 NOWRAP>물품인도기일</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtMoveDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22X1" ALT="물품인도기일"> </OBJECT>');</SCRIPT>
											</TD>															
										</TR>
										<TR>
											<TD CLASS=TD5>결제방법</TD>
											<TD CLASS=TD6>
												<INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=5 TAG="22XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnPayTermsOnClick()">&nbsp;
												<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="14">
											</TD>
											<TD CLASS=TD5 NOWRAP>분할인도여부</TD>
											<TD CLASS=TD6 NOWRAP><TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=0><TR><TD WIDTH=30%><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="2X" VALUE="Y" CHECKED ID="rdoPartailShip1"><LABEL FOR="rdoPartailShip1">Y</LABEL>&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="2X" VALUE="N" ID="rdoPartailShip2"><LABEL FOR="rdoPartailShip2">N</LABEL></TD></TR></TABLE></TD>									
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>결제기간</TD>
											<TD CLASS=TD6 NOWRAP>
											
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtPayDur" ALT="결제기간" style="HEIGHT: 20px; WIDTH: 50px" tag="21X6Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;<LABEL>일</LABEL></TD>
											<TD CLASS=TD5 NOWRAP>선통지참조</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRef" ALT="선통지참조사항" TYPE=TEXT MAXLENGTH=120 SIZE=30 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>개설신청인</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="개설신청인">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>									
											<TD CLASS=TD5 NOWRAP>수혜자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수혜자">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>AMEND일</TD>
											<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime name=txtAmendDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24" ALT="AMEND일"> </OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS=TD5 NOWRAP>영업그룹</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="영업그룹">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
										</TR>
										<%Call SubFillRemBodyTD5656(5)%>
									</TABLE>
								</DIV>
			
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>서류제시기간</TD>
											<TD CLASS=TD656 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtFileDt" style="HEIGHT: 20px; WIDTH: 50px" tag="21X7Z" ALT="서류제시기간" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;<LABEL>일</LABEL>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>서류제시기간 참조</TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtFileDtTxt" ALT="서류제시기간 참조" TYPE=TEXT MAXLENGTH=35 SIZE=70 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>구비서류</TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtDoc1" TYPE=TEXT MAXLENGTH=65 SIZE=70 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtDoc2" TYPE=TEXT MAXLENGTH=65 SIZE=70 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtDoc3" TYPE=TEXT MAXLENGTH=65 SIZE=70 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtDoc4" TYPE=TEXT MAXLENGTH=65 SIZE=70 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtDoc5" TYPE=TEXT MAXLENGTH=65 SIZE=70 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>개설은행앞 정보</TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtBankTxt" ALT="개설은행앞 정보" TYPE=TEXT MAXLENGTH=35 SIZE=70 TAG="21X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>기타참조</TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtEtcRef" ALT="기타참조사항" TYPE=TEXT MAXLENGTH=35 SIZE=70 TAG="21X"></TD>
										</TR>
										<%Call SubFillRemBodyTD656(12)%>
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
						<TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck(0)">LOCAL L/C내역등록</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(1)">AMEND등록</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(2)">판매경비등록</A></TD>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = -1></IFRAME></TD>
			</TR>
		</TABLE>
		<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtHLCNo" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtIncoTerms" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtPrevNext" TAG="24" TABINDEX = -1>
		<INPUT TYPE=HIDDEN NAME="txtSONoFlg" TAG="24" TABINDEX = -1>  
		<INPUT TYPE=HIDDEN NAME="txtExchRateOp" TAG="24" TABINDEX = -1>  
	</FORM>
	<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
		<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
	</DIV>
</BODY>
</HTML>
