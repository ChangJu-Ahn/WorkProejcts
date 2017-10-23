
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3211ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open L/C Header 등록 ASP													*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/03/20																*
'*  8. Modified date(Last)  : 2003/05/19																*
'*  9. Modifier (First)     : Sun-joung Lee																*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/03/22		: Coding Start											*
'********************************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'********************************************  1.1 Inc 선언  ********************************************
-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--
'============================================  1.1.2 공통 Include  ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 

<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBS">
	Option Explicit
	
<!--
'********************************************  1.2 Global 변수/상수 선언  *******************************
'*	Description : 1. Constant는 반드시 대문자 표기														*
'********************************************************************************************************
-->

	Const BIZ_PGM_QRY_ID = "m3211mb1.asp"
	Const BIZ_PGM_SAVE_ID = "m3211mb2.asp"
	Const BIZ_PGM_DEL_ID = "m3211mb3.asp"
	Const BIZ_PGM_POQRY_ID = "m3211mb4.asp"
	Const LC_DETAIL_ENTRY_ID = "m3212ma1"
	Const LCAMEND_HDR_ENTRY_ID = "m3221ma1"
	Const CHARGE_HDR_ENTRY_ID = "m6111ma2"
	
	Const TAB1 = 1
	Const TAB2 = 2
	Const TAB3 = 3
	
	Const gstrLCTypeMajor 		= "S9000"				'신용장유형 
	Const gstrTransportMajor 	= "B9009"				'운송방법 
	Const gstrFreightMajor 		= "S9007"				'운임지불여부 
	Const gstrCreditCoreMajor 	= "S9003"				'신용공여주체 
	Const gstrLoadingPortMajor	= "B9092"				'선적항	
	Const gstrDisChgePortMajor	= "B9092"				'도착항 
	Const gstrFundTypeMajor		= "S9005"				'자금종류 
	Const gstrOriginMajor		= "B9094"				'원산지 
	Const gstrDeliveryPlceMajor	= "B9095"				'인도장소 
	
	Dim lgBlnFlgChgValue
	Dim lgIntGrpCount
	Dim lgIntFlgMode

	Dim gSelframeFlg
	Dim gblnWinEvent
	DIM SCheck 
	
	Dim serverDate
	Dim iCurDate

	serverDate = "<%=GetSvrDate%>"
	iCurDate = UniConvDateAToB(serverDate, Parent.gServerDateFormat, Parent.gDateFormat)

<!--
'==========================================  2.1.1 InitVariables()  =====================================
-->
Function InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE
	lgBlnFlgChgValue = False
	lgIntGrpCount = 0
		
	gblnWinEvent = False
End Function
	
<!--
'==========================================  2.2.1 SetDefaultVal()  =====================================
-->
Sub SetDefaultVal()		
	Call SetToolbar("1110000000001111")    
	frm1.txtReqDt.text = iCurDate
	frm1.txtExpiryDt.text = iCurDate
	frm1.txtLatestShipDt.text = iCurDate
	Call ClickTab1()
	frm1.txtLCNo.focus	
	Set gActiveElement = document.activeElement
End Sub
	
<!--
'==========================================  2.2.2 LoadInfTB19029()  ====================================
-->
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub
	
<!--
'===========================================  2.3.1 Tab Click 처리  =====================================
-->

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

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenOrigin()  +++++++++++++++++++++++++++++++++++++
-->
Function OpenOrigin()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "원산지"
	arrParam(1) = "B_Minor"
	arrParam(2) = Trim(frm1.txtOrigin.Value)
	'arrParam(3) = Trim(frm1.txtOriginNm.Value)
	arrParam(4) = "MAJOR_CD= " & FilterVar(gstrOriginMajor, "''", "S") & ""
	arrParam(5) = "원산지"

	arrField(0) = "Minor_CD"
	arrField(1) = "Minor_NM"

	arrHeader(0) = "원산지"
	arrHeader(1) = "원산지명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtOrigin.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtOrigin.Value = arrRet(0)
		frm1.txtOriginNm.Value = arrRet(1)
		lgBlnFlgChgValue = true	
		frm1.txtOrigin.focus
	End If
	Set gActiveElement = document.activeElement
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenDischgePort()  +++++++++++++++++++++++++++++++++++++
-->
Function OpenDischgePort()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "도착항"
	arrParam(1) = "B_Minor"
	arrParam(2) = Trim(frm1.txtDischgePort.Value)
	'arrParam(3) = Trim(frm1.txtDischgePortNm.Value)
	arrParam(4) = "MAJOR_CD= " & FilterVar(gstrDisChgePortMajor, "''", "S") & ""
	arrParam(5) = "도착항"

	arrField(0) = "Minor_CD"
	arrField(1) = "Minor_NM"

	arrHeader(0) = "도착항"
	arrHeader(1) = "도착항명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtdischgePort.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		lgBlnFlgChgValue = True	
		frm1.txtdischgePort.Value = arrRet(0)
		frm1.txtDischgePortNm.Value = arrRet(1)
		frm1.txtdischgePort.focus
	End If
	Set gActiveElement = document.activeElement
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenBankPop()  ++++++++++++++++++++++++++++++++++++++++
-->
Function OpenBankPop(strBankCd, strBankNm, strPopPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "은행"	
	arrParam(1) = "B_BANK"
	arrParam(2) = Trim(strBankCd)
'		arrParam(3) = Trim(strBankNm)
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

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenLCNoPop()  ++++++++++++++++++++++++++++++++++++++++
-->
Function OpenLCNoPop()
	Dim strRet,IntRetCD
	Dim iCalledAspName
		
	If gblnWinEvent = True Or UCase(frm1.txtLCNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("M3211PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3211PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
		frm1.txtLCNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtLCNo.value = strRet
		frm1.txtLCNo.focus
	End If	
	Set gActiveElement = document.activeElement
End Function


<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenPORef()  ++++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenPORef()																					+
'+	Description : P/O Reference Window Call																+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenPORef()
	Dim strRet,IntRetCD		
	Dim iCalledAspName 	
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "X", "X", "X")
		Exit function
	End If
		
	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("M3111RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111RA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	gblnWinEvent = False

	If strRet(0) = "" Then
		Call ClickTab1()
		frm1.txtLCDocNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		Call SetPORef(strRet)	
		frm1.txtLCDocNo.focus		    
	End If
	Set gActiveElement = document.activeElement
End Function
	
<!--
'++++++++++++++++++++++++++++++++++++++++++++  OpenNotifyParty()  ++++++++++++++++++++++++++++++++++++++++
-->
Function OpenNotifyParty()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "통지처"
	arrParam(1) = "B_BIZ_PARTNER"
	arrParam(2) = frm1.txtNotifyParty.value
'		arrParam(3) = Trim(strBizPartnerNM)
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "통지처"

	arrField(0) = "BP_CD"
	arrField(1) = "BP_NM"

	arrHeader(0) = "통지처"
	arrHeader(1) = "통지처명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtNotifyParty.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		lgBlnFlgChgValue = True
		frm1.txtNotifyParty.value = arrRet(0)
		frm1.txtNotifyPartyNm.value = arrRet(1)
		frm1.txtNotifyParty.focus
		Set gActiveElement = document.activeElement
	End If
		
End Function
<!--
'++++++++++++++++++++++++++++++++++++++++++++  OpenBizPartner()  ++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenBizPartner()																				+
'+	Description : Business Partner PopUp Window Call													+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenBizPartner(strBizPartnerCD, strBizPartnerNM, strPopPos, strPopPosNm)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = strPopPos
	arrParam(1) = "B_BIZ_PARTNER"
	arrParam(2) = Trim(strBizPartnerCD)
	arrParam(4) = ""
	arrParam(5) = strPopPos

	arrField(0) = "BP_CD"
	arrField(1) = "BP_NM"

	arrHeader(0) = strPopPos
	arrHeader(1) = strPopPosNm

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizPartner(strPopPos, arrRet)
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenCountry()  +++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenCountry()																				+
'+	Description : Country PopUp Window Call																+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenCountry()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "국가"
	arrParam(1) = "B_COUNTRY"
	arrParam(2) = Trim(frm1.txtOriginCntry.value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "국가"

	arrField(0) = "COUNTRY_CD"
	arrField(1) = "COUNTRY_NM"

	arrHeader(0) = "국가"
	arrHeader(1) = "국가명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtOriginCntry.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		lgBlnFlgChgValue = True
		frm1.txtOriginCntry.Value = arrRet(0)
		frm1.txtOriginCntry.focus
		Set gActiveElement = document.activeElement
	End If
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++  OpenPurGrp()  +++++++++++++++++++++++++++++++++++++++
'+	Name : OpenPurGrp()																					+
'+	Description : Sales Group PopUp Window Call															+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "수입담당"
	arrParam(1) = "B_PURCHASE_GROUP"
	arrParam(2) = Trim(frm1.txtPurGrp.value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "수입담당"

	arrField(0) = "PUR_GRP"
	arrField(1) = "PUR_GRP_NM"

	arrHeader(0) = "수입담당"
	arrHeader(1) = "수입담당명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtPurGrp.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		lgBlnFlgChgValue = True
		frm1.txtPurGrp.value = arrRet(0)
		frm1.txtPurGrpNm.value = arrRet(1)
		frm1.txtPurGrp.focus
		Set gActiveElement = document.activeElement
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenMinorCd()  +++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenMinorCd()																				+
'+	Description : Minor Code PopUp Window Call															+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strPopPosNm, strMajorCd)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = strPopPos
	arrParam(1) = "B_Minor"
	arrParam(2) = Trim(strMinorCD)
	'arrParam(3) = Trim(strMinorNM)
	arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""
	arrParam(5) = strPopPos

	arrField(0) = "Minor_CD"
	arrField(1) = "Minor_NM"

	arrHeader(0) = strPopPos
	arrHeader(1) = strPopPosNm

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMinorCd(strMajorCd, arrRet)
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenPayType()  +++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenPayType()																				+
'+	Description : PayType Code PopUp Window Call															+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenPayType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Or UCase(frm1.txtPayTerms.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	gblnWinEvent = True
		
	arrParam(0) = "결제방법"					
	arrParam(1) = "B_Minor,b_configuration"		
	arrParam(2) = Trim(frm1.txtPayTerms.Value)		
	arrParam(4) = "b_minor.Major_Cd=" & FilterVar("B9004", "''", "S") & "" _
						& " and b_minor.minor_cd=b_configuration.minor_cd" _
						& " AND b_configuration.REFERENCE = " & FilterVar("M", "''", "S") & " "
							
	arrParam(5) = "결제방법"					
		
	arrField(0) = "b_minor.Minor_Cd"			
	arrField(1) = "b_minor.Minor_Nm"			
	arrField(2) = "b_configuration.REFERENCE"
	    
	arrHeader(0) = "결제방법"				
	arrHeader(1) = "결제방법명"				
	arrHeader(2) = "Reference"
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtPayTerms.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPayTerms.Value = arrRet(0)
		frm1.txtPayTermsNm.Value = arrRet(1)
		lgBlnFlgChgValue 		= True
		frm1.txtPayTerms.focus
		Set gActiveElement = document.activeElement
	End If	
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  btnPayTermsOnClick()  +++++++++++++++++++++++++++++++++++++++++
'+	Name : btnPayTermsOnClick()																				+
'+	Description : Call other function															+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Sub btnPayTermsOnClick()
	Call OpenPayType()
End Sub

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetPORef()  ++++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetPORef()																					+
'+	Description : Set Return array from S/O Reference Window											+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function SetPORef(strRet)
	Dim strVal

	Call ggoOper.ClearField(Document, "A")
	Call SetRadio()
	Call SetDefaultVal

	frm1.txtPONo.value = strRet(0)
	frm1.hdnXchRtOp.value = strRet(1)
		
	If LayerShowHide(1) = False Then
	    Exit Function
	End If

	strVal = BIZ_PGM_POQRY_ID & "?txtPONo=" & Trim(frm1.txtPONo.value)
    strVal = strVal & "&txtCurrency=" & Trim(Parent.gCurrency)
        
	Call RunMyBizASP(MyBizASP, strVal)
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetBank()  +++++++++++++++++++++++++++++++++++++++++++++
-->
Function SetBank(strPopPos, arrRet)
	Select Case UCase(strPopPos)
		Case "ADVBANK"
			frm1.txtAdvBank.Value = arrRet(0)
			frm1.txtAdvBankNm.Value = arrRet(1)
			frm1.txtAdvBank.focus
				
		Case "OPENBANK"
			frm1.txtOpenBank.Value = arrRet(0)
			frm1.txtOpenBankNm.Value = arrRet(1)
			frm1.txtOpenBank.focus
				
		Case "PAYBANK"
			frm1.txtPayBank.Value = arrRet(0)
			frm1.txtPayBankNm.Value = arrRet(1)
			frm1.txtPayBank.focus
				
		Case "RENEGOBANK"
			frm1.txtRenegoBank.Value = arrRet(0)
			frm1.txtRenegoBankNm.Value = arrRet(1)
			frm1.txtRenegoBank.focus
				
		Case "CONFIRMBANK"
			frm1.txtConfirmBank.Value = arrRet(0)
			frm1.txtConfirmBankNm.Value = arrRet(1)
			frm1.txtConfirmBank.focus
				
		Case Else
	End Select
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = True
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++  SetBizPartner()  ++++++++++++++++++++++++++++++++++++++++++
-->
Function SetBizPartner(strPopPos, arrRet)
	Select Case (strPopPos)
		Case "통지처"
			frm1.txtNotifyParty.Value = arrRet(0)
			frm1.txtNotifyPartyNm.Value = arrRet(1)
			frm1.txtNotifyParty.focus
				
		Case "대행자"
			frm1.txtAgent.Value = arrRet(0)
			frm1.txtAgentNm.Value = arrRet(1)
			frm1.txtAgent.focus
				
		Case "제조자"
			frm1.txtManufacturer.Value = arrRet(0)
			frm1.txtManufacturerNm.Value = arrRet(1)
			frm1.txtManufacturer.focus
				
		Case Else
	End Select
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = True
End Function
	
<!--
'+++++++++++++++++++++++++++++++++++++++++++++  SetMinorCd()  +++++++++++++++++++++++++++++++++++++++++++
-->
Function SetMinorCd(strMajorCd, arrRet)
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
				
		Case gstrDeliveryPlceMajor
			frm1.txtDeliveryPlce.Value = arrRet(0)
			frm1.txtDeliveryPlceNm.Value = arrRet(1)
			frm1.txtDeliveryPlce.focus

		Case gstrLoadingPortMajor
			frm1.txtLoadingPort.Value = arrRet(0)
			frm1.txtLoadingPortNm.Value = arrRet(1)
			frm1.txtLoadingPort.focus
				
		Case gstrFundTypeMajor
			frm1.txtfundType.Value = arrRet(0)
			frm1.txtfundTypeNm.Value = arrRet(1)
			frm1.txtfundType.focus
		Case Else
	End Select
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = True
End Function


'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

ggoOper.FormatFieldByObjectOfCur frm1.txtDocAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
ggoOper.FormatFieldByObjectOfCur frm1.txtXchRate, frm1.txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		

End Sub
'===================================== changeTag()  ================================================
Function changeTag()

with frm1

if Trim(frm1.txtLCAmendSeq.value) <> "" And Trim(frm1.txtLCAmendSeq.value) <> "0" then
	'tab1
	ggoOper.SetReqAttr	.txtLCDocNo, "Q"
	ggoOper.SetReqAttr	.chkPoNoCnt, "Q"
	ggoOper.SetReqAttr	.txtLCType, "Q"
	ggoOper.SetReqAttr	.txtReqDt, "Q"
	ggoOper.SetReqAttr	.txtOpenDt, "Q"
	ggoOper.SetReqAttr	.txtAdvBank, "Q"
	ggoOper.SetReqAttr	.txtOpenBank, "Q"
	ggoOper.SetReqAttr	.txtExpiryDt, "Q"
	ggoOper.SetReqAttr	.txtLatestShipDt, "Q"
	ggoOper.SetReqAttr	.txtShipment, "Q"
	ggoOper.SetReqAttr	.txtXchRate, "Q"
	ggoOper.SetReqAttr	.txtDocAmt, "Q"
	ggoOper.SetReqAttr	.txtTransport, "Q"
	ggoOper.SetReqAttr	.txtPaytermstxt, "Q"
	ggoOper.SetReqAttr	.txtDeliveryPlce, "Q"
	ggoOper.SetReqAttr	.txttolerance, "Q"
	ggoOper.SetReqAttr	.txtLoadingPort, "Q"
	ggoOper.SetReqAttr	.txtDischgePort, "Q"
	ggoOper.SetReqAttr	.txtOrigin, "Q"
	ggoOper.SetReqAttr	.rdoPartailShip1, "Q"
	ggoOper.SetReqAttr	.rdoPartailShip2, "Q"
	ggoOper.SetReqAttr	.rdoTranshipment1, "Q"
	ggoOper.SetReqAttr	.rdoTranshipment2, "Q"
	ggoOper.SetReqAttr	.rdoChargeCd1, "Q"
	ggoOper.SetReqAttr	.rdoChargeCd2, "Q"
	ggoOper.SetReqAttr	.txtPayTerms, "Q"
	ggoOper.SetReqAttr	.txtPayDur, "Q"
	'tab2
	ggoOper.SetReqAttr	.txtFileDt, "Q"
	ggoOper.SetReqAttr	.txtFileDtTxt, "Q"
	ggoOper.SetReqAttr	.txtInvCnt, "Q"
	ggoOper.SetReqAttr	.txtPackList, "Q"
	ggoOper.SetReqAttr	.rdoBLAwFlg1, "Q"
	ggoOper.SetReqAttr	.rdoBLAwFlg2, "Q"
	ggoOper.SetReqAttr	.txtFreight, "Q"
	ggoOper.SetReqAttr	.txtNotifyParty, "Q"
	ggoOper.SetReqAttr	.txtConsignee, "Q"
	ggoOper.SetReqAttr	.chkCertOriginFlg, "Q"
	ggoOper.SetReqAttr	.rdoTransfer1, "Q"
	ggoOper.SetReqAttr	.rdoTransfer2, "Q"
	ggoOper.SetReqAttr	.txtOriginCntry, "Q"
	ggoOper.SetReqAttr	.txtInsurPolicy, "Q"
	ggoOper.SetReqAttr	.txtDoc1, "Q"
	ggoOper.SetReqAttr	.txtDoc2, "Q"
	ggoOper.SetReqAttr	.txtDoc3, "Q"
	ggoOper.SetReqAttr	.txtDoc4, "Q"
	ggoOper.SetReqAttr	.txtDoc5, "Q"
	'tab3
	ggoOper.SetReqAttr	.txtPayBank, "Q"
	ggoOper.SetReqAttr	.txtRenegoBank, "Q"
	ggoOper.SetReqAttr	.txtConfirmBank, "Q"
	ggoOper.SetReqAttr	.txtBankTxt, "Q"
	ggoOper.SetReqAttr	.txtCreditCore, "Q"
	ggoOper.SetReqAttr	.txtfundType, "Q"
	ggoOper.SetReqAttr	.txtLmtXchRate, "Q"
	ggoOper.SetReqAttr	.txtChargeTxt, "Q"
	ggoOper.SetReqAttr	.txtLmtAmt, "Q"
	ggoOper.SetReqAttr	.txtTransportComp, "Q"
	ggoOper.SetReqAttr	.txtAgent, "Q"
	ggoOper.SetReqAttr	.txtManufacturer, "Q"
	ggoOper.SetReqAttr	.txtAdvDt, "Q"
	ggoOper.SetReqAttr	.txtRemark, "Q"
	Call SetToolbar("1110000000001111")
else
	
	Call changeLcDocNo()
	Call changeLcDocNoSet()
	Call setDefaultTag()
	Call changeXchRtTag()
end if
end with

End Function

Function changeXchRtTag()
	
if Trim(frm1.txtCurrency.value) <> Parent.gCurrency then
	Call ggoOper.SetReqAttr(frm1.txtXchRate, "N")
else
	Call ggoOper.SetReqAttr(frm1.txtXchRate, "Q")
	frm1.txtXchRate.Text = "1"
end if
	
End Function


'================================= setDefaultTag()  ============================================
sub setDefaultTag()
with frm1
	'tab1
	ggoOper.SetReqAttr	.chkPoNoCnt, "D"
	ggoOper.SetReqAttr	.txtLCType, "N"
	ggoOper.SetReqAttr	.txtReqDt, "N"
	ggoOper.SetReqAttr	.txtAdvBank, "N"
	ggoOper.SetReqAttr	.txtOpenBank, "N"
	ggoOper.SetReqAttr	.txtExpiryDt, "N"
	ggoOper.SetReqAttr	.txtLatestShipDt, "N"
	ggoOper.SetReqAttr	.txtShipment, "D"
	ggoOper.SetReqAttr	.txtXchRate, "N"
	ggoOper.SetReqAttr	.txtDocAmt, "N"
	ggoOper.SetReqAttr	.txtTransport, "N"
	ggoOper.SetReqAttr	.txtPaytermstxt, "D"
	ggoOper.SetReqAttr	.txtDeliveryPlce, "N"
	ggoOper.SetReqAttr	.txttolerance, "D"
	ggoOper.SetReqAttr	.txtLoadingPort, "N"
	ggoOper.SetReqAttr	.txtDischgePort, "N"
	ggoOper.SetReqAttr	.txtOrigin, "D"
	ggoOper.SetReqAttr	.rdoPartailShip1, "D"
	ggoOper.SetReqAttr	.rdoPartailShip2, "D"
	ggoOper.SetReqAttr	.rdoTranshipment1, "D"
	ggoOper.SetReqAttr	.rdoTranshipment2, "D"
	ggoOper.SetReqAttr	.rdoChargeCd1, "D"
	ggoOper.SetReqAttr	.rdoChargeCd2, "D"
	ggoOper.SetReqAttr	.txtPayDur, "N"
	'tab2
	ggoOper.SetReqAttr	.txtFileDt, "D"
	ggoOper.SetReqAttr	.txtFileDtTxt, "D"
	ggoOper.SetReqAttr	.txtInvCnt, "D"
	ggoOper.SetReqAttr	.txtPackList, "D"
	ggoOper.SetReqAttr	.rdoBLAwFlg1, "D"
	ggoOper.SetReqAttr	.rdoBLAwFlg2, "D"
	ggoOper.SetReqAttr	.txtFreight, "D"
	ggoOper.SetReqAttr	.txtNotifyParty, "D"
	ggoOper.SetReqAttr	.txtConsignee, "D"
	ggoOper.SetReqAttr	.chkCertOriginFlg, "D"
	ggoOper.SetReqAttr	.rdoTransfer1, "D"
	ggoOper.SetReqAttr	.rdoTransfer2, "D"
	ggoOper.SetReqAttr	.txtOriginCntry, "D"
	ggoOper.SetReqAttr	.txtInsurPolicy, "D"
	ggoOper.SetReqAttr	.txtDoc1, "D"
	ggoOper.SetReqAttr	.txtDoc2, "D"
	ggoOper.SetReqAttr	.txtDoc3, "D"
	ggoOper.SetReqAttr	.txtDoc4, "D"
	ggoOper.SetReqAttr	.txtDoc5, "D"
	'tab3
	ggoOper.SetReqAttr	.txtPayBank, "D"
	ggoOper.SetReqAttr	.txtRenegoBank, "D"
	ggoOper.SetReqAttr	.txtConfirmBank, "D"
	ggoOper.SetReqAttr	.txtBankTxt, "D"
	ggoOper.SetReqAttr	.txtCreditCore, "D"
	ggoOper.SetReqAttr	.txtfundType, "D"
	ggoOper.SetReqAttr	.txtLmtXchRate, "D"
	ggoOper.SetReqAttr	.txtChargeTxt, "D"
	ggoOper.SetReqAttr	.txtLmtAmt, "D"
	ggoOper.SetReqAttr	.txtTransportComp, "D"
	ggoOper.SetReqAttr	.txtAgent, "D"
	ggoOper.SetReqAttr	.txtManufacturer, "D"
	ggoOper.SetReqAttr	.txtAdvDt, "D"
	ggoOper.SetReqAttr	.txtRemark, "D"
End with
End Sub
<!--
'============================================ ValidDateCheckLocal()  ====================================
-->
Function ValidDateCheckLocal(pObjFromDt, pObjToDt)

	ValidDateCheckLocal = False

	If Len(Trim(pObjToDt.Text)) And Len(Trim(pObjFromDt.Text)) Then

		If UniConvDateToYYYYMMDD(pObjFromDt.Text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(pObjToDt.Text,Parent.gDateFormat,"") Then
			ClickTab1()
			Call DisplayMsgBox("970023","X", pObjToDt.Alt, pObjFromDt.Alt)
			pObjToDt.Focus
	        Set gActiveElement = document.activeElement                            
				
			Exit Function
		End If

	End If

	ValidDateCheckLocal = True

End Function
<!--
'============================================ changeLcDocNo()  ======================================
-->
Sub changeLcDocNo()

	if Trim(frm1.txtLCDocNo.Value) <> "" or Trim(frm1.txtOpenDt.Text) <> "" then
		ggoOper.SetReqAttr	frm1.txtOpenDt, "N"
		ggoOper.SetReqAttr	frm1.txtLCDocNo, "N"
	else
		ggoOper.SetReqAttr	frm1.txtOpenDt, "D"
		ggoOper.SetReqAttr	frm1.txtLCDocNo, "D"
	end if

End Sub


Sub changeLcDocNoSet()

	if Trim(frm1.txtLCDocNo.Value) <> "" then
		ggoOper.SetReqAttr	frm1.txtLCDocNo, "N"
	else
		ggoOper.SetReqAttr	frm1.txtLCDocNo, "D"
	end if

End Sub
<!--
'============================================ OpenLCwin()  ======================================
'=	Name : OpenLCwin()																					=
'=	Description :
'========================================================================================================
-->
Function OpenLCwin()
	Dim strRet
		
	strRet = window.open("../M_EDI2/Main/main.asp", "", _
			"height=600,width=820,left=50,top=50,status=no,resizable=no,scrollbars=no,toolbar=no,menubar=no,location=no")
		        		
End Function

<!--
'============================================  2.5.1 OpenCookie()  ======================================
-->
Function OpenCookie()
Dim strTemp
	strTemp = ReadCookie("LCNo")				'L/C관리번호 
	frm1.txtLCNo.Value = strTemp
	WriteCookie "LCNo", ""
		
	If Trim(strTemp) <> "" Then
		Call dbquery()
		'Call MainQuery()
	End If
		
End Function

<!--
'=============================================  2.5.1 LoadLCDtl()  ======================================
-->
Function LoadLCDtl()
	Dim strDtlOpenParam
	Dim IntRetCD

    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End if
	    	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    	
	WriteCookie "LCNo", UCase(Trim(frm1.txtLCNo.value))				'L/C관리번호 
		
	PgmJump(LC_DETAIL_ENTRY_ID)

End Function

<!--
'=============================================  2.5.1 LoadLCAmendHdr()  ======================================
-->
Function LoadLCAmendHdr()
	Dim strHdrOpenParam
	Dim IntRetCD
	Dim strVal

    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
		
	if lgBlnFlgChgValue = True	then
		Call DisplayMsgBox("189217", "X", "X", "X")
		Exit function
	End if
    	
    If len(Trim(frm1.txtOpenDt.text)) < 1 Then
        Call DisplayMsgBox("173428", "X", "X", "X")
        Exit Function
    End If
	WriteCookie "LCNo", UCase(Trim(frm1.txtLCNo.value))
	PgmJump(LCAMEND_HDR_ENTRY_ID)

End Function

<!--
'=============================================  2.5.1 LoadChargeHdr()  ======================================
-->
Function LoadChargeHdr()
	
	Dim IntRetCD
		
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
		
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	    	
	WriteCookie "Process_Step", "VL"
	WriteCookie "Po_No", UCase(Trim(frm1.txtLCNo.value))
	WriteCookie "Pur_Grp", UCase(Trim(frm1.txtPurGrp.value))
	
	PgmJump(CHARGE_HDR_ENTRY_ID)
		
End Function
	
<!--
'==============================================  2.5.4 SetRadio()  ======================================
'=	Event Name : SetRadio																				=
'=	Event Desc :																						=
'========================================================================================================
-->
Function SetRadio()
	frm1.rdoPartailShip1.checked = True
	frm1.rdoTranshipment1.checked = True
	frm1.rdoBLAwFlg1.checked = True
	frm1.rdoTransfer1.checked = True
	frm1.rdoChargeCd1.checked = True
End Function
<!--
'==============================================  2.5.4 setAmt()  ======================================
'=	Event Name : setAmt()																				=
'=	Event Desc : 13차 추가																						=
'========================================================================================================
-->
Function setAmt()
    Dim SumTotal
    
    '-- Issue 10239 by Byun Jee Hyun 2005-09-23
    if trim(frm1.txtCurrency.value) <> Trim(Parent.gCurrency) then
		if Trim(frm1.hdnXchRtOp.value) = "/" then
			if UNICDbl(frm1.txtXchRate.text) = 0 then
				SumTotal =  0
			else
				SumTotal = UNICDbl(frm1.txtDocAmt.text) / UNICDbl(frm1.txtXchRate.text)   
				frm1.txtLocAmt.text = UNIFormatNumberByCurrecny(CStr(SumTotal), Parent.gCurrency, Parent.ggAmtOfMoneyNo)
			end if
		else
			SumTotal = UNICDbl(frm1.txtDocAmt.text) * UNICDbl(frm1.txtXchRate.text)
			frm1.txtLocAmt.text = UNIFormatNumberByCurrecny(CStr(SumTotal), Parent.gCurrency, Parent.ggAmtOfMoneyNo)
		end if
    else
		frm1.txtLocAmt.text = frm1.txtDocAmt.text
    end if
 

	'If Trim(frm1.hdnXchRtOp.value) = "*" then
'
 '       SumTotal = UNICDbl(frm1.txtDocAmt.text) * UNICDbl(frm1.txtXchRate.text)
'
 '       frm1.txtLocAmt.text = UNIFormatNumberByCurrecny(CStr(SumTotal), Parent.gCurrency, Parent.ggAmtOfMoneyNo)
'	elseif Trim(frm1.hdnXchRtOp.value) = "/" then
 '       if UNICDbl(frm1.txtXchRate.text) = 0 then
'		    SumTotal =  0
'		else
'		    SumTotal = UNICDbl(frm1.txtDocAmt.text) / UNICDbl(frm1.txtXchRate.text)   
'		end if
'		frm1.txtLocAmt.text = UNIFormatNumberByCurrecny(CStr(SumTotal), Parent.gCurrency, Parent.ggAmtOfMoneyNo)
'	else
 '       frm1.txtLocAmt.text = frm1.txtDocAmt.text
'	End If
				
End Function
<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
Sub Form_Load()
	Call LoadInfTB19029
	Call AppendNumberRange("0","0","99")
	Call AppendNumberRange("1","0","9999")
	'Call AppendNumberRange("2","0","999999999999")
	'Call AppendNumberRange("3","0","99999999")
	Call AppendNumberPlace("7","2","0")
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")
		
   	'Call SetToolbar("1110100000001111")    
	Call SetDefaultVal
	Call InitVariables
	'gSelframeFlg = TAB1
		
	'Call changeTabs(TAB1)
		
	gIsTab     = "Y"
	gTabMaxCnt = 3
	SCheck = TRUE
	Call OpenCookie()
		
End Sub
	
<!--
'=========================================  3.1.2 Form_QueryUnload()  ===================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
	   
End Sub
	
<!--
'========================================================================================================
'=	Event Name : chkPoNoCnt_onpropertychange															=
'========================================================================================================
-->
Sub chkPoNoCnt_onpropertychange()
   lgBlnFlgChgValue = true	
End Sub
<!--
'==========================================================================================
'   Event Name : txtReqDt
'==========================================================================================
-->
Sub txtReqDt_DblClick(Button)
   if Button = 1 then
   	frm1.txtReqDt.Action = 7
   	Call SetFocusToDocument("M")
   	frm1.txtReqDt.focus
   End if
End Sub

Sub txtReqDt_Change()
   lgBlnFlgChgValue = true	
End Sub
<!--
'==========================================================================================
'   Event Name : txtOpenDt
'==========================================================================================
-->
Sub txtOpenDt_DblClick(Button)
   if Button = 1 then
   	frm1.txtOpenDt.Action = 7
   	Call SetFocusToDocument("M")
   	frm1.txtOpenDt.focus
   End if
End Sub

Sub txtOpenDt_Change()
   Call changeLcDocNo()
   lgBlnFlgChgValue = true	

   if Trim(frm1.txtOpenDt.Text) = ""  then
   	Exit Sub
   End if	
	
   IF SCheck = FALSE THEN
   	SCheck = TRUE
   	EXIT Sub
   END IF
	
	if Trim(frm1.txtCurrency.value) <> "" and Trim(frm1.txtCurrency.value) <> Parent.gCurrency  then
   	Call ChangeCurOrDt()
   elseif Trim(frm1.txtCurrency.value) = Parent.gCurrency  then
   	frm1.txtXchRate.text = UNIFormatNumber(1,ggExchRate.DecPoint, -2, 0, Parent.ggExchRate.RndPolicy, Parent.ggExchRate.RndUnit)
   end if	
End Sub
<!--
'==========================================================================================
'   Event Name : txtExpiryDt
'==========================================================================================
-->
Sub txtExpiryDt_DblClick(Button)
   if Button = 1 then
   	frm1.txtExpiryDt.Action = 7
   	Call SetFocusToDocument("M")
   	frm1.txtExpiryDt.focus
   End if
End Sub

<!--
'==========================================================================================
'   Event Name : txtExpiryDt
'   Event Desc :
'==========================================================================================
-->
Sub txtExpiryDt_Change()
   lgBlnFlgChgValue = true	
End Sub
<!--
'==========================================================================================
'   Event Name : txtLatestShipDt
'   Event Desc :
'==========================================================================================
-->
Sub txtLatestShipDt_DblClick(Button)
if Button = 1 then
	frm1.txtLatestShipDt.Action = 7
	Call SetFocusToDocument("M")
	frm1.txtLatestShipDt.focus
End if
End Sub
<!--
'==========================================================================================
'   Event Name : txtAdvDt
'==========================================================================================
-->
Sub txtAdvDt_DblClick(Button)
if Button = 1 then
	frm1.txtAdvDt.Action = 7
	Call SetFocusToDocument("M")
	frm1.txtAdvDt.focus
End if
End Sub

Sub txtAdvDt_Change()
lgBlnFlgChgValue = true	
End Sub
<!--
'==========================================================================================
'   Event Name : txtAmendDt
'==========================================================================================
-->
Sub txtAmendDt_DblClick(Button)
if Button = 1 then
	frm1.txtAmendDt.Action = 7
	Call SetFocusToDocument("M")
	frm1.txtAmendDt.focus
End if
End Sub

Sub txtAmendDt_Change()
lgBlnFlgChgValue = true	
End Sub

<!--
'==========================================================================================
'   Event Name : txtLatestShipDt
'==========================================================================================
-->
Sub txtLatestShipDt_Change()
lgBlnFlgChgValue = true	
End Sub
'=================================  ChangeCurOrDt()  ==================================================
Function ChangeCurOrDt()
   
   Dim strVal
   Err.Clear                                                               '☜: Protect system from crashing
	
   if Trim(frm1.txtCurrency.value) = Parent.gCurrency  then
       frm1.txtXchRate.text = UNIFormatNumber(1,ggExchRate.DecPoint, -2, 0, Parent.ggExchRate.RndPolicy, Parent.ggExchRate.RndUnit)
       Exit Function
   end if


   With frm1
		
   	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & "LookupDailyExRt"			
   	strVal = strVal & "&txtOpenDt=" & Trim(.txtOpenDt.text)	
       strVal = strVal & "&txtCurrency=" & Trim(.txtCurrency.value)	
				
   End With
	
   If  LayerShowHide(1) = False Then
     	Exit Function
   End If

   Call RunMyBizASP(MyBizASP, strVal)
        
End Function

'==============================  Change_Event =================================================
Sub txtDocAmt_Change()
	Call setAmt()
	lgBlnFlgChgValue = true	
End Sub

Sub txtXchRate_Change()
	Call setAmt()
	lgBlnFlgChgValue = true	
End Sub

Sub txtLocAmt_Change()
lgBlnFlgChgValue = true	
End Sub

Sub txtPayDur_Change()
lgBlnFlgChgValue = true	
End Sub

Sub txttolerance_Change()
lgBlnFlgChgValue = true	
End Sub

Sub txtFileDt_Change()
lgBlnFlgChgValue = true	
End Sub

Sub txtInvCnt_Change()
lgBlnFlgChgValue = true	
End Sub

Sub txtPackList_Change()
lgBlnFlgChgValue = true	
End Sub

Sub txtLmtXchRate_Change()
lgBlnFlgChgValue = true	
End Sub

Sub txtLmtAmt_Change()
lgBlnFlgChgValue = true	
End Sub

<!--
'=========================================  3.2.1 btnLCNo_OnClick()  ====================================
-->
Sub btnLCNoOnClick()
	If frm1.txtLCNo.readOnly <> True Then
		Call OpenLCNoPop()
	End If
End Sub

<!--
'========================================  3.2.2 btnAdvBank_OnClick()  ==================================
-->
Sub btnAdvBankOnClick()
	If frm1.txtAdvBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtAdvBank.value,	frm1.txtAdvBankNm.value, "ADVBANK")
	End If
End Sub

<!--
'=======================================  3.2.3 btnOpenBank_OnClick()  ==================================
-->
Sub btnOpenBankOnClick()
	If frm1.txtOpenBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtOpenBank.value, frm1.txtOpenBankNm.value, "OPENBANK")
	End If
End Sub

<!--
'========================================  3.2.4 btnPayBank_OnClick()  ==================================
-->
Sub btnPayBankOnClick()
	If frm1.txtPayBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtPayBank.value, frm1.txtPAyBankNm.value, "PAYBANK")
	End If
End Sub

<!--
'=====================================  3.2.5 btnRenegoBank_OnClick()  ==================================
-->
Sub btnRenegoBankOnClick()
	If frm1.txtRenegoBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtRenegoBank.value, frm1.txtRenegoBankNm.value, "RENEGOBANK")
	End If
End Sub

<!--
'====================================  3.2.6 btnConfirmBankOnClick()  ==================================
-->
Sub btnConfirmBankOnClick()
	If frm1.txtConfirmBank.readOnly <> True Then
		Call OpenBankPop(frm1.txtConfirmBank.value, frm1.txtConfirmBankNm.value, "CONFIRMBANK")
	End If
End Sub

<!--
'===================================  3.2.7 btnNotifyPartyOnClick()  ===================================
-->
Sub btnNotifyPartyOnClick()
	If frm1.txtNotifyParty.readOnly <> True Then
		'Call OpenBizPartner(frm1.txtNotifyParty.value, frm1.txtNotifyPartyNm.value, "통지처","통지처명")
		Call OpenNotifyParty()
	End If
End Sub
<!--
'======================================  3.2.8 btnAgentOnClick()  ======================================
-->
Sub btnAgentOnClick()
	If frm1.txtAgent.readOnly <> True Then
		Call OpenBizPartner(frm1.txtAgent.value, frm1.txtAgentNm.value, "대행자", "대행자명")
	End If
End Sub
<!--
'==================================  3.2.9 btnManufacturerOnClick()  ===================================
-->
Sub btnManufacturerOnClick()
	If frm1.txtManufacturer.readOnly <> True Then
		Call OpenBizPartner(frm1.txtManufacturer.value, frm1.txtManufacturerNm.value, "제조자", "제조자명")
	End If
End Sub
<!--
'===================================  3.2.10 btnOriginCntryOnClick()  ==================================
-->
Sub btnOriginCntryOnClick()
	If frm1.txtOriginCntry.readOnly <> True Then
		Call OpenCountry()
	End If
End Sub
<!--
'===================================  3.2.11 btnPurGrpOnClick()  ===================================
-->
Sub btnPurGrpOnClick()
	If frm1.txtPurGrp.readOnly <> True Then
		Call OpenPurGrp()
	End If
End Sub
<!--
'=====================================  3.2.12 btnLCTypeOnClick()  =====================================
-->
Sub btnLCTypeOnClick()
	If frm1.txtLCType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtLCType.value, frm1.txtLCTypeNm.value, "L/C유형", "L/C유형명", gstrLCTypeMajor)
	End If
End Sub
<!--
'====================================  3.2.13 btnTransportOnClick()  ===================================
-->
Sub btnTransportOnClick()
	If frm1.txtTransport.readOnly <> True Then
		Call OpenMinorCd(frm1.txtTransport.value, frm1.txtTransportNm.value, "운송방법", "운송방법명", gstrTransportMajor)
	End If
End Sub
<!--
'=====================================  3.2.14 btnFreightOnClick()  ====================================
-->
Sub btnFreightOnClick()
	If frm1.txtFreight.readOnly <> True Then
		Call OpenMinorCd(frm1.txtFreight.value, frm1.txtFreightNm.value, "운임지불형태", "운임지불형태명", gstrFreightMajor)
	End If
End Sub
<!--
'===================================  3.2.15 btnCreditCoreOnClick()  ===================================
-->
Sub btnCreditCoreOnClick()
	If frm1.txtCreditCore.readOnly <> True Then
		Call OpenMinorCd(frm1.txtCreditCore.value, frm1.txtCreditCoreNm.value, "신용공여주체", "신용공여주체명", gstrCreditCoreMajor)
	End If
End Sub
<!--
'===================================  3.2.15 btnLoadingPortOnClick()  ===================================
-->
Sub btnLoadingPortOnClick()
	If frm1.txtLoadingPort.readOnly <> True Then
		Call OpenMinorCd(frm1.txtLoadingPort.value, frm1.txtLoadingPortNm.value, "선적항", "선적항명", gstrLoadingPortMajor)
	End If
End Sub
<!--
'===================================  3.2.15 btnTransportOnClick()  ===================================
-->
Sub btnTransportOnClick()
	If frm1.txtTransport.readOnly <> True Then
		Call OpenMinorCd(frm1.txtTransport.value, frm1.txtTransportNm.value, "운송방법", "운송방법명", gstrTransportMajor)
	End If
End Sub
	
<!--
'===================================  3.2.15 btnDeliveryPlceOnClick()  =================================
-->
Sub btnDeliveryPlceOnClick()
	If frm1.txtDeliveryPlce.readOnly <> True Then
		Call OpenMinorCd(frm1.txtDeliveryPlce.value, frm1.txtDeliveryPlce.value, "인도장소", "인도장소명", gstrDeliveryPlceMajor)
	End If
End Sub

<!--
'===================================  3.2.15 btnDischgePortOnClick()  ===================================
-->
Sub btnDischgePortOnClick()
	If frm1.txtDischgePort.readOnly <> True Then
		Call OpenDischgePort()
	End If
End Sub

<!--
'===================================  3.2.15 btnFundTypeOnClick()  ===================================
-->
Sub btnFundTypeOnClick()
	If frm1.txtFundType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtFundType.value, frm1.txtFundTypeNm.value, "자금종류", "자금종류명", gstrFundTypeMajor)
	End If
End Sub
	
<!--
'===================================  3.2.15 btnOriginOnClick()  ===================================
-->
Sub btnOriginOnClick()
	If frm1.txtOrigin.readOnly <> True Then
		Call OpenOrigin()
	End If
End Sub
<!--
'===============================  3.2.16 rdoPartailShip_OnPropertyChange()  =============================
-->
Sub rdoPartailShip1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoPartailShip2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

<!--
'===============================  3.2.17 rdoTranshipment_OnPropertyChange()  ============================
-->
Sub rdoTranshipment1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoTranshipment2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

<!--
'======================================  3.2.19 chkPackListOnClick()  ==================================
-->
Sub chkPackListOnClick()
	lgBlnFlgChgValue = True
End Sub

<!--
'=================================  3.2.20 rdoBLAwFlg_OnPropertyChange()  ===============================
-->
Sub rdoBLAwFlg1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoBLAwFlg2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub chkCertOriginFlg_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

<!--
'====================================  3.2.21 chkCertOriginFlgOnClick()  ===============================
-->
Sub chkCertOriginFlgOnClick()
	lgBlnFlgChgValue = True
End Sub

<!--
'================================  3.2.23 rdoTransferg_OnPropertyChange()  ==============================
-->
Sub rdoTransfer1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoTransfer2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

<!--
'================================  3.2.24 rdoChargeCd_OnPropertyChange()  ===============================
-->
Sub rdoChargeCd1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoChargeCd2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

<!--
'=========================================  5.1.1 FncQuery()  ===========================================
-->
Function FncQuery()
	Dim IntRetCD

	FncQuery = False

	Err.Clear

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If


	If Not chkField(Document, "1") Then
	    If gPageNo > 0 Then
	        gSelframeFlg = gPageNo
	    End If
		        
	    Exit Function
	End If 

		
	Call ggoOper.ClearField(Document, "2")
	Call SetRadio()

	If DbQuery = False Then Exit Function

	FncQuery = True
	SCheck = FALSE
	Set gActiveElement = document.activeElement

End Function
	
<!--
'===========================================  5.1.2 FncNew()  ===========================================
-->
Function FncNew()
	Dim IntRetCD 

	FncNew = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

		
	Call ClickTab1()
	Call ggoOper.ClearField(Document, "A")
	Call SetRadio()
	Call ggoOper.LockField(Document, "N")
	Call changeXchRtTag()
		
	Call SetDefaultVal
	Call InitVariables
	Call setDefaultTag()

	frm1.txtLCNo.focus
	Set gActiveElement = document.activeElement
		
	FncNew = True
End Function
	
<!--
'===========================================  5.1.3 FncDelete()  ========================================
-->
Function FncDelete()
	Dim IntRetCD
		
	FncDelete = False
		
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")
	If IntRetCD = vbNo Then
		Exit Function
	End If
		
	If DbDelete = False Then Exit Function

	FncDelete = True
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.4 FncSave()  ==========================================
-->
Function FncSave()
	Dim IntRetCD
		
	FncSave = False
		
	Err.Clear
		
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
	    Exit Function
	End If
		
		
    If Not chkField(Document, "2") Then
        If gPageNo > 0 Then
            gSelframeFlg = gPageNo
        End If
	        
        Exit Function
    End If
	    
	if Trim(UNICDbl(frm1.txtDocAmt.text)) = "" Or Trim(UNICDbl(frm1.txtDocAmt.text)) = "0" then
		Call DisplayMsgBox("970021", "X","개설금액", "X")
		Call ClickTab1()
		frm1.txtDocAmt.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
	    
    If ValidDateCheckLocal(frm1.txtReqDt, frm1.txtOpenDt) = False Then Exit Function
	If ValidDateCheckLocal(frm1.txtReqDt, frm1.txtLatestShipDt) = False Then Exit Function
    If ValidDateCheckLocal(frm1.txtOpenDt, frm1.txtLatestShipDt) = False Then Exit Function
    If ValidDateCheckLocal(frm1.txtLatestShipDt, frm1.txtExpiryDt) = False Then Exit Function
	    
	If DbSave = False Then Exit Function
		
	FncSave = True
	SCheck = FALSE
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.5 FncCopy()  ==========================================
-->
Function FncCopy()
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	lgIntFlgMode = Parent.OPMD_CMODE

	Call ggoOper.ClearField(Document, "1")
	Call ggoOper.LockField(Document, "N")
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.6 FncCancel()  ========================================
-->
Function FncCancel() 
	On Error Resume Next
	Set gActiveElement = document.activeElement
End Function

<!--
'==========================================  5.1.7 FncInsertRow()  ======================================
-->
Function FncInsertRow()
	On Error Resume Next
	Set gActiveElement = document.activeElement
End Function
<!--
'==========================================  5.1.8 FncDeleteRow()  ======================================
-->
Function FncDeleteRow()
	On Error Resume Next
	Set gActiveElement = document.activeElement
End Function

<!--
'============================================  5.1.9 FncPrint()  ========================================
-->
Function FncPrint()
   Call parent.FncPrint()
   Set gActiveElement = document.activeElement
End Function
<!--
'============================================  5.1.10 FncPrev()  ========================================
-->
Function FncPrev() 
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	ElseIf lgPrevNo = "" Then
		Call DisplayMsgBox("900011", "X", "X", "X")
	End If
	Set gActiveElement = document.activeElement
End Function

<!--
'============================================  5.1.11 FncNext()  ========================================
-->
Function FncNext()
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")
	End If
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.12 FncExcel()  ========================================
-->
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLE)
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.13 FncFind()  =========================================
-->
Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE, True)
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.14 FncExit()  =========================================
-->
Function FncExit()
	Dim IntRetCD

	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
	Set gActiveElement = document.activeElement
End Function

<!--
'=============================================  5.2.1 DbQuery()  ========================================
-->
Function DbQuery()
	Dim strVal

	Err.Clear

	DbQuery = False

	if LayerShowHide(1) =false then
	    exit Function
	end if

	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001
	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)

	Call RunMyBizASP(MyBizASP, strVal)
	
	DbQuery = True
End Function
	
<!--
'=============================================  5.2.2 DbSave()  =========================================
-->
Function DbSave()
	Dim strVal

	Err.Clear

	DbSave = False

	if LayerShowHide(1) =false then
	    exit Function
	end if

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
			
		if .chkPoNoCnt.checked = true then
			.hdnPoNoCnt.Value = "1"
		else
			.hdnPoNoCnt.Value = "0"
		End if
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	End With

	DbSave = True
End Function
	
<!--
'=============================================  5.2.3 DbDelete()  =======================================
-->
Function DbDelete()
	Dim strVal

	Err.Clear

	DbDelete = False

	strVal = BIZ_PGM_DEL_ID & "?txtMode=" & Parent.UID_M0003
	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)

	Call RunMyBizASP(MyBizASP, strVal)

	DbDelete = True
End Function
	
<!--
'=============================================  5.2.4 DbQueryOk()  ======================================
-->
Function DbQueryOk()
		
	lgIntFlgMode = Parent.OPMD_UMODE

	Call ggoOper.LockField(Document, "Q")
	Call SetToolbar("11111000000111")
	
	Call ClickTab1()
	Call changeTag()
        
	lgBlnFlgChgValue = False
	frm1.txtLCDocNo.focus	
	Set gActiveElement = document.activeElement

End Function

<!--
'=============================================  5.2.4 RefOk()  ======================================
-->
Function RefOk()
	Call SetToolbar("1110100000001111")
	Call changeXchRtTag()
End Function	
<!--
'=============================================  5.2.5 DbSaveOk()  =======================================
-->
Function DbSaveOk()
	Call InitVariables
	Call MainQuery()
End Function	
<!--
'=============================================  5.2.6 DbDeleteOk()  =====================================
-->
Function DbDeleteOk()													<%'☆: 삭제 성공후 실행 로직 %>
	Call FncNew()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C 일반정보</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
	 				</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>구비서류</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>은행 및 기타</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPORef" >발주참조</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD <%=HEIGHT_TYPE_02%>></TD>
					</TR>
					<TR>
						<TD HEIGHT=20 WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
									<TR>
										<TD CLASS=TD5>L/C관리번호</TD>
										<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtLCNo"  SIZE=32 MAXLENGTH=18 TAG="12XXXU" ALT="L/C관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCNo" ALIGN=top TYPE="BUTTON" onclick="vbscript:btnLCNoOnClick()"></TD>
										<TD CLASS=TD6>&nbsp;</TD>
										<TD CLASS=TD6>&nbsp;</TD>
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
							<!-- 첫번째 탭 내용 -->
							<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5>L/C관리번호</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtLCNo1"  SIZE=34 MAXLENGTH=18 TAG="25XXXU" ALT="L/C관리번호"></TD>
									<TD CLASS=TD5>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>L/C번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="L/C번호" TYPE=TEXT MAXLENGTH=35 SIZE=20  TAG="21XXXU" OnChange="VBScript:changeLcDocNo()"> -
														 <INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>발주번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPONo" ALT="발주번호" TYPE=TEXT MAXLENGTH=18 SIZE=20  TAG="24XXXU">
														 <INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="21" CHECKED ID="chkPoNoCnt">발주번호지정</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>L/C유형</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCType" SIZE=10  MAXLENGTH=5 TAG="22XXXU" ALT="L/C유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" onclick="vbscript:btnLCTypeOnClick()">
														 <INPUT TYPE=TEXT NAME="txtLCTypeNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>개설신청일</TD>
									<TD CLASS=TD6 NOWRAP>
										<Table Cellspacing=0 Cellpadding=0>
											<TR>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=개설신청일 NAME="txtReqDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD NOWRAP>
													&nbsp;개설일&nbsp;
												</TD>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=개설일 NAME="txtOpenDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</Table>
									</TD>														 
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>통지은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdvBank" SIZE=10  MAXLENGTH=10 TAG="22XXXU" ALT="통지은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAdvBank" align=top TYPE="BUTTON" onclick="vbscript:btnAdvBankOnClick()">
														 <INPUT TYPE=TEXT NAME="txtAdvBankNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>개설은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10  MAXLENGTH=10 TAG="22XXXU" ALT="개설은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenBank" align=top TYPE="BUTTON" onclick="vbscript:btnOpenBankOnClick()">
														 <INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>유효일</TD>
									<TD CLASS=TD6 NOWRAP>
										<Table Cellspacing=0 Cellpadding=0>
											<TR>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=유효일 NAME="txtExpiryDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD NOWRAP>
													&nbsp;최종선적기일
												</TD>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=최종선적기일 NAME="txtLatestShipDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</Table>
									</TD>
									<TD CLASS=TD5 NOWRAP>선적기일참조</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtShipment" ALT="선적기일참조" TYPE=TEXT MAXLENGTH=70 SIZE=34 TAG="21X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>화폐</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10  MAXLENGTH=3 TAG="24XXXU" ALT="화폐"></TD>
									<TD CLASS=TD5 NOWRAP>환율</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=환율 NAME="txtXchRate" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="22X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=개설금액 NAME="txtDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="22X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
									<TD CLASS=TD5 NOWRAP>자국금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=자국금액 NAME="txtLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>가격조건</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoterms" SIZE=10  MAXLENGTH=5 TAG="24XXXU" ALT="가격조건"></TD>
									<TD CLASS=TD5 NOWRAP>운송방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransport" SIZE=10  MAXLENGTH=5 TAG="22XXXU" ALT="운송방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransport" align=top TYPE="BUTTON" onclick="vbscript:btnTransportOnClick()">
														 <INPUT TYPE=TEXT NAME="txtTransportNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>결제방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10  MAXLENGTH=5 TAG="22XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" onclick="vbscript:btnPayTermsOnClick()">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>대금결제참조</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPaytermstxt" ALT="대금결제참조" TYPE=TEXT MAXLENGTH=120 SIZE=34 TAG="21X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>결제기간</TD>
									<TD CLASS=TD6 NOWRAP>
										<table cellpadding=0 cellspacing=0 >
											<TR>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="결제기간" NAME="txtPayDur" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="21X70" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</table>
									</TD>
									<TD CLASS=TD5 NOWRAP>인도장소</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeliveryPlce" ALT="인도장소" TYPE=TEXT MAXLENGTH=5 SIZE=10  TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeliveryPlce" align=top TYPE="BUTTON" onclick="vbscript:btnDeliveryPlceOnClick()">
														 <INPUT NAME="txtDeliveryPlceNm" ALT="인도장소" TYPE=TEXT  SIZE=20 TAG="24X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>과부족허용율</TD>
									<TD CLASS=TD6 NOWRAP>
										<table cellpadding=0 cellspacing=0 >
											<TR>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=과부족허용율 NAME="txttolerance" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>
												<TD NOWRAP>
													(%)
												</TD>
											</TR>
										</table>
									</TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>선적항</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoadingPort" ALT="선적항" TYPE=TEXT MAXLENGTH=5 SIZE=10  TAG="22NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingPort" align=top TYPE="BUTTON" onclick="vbscript:btnLoadingPortOnClick()">
														 <INPUT TYPE=TEXT NAME="txtLoadingPortNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>도착항</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDischgePort" ALT="도착항" TYPE=TEXT MAXLENGTH=5 SIZE=10  TAG="22NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgePort" align=top TYPE="BUTTON" onclick="vbscript:btnDischgePortOnClick()">
														 <INPUT TYPE=TEXT NAME="txtDischgePortNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>분할선적허용</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="21" VALUE="Y" CHECKED ID="rdoPartailShip1"><LABEL FOR="rdoPartailShip1">Y</LABEL>
														 <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="21" VALUE="N" ID="rdoPartailShip2"><LABEL FOR="rdoPartailShip2">N</LABEL>&nbsp;&nbsp;
														 환적허용
														 <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTranshipment" TAG="21" CHECKED VALUE="Y" ID="rdoTranshipment1"><LABEL FOR="rdoTranshipment1">Y</LABEL>
														 <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTranshipment" TAG="21" VALUE="N" ID="rdoTranshipment2"><LABEL FOR="rdoTranshipment2">N</LABEL>&nbsp;&nbsp;</TD>
									<TD CLASS=TD5 NOWRAP>원산지</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOrigin" ALT="원산지" TYPE=TEXT MAXLENGTH=5 SIZE=10  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrigin" align=top TYPE="BUTTON" onclick="vbscript:btnOriginOnClick()">
														 <INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>구매그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10  MAXLENGTH=4 TAG="24XXXU" ALT="구매그룹">&nbsp;&nbsp;&nbsp;&nbsp;
														 <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>수수료 부담자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoChargeCd" TAG="21" VALUE="Y" CHECKED ID="rdoChargeCd1"><LABEL FOR="rdoChargeCd1">APPLICANT</LABEL>
														 <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoChargeCd" TAG="21" VALUE="N" ID="rdoChargeCd2"><LABEL FOR="rdoChargeCd2">Beneficiary</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>수출자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="24XXXU" ALT="수출자">&nbsp;&nbsp;&nbsp;&nbsp;
														 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>구매조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=10  MAXLENGTH=4 TAG="24XXXU" ALT="구매조직">&nbsp;&nbsp;&nbsp;&nbsp;
														 <INPUT TYPE=TEXT NAME="txtPurOrgNm" SIZE=20 TAG="24"></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(1)%>
							</TABLE>
							</DIV>
						<!--</TD>
					</TR>
				</TABLE>
						 두번째 탭 내용 -->
						<!--<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>-->
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>서류제시기간</TD>
									<TD CLASS=TD6 NOWRAP>
										<table cellpadding=0 cellspacing=0 >
											<TR>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=서류제시기간 NAME="txtFileDt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="21X70" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>
												<TD NOWRAP>
													DAYS
												</TD>
											</TR>
										</table>
									</TD>
									<TD CLASS=TD5 NOWRAP>서류제시기간참조</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFileDtTxt" ALT="서류제시기간참조" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="21X"></TD>
								</TR>
								
								<TR>
									<TD CLASS=TD5 NOWRAP>상업송장</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=상업송장 NAME="txtInvCnt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="21X70" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>포장명세서</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=포장명세서 NAME="txtPackList" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 80px" tag="21X71" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>B/L형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoBLAwFlg" TAG="21" VALUE="Y" CHECKED ID="rdoBLAwFlg1"><LABEL FOR="rdoBLAwFlg1">BILL OF LADING</LABEL>
														 <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoBLAwFlg" TAG="21" VALUE="N" ID="rdoBLAwFlg2"><LABEL FOR="rdoBLAwFlg2">AIR WAY BILL</LABEL></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>운임지불형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFreight" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="운임지불형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFreight" align=top TYPE="BUTTON" onclick="vbscript:btnFreightOnClick()">
														 <INPUT TYPE=TEXT NAME="txtFreightNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>통지처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNotifyParty" SIZE=10  MAXLENGTH=10 TAG="21XXXU" ALT="통지처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNotifyParty" align=top TYPE="BUTTON" onclick="vbscript:btnNotifyPartyOnClick()">
														 <INPUT TYPE=TEXT NAME="txtNotifyPartyNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>수탁자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConsignee" ALT="수탁자" TYPE=TEXT MAXLENGTH=30 SIZE=34 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP>원산지증명서필요</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="chkCertOriginFlg" TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" ID="chkCertOriginFlg"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>양도가능여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTransfer" TAG="21" VALUE="Y" CHECKED ID="rdoTransfer1"><LABEL FOR="rdoTransfer1">Y</LABEL>
														 <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTransfer" TAG="21" VALUE="N" ID="rdoTransfer2"><LABEL FOR="rdoTransfer2">N</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>원산지국가</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOriginCntry" ALT="원산지국가" TYPE=TEXT MAXLENGTH=3 SIZE=10  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOriginCntry" align=top TYPE="BUTTON" onclick="vbscript:btnOriginCntryOnClick()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>부보조건</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInsurPolicy" ALT="보험부보조건" TYPE=TEXT MAXLENGTH=30 SIZE=34 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>기타서류</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc1" ALT="기타서류" TYPE=TEXT MAXLENGTH=65 SIZE=34 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc2" ALT="기타서류" TYPE=TEXT MAXLENGTH=65 SIZE=34 TAG="21X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc3" ALT="기타서류" TYPE=TEXT MAXLENGTH=65 SIZE=34 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc4" ALT="기타서류" TYPE=TEXT MAXLENGTH=65 SIZE=34 TAG="21X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc5" ALT="기타서류" TYPE=TEXT MAXLENGTH=65 SIZE=34 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(5)%>
							</TABLE>
						</DIV>
						<!-- 세번째 탭 내용 -->
						<!--<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>-->
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>지급은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayBank" SIZE=10  MAXLENGTH=10 TAG="21XXXU" ALT="지급은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayBank" align=top TYPE="BUTTON" onclick="vbscript:btnPayBankOnClick()">
														 <INPUT TYPE=TEXT NAME="txtPayBankNm" SIZE=20 TAG="24X"></TD>
									<TD CLASS=TD5 NOWRAP>RENEGO은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRenegoBank" SIZE=10  MAXLENGTH=10 TAG="21XXXU" ALT="RENEGO은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRenegoBank" align=top TYPE="BUTTON" onclick="vbscript:btnRenegoBankOnClick()">
														 <INPUT TYPE=TEXT NAME="txtRenegoBankNm" SIZE=20 TAG="24X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>확인은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtConfirmBank" SIZE=10  MAXLENGTH=10 TAG="21XXXU" ALT="확인은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConfirmBank" align=top TYPE="BUTTON" onclick="vbscript:btnConfirmBankOnClick()">
														 <INPUT TYPE=TEXT NAME="txtConfirmBankNm" SIZE=20 TAG="24X"></TD>
									<TD CLASS=TD5 NOWRAP>은행지시사항</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBankTxt" SIZE=34  MAXLENGTH=70 TAG="21X" ALT="은행지시사항"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>신용공여주체</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCreditCore" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="신용공여주체"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCreditCore" align=top TYPE="BUTTON" onclick="vbscript:btnCreditCoreOnClick()">
														 <INPUT TYPE=TEXT NAME="txtCreditCoreNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>자금종류</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtfundType" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="자금종류"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFundType" align=top TYPE="BUTTON" onclick="vbscript:btnFundTypeOnClick()">
														 <INPUT TYPE=TEXT NAME="txtfundTypeNm" SIZE=20 TAG="24X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>금융한도액환산율</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=금융한도액환산율 NAME="txtLmtXchRate" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>수수료참조사항</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtChargeTxt" SIZE=34 MAXLENGTH=30 TAG="21X" ALT="수수료 참조사항"></TD>
								</TR>
								<TR>				 
									<TD CLASS=TD5 NOWRAP>환산금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=환산금액 NAME="txtLmtAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="21X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
<!--					<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=환산금액 NAME="txtLmtAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>-->
									<TD CLASS=TD5 NOWRAP>운송회사</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTransportComp" ALT="운송회사" TYPE=TEXT MAXLENGTH=30 SIZE=34 TAG="21X"></TD>					 
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>대행자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10  MAXLENGTH=10 TAG="21XXXU" ALT="대행자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAgent" align=top TYPE="BUTTON" onclick="vbscript:btnAgentOnClick()">
														 <INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>제조자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10  MAXLENGTH=10 TAG="21XXXU" ALT="제조자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnManufacturer" align=top TYPE="BUTTON" onclick="vbscript:btnManufacturerOnClick()">
														 <INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>접수일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=접수일 NAME="txtAdvDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>AMEND일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=AMEND일 NAME="txtAmendDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								</TR><!--
								<TR>
									<TD CLASS=TD5 NOWRAP>통지번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAdvNo" ALT="통지번호" TYPE=TEXT MAXLENGTH=35 SIZE=20 STYLE="Text-Transform: uppercase" TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP>선통지참조사항</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPreAdvRef" ALT="선통지참조사항" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="21X"></TD>
								</TR>-->
								<TR>
									<TD CLASS=TD5 NOWRAP>기타참조</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRemark" ALT="기타참조" TYPE=TEXT MAXLENGTH=70 SIZE=34 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP>개설의뢰인</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10  MAXLENGTH=10 TAG="24XXXU" ALT="개설의뢰인">&nbsp;&nbsp;&nbsp;&nbsp;
														 <INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(11)%>
							</TABLE>
						</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				<TD WIDTH=10>&nbsp;</TD>
				<!--<td align="Left"><a><button name="btnLCwin" id="btnLCwin" class="clsmbtn" ONCLICK="OpenLCwin()">L/C 개설신청</button></a></td> -->
				<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadLCDtl()">L/C내역등록</A>&nbsp;|&nbsp;<A href="vbscript:LoadLCAmendHdr()">AMEND등록</A>&nbsp;|&nbsp;<A href="vbscript:LoadChargeHdr()">경비등록</A></TD>
				<TD WIDTH=10>&nbsp;</TD>
			</TABLE>
		</TD>
	</TR>
	<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex= -1></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHLCNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNoCnt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnXchRtOp" tag="24">   
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" tabindex = -1></IFRAME>
</DIV>
</BODY>
</HTML>
