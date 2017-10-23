
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4211ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수입통관 등록 ASP															*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2001/11/16																*
'*  9. Modifier (First)     : Cho Song Hyon																*
'* 10. Modifier (Last)      : Jin-hyun Shin																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'********************************************  1.1 Inc 선언  ********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBS">
Option Explicit	

	Const BIZ_PGM_QRY_ID = "m4211mb1_KO441.asp"	
	Const BIZ_PGM_SAVE_ID = "m4211mb2_KO441.asp"	
	Const BIZ_PGM_DEL_ID = "m4211mb3_KO441.asp"	
	Const BIZ_PGM_BLQRY_ID = "m4211mb4_KO441.asp"	
	'Const BIZ_PGM_CCHQRY_ID = "m4211mb5.asp"
	Const CC_DETAIL_ENTRY_ID = "m4212ma1"
	Const CHARGE_HDR_ENTRY_ID = "m6111ma2"	

	Const TAB1 = 1
	Const TAB2 = 2

	Const gstrCustomsMajor = "S9013"		'징수형태 
	Const gstrCollectTypeMajor = "M4201"	'통관계획 
	Const gstrCCtypeMajor = "M9000"			'신고구분 
	Const gstrIDTypeMajor = "M9001"			'거래구분 
	Const gstrIPTypeMajor = "M9002"			'수입종류 
	Const gstrImportTypeMajor = "M9003"		'포장형태 
	Const gstrPackingTypeMajor = "B9007"	'도착항 
	Const gstrDischgePortMajor = "B9092"	'운송방법 
	Const gstrTransportMajor = "B9009"		'선적항 
	Const gstrLoadingPortMajor = "B9092"	'원산지 
	Const gstrOriginMajor = "B9094"			'VAT유형 
	Const gstrVatTypeMajor = "B9001"

	<!-- #Include file="../../inc/lgvariables.inc" -->
	Dim gSelframeFlg					
	Dim gblnWinEvent
	Dim EndDate		
	
	EndDate = "<%=GetSvrDate%>"											'⊙: DB의 현재 날짜를 받아와서 시작날짜에 사용한다.
	EndDate = UNIConvDateAtoB(EndDate, parent.gServerDateFormat, parent.gDateFormat)

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
	frm1.txtIDDt.Text = EndDate
	frm1.txtIDReqDt.Text = EndDate
	frm1.txtDischgeDt.Text = EndDate
	Call SetToolBar("1110000000001111")
	frm1.chkBLNo.checked = False
	Call ClickTab1()
	frm1.txtCCNo.focus
	Set gActiveElement = document.activeElement
End Sub
	
<!--
'==========================================  2.2.2 LoadInfTB19029()  ====================================
-->
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
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
<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenCcNoPop()  ++++++++++++++++++++++++++++++++++++++
-->

Function OpenCcNoPop()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Or UCase(frm1.txtCCNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("M4211PA1_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4211PA1_KO441", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
		frm1.txtCCNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtCCNo.value = strRet
		frm1.txtCCNo.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenBlRef()  ++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenBlRef()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "X", "X", "X")
		Exit function
	End If	
	
	gblnWinEvent = True
	
	iCalledAspName = AskPRAspName("M5211RA1_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M5211RA1_KO441", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	
	If strRet(0) = "" Then
		frm1.txtCCNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		Call SetBlRef(strRet)
	End If
		
		
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenCountry()  +++++++++++++++++++++++++++++++++++++++++
-->
Function OpenCountry(strCntryCD, strPopPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "국가"						
	arrParam(1) = "B_COUNTRY"					
	arrParam(2) = Trim(strCntryCD)				
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
		Exit Function
	Else
		Call SetCountry(strPopPos, arrRet)
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++  OpenBizPartner()  ++++++++++++++++++++++++++++++++++++++++
-->
Function OpenBizPartner(strBizPartnerCD, strBizPartnerNM, strPopPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = strPopPos						
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = Trim(strBizPartnerCD)			
'		arrParam(3) = Trim(strBizPartnerNM)			
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "				
	arrParam(5) = strPopPos						

	arrField(0) = "BP_CD"						
	arrField(1) = "BP_NM"						

	arrHeader(0) = strPopPos					
	arrHeader(1) = strPopPos & "명"			

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
'++++++++++++++++++++++++++++++++++++++++++++++++  OpenUnit()  ++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "중량단위"					
	arrParam(1) = "B_UNIT_OF_MEASURE"			
	arrParam(2) = Trim(frm1.txtWeightUnit.value)
	arrParam(3) = ""							
	arrParam(4) = "DIMENSION=" & FilterVar("WT", "''", "S") & ""				
	arrParam(5) = "중량단위"					

	arrField(0) = "UNIT"						
	arrField(1) = "UNIT_NM"					

	arrHeader(0) = "중량단위"			
	arrHeader(1) = "중량단위명"			

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtWeightUnit.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtWeightUnit.Value = arrRet(0)
		frm1.txtWeightUnit.focus
		lgBlnFlgChgValue = True
		Set gActiveElement = document.activeElement
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenMinorCd()  +++++++++++++++++++++++++++++++++++++++++
-->
Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strMajorCd)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = strPopPos							
	arrParam(1) = "B_Minor"							
	arrParam(2) = Trim(strMinorCD)					
'		arrParam(3) = Trim(strMinorNM)					
	arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""	
	arrParam(5) = strPopPos							

	arrField(0) = "Minor_CD"						
	arrField(1) = "Minor_NM"						

	arrHeader(0) = strPopPos						
	arrHeader(1) = strPopPos & "명"				

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
'++++++++++++++++++++++++++++++++++++++++++++++  OpenPortCd()  +++++++++++++++++++++++++++++++++++++++++
-->
Function OpenPortCd(Byval iOpenMinor)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "통관등록팝업"					
	arrParam(1) = "B_MINOR"								
	arrParam(4) = "MAJOR_CD=" & FilterVar("B9092", "''", "S") & ""							
    arrField(0) = "MINOR_CD"							
    arrField(1) = "MINOR_NM"							
	    	    
	Select Case iOpenMinor
    	Case 1
    		arrParam(0) = "도착항"						
    		arrParam(2) = Trim(frm1.txtDischgePortCd.value)	
    '		arrParam(3) = Trim(frm1.txtDischgePortNm.value)	
    		arrParam(5) = "도착항"						
    		    	    
    	    arrHeader(0) = "도착항"						
    	    arrHeader(1) = "도착항명"					
    	Case 2
    		arrParam(0) = "선적항"						
    		arrParam(2) = Trim(frm1.txtLoadingPort.value)	
    '		arrParam(3) = Trim(frm1.txtLoadingPortNm.value)	
    		arrParam(5) = "선적항"						
    		    	    
    	    arrHeader(0) = "선적항"						
    	    arrHeader(1) = "선적항명"					
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPortCd(iOpenMinor, arrRet)
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenVatType()  +++++++++++++++++++++++++++++++++++++++++
-->
Function OpenVatType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "부가세유형"							
	arrParam(1) = "B_MINOR Minor,B_CONFIGURATION Config"	
	arrParam(2) = Trim(frm1.txtVatType.value)				
	arrParam(3) = ""										
	arrParam(4) = "Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD" _
					& " And Config.MINOR_CD = Minor.MINOR_CD" _
					& " And Config.SEQ_NO = 1"				
	arrParam(5) = "부가세유형"							
		
    arrField(0) = "Minor.MINOR_CD"							
    arrField(1) = "Minor.MINOR_NM"							
    arrField(2) = "Config.REFERENCE"						
	    	    
    arrHeader(0) = "부가세코드"							
    arrHeader(1) = "부가세명"							
	arrHeader(2) = "부가세율"							

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtVatType.focus
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtVatType.Value = arrRet(0)
		frm1.txtVatRate.Text = arrRet(2)	
		frm1.txtVatType.focus
		Set gActiveElement = document.activeElement	
		lgBlnFlgChgValue = True
	End If
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++  SetMinorCd()  +++++++++++++++++++++++++++++++++++++++++++
-->
Function SetPortCd(Byval iSetMinor, arrRet)
	Select Case iSetMinor
		'도착항 
		Case 1
			frm1.txtDischgePortCd.Value = arrRet(0)
			frm1.txtDischgePortNm.Value = arrRet(1)
			frm1.txtDischgePortCd.focus
		'선적항 
		Case 2
			frm1.txtLoadingPort.Value = arrRet(0)
			frm1.txtLoadingPortNm.Value = arrRet(1)
			frm1.txtLoadingPort.focus
	End Select
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = True
End Function
	
<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetBlRef()  ++++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetBlRef()																					+
'+	Description : Set Return array from B/L Reference Window											+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function SetBlRef(strRet)
	Dim strVal
			
	Call ggoOper.ClearField(Document, "A")					
	Call SetDefaultVal
	frm1.txtBlNo.value =  strRet(0)
	frm1.txtBLDocNo.value =  strRet(1)
	If Trim(frm1.txtBLDocNo.value) <> "" Then 
		frm1.chkBLNo.checked = True
	End if
		
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 

	strVal = BIZ_PGM_BLQRY_ID & "?txtBlNo=" & Trim(frm1.txtBlNo.value)

	Call RunMyBizASP(MyBizASP, strVal)								
		
End Function
	
<!--
'+++++++++++++++++++++++++++++++++++++++++++++  SetCountry()  +++++++++++++++++++++++++++++++++++++++++++
-->
Function SetCountry(strPopPos, arrRet)
	Select Case (strPopPos)
		Case "선박국적"
			frm1.txtVesselCntry.Value = arrRet(0)
			frm1.txtVesselCntry.focus

		Case "적출국가"
			frm1.txtLoadingCntry.Value = arrRet(0)
			frm1.txtLoadingCntry.focus

		Case "원산지국가"
			frm1.txtOriginCntry.Value = arrRet(0)
			frm1.txtOriginCntry.focus
	End Select
			Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = True
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++  SetBizPartner()  ++++++++++++++++++++++++++++++++++++++++++
-->
Function SetBizPartner(strPopPos, arrRet)
	Select Case (strPopPos)
		Case "신고자"
			frm1.txtReporterCd.Value = arrRet(0)
			frm1.txtReporterNm.Value = arrRet(1)
			frm1.txtReporterCd.focus
				
		Case "납세의무자"
			frm1.txtTaxPayerCd.Value = arrRet(0)
			frm1.txtTaxPayerNm.Value = arrRet(1)
			frm1.txtTaxPayerCd.focus
	End Select
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = True
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++  SetMinorCd()  +++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetMinorCd()																					+
'+	Description : Set Return array from Minor Code PopUp Window											+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function SetMinorCd(strMajorCd, arrRet)
	Select Case strMajorCd
		Case gstrCustomsMajor
			frm1.txtCustoms.Value = arrRet(0)
			frm1.txtCustomsNm.Value = arrRet(1)
			frm1.txtCustoms.focus
		Case gstrCollectTypeMajor
			frm1.txtCollectType.Value = arrRet(0)
			frm1.txtCollectTypeNm.Value = arrRet(1)
			frm1.txtCollectType.focus
		Case gstrCCtypeMajor
			frm1.txtCCtypeCd.Value = arrRet(0)
			frm1.txtCCtypeNm.Value = arrRet(1)
			frm1.txtCCtypeCd.focus
		Case gstrIDTypeMajor
			frm1.txtIDType.Value = arrRet(0)
			frm1.txtIDTypeNm.Value = arrRet(1)
			frm1.txtIDType.focus
		Case gstrIPTypeMajor
			frm1.txtIPType.Value = arrRet(0)
			frm1.txtIPTypeNm.Value = arrRet(1)
			frm1.txtIPType.focus
		Case gstrImportTypeMajor
			frm1.txtImportType.Value = arrRet(0)
			frm1.txtImportTypeNm.Value = arrRet(1)
			frm1.txtImportType.focus
		Case gstrPackingTypeMajor
			frm1.txtPackingType.Value = arrRet(0)
			frm1.txtPackingTypeNm.Value = arrRet(1)
			frm1.txtPackingType.focus
		Case gstrDischgePortMajor
			frm1.txtDischgePortCd.Value = arrRet(0)
			frm1.txtDischgePortNm.Value = arrRet(1)
			frm1.txtDischgePortCd.focus
		Case gstrTransportMajor
			frm1.txtTransport.Value = arrRet(0)
			frm1.txtTransportNm.Value = arrRet(1)
			frm1.txtTransport.focus
		Case gstrLoadingPortMajor
			frm1.txtLoadingPort.Value = arrRet(0)
			frm1.txtLoadingPortNm.Value = arrRet(1)
			frm1.txtLoadingPort.focus
		Case gstrOriginMajor
			frm1.txtOrigin.Value = arrRet(0)
			frm1.txtOriginNm.Value = arrRet(1)
			frm1.txtOrigin.focus
		Case gstrVatTypeMajor
			frm1.txtVatType.Value = arrRet(0)
			frm1.txtVatType.focus

	End Select
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = True
End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtCIFDocAmt, "USD", parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		
		
		
	End With

End Sub
<!--
'++++++++++++++++++++++++++++++++++++++++++++  dbBlQueryok()  *++++++++++++++++++++++++++++++++++++++++++
-->
Function dbBlQueryok()
	Call SetToolBar("1110100000001111")
End Function	
<!--
'=============================================  2.5.1 LoadCcDtl()  ======================================
-->
Function LoadCcDtl()

	'통관관리번호 
	WriteCookie "CCNo", UCase(Trim(frm1.txtCCNo.value))
	PgmJump(CC_DETAIL_ENTRY_ID)

End Function
	
<!--
'=============================================  2.5.1 LoadChargeHdr()  ======================================
-->
Function LoadChargeHdr()

	Dim IntRetCD

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                              
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End if
	    	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	'Process Step
	WriteCookie "Process_Step", "VD"
	'통관관리번호 
	WriteCookie "Po_No", UCase(Trim(frm1.txtCCNo.value))
	'면허번호 
	'구매그룹(수입담당)
	WriteCookie "Pur_Grp", UCase(Trim(frm1.txtPurGrp.value))
	'화폐 
	'WriteCookie "Currency", UCase(Trim(frm1.txtCurrency.value))
	'환율 
	'WriteCookie "XchRate", UCase(Trim(frm1.txtXchRate.Text))

	PgmJump(CHARGE_HDR_ENTRY_ID)

End Function
		
<!--
'============================================  2.5.2 OpenCookie()  ======================================
-->
Function OpenCookie(ByVal Kubun)

	frm1.txtCCNo.value = ReadCookie("CCNo")
	WriteCookie "CCNo", ""

	If UCase(Trim(frm1.txtCCNo.value)) <> "" Then 
		Call MainQuery
	End If

End Function

<!--
'==========================================  2.5.3 CookiePage()  ======================================
-->
Function CookiePage(Byval Kubun)

	Const CookieSplit = 4875						
	Dim strTemp, arrVal
	Dim IntRetCD

	If Kubun = 1 Then

	    If lgIntFlgMode <> Parent.OPMD_UMODE Then          
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If

		WriteCookie CookieSplit , frm1.txtCCNo.value 
		
		Call PgmJump(CC_DETAIL_ENTRY_ID)
		
	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtCCNo.value =  arrVal(0) 

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()
					
		WriteCookie CookieSplit , ""

	End IF

End Function

<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
Sub Form_Load()
	Call LoadInfTB19029							
	Call AppendNumberRange("0","0","9999999999")
	Call AppendNumberPlace("7","2","0")	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")				
	Call SetDefaultVal
	Call InitVariables
	Call CookiePage(0)
	'Call changeTabs(TAB1)
	gIsTab     = "Y" 
    gTabMaxCnt = 2   

End Sub
	
<!--
'=========================================  3.1.2 Form_QueryUnload()  ===================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
	   
End Sub	
<!--
'=====================================  3.2.2 btnCustoms_OnClick()  ====================================
-->
Sub btnCustoms_Click()
	If frm1.txtCustoms.readOnly <> True Then
		Call OpenMinorCd(frm1.txtCustoms.value, frm1.txtCustomsNm.value, "세관", gstrCustomsMajor)
	End If
End Sub

<!--
'===================================  3.2.3 btnCollectType_OnClick()  ====================================
-->
Sub btnCollectType_Click()
	If frm1.txtCollectType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtCollectType.value, frm1.txtCollectTypeNm.value, "징수형태", gstrCollectTypeMajor)
	End If
End Sub
<!--
'======================================  3.2.4 btnReporter_OnClick()  ===================================
-->
Sub btnReporter_Click()
	If frm1.txtReporterCd.readOnly <> True Then
		Call OpenBizPartner(frm1.txtReporterCd.value, frm1.txtReporterNm.value, "신고자")
	End If
End Sub

<!--
'======================================  3.2.5 btnTaxPayer_OnClick()  ===================================
-->
Sub btnTaxPayer_Click()
	If frm1.txtTaxPayerCd.readOnly <> True Then
		Call OpenBizPartner(frm1.txtTaxPayerCd.value, frm1.txtTaxPayerNm.value, "납세의무자")
	End If
End Sub

<!--
'======================================  3.2.6 btnCCtype_OnClick()  ===================================
-->
Sub btnCCtype_Click()
	If frm1.txtCCtypeCd.readOnly <> True Then
		Call OpenMinorCd(frm1.txtCCtypeCd.value, frm1.txtCCtypeNm.value, "통관계획", gstrCCtypeMajor)
	End If
End Sub

<!--
'=====================================  3.2.7 btnIDType_OnClick()  ======================================
-->
Sub btnIDType_Click()
	If frm1.txtIDType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtIDType.value, frm1.txtIDTypeNm.value, "신고구분", gstrIDTypeMajor)
	End If
End Sub

<!--
'===================================  3.2.8 btnIPType_OnClick()  ====================================
-->
Sub btnIPType_Click()
	If frm1.txtIPType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtIPType.value, frm1.txtIPTypeNm.value, "거래구분", gstrIPTypeMajor)
	End If
End Sub

<!--
'===================================  3.2.9 btnImportType_OnClick()  ====================================
-->
Sub btnImportType_Click()
	If frm1.txtImportType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtImportType.value, frm1.txtImportTypeNm.value, "수입종류", gstrImportTypeMajor)
	End If
End Sub

<!--
'====================================  3.2.10 btnWeightUnit_OnClick()  ===================================
-->
Sub btnWeightUnit_Click()
	If frm1.txtWeightUnit.readOnly <> True Then
		Call OpenUnit()
	End If
End Sub

<!--
'===================================  3.2.11 btnPackingType_OnClick()  ==================================
-->
Sub btnPackingType_Click()
	If frm1.txtPackingType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtPackingType.value, frm1.txtPackingTypeNm.value, "포장형태", gstrPackingTypeMajor)
	End If
End Sub

<!--
'===================================  3.2.12 btnDischgePort_OnClick()  ==================================
-->
Sub btnDischgePort_Click()
	If frm1.txtDischgePortCd.readOnly <> True Then
'			Call OpenMinorCd(frm1.txtDischgePortCd.value, frm1.txtDischgePortNm.value, "도착항", gstrDischgePortMajor)
		Call OpenPortCd(1)
	End If
End Sub

<!--
'===================================  3.2.13 btnTransport_OnClick()  ====================================
-->
Sub btnTransport_Click()
	If frm1.txtTransport.readOnly <> True Then
		Call OpenMinorCd(frm1.txtTransport.value, frm1.txtTransportNm.value, "운송방법", gstrTransportMajor)
	End If
End Sub

<!--
'===================================  3.2.14 btnVesselCntry_OnClick()  ==================================
-->
Sub btnVesselCntry_Click()
	If frm1.txtTransport.readOnly <> True Then
		Call OpenCountry(frm1.txtVesselCntry.value, "선박국적")
	End If
End Sub

<!--
'==================================  3.2.15 btnLoadingPort_OnClick()  ===================================
-->
Sub btnLoadingPort_Click()
	If frm1.txtLoadingPort.readOnly <> True Then
'			Call OpenMinorCd(frm1.txtLoadingPort.value, frm1.txtLoadingPortNm.value, "선적항", gstrLoadingPortMajor)
		Call OpenPortCd(2)
	End If
End Sub

<!--
'===================================  3.2.16 btnLoadingCntry_OnClick()  =================================
-->
Sub btnLoadingCntry_Click()
	If frm1.txtLoadingCntry.readOnly <> True Then
		Call OpenCountry(frm1.txtLoadingCntry.value, "적출국가")
	End If
End Sub

<!--
'===================================  3.2.17 btnOrigin_OnClick()  =======================================
-->
Sub btnOrigin_Click()
	If frm1.txtOrigin.readOnly <> True Then
		Call OpenMinorCd(frm1.txtOrigin.value, frm1.txtOriginNm.value, "원산지", gstrOriginMajor)
	End If
End Sub

<!--
'===================================  3.2.18 btnOriginCntry_OnClick()  ==================================
-->
Sub btnOriginCntry_Click()
	If frm1.txtOriginCntry.readOnly <> True Then
		Call OpenCountry(frm1.txtOriginCntry.value, "원산지국가")
	End If
End Sub

<!--
'=====================================  3.2.19 btnVatType_OnClick()  ====================================
-->
Sub btnVatType_Click()
	If frm1.txtVatType.readOnly <> True Then
		Call OpenVatType()
	End If
End Sub

<!--
'==========================================================================================
'   Event Name : OCX_Change()
'==========================================================================================
-->
Sub txtIDDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtIDReqDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtIPDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtDischgeDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtPutDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtGrossWeight_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtTotPackingCnt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtNetWeight_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtDocAmt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtUSDXchRate_Change()
	lgBlnFlgChgValue = True
End Sub


Sub txtLoadingDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtInspectDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtOutputDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtCustomsExpDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtPaymentDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtDvryDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtTaxBillDt_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtTariffTax_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtTariffRate_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtVatRate_Change()
	lgBlnFlgChgValue = True
End Sub
Sub txtVatAmt_Change()
	lgBlnFlgChgValue = True
End Sub
'Sub txtAddLocAmt_Change()
'	lgBlnFlgChgValue = True
'End Sub
'Sub txtReduLocAmt_Change()
'	lgBlnFlgChgValue = True
'End Sub

<!--
'==========================================================================================
'   Event Name : OCX_DbClick()
'==========================================================================================
-->
Sub txtIDDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIDDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIDDt.focus
	End If
End Sub
Sub txtIDReqDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIDReqDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIDReqDt.focus
	End If
End Sub
Sub txtIPDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIPDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIPDt.focus
	End If
End Sub
Sub txtDischgeDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDischgeDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDischgeDt.focus
	End If
End Sub
Sub txtPutDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPutDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPutDt.focus
	End If
End Sub
Sub txtLoadingDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtLoadingDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLoadingDt.focus
	End If
End Sub
Sub txtInspectDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtInspectDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtInspectDt.focus
	End If
End Sub
Sub txtOutputDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtOutputDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtOutputDt.focus
	End If
End Sub
Sub txtCustomsExpDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtCustomsExpDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtCustomsExpDt.focus
	End If
End Sub
Sub txtPaymentDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPaymentDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPaymentDt.focus
	End If
End Sub
Sub txtDvryDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDvryDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDvryDt.focus
	End If
End Sub
Sub txtTaxBillDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtTaxBillDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtTaxBillDt.focus
	End If
End Sub

<!--
'=======================================  chkBLNo_onpropertychange()  ======================================
-->
Sub chkBLNo_onpropertychange()
	lgBlnFlgChgValue = true	
End Sub
<!--
'=========================================  5.1.1 FncQuery()  ===========================================
-->
Function FncQuery()
	Dim IntRetCD

	FncQuery = False

	Err.Clear		

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "2")					
	Call InitVariables										

	If Not chkField(Document, "1") Then	Exit Function
	If DbQuery = False Then Exit Function

	FncQuery = True	
	Set gActiveElement = document.activeElement
End Function
	
<!--
'===========================================  5.1.2 FncNew()  ===========================================
-->
Function FncNew()
	Dim IntRetCD 

	FncNew = False  

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'Call ClickTab1()
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")
	Call SetDefaultVal
	Call InitVariables					
		
	frm1.txtCCNo.focus
	Set gActiveElement = document.activeElement
		
	FncNew = True						
End Function
	
<!--
'===========================================  5.1.3 FncDelete()  ========================================
-->
Function FncDelete()
	FncDelete = False					
		
	Dim IntRetCD
		
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")
	    
    If IntRetCD = vbNo Then Exit Function
	    
	If lgIntFlgMode <> Parent.OPMD_UMODE Then					
		Call DisplayMsgBox("900002","X","X","X")
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
		
	If frm1.txtCCNo.value  <> "" Then
		If lgBlnFlgChgValue = False Then						
	    IntRetCD = DisplayMsgBox("900001","X","X","X")			
		    Exit Function
		End If
	End If
		
    If Not chkField(Document, "2") Then	
        If gPageNo > 0 Then
            gSelframeFlg = gPageNo
        End If
        Exit Function
    End If 
    '환율을 protect로 set
     'if Trim(UNICDbl(frm1.txtXchRate.text)) = "" Or Trim(UNICDbl(frm1.txtXchRate.text)) = "0" then
	'	Call DisplayMsgBox("970021", "X","환율", "X")
	'	Call ClickTab1()
	'	frm1.txtXchRate.focus
	'	Set gActiveElement = document.activeElement
	'	Exit Function
	 'End if
		
	 if Trim(UNICDbl(frm1.txtGrossWeight.text)) = "" Or Trim(UNICDbl(frm1.txtGrossWeight.text)) = "0" then
		Call DisplayMsgBox("970021", "X","총중량", "X")
		Call ClickTab1()
		frm1.txtGrossWeight.focus
		Set gActiveElement = document.activeElement
		Exit Function
	 End if
	If DbSave = False Then Exit Function
		
	FncSave = True	
	Set gActiveElement = document.activeElement	
End Function

<!--
'===========================================  5.1.5 FncCopy()  ==========================================
-->
Function FncCopy()
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	lgIntFlgMode = Parent.OPMD_CMODE									

	Call ggoOper.ClearField(Document, "1")		
	Call ggoOper.LockField(Document, "N")		

	frm1.txtInPutCCNo.value = ""
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.6 FncCancel()  ========================================
-->
Function FncCancel() 
	On Error Resume Next	
End Function

<!--
'==========================================  5.1.7 FncInsertRow()  ======================================
-->
Function FncInsertRow()
	On Error Resume Next	
End Function
<!--
'==========================================  5.1.8 FncDeleteRow()  ======================================
-->
Function FncDeleteRow()
	On Error Resume Next	
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
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	ElseIf lgPrevNo = "" Then			
		Call DisplayMsgBox("900011","X","X","X")
	End If
	Set gActiveElement = document.activeElement
End Function

<!--
'============================================  5.1.11 FncNext()  ========================================
-->
Function FncNext()
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then	
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	ElseIf lgNextNo = "" Then			
		Call DisplayMsgBox("900012","X","X","X")
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
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
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

	If LayerShowHide(1) = False Then Exit Function
	
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001		
	strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo.value)

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

	If LayerShowHide(1) = False Then Exit Function
	
	With frm1
		.txtMode.value = Parent.UID_M0002							
		.txtFlgMode.value = lgIntFlgMode
			
		If .chkBLNo.checked = True Then
			.txtChkBLNo.value = "Y"
		Else
			.txtChkBLNo.value = "N"
		End If

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
	strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo.value)

	Call RunMyBizASP(MyBizASP, strVal)						

	DbDelete = True											
End Function

<!--
'=============================================  5.2.4 DbQueryOk()  ======================================
-->
Function DbQueryOk()										
	lgIntFlgMode = Parent.OPMD_UMODE								
	Call ggoOper.LockField(Document, "Q")					
	Call SetToolBar("11111000000111")
	lgBlnFlgChgValue = False
	frm1.txtCCNo.focus
	Call ClickTab1()
		
	Set gActiveElement = document.activeElement 
End Function
	
<!--
'=============================================  5.2.5 DbSaveOk()  =======================================
-->
Function DbSaveOk()	
	Call InitVariables
	'Call MainQuery()
	Call FncQuery()
End Function
	
<!--
'=============================================  5.2.6 DbDeleteOk()  =====================================
-->
Function DbDeleteOk()	
	lgBlnFlgChgValue = False
	Call MainNew()
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수입신고1</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수입신고2</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenBlRef">B/L참조</A></TD>
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
									<TD CLASS=TD5 NOWRAP>통관관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCNo" SIZE=32 MAXLENGTH=18 TAG="12XXXU"  ALT="통관관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCcNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCcNoPop()"></TD>
									<TD CLASS=TD6></TD>
									<TD CLASS=TD6></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100%>
					
						<!-- 첫번째 탭 내용 -->
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>통관관리번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtInPutCCNo" SIZE=34 MAXLENGTH=18 TAG="25" STYLE="Text-Transform: uppercase" ALT="통관관리번호"></TD>
								<TD CLASS=TD5></TD>
								<TD CLASS=TD6></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>신고번호</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP><INPUT NAME="txtIDNo" ALT="신고번호" TYPE=TEXT MAXLENGTH=20 SIZE=20  TAG="21XXXU"></TD>
											<TD NOWRAP>&nbsp;신고일&nbsp;</TD>
											<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtIDDt" CLASS=FPDTYYYYMMDD tag="22X1" ALT="신고일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>신고요청일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtIDReqDt" CLASS=FPDTYYYYMMDD tag="22X1" ALT="신고요청일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>면허번호</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP><INPUT NAME="txtIPNo" ALT="면허번호" TYPE=TEXT MAXLENGTH=20 SIZE=20 TAG="21XXXU"></TD>
											<TD NOWRAP>&nbsp;면허일&nbsp;</TD>
											<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtIPDt" style="HEIGHT: 20px; WIDTH: 100px" tag="21X1" ALT="면허일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>B/L번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBLDocNo" ALT="B/L번호" TYPE=TEXT MAXLENGTH=35 SIZE=25  TAG="24XXXU"><INPUT TYPE=CHECKBOX NAME="chkBLNo" tag="21" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid"><LABEL FOR="chkBLNo">B/L번호지정</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>세관</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCustoms" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="세관"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCustoms" align=top TYPE="BUTTON" ONCLICK="vbscript:btnCustoms_Click()" >&nbsp;<INPUT TYPE=TEXT NAME="txtCustomsNm" SIZE=20 TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>도착일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 NAME="txtDischgeDt" CLASS=FPDTYYYYMMDD tag="22X1" ALT="도착일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>반입번호</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP>
												<INPUT NAME="txtInputNo" TYPE=TEXT MAXLENGTH=35  SIZE=20 TAG="21XXXU">
											</TD>
											<TD NOWRAP>
												&nbsp;반입일&nbsp;
											</TD>
											<TD NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime5 NAME="txtPutDt" CLASS=FPDTYYYYMMDD tag="21X1" ALT="반입일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>징수형태</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCollectType" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="징수형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCollectType" align=top TYPE="BUTTON" ONCLICK="vbscript:btnCollectType_Click()">&nbsp;<INPUT TYPE=TEXT NAME="txtCollectTypeNm" SIZE=20 TAG="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>신고자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReporterCd" SIZE=10  MAXLENGTH=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReporter" align=top TYPE="BUTTON" ONCLICK="vbscript:btnReporter_Click()">&nbsp;<INPUT TYPE=TEXT NAME="txtReporterNm" SIZE=20 TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>납세의무자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTaxPayerCd" SIZE=10  MAXLENGTH=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxPayer" align=top TYPE="BUTTON" ONCLICK="vbscript:btnTaxPayer_Click()">&nbsp;<INPUT TYPE=TEXT NAME="txtTaxPayerNm" SIZE=20 TAG="24"></TD>
							</TR>								
							<TR>
								<TD CLASS=TD5 NOWRAP>통관계획</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCtypeCd" SIZE=10  MAXLENGTH=5 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCCType" align=top TYPE="BUTTON" ONCLICK="vbscript:btnCCtype_Click()">&nbsp;<INPUT TYPE=TEXT NAME="txtCCtypeNm" SIZE=20 TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>신고구분</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIDType" SIZE=10  MAXLENGTH=4 TAG="21XXXU" ALT="신고구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIDType" align=top TYPE="BUTTON" ONCLICK="vbscript:btnIDType_Click()">&nbsp;<INPUT TYPE=TEXT NAME="txtIDTypeNm" SIZE=20 TAG="24"></TD>
							</TR>								
							<TR>
								<TD CLASS=TD5 NOWRAP>거래구분</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIPType" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="거래구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIPType" align=top TYPE="BUTTON" ONCLICK="vbscript:btnIPType_Click()">&nbsp;<INPUT TYPE=TEXT NAME="txtIPTypeNm" SIZE=20 TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>수입종류</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtImportType" SIZE=10  MAXLENGTH=5 TAG="21XXXU" ALT="수입종류"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnImportType" align=top TYPE="BUTTON" ONCLICK="vbscript:btnImportType_Click()">&nbsp;<INPUT TYPE=TEXT NAME="txtImportTypeNm" SIZE=20 TAG="24"></TD>
							</TR>								
							<TR>
								<TD CLASS=TD5 NOWRAP>총중량</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle13 NAME="txtGrossWeight" style="HEIGHT: 20px; WIDTH: 160px" ALT="총중량" tag="22X3Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>총포장개수</TD>
								<TD CLASS=TD6 NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle14 NAME="txtTotPackingCnt" ALT="총포장개수" style="HEIGHT: 20px; WIDTH: 160px" tag="21X7Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;
											</td>
										</tr>
									</table>
								</td>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>순중량</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
											<TD NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle14 NAME="txtNetWeight" style="HEIGHT: 20px; WIDTH: 80px" tag="24X30" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
											<TD NOWRAP>
												&nbsp;중량단위
											</TD>
											<TD>
												<INPUT NAME="txtWeightUnit" ALT="중량단위" TYPE=TEXT MAXLENGTH=3 SIZE=10  TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWeightUnit" align=top TYPE="BUTTON" ONCLICK="vbscript:btnWeightUnit_Click()">
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>포장형태</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPackingType" SIZE=10  MAXLENGTH=4 TAG="21XXXU" ALT="포장형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPackingType" align=top TYPE="BUTTON" ONCLICK="vbscript:btnPackingType_Click()">
													 <INPUT TYPE=TEXT NAME="txtPackingTypeNm" SIZE=20 TAG="24"></TD>
							</TR>								
							<TR>
								<TD CLASS=TD5 NOWRAP>도착항</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDischgePortCd" ALT="도착항" TYPE=TEXT MAXLENGTH=5 SIZE=10  TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgePort" align=top TYPE="BUTTON" ONCLICK="vbscript:btnDischgePort_Click()">&nbsp;<INPUT NAME="txtDischgePortNm" TYPE=TEXT SIZE=20 TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>운송방법</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransport" SIZE=10  MAXLENGTH=5 TAG="22XXXU" ALT="운송방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransport" align=top TYPE="BUTTON" ONCLICK="vbscript:btnTransport_Click()">&nbsp;<INPUT TYPE=TEXT NAME="txtTransportNm" SIZE=20 TAG="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>결제방법</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTerms" TYPE="Text" MAXLENGTH="5" SIZE=10 STYLE=" Text-Transform: uppercase" tag="24">&nbsp;&nbsp;&nbsp;&nbsp;
													 <INPUT NAME="txtPayTermsNm" TYPE="Text" MAXLENGTH="20" SIZE=20 STYLE=" Text-Transform: uppercase" tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>결제기간</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayDur" TYPE="Text" MAXLENGTH="3" SIZE=3 STYLE=" Text-Transform: uppercase" tag="24"><LABEL>일</LABEL>
									&nbsp;&nbsp;&nbsp;가격조건<INPUT NAME="txtIncoterms" TYPE="Text" MAXLENGTH="5" SIZE=12 STYLE=" Text-Transform: uppercase" tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>화폐</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP>
												<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>환율</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtXchRate" style="HEIGHT: 20px; WIDTH: 160px" tag="24X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>통관금액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtDocAmt" style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>자국금액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtLocAmt" style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>CIF금액(US$)</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="CIF금액(US$)" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtCIFDocAmt" style="HEIGHT: 20px; WIDTH: 250px" tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>USD환율</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="USD환율" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtUSDXchRate" style="HEIGHT: 20px; WIDTH: 160px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
								<TD CLASS=TD5 NOWRAP>CIF자국금액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 NAME="txtCIFLocAmt" style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							<!--<TR>
								<TD CLASS=TD5 NOWRAP>USD환율</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtUSDXchRate" style="HEIGHT: 20px; WIDTH: 120px" tag="21X5" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
											<TD NOWRAP>
												<LABEL>CIF금액(US$)</LABEL>
											</TD>
											<TD NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtCIFDocAmt" style="HEIGHT: 20px; WIDTH: 120px" tag="21X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>CIF원화금액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 NAME="txtCIFLocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>-->
							<TR>
								<TD CLASS=TD5 NOWRAP>수출자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="24XXXU" ALT="수출자">&nbsp;&nbsp;&nbsp;&nbsp;
													 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>수입자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10  MAXLENGTH=10 TAG="24XXXU" ALT="수입자">&nbsp;&nbsp;&nbsp;&nbsp;
													 <INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
							</TR>
							<%Call SubFillRemBodyTD5656(0)%>
						</TABLE>
						</DIV>

						<!-- 두번째 탭 내용 -->
						<DIV ID="TabDiv" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>VESSEL명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVesselNm" ALT="VESSEL명" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="21"></TD>
								<TD CLASS=TD5 NOWRAP>선박국적</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVesselCntry" TYPE=TEXT MAXLENGTH=3 SIZE=10  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVesselCntry" align=top TYPE="BUTTON" ONCLICK="vbscript:btnVesselCntry_Click()"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>선적항</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoadingPort" ALT="선적항" TYPE=TEXT MAXLENGTH=5 SIZE=10  TAG="21"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingPort" align=top TYPE="BUTTON" ONCLICK="vbscript:btnLoadingPort_Click()">&nbsp;<INPUT NAME="txtLoadingPortNm" TYPE=TEXT SIZE=20 TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>적출국가</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<INPUT NAME="txtLoadingCntry" TYPE=TEXT MAXLENGTH=3 SIZE=10  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingCntry" align=top TYPE="BUTTON" ONCLICK="vbscript:btnLoadingCntry_Click()">
											</TD>
											<TD>
												&nbsp;선적일
											</TD>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime6 NAME="txtLoadingDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>장치확인번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeviceNo" ALT="장치확인번호" TYPE=TEXT MAXLENGTH=20 SIZE=34  TAG="21XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>반입장소</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDevicePlce" ALT="반입장소" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="21X"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>포장번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPackingNo" ALT="포장번호" TYPE=TEXT MAXLENGTH=20 SIZE=20 TAG="21XXXU">
								<TD CLASS=TD5 NOWRAP>조사란</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtExamTxt" ALT="조사란" TYPE=TEXT MAXLENGTH=30 SIZE=34 TAG="21X"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>원산지</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOrigin" ALT="원산지" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrigin" align=top TYPE="BUTTON" ONCLICK="vbscript:btnOrigin_Click()">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=20 TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>원산지국가</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOriginCntry" TYPE=TEXT MAXLENGTH=3 SIZE=10  TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOriginCntry" align=top TYPE="BUTTON" ONCLICK="vbscript:btnOriginCntry_Click()"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>검사일</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime7 NAME="txtInspectDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
											<TD NOWRAP>
												&nbsp;반출일
											</TD>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime8 NAME="txtOutputDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>납부서번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPaymentNo" ALT="납부서번호" TYPE=TEXT MAXLENGTH=20 SIZE=34  TAG="21XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>세관만기일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime9 NAME="txtCustomsExpDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>납부일</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime10 NAME="txtPaymentDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
											<TD NOWRAP>
												&nbsp;납기일
											</TD>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime11 NAME="txtDvryDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>계산서번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBillNo" ALT="계산서번호" TYPE=TEXT MAXLENGTH=20 SIZE=34  TAG="21XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>계산서발행일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime12 NAME="txtTaxBillDt" style="HEIGHT: 20px; WIDTH: 100px" tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>관세</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 NAME="txtTariffTax" style="HEIGHT: 20px; WIDTH: 150px" tag="21X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
											<TD NOWRAP>
												&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle8 NAME="txtTariffRate" style="HEIGHT: 20px; WIDTH: 150px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
											<TD>
												%
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>VAT</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD NOWRAP>
												<INPUT NAME="txtVatType" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON" ONCLICK="vbscript:btnVatType_Click()">
											</TD>
											<TD NOWRAP>
												&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle9 NAME="txtVatRate" style="HEIGHT: 20px; WIDTH: 150px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
											<TD>
												%
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>VAT금액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle10 NAME="txtVatAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="21X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
<!--								<TD CLASS=TD5 NOWRAP>가산금액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle11 NAME="txtAddLocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="21X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>공제금액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle12 NAME="txtReduLocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="21X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>								
-->								<TD CLASS=TD5 NOWRAP>L/C번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCDocNo" SIZE=29 MAXLENGTH=35 TAG="24XXXU"><LABEL>-</LABEL><INPUT TYPE=TEXT NAME="txtLCAmendSeq" SIZE=3 MAXLENGTH=3 TAG="24XXXU" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>대행자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgentCd" SIZE=10  MAXLENGTH=10 TAG="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>제조자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturerCd" SIZE=10  MAXLENGTH=10 TAG="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>구매그룹</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10  MAXLENGTH=4 TAG="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="24"></TD>
								<TD CLASS=TD5 NOWRAP>구매조직</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=10  MAXLENGTH=4 TAG="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtPurOrgNm" SIZE=20 TAG="24"></TD>
							</TR>
							<%Call SubFillRemBodyTD5656(6)%>
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
				<TD WIDTH=* ALIGN=RIGHT><A href="VBSCRIPT:CookiePage(1)">수입통관내역등록</A>&nbsp;|&nbsp;<A href="vbscript:LoadChargeHdr()">경비등록</A></TD>
				<TD WIDTH=50>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=100 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHCCNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtLcNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtLcType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtLcOpenDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtChkBLNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtBlNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDiv" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>	
</BODY>
</HTML>
