
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3211ma2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Import Local L/C Header 등록 ASP											*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/17																*
'*  8. Modified date(Last)  : 2003/05/19																*
'*  9. Modifier (First)     : Sun-joung Lee																*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/03/22 : Coding Start												*
'********************************************************************************************************
-->
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
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
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

	Const BIZ_PGM_QRY_ID = "m3211mb5.asp"			
	Const BIZ_PGM_SAVE_ID = "m3211mb5.asp"			
	Const BIZ_PGM_DEL_ID = "m3211mb5.asp"			
	Const BIZ_PGM_POQRY_ID = "m3211mb8.asp"			
	Const BIZ_PGM_GRQRY_ID = "m3211mb8.asp"			
	Const LC_DETAIL_ENTRY_ID = "m3212ma2"
	Const AMEND_ENTRY_ID = "m3221ma2"
	Const CHARGE_HDR_ENTRY_ID = "m6111ma2"			
	Const BIZ_PGM_CAL_AMT_ID = "m3211mb9.asp"
	
	Const TAB1 = 1
	Const TAB2 = 2
	
	Dim lgBlnFlgChgValue				
	Dim lgIntGrpCount					
	Dim lgIntFlgMode					

	Dim gSelframeFlg					
	Dim gblnWinEvent					
	Dim SCheck	
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

<!--
'============================================ OpenPayType()  ======================================
-->
Function OpenPayType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	'if Trim(frm1.txtPayTerms.Value) = "" then
	'		Call DisplayMsgBox("17a002", Parent.VB_YES_NO,"결제방법", "X")
	'		Exit Function
	'	End if

	If gblnWinEvent = True Or UCase(frm1.txtPayTerms.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	gblnWinEvent = True
		
	arrParam(0) = "결제방법"					
	arrParam(1) = "B_Minor,b_configuration"		

	arrParam(2) = Trim(frm1.txtPayTerms.Value)	
	'arrParam(3) = Trim(frm1.txtPayNm.Value)	
		
	arrParam(4) = "b_minor.Major_Cd=" & FilterVar("B9004", "''", "S") & "" _
	   				& " and b_minor.minor_cd=b_configuration.minor_cd" _
	   				& " AND b_configuration.REFERENCE = " & FilterVar("L", "''", "S") & " "
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
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenBankPop()  ++++++++++++++++++++++++++++++++++++++++
-->
Function OpenBankPop(strBankCd, strBankNm, strPopPos)
   Dim arrRet
   Dim arrParam(5), arrField(6), arrHeader(6)

   If gblnWinEvent = True Or UCase(frm1.txtAdvBank.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

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
'+++++++++++++++++++++++++++++++++++++++++++++  OpenPurGrp()  +++++++++++++++++++++++++++++++++++++++++++
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
   		frm1.txtPurGrp.value = arrRet(0)
		frm1.txtPurGrpNm.value = arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtPurGrp.focus
		Set gActiveElement = document.activeElement
   End If
End Function
	
<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenLCNoPop()  ++++++++++++++++++++++++++++++++++++++++
-->
Function OpenLCNoPop()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
			
	If gblnWinEvent = True Or UCase(frm1.txtLCNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
			
	gblnWinEvent = True
			
	iCalledAspName = AskPRAspName("M3211PA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3211PA2", "X")
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
		Set gActiveElement = document.activeElement
	End If	
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenPORef()  ++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenPORef()
				
	Dim strRet,IntRetCD
	Dim iCalledAspName
			
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "X", "X", "X")
		Exit function
	End If
			
	iCalledAspName = AskPRAspName("M3111RA3")		

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111RA3", "X")
		gblnWinEvent = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
			
	gblnWinEvent = False

	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		lgIntFlgMode = Parent.OPMD_CMODE
		Call FncNew()
	End If 
			
	If strRet(0) = "" Then
		Call ClickTab1()
		frm1.txtLCDocNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else			
		Call SetPORef(strRet)
		frm1.txtLCDocNo.focus
		Set gActiveElement = document.activeElement		
	End If
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenGRRef()  ++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenGRRef()
				
	Dim strRet,IntRetCD
	Dim iCalledAspName
			
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "X", "X", "X")
		Exit function
	End If
			
	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True

	iCalledAspName = AskPRAspName("M4111RA3")		
			
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4111RA3", "X")
		gblnWinEvent = False
		Exit Function
	End If
			
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		lgIntFlgMode = Parent.OPMD_CMODE
		Call FncNew()
	End If
			
	If strRet = "" Then
		frm1.txtLCDocNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else						
		Call SetGRRef(strRet)	
		frm1.txtLCDocNo.focus
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
   arrHeader(1) = strPopPos					

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
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenOLcDocNoPop()  ++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenOLcDocNoPop()																				+
'+	Description : Master L/C No PopUp Call																	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenOLcNoPop()
   Dim arrRet
   Dim arrParam(5), arrField(6), arrHeader(6)

   If gblnWinEvent = True Or UCase(frm1.txtOLCNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

   gblnWinEvent = True

   arrParam(0) = "근거서류번호"				
   arrParam(1) = "S_LC_HDR"					
   arrParam(2) = Trim(frm1.txtOLCNo.value)		
   arrParam(3) = ""							
   arrParam(4) = "lc_kind = " & FilterVar(Trim(frm1.txtOLcKind.value), " " , "S") & " "
   arrParam(5) = "근거서류번호"				

   arrField(0) = "LC_NO"						
   arrField(1) = "LC_DOC_No"					

   arrHeader(0) = "LC관리번호"				
   arrHeader(1) = "LC번호"					

   arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
   		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

   gblnWinEvent = False

   If arrRet(0) = "" Then
		frm1.txtOLcNo.focus
		Set gActiveElement = document.activeElement		
   		Exit Function
   Else
		frm1.txtOLcNo.value = arrRet(0)
		frm1.txtHOLCDocNo.value = arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtOLcNo.focus
		Set gActiveElement = document.activeElement	
   End If
End Function	

<!--
'+++++++++++++++++++++++++++++++++++++++ 2.4.5  OpenBizPartner() ++++++++++++++++++++++++++++++++++++++++
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
'++++++++++++++++++++++++++++++++++++++++++++++  SetPORef()  ++++++++++++++++++++++++++++++++++++++++++++
-->
Function SetPORef(strRet)
   Dim strVal

   Call ggoOper.ClearField(Document, "2")		
   Call SetDefaultVal
		
   frm1.txtPONo.value = strRet(0)
   frm1.hdnXchRtOp.value = strRet(1)
   frm1.chkPONoFlg.checked = True

   strVal = BIZ_PGM_POQRY_ID & "?txtPONo=" & Trim(frm1.txtPONo.value)	
   strVal = strVal & "&txtCurrency=" & Trim(Parent.gCurrency)

   Call RunMyBizASP(MyBizASP, strVal)		
		
End Function
<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetGRRef()  ++++++++++++++++++++++++++++++++++++++++++++
-->
Function SetGRRef(strRet)
   Dim strVal

   Call ggoOper.ClearField(Document, "2")								
   Call SetDefaultVal

   strRet = split(strRet,"&")
        
   frm1.txtPONo.value = strRet(0)
   frm1.hdnXchRtOp.value = strRet(1)
   frm1.chkPONoFlg.checked = True

   strVal = BIZ_PGM_GRQRY_ID & "?txtPONo=" & Trim(frm1.txtPONo.value)	
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
		Case "S9001"
			frm1.txtLCType.Value = arrRet(0)
			frm1.txtLCTypeNm.Value = arrRet(1)
			frm1.txtLCType.focus

		Case "S9002"
			frm1.txtOLcKind.Value = arrRet(0)
			frm1.txtOLcKindNm.Value = arrRet(1)
			frm1.txtOLcKind.focus
					
		Case Else
	End Select
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = True
End Function	

<!--
'++++++++++++++++++++++++++++++++++++++++ 2.5.6 SetBizPartner() +++++++++++++++++++++++++++++++++++++++++
-->
Function SetBizPartner(strPopPos, arrRet)
   Select Case (strPopPos)
   	Case "대행자"
   		frm1.txtAgent.Value = arrRet(0)
   		frm1.txtAgentNm.Value = arrRet(1)
				
   	Case "제조자"
   		frm1.txtManufacturer.Value = arrRet(0)
   		frm1.txtManufacturerNm.Value = arrRet(1)
				
   	Case Else
   End Select

   lgBlnFlgChgValue = True
End Function	

'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

   ggoOper.FormatFieldByObjectOfCur frm1.txtDocAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
   ggoOper.FormatFieldByObjectOfCur frm1.txtXchRate, frm1.txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
End Sub
'============================================ changeTag()  ====================================
Function changeTag()

	With frm1

	If Trim(frm1.txtLCAmendSeq.value) <> "" And Trim(frm1.txtLCAmendSeq.value) <> "0" Then
	   Call SetToolbar("1110000000001111")
	   'tab1
	   ggoOper.SetReqAttr	.txtLCNo1, "Q"
	   ggoOper.SetReqAttr	.txtLCDocNo, "Q"
	   ggoOper.SetReqAttr	.chkPONoFlg, "Q"
	   ggoOper.SetReqAttr	.txtLCType, "Q"
	   ggoOper.SetReqAttr	.txtReqDt, "Q"
	   ggoOper.SetReqAttr	.txtOpenDt, "Q"
	   ggoOper.SetReqAttr	.txtOpenBank, "Q"
	   ggoOper.SetReqAttr	.txtAdvBank, "Q"
	   ggoOper.SetReqAttr	.txtExpiryDt, "Q"
	   ggoOper.SetReqAttr	.txtLatestShipDt, "Q"
	   ggoOper.SetReqAttr	.txtShipment, "Q"
	   ggoOper.SetReqAttr	.txtXchRate, "Q"
	   ggoOper.SetReqAttr	.txtDocAmt, "Q"
	   ggoOper.SetReqAttr	.txtOLcKind, "Q"
	   ggoOper.SetReqAttr	.rdoPartailShip1, "Q"
	   ggoOper.SetReqAttr	.rdoPartailShip2, "Q"
	   ggoOper.SetReqAttr	.txtFileDt, "Q"
	   ggoOper.SetReqAttr	.txtFileDtTxt, "Q"
	   ggoOper.SetReqAttr	.txtPayTerms, "Q"
	   ggoOper.SetReqAttr	.txtPaytermstxt, "Q"
	   ggoOper.SetReqAttr	.txtAdvDt, "Q"
	   'tab2
	   ggoOper.SetReqAttr	.txtDoc1, "Q"
	   ggoOper.SetReqAttr	.txtDoc2, "Q"
	   ggoOper.SetReqAttr	.txtDoc3, "Q"
	   ggoOper.SetReqAttr	.txtDoc4, "Q"
	   ggoOper.SetReqAttr	.txtDoc5, "Q"
	   ggoOper.SetReqAttr	.txtBankTxt, "Q"
	   ggoOper.SetReqAttr	.txtRemark, "Q"
	   ggoOper.SetReqAttr	.txtOLCNo, "Q"
	   ggoOper.SetReqAttr	.txtAgent, "Q"
	   ggoOper.SetReqAttr	.txtManufacturer, "Q"

	Else
	   Call changeLcDocNo()	
	   Call changeXchRtTag()
	   Call setDefaultTag()
	End If

	End With

End Function


'============================================ setDefaultTag()  ====================================
Sub setDefaultTag()
	With frm1
	   'tab1
	   ggoOper.SetReqAttr	.chkPONoFlg, "D"
	   ggoOper.SetReqAttr	.txtLCType, "N"
	   ggoOper.SetReqAttr	.txtReqDt, "N"
	   ggoOper.SetReqAttr	.txtOpenBank, "N"
	   ggoOper.SetReqAttr	.txtAdvBank, "N"
	   ggoOper.SetReqAttr	.txtExpiryDt, "N"
	   ggoOper.SetReqAttr	.txtLatestShipDt, "N"
	   ggoOper.SetReqAttr	.txtShipment, "D"
	   ggoOper.SetReqAttr	.txtDocAmt, "N"
	   ggoOper.SetReqAttr	.txtOLcKind, "N"
	   ggoOper.SetReqAttr	.rdoPartailShip1, "D"
	   ggoOper.SetReqAttr	.rdoPartailShip2, "D"
	   ggoOper.SetReqAttr	.txtFileDt, "D"
	   ggoOper.SetReqAttr	.txtFileDtTxt, "D"
	   ggoOper.SetReqAttr	.txtPayTerms, "N"
	   ggoOper.SetReqAttr	.txtPaytermstxt, "D"
	   ggoOper.SetReqAttr	.txtAdvDt, "D"
	   'tab2
	   ggoOper.SetReqAttr	.txtDoc1, "D"
	   ggoOper.SetReqAttr	.txtDoc2, "D"
	   ggoOper.SetReqAttr	.txtDoc3, "D"
	   ggoOper.SetReqAttr	.txtDoc4, "D"
	   ggoOper.SetReqAttr	.txtDoc5, "D"
	   ggoOper.SetReqAttr	.txtBankTxt, "D"
	   ggoOper.SetReqAttr	.txtRemark, "D"
	   ggoOper.SetReqAttr	.txtOLCNo, "D"
	   ggoOper.SetReqAttr	.txtAgent, "D"
	   ggoOper.SetReqAttr	.txtManufacturer, "D"
	End With
End Sub
<!--
'===========================================  changeXchRtTag()  =======================================
-->
Function changeXchRtTag()

	If Trim(frm1.txtCurrency.value) <> Parent.gCurrency Then
	   Call ggoOper.SetReqAttr(frm1.txtXchRate, "N")
	Else
	   Call ggoOper.SetReqAttr(frm1.txtXchRate, "Q")
	   frm1.txtXchRate.Text = "1"
	End If
	
End Function
<!--
'===========================================  2.5.1 CookiePage()  =======================================
-->
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877		
	Dim strTemp, arrVal
	Dim IntRetCD

	If Kubun = 1 Then

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
		    
		WriteCookie CookieSplit , frm1.txtLCNo.value			

		Call PgmJump(LC_DETAIL_ENTRY_ID)
				
	ElseIf Kubun = 0 Then

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
				
	End If

End Function	
<!--
'=============================================  2.5.2 LoadAmend()  ======================================
-->
Function LoadAmend()
	Dim strAmendParam
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

	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & "ListLcDtl"
	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)
			
	Call RunMyBizASP(MyBizASP, strVal)

End Function

'=============================================  DbJumpQueryOK()  ======================================
Function DbJumpQueryOK()
	WriteCookie "txtLCNo", UCase(Trim(frm1.txtLCNo.value))
	PgmJump(AMEND_ENTRY_ID)
End Function
<!--
'=============================================  2.5.1 LoadChargeHdr()  ==================================
-->
Function LoadChargeHdr()
	Dim strHdrOpenParam
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

	WriteCookie "TmpNo",	UCase(Trim(frm1.txtLCNo.value))
	WriteCookie "LCNo",		UCase(Trim(frm1.txtLCNo.value))
	WriteCookie "Process_Step", "VO"
	WriteCookie "Po_No",	UCase(Trim(frm1.txtLCNo.value))		
	WriteCookie "Pur_Grp",	UCase(Trim(frm1.txtPurGrp.value))		
	WriteCookie "BasNo",	UCase(Trim(frm1.txtLCNo.value))

	PgmJump(CHARGE_HDR_ENTRY_ID)

End Function

<!--
'============================================  2.5.3 OpenCookie()  ======================================
-->
Function OpenCookie()	
   frm1.txtLCNo.value = ReadCookie("txtLCNo.value")
   WriteCookie "txtLCNo.value", ""
End Function
	
'==============================================  2.5.4 setAmt()  ======================================
'=	Event Name : setAmt()																				=
'=	Event Desc : 13차 추가																						=
'========================================================================================================
-->
Function setAmt()
 
   Dim SumTotal
        
   		If Trim(frm1.hdnXchRtOp.value) = "*" Then
               SumTotal = UNICDbl(frm1.txtDocAmt.text) * UNICDbl(frm1.txtXchRate.text)
               frm1.txtLocAmt.text = UNIFormatNumberByCurrecny(CStr(SumTotal),Parent.gCurrency, Parent.ggAmtOfMoneyNo)
   		ElseIf Trim(frm1.hdnXchRtOp.value) = "/" Then
               If UNICDbl(frm1.txtXchRate.text) = 0 Then
                   SumTotal = 0
               Else
                   SumTotal = UNICDbl(frm1.txtDocAmt.text) / UNICDbl(frm1.txtXchRate.text)
   			End If
   			frm1.txtLocAmt.text = UNIFormatNumberByCurrecny(CStr(SumTotal),Parent.gCurrency, Parent.ggAmtOfMoneyNo)
   		Else 
   			frm1.txtLocAmt.text = frm1.txtDocAmt.text
   		End If
				
End Function		
<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
Sub Form_Load()
		
	Call LoadInfTB19029									
	Call AppendNumberRange("0","0","99")
	Call AppendNumberPlace("7","2","0")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")				
	Call changeTabs(TAB1)
	Call SetDefaultVal()
					
	Call InitVariables
	gSelframeFlg = TAB1
			
	gIsTab     = "Y"
	gTabMaxCnt = 2
	SCheck = TRUE

	Call CookiePage(0)
			
End Sub
	
<!--
'=========================================  3.1.2 Form_QueryUnload()  ===================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
	   
End Sub
	

<!--
'========================================================================================================
'=	Event Name : chkPONoFlg_onpropertychange															=
'========================================================================================================
-->
Sub chkPONoFlg_onpropertychange()
	lgBlnFlgChgValue = true	
End Sub
<!--
'=========================================  3.2.1 btnPayTermsOnClick()  ====================================
-->
Sub btnPayTermsOnClick()
   Call OpenPayType()
End Sub	
<!--
'=========================================  3.2.1 btnLCNoOnClick()  ====================================
-->
Sub btnLCNoOnClick()
   Call OpenLCNoPop()
End Sub
<!--
'========================================  3.2.2 btnAdvBankOnClick()  ==================================
-->
Sub btnAdvBankOnClick()
   Call OpenBankPop(frm1.txtAdvBank.value, frm1.txtAdvBankNm.value, "ADVBANK")
End Sub

<!--
'=======================================  3.2.3 btnOpenBankOnClick()  ==================================
-->
Sub btnOpenBankOnClick()
   Call OpenBankPop(frm1.txtOpenBank.value, frm1.txtOpenBankNm.value, "OPENBANK")
End Sub

<!--
'===================================  3.2.11 btnPurGrpOnClick()  =======================================
-->
Sub btnPurGrpOnClick()
   If frm1.txtPurGrp.readOnly <> True Then
   	Call OpenPurGrp()
   End If
End Sub

<!--
'=========================================  3.2.1 btnOLcDocNoOnClick()  ================================
-->
Sub btnOLcNoOnClick()
   Call OpenOLCNoPop()
End Sub
	
<!--
'=====================================  3.2.12 btnLCTypeOnClick()  =====================================
-->
Sub btnLCTypeOnClick()
	If frm1.txtLCType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtLCType.value, frm1.txtLCTypeNm.value, "LOCAL L/C유형", "S9001")
	End If
End Sub

<!--
'=====================================  3.2.12 btnOLCKindOnClick()  ====================================
-->
Sub btnOLCKindOnClick()
	If frm1.txtOLcKind.readOnly <> True Then
		Call OpenMinorCd(frm1.txtOLcKind.value, frm1.txtOLcKindNm.value, "근거서류유형", "S9002")
	End If
End Sub

<!--
'=================================  3.2.6 rdoPartailShip_OnPropertyChange()  ============================
-->
Sub rdoPartailShip1_OnPropertyChange()
   lgBlnFlgChgValue = True
End Sub

Sub rdoPartailShip2_OnPropertyChange()
   lgBlnFlgChgValue = True
End Sub

<!--
'======================================  3.2.8 btnAgentOnClick()  ======================================
-->
Sub btnAgentOnClick()
	If frm1.txtAgent.readOnly <> True Then
		Call OpenBizPartner(frm1.txtAgent.value, frm1.txtAgentNm.value, "대행자")
	End If
End Sub

<!--
'==================================  3.2.9 btnManufacturerOnClick()  ===================================
-->
Sub btnManufacturerOnClick()
	If frm1.txtManufacturer.readOnly <> True Then
		Call OpenBizPartner(frm1.txtManufacturer.value, frm1.txtManufacturerNm.value, "제조자")
	End If
End Sub			
		
<!--
'================================  3.2.33 txtFileDt_Change()  ===========================================
-->
Sub txtFileDt_Change()
	lgBlnFlgChgValue = True
End Sub	
<!--
'================================  3.2.33 txtReqDt_Change()  ===========================================
-->
Sub txtReqDt_Change()
	lgBlnFlgChgValue = True
End Sub	
<!--
'================================  3.2.33 txtOpenDt_Change()  ===========================================
-->
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

		
	'Call GetPayDt()
				
	If Trim(frm1.txtCurrency.value) <> "" And Trim(frm1.txtCurrency.value) <> Parent.gCurrency  Then
	   Call ChangeCurOrDt()
	ElseIf Trim(frm1.txtCurrency.value) = Parent.gCurrency  Then
	   frm1.txtXchRate.text = UNIFormatNumber(1,ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	End If
		
	lgBlnFlgChgValue = true	
End Sub
<!--
'================================   ChangeCurOrDt()  ===========================================
-->
Function ChangeCurOrDt()
   
	Dim strVal
	Err.Clear                                                               '☜: Protect system from crashing
		
	'frm1.hdnDefaultFlg.value = "Y"

	'if Trim(frm1.hdnchangeflg.value) = "Y" then
	'    frm1.hdnchangeflg.value = "N"
	'    Exit Function
	'end if
		
	If Trim(frm1.txtCurrency.value) = Parent.gCurrency  Then
	   frm1.txtXchRate.text = UNIFormatNumber(1,ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	   Exit Function
	End If


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
<!--
'================================  3.2.33 txtExpiryDt_Change()  =========================================
-->
Sub txtExpiryDt_Change()
	lgBlnFlgChgValue = True
End Sub	
<!--
'================================  3.2.33 txtLatestShipDt_Change()  =====================================
-->
Sub txtLatestShipDt_Change()
	lgBlnFlgChgValue = True
End Sub	
<!--
'================================  3.2.33 txtDocAmt_Change()  ===========================================
-->
Sub txtDocAmt_Change()
	Call setAmt()
	lgBlnFlgChgValue = True
End Sub	
<!--
'================================  3.2.33 txtXchRate_Change()  ==========================================
-->
Sub txtXchRate_Change()
	Call setAmt()
	lgBlnFlgChgValue = True
End Sub	
<!--
'================================  3.2.33 txtLocAmt_Change()  ===========================================
-->
Sub txtLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub	
  
<!--
'================================  3.2.33 txtAdvDt_Change()  ============================================
-->
Sub txtAdvDt_Change()
	lgBlnFlgChgValue = True
End Sub	
<!--
'==========================================================================================
'   Event Name : OCX_DbClick()
'==========================================================================================
-->
Sub txtAdvDt_DblClick(Button)
	If Button = 1 Then
	   frm1.txtAdvDt.Action = 7
	   Call SetFocusToDocument("M")
	   frm1.txtAdvDt.focus
	End If
End Sub
Sub txtReqDt_DblClick(Button)
	If Button = 1 Then
	   frm1.txtReqDt.Action = 7
	   Call SetFocusToDocument("M")
	   frm1.txtReqDt.focus
	End If
End Sub

Sub txtExpiryDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtExpiryDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtExpiryDt.focus
	End If
End Sub
Sub txtLatestShipDt_DblClick(Button)
	If Button = 1 Then
	   frm1.txtLatestShipDt.Action = 7
	   Call SetFocusToDocument("M")
	   frm1.txtLatestShipDt.focus
	End If
End Sub
Sub txtOpenDt_DblClick(Button)
	If Button = 1 Then
	   frm1.txtOpenDt.Action = 7
	   Call SetFocusToDocument("M")
	   frm1.txtOpenDt.focus
	End If
End Sub
Sub txtMoveDt_DblClick(Button)
	If Button = 1 Then
	   frm1.txtMoveDt.Action = 7
	   Call SetFocusToDocument("M")
	   frm1.txtMoveDt.focus
	End If
End Sub
Sub txtAmendDt_DblClick(Button)
	If Button = 1 Then
	   frm1.txtAmendDt.Action = 7
	   Call SetFocusToDocument("M")
	   frm1.txtAmendDt.focus
	End If
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

	Call ggoOper.ClearField(Document, "2")	
	Call InitVariables			
			
	If Not chkField(Document, "1") Then		
		Exit Function
	End If

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
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")	<% '⊙: "Will you destory previous data" %>
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "A")					
	Call ggoOper.LockField(Document, "N")					
	Call SetDefaultVal
	Call InitVariables										
	Call changeXchRtTag()
	Call setDefaultTag()

	FncNew = True	
	Set gActiveElement = document.activeElement										
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

   If UniConvDateToYYYYMMDD(frm1.txtOpenDt.Text,Parent.gDateFormat,"") < UniConvDateToYYYYMMDD(frm1.txtReqDt.Text,Parent.gDateFormat,"") And (frm1.txtOpenDt.Text<>"" Or frm1.txtOpenDt.Text <> Null) Then
   		Call DisplayMsgBox("970023", "X","개설일","개설신청일")
   		Call ClickTab1()
   		frm1.txtOpenDt.focus
   		Set gActiveElement = document.activeElement		
   		Exit Function
   End if

   If UniConvDateToYYYYMMDD(frm1.txtExpiryDt.Text,Parent.gDateFormat,"") < UniConvDateToYYYYMMDD(frm1.txtOpenDt.Text,Parent.gDateFormat,"") And (frm1.txtOpenDt.Text<>"" Or frm1.txtOpenDt.Text <> Null) Then
   		Call DisplayMsgBox("970023", "X","유효일","개설일")
   		Call ClickTab1()
   		frm1.txtExpiryDt.focus
   		Set gActiveElement = document.activeElement		
   		Exit Function
   End if

   If UniConvDateToYYYYMMDD(frm1.txtLatestShipDt.Text,Parent.gDateFormat,"") < UniConvDateToYYYYMMDD(frm1.txtOpenDt.Text,Parent.gDateFormat,"") And (frm1.txtOpenDt.Text<>"" Or frm1.txtOpenDt.Text <> Null) Then
   		Call DisplayMsgBox("970023", "X","인도일자","개설일")
   		Call ClickTab1()
   		frm1.txtLatestShipDt.focus
   		Set gActiveElement = document.activeElement		
   		Exit Function
   End if

   If UniConvDateToYYYYMMDD(frm1.txtExpiryDt.Text,Parent.gDateFormat,"") < UniConvDateToYYYYMMDD(frm1.txtLatestShipDt.Text,Parent.gDateFormat,"") Then
   		Call DisplayMsgBox("970023", "X","유효일","인도일자")
   		Call ClickTab1()
   		frm1.txtExpiryDt.focus
   		Set gActiveElement = document.activeElement		
   		Exit Function
   End if

   If UniConvDateToYYYYMMDD(frm1.txtAdvDt.Text,Parent.gDateFormat,"") <> "" And UniConvDateToYYYYMMDD(frm1.txtAdvDt.Text,Parent.gDateFormat,"") < UniConvDateToYYYYMMDD(frm1.txtOpenDt.Text,Parent.gDateFormat,"") And (frm1.txtOpenDt.Text<>"" Or frm1.txtOpenDt.Text <> Null) Then
   		Call DisplayMsgBox("970023", "X","통지일","개설일")
   		Call ClickTab1()
   		frm1.txtAdvDt.focus
   		Set gActiveElement = document.activeElement		
   		Exit Function
   End if

   If UniConvDateToYYYYMMDD(frm1.txtExpiryDt.Text,Parent.gDateFormat,"") < UniConvDateToYYYYMMDD(frm1.txtReqDt.Text,Parent.gDateFormat,"")  Then
   		Call DisplayMsgBox("970023", "X","유효일","개설신청일")
   		Call ClickTab1()
   		frm1.txtExpiryDt.focus
   		Set gActiveElement = document.activeElement		
   		Exit Function
   End if

   If UniConvDateToYYYYMMDD(frm1.txtLatestShipDt.Text,Parent.gDateFormat,"") < UniConvDateToYYYYMMDD(frm1.txtReqDt.Text,Parent.gDateFormat,"")  Then
   		Call DisplayMsgBox("970023", "X","인도일자","개설신청일")			
   		Call ClickTab1()
   		frm1.txtLatestShipDt.focus
   		Set gActiveElement = document.activeElement		
   		Exit Function
   End if

'	frm1.txtLocAmt.text = UNICDbl(frm1.txtXchRate.text) * UNICDbl(frm1.txtDocAmt.text)
   'msgbox frm1.txtDocAmt.Text

   If UNICDbl(frm1.txtDocAmt.Text) <= 0 Then
   		Call DisplayMsgBox("203425", "X", "X", "X")
   		Call ClickTab1()
   		frm1.txtDocAmt.focus
   		Set gActiveElement = document.activeElement		
   		Exit Function
   End if
			
   'If UNICDbl(frm1.txtXchRate.Text) <= 0 Then
   '	Call DisplayMsgBox("200095", "X", "X", "X")
   '	Call ClickTab1()
   '	frm1.txtXchRate.focus
   '	Set gActiveElement = document.activeElement		
   '	Exit Function
   'End if
		
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
		
   If LayerShowHide(1) = False Then
       Exit Function
   End If
		
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
			
	If frm1.chkPONoFlg.checked = True Then
		frm1.txtPONoFlg.value = "Y"		
	Else 
		frm1.txtPONoFlg.value = "N"	
	End If		 
			
	If LayerShowHide(1) = False Then
	    Exit Function
	End If
			
	'Call txtDocAmt_OnBlur
			
	With frm1
		.txtMode.value = Parent.UID_M0002				
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
	
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
				
	If frm1.txtPONo.value <> "" Then
		frm1.chkPONoFlg.checked=True
	Else
		frm1.chkPONoFlg.checked=False
	End if
			
	Call setAmt()
	Call changeTag()
	lgBlnFlgChgValue = False
	frm1.txtLCDocNo.focus	
	Set gActiveElement = document.activeElement
			
End Function
<!--
'=============================================  5.2.4 DbQueryOk()  ======================================
-->
Function RefOk()
			
	Call SetToolbar("1110100000001111")
	If frm1.txtHLcFlg.value = "N" Then
		frm1.txtPayTerms.value = ""
		frm1.txtPayTermsNm.value = ""
		ggoOper.SetReqAttr	frm1.txtPayTerms, "N"		
	End if
			
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
Function DbDeleteOk()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>LOCAL L/C</font></td>
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
					<TD WIDTH=* align=right><A href="vbscript:OpenPORef" >발주참조</A><!--&nbsp;|&nbsp;<A href="vbscript:OpenGRRef" >입고참조</A>--></TD>
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
										<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtLCNo" SIZE=32 MAXLENGTH=18 TAG="12XXXU" ALT="LOCAL L/C관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCNo" ALIGN=top TYPE="BUTTON" onclick="vbscript:btnLCNoOnClick()"></TD>
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
						<TD WIDTH=100% VALIGN=TOP>
							<!-- 첫번째 탭 내용 -->
							<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>LOCAL L/C관리번호</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtLCNo1" SIZE=34 MAXLENGTH=18 TAG="25XXXU" ALT="LOCAL L/C관리번호"></TD>
									<TD CLASS=TD5>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>LOCAL L/C번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LOCAL L/C번호" TYPE=TEXT MAXLENGTH=35 SIZE=30 TAG="21XXXU" OnChange="VBScript:changeLcDocNo()">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>발주번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPONo" ALT="발주번호" TYPE=TEXT MAXLENGTH=18 SIZE=20  TAG="24XXXU" >
														 <INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkPONoFlg" ID="chkPONoFlg">
														 <LABEL FOR="chkPONoFlg">발주번호지정</LABEL>
									</TD>
								</TR>
								<TR>									
									<TD CLASS=TD5 NOWRAP>LOCAL L/C유형</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCType" SIZE=10  MAXLENGTH=5 TAG="22XXXU" ALT="LOCAL L/C유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" onclick="vbscript:btnLCTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtLCTypeNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>개설신청일</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtReqDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="개설신청일"></OBJECT>');</SCRIPT></TD>
												<TD NOWRAP>&nbsp;개설일</TD>
												<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtOpenDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="개설일"></OBJECT>');</SCRIPT></TD>
											</TR>
										</TABLE>				
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10  MAXLENGTH=10 TAG="22XXXU" ALT="개설은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenBank" align=top TYPE="BUTTON" onclick="vbscript:btnOpenBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>통지은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdvBank" SIZE=10  MAXLENGTH=10 TAG="22XXXU" ALT="통지은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAdvBank" align=top TYPE="BUTTON" onclick="vbscript:btnAdvBankOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtAdvBankNm" SIZE=20 TAG="24"></TD>									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>유효일</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtExpiryDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="유효일"></OBJECT>');</SCRIPT></TD>
												<TD NOWRAP>&nbsp;인도일자</TD>
												<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtLatestShipDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="인도일자"></OBJECT>');</SCRIPT></TD>
											</TR>
										</TABLE>	
									</TD>
									<TD CLASS=TD5 NOWRAP>인도기한 참조</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtShipment" SIZE=34 MAXLENGTH=30 TAG="21X" ALT="인도기한 참조"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>화폐</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=12  MAXLENGTH=3 TAG="24XXXU" ALT="화폐"></TD>												
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>환율</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchRate" style="HEIGHT: 20px; WIDTH: 250px" tag="22X5Z" ALT="환율" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								</TR> 
								<TR>
									<TD CLASS=TD5 NOWRAP>개설금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtDocAmt" style="HEIGHT: 20px; WIDTH: 250px" tag="22X2Z" ALT="개설금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>원화금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtLocAmt" style="HEIGHT: 20px; WIDTH: 250px" tag="24X2Z" ALT="원화금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								</TR>							
								 
								<TR>
									<TD CLASS=TD5 NOWRAP>근거서류유형</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOLcKind" SIZE=10  MAXLENGTH=5 TAG="22XXXU" ALT="근거서류유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOLcKind" align=top TYPE="BUTTON" onclick="vbscript:btnOLcKindOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtOLcKindNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>분할인도여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=0>
											<TR>
												<TD WIDTH=30%>
													<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="21" VALUE="Y" CHECKED ID="rdoPartailShip1">
													<LABEL FOR="rdoPartailShip1">Y</LABEL>&nbsp;&nbsp;&nbsp;
													<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="21" VALUE="N" ID="rdoPartailShip2">
													<LABEL FOR="rdoPartailShip2">N</LABEL>
												</TD>
											</TR>
										</TABLE>
									</TD>	
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>서류제시기간</TD>
									<TD CLASS=TD6 NOWRAP>
										<Table cellpadding=0 cellspacing=0>
												<TR>
													<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="서류제시기간" NAME="txtFileDt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 style="HEIGHT: 20px; WIDTH: 95px" tag="21X70" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
													</TD>
													<TD NOWRAP>
														&nbsp;Days
													</TD>
													<TD NOWRAP>
													</TD>
												</TR>
										</Table>
									</TD>
									<TD CLASS=TD5 NOWRAP>서류제시기간참조</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFileDtTxt" ALT="서류제시기간참조" TYPE=TEXT MAXLENGTH=34 SIZE=34 TAG="21X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>결제방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10  MAXLENGTH=5 TAG="22XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" onclick="vbscript:btnPayTermsOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>결제기간</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayDur" ALT="결제기간" TYPE=TEXT MAXLENGTH=2 SIZE=12 STYLE="TEXT-ALIGN: right" TAG="24X">&nbsp;Days</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>대금결제참조</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPaytermstxt" ALT="대금결제참조" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP>
									<TD CLASS=TD6 NOWRAP>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>통지일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtAdvDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="통지일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>AMEND일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtAmendDt" CLASS=FPDTYYYYMMDD tag="24" Title="FPDATETIME" ALT="AMEND일"></OBJECT>');</SCRIPT></TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>구매그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=12 MAXLENGTH=4 TAG="24XXXU" ALT="구매그룹">&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>수혜자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=12  MAXLENGTH=10 TAG="24XXXU" ALT="수혜자">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>구매조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=12  MAXLENGTH=4 TAG="24XXXU" ALT="구매조직">&nbsp;<INPUT TYPE=TEXT NAME="txtPurOrgNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>개설의뢰인</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=12  MAXLENGTH=10 TAG="24XXXU" ALT="개설의뢰인">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(2)%>
							</TABLE>
							</DIV>
							<!-- 두번째 탭 내용 -->
							<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>구비서류</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc1" ALT="구비서류" TYPE=TEXT MAXLENGTH=65 SIZE=35 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc2" ALT="구비서류" TYPE=TEXT MAXLENGTH=65 SIZE=35 TAG="21X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc3" ALT="구비서류" TYPE=TEXT MAXLENGTH=65 SIZE=35 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc4" ALT="구비서류" TYPE=TEXT MAXLENGTH=65 SIZE=35 TAG="21X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc5" ALT="구비서류" TYPE=TEXT MAXLENGTH=65 SIZE=35 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설은행앞 정보</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBankTxt" ALT="개설은행앞 정보" TYPE=TEXT MAXLENGTH=70 SIZE=35 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>기타참조사항</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRemark" ALT="기타참조사항" TYPE=TEXT MAXLENGTH=70 SIZE=35 TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>근거서류번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOLCNo" ALT="근거서류번호" TYPE=TEXT MAXLENGTH=18 SIZE=20  TAG="21XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOLcNo" align=top TYPE="BUTTON" onclick="vbscript:btnOLcNoOnClick()"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>근거서류발생일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtOLcOpenDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="근거서류발생일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>근거서류유효일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtOLcExpiryDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="근거서류유효일"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>대행자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10  MAXLENGTH=10 TAG="21XXXU" ALT="대행자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAgent" align=top TYPE="BUTTON" onclick="vbscript:btnAgentOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>제조자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10  MAXLENGTH=10 TAG="21XXXU" ALT="제조자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnManufacturer" align=top TYPE="BUTTON" onclick="vbscript:btnManufacturerOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(11)%>
							</TABLE>
						</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TD WIDTH=10>&nbsp;</TD>
				<TD WIDTH=* ALIGN=RIGHT><A href="VBSCRIPT:CookiePage(1)">LOCAL L/C내역등록</A>&nbsp;|&nbsp;<A href="vbscript:LoadAmend()">AMEND등록</A>&nbsp;|&nbsp;<A href="vbscript:LoadChargeHdr()">경비등록</A></TD>
				<TD WIDTH=10>&nbsp;</TD>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="m3211mb6.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex = -1></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHLcNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHOLCDocNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtGRNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPONoFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="txtMultiDiv" tag="24">
<INPUT TYPE=HIDDEN NAME="txtchkPONoFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHLcFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnXchRtOp" tag="24"> 
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" tabindex = -1></IFRAME>
</DIV>
</BODY>
</HTML>
