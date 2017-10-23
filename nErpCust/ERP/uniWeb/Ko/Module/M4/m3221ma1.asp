<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3221ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open L/C Amend 등록 ASP													*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/03/31																*
'*  8. Modified date(Last)  : 2003/05/21																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      : Lee Eun Hee																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/31 : 화면 design												*
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 

<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>
<Script Language="VBS">

Option Explicit	


Const BIZ_PGM_QRY_ID = "m3221mb1.asp" 
Const BIZ_PGM_SAVE_ID = "m3221mb2.asp" 
Const BIZ_PGM_DEL_ID = "m3221mb2.asp" 
Const BIZ_PGM_LCQRY_ID = "m3221mb4.asp"		 
Const LCAMEND_DETAIL_ENTRY_ID = "m3222ma1"
Const CHARGE_HDR_ENTRY_ID = "m6111ma2"		 
	
Const gstrTransportMajor 	= "B9009"				'운송방법 
Const gstrLoadingPortMajor	= "B9092"				'선적항	
Const gstrDisChgePortMajor	= "B9092"				'도착항 

Const TAB1 = 1
Const TAB2 = 2


Dim lgBlnFlgChgValue			 
Dim lgIntGrpCount			 
Dim lgIntFlgMode					 
Dim lgLCNo						 
	
Dim gSelframeFlg					 
Dim gblnWinEvent					 
	
Dim StartDate
Dim EndDate

EndDate = "<%=GetSvrDate%>"
EndDate = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)

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
	frm1.txtAmendDt.text 		= EndDate
	frm1.txtBeExpiryDt.text 	= EndDate
	frm1.txtBeLatestShipDt.text = EndDate
	frm1.txtOpenDt.text 		= EndDate
	frm1.txtAmendAmt.text 		= UNIFormatNumber(CStr(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	frm1.txtAtDocAmt.text 		= UNIFormatNumber(CStr(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	frm1.txtBeDocAmt.text 		= UNIFormatNumber(CStr(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	Call SetToolbar("1110000000001111")	
	Call ClickTab1()
	frm1.txtLCAmdNo.focus
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
'++++++++++++++++++++++++++++++++++++++++++++++  OpenLCAmdNoPop()  ++++++++++++++++++++++++++++++++++++++
'+	Name : OpenLCAmdNoPop()																				+
'+	Description : Master L/C Amend No PopUp Call														+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenLCAmdNoPop()
	Dim strRet,IntRetCD
	Dim iCalledAspName
		
	If gblnWinEvent = True Or UCase(frm1.txtLCAmdNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("M3221PA1")		
		
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3221PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
		frm1.txtLCAmdNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtLCAmdNo.value = strRet
		frm1.txtLCAmdNo.focus
		Set gActiveElement = document.activeElement
	End If	
End Function
	
<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenLCRef()  ++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenLCRef()
	Dim strRet
	Dim IntRetCD
	Dim arrParam(1)
	Dim iCalledAspName
		
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "X", "X", "X")
		Exit function
	End If

	arrParam(0) = ""
		
	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True

	iCalledAspName = AskPRAspName("M3211RA1")		
		
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3211RA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	gblnWinEvent = False

	If strRet = "" Then			
		Exit Function
	Else			
		Call ggoOper.ClearField(Document, "A")							 
		Call SetRadio()
		Call SetDefaultVal
		Call SetLCRef(strRet)
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenMinorCd()  +++++++++++++++++++++++++++++++++++++++++
-->
Function OpenMinorCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "운송방법"							 
	arrParam(1) = "B_Minor"									 
	arrParam(2) = Trim(frm1.txtAtTransport.Value)				 
	arrParam(4) = "MAJOR_CD= " & FilterVar(gstrTransportMajor, "''", "S") & ""		 
	arrParam(5) = "운송방법"								 

	arrField(0) = "Minor_CD"									 
	arrField(1) = "Minor_NM"								 

	arrHeader(0) = "운송방법"								 
	arrHeader(1) = "운송방법명"								 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtAtTransport.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtAtTransport.Value = arrRet(0)
		frm1.txtAtTransportNm.Value = arrRet(1)
		frm1.txtAtTransport.focus
		Set gActiveElement = document.activeElement
		lgBlnFlgChgValue = True
	End If
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenDischgePort()  +++++++++++++++++++++++++++++++++++++++++
-->
Function OpenDischgePort()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "도착항"							
	arrParam(1) = "B_Minor"								
	arrParam(2) = Trim(frm1.txtAtDischgePort.Value)		
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
		frm1.txtAtDischgePort.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtAtDischgePort.Value 	= arrRet(0)
		frm1.txtAtDischgePortNm.Value 	= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtAtDischgePort.focus
		Set gActiveElement = document.activeElement
	End If
		
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenLoadingPort()  +++++++++++++++++++++++++++++++++++++++++
-->
Function OpenLoadingPort()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "선적항"								
	arrParam(1) = "B_Minor"									
	arrParam(2) = Trim(frm1.txtAtLoadingPort.Value)			
'		arrParam(3) = Trim(frm1.txtAtLoadingPortNm.Value)		
	arrParam(4) = "MAJOR_CD= " & FilterVar(gstrLoadingPortMajor, "''", "S") & ""	
	arrParam(5) = "선적항"								

	arrField(0) = "Minor_CD"								
	arrField(1) = "Minor_NM"								

	arrHeader(0) = "선적항"								
	arrHeader(1) = "선적항명"							

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtAtLoadingPort.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtAtLoadingPort.Value = arrRet(0)
		frm1.txtAtLoadingPortNm.Value = arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtAtLoadingPort.focus
		Set gActiveElement = document.activeElement
	End If
		
End Function
<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetLCRef()  ++++++++++++++++++++++++++++++++++++++++++++
-->
Function SetLCRef(strRet)
	Dim strVal

	if LayerShowHide(1) =false then
	    exit Function
	end if

	strVal = BIZ_PGM_LCQRY_ID & "?txtMode=" & Parent.UID_M0001		
	strVal = strVal & "&txtLCNo=" & UCase(Trim(strRet))
		
	Call RunMyBizASP(MyBizASP, strVal)						
End Function

'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtAmendAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtAtDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtBeDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtXchRt, .txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		
	End With

End Sub
<!--
'============================================ ValidDateCheckLocal()  ======================================
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
'=============================================  2.5.1 LoadLCAmendDtl()  ======================================
-->
Function LoadLCAmendDtl()
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

	WriteCookie "LCAmdNo", UCase(Trim(frm1.txtLCAmdNo.value))
		
	PgmJump(LCAMEND_DETAIL_ENTRY_ID)

End Function
	
<!--
'=============================================  2.5.1 LoadChargeHdr()  ======================================
-->
Function LoadChargeHdr()
	Dim strHdrOpenParam
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
    	
    WriteCookie "Process_Step" , "VA"
	WriteCookie "Po_No" , Trim(frm1.txtLCAmdNo.value)
	WriteCookie "Pur_Grp", Trim(frm1.txtPurGrp.Value)
	WriteCookie "Po_Cur", Trim(frm1.txtCurrency.Value)
	'WriteCookie "Po_Xch", Trim(frm1.hdnXch.Value)
		
	PgmJump(CHARGE_HDR_ENTRY_ID)

End Function
	
<!--
'============================================  2.5.1 OpenCookie()  ======================================
-->
Function OpenCookie()
	Dim strLCNo, strLCAmdNo
	strLCNo = ReadCookie("LCNo")
	strLCAmdNo = ReadCookie("LCAmdNo")
	frm1.txtLCNo.value = strLCNo
	frm1.txtLCAmdNo.value = strLCAmdNo

	WriteCookie "LCNo", ""
	WriteCookie "LCAmdNo", ""

	If frm1.txtLCNo.value <> "" Then
		Call SetLCRef(frm1.txtLCNo.value)
	ElseIf frm1.txtLCAmdNo.value <> "" Then
		Call MainQuery()
	End If
End Function

<!--
'==============================================  2.5.3 SetRadio()  ======================================
-->
Function SetRadio()
	frm1.rdoAtDocAmt1.checked = True
	frm1.rdoAtTranshipment1.checked = True
	frm1.rdoAtPartialShip1.checked = True
	frm1.rdoAtTransfer1.checked = True
End Function
<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")		
	Call SetDefaultVal
	Call SetToolbar("1110100000001111")		
	Call InitVariables
		
	gIsTab     = "Y"
	gTabMaxCnt = 2

	Call OpenCookie()

	gSelframeFlg = TAB1
		
End Sub
<!--
'=========================================  3.1.2 Form_QueryUnload()  ===================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub
<!--
'==========================================================================================
'   Event Name : Ocx Event  	 
'==========================================================================================
-->
Sub txtAmendReqDt_DblClick(Button)
	If Button = 1 then
		frm1.txtAmendReqDt.Action = 7
		Call SetFocusToDocument("M")
   		frm1.txtAmendReqDt.focus
	End If
End Sub

Sub txtAmendReqDt_Change()
	lgBlnFlgChgValue = true	
End Sub

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

Sub txtAtExpiryDt_DblClick(Button)
	if Button = 1 then
		frm1.txtAtExpiryDt.Action = 7
		Call SetFocusToDocument("M")
   		frm1.txtAtExpiryDt.focus
	End if
End Sub

Sub txtAtExpiryDt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtAtLatestShipDt_DblClick(Button)
	if Button = 1 then
		frm1.txtAtLatestShipDt.Action = 7
		Call SetFocusToDocument("M")
   		frm1.txtAtLatestShipDt.focus
	End if
End Sub

Sub txtAtLatestShipDt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtAmendAmt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtXchRt_Change()
	lgBlnFlgChgValue = true	
End Sub
<!--
'======================================  3.2.1 btnLCAmdNoOnClick()  ====================================
-->
Sub btnLCAmdNoOnClick()
	Call OpenLCAmdNoPop()
End Sub
<!--
'====================================  3.2.2 btnAtTransportOnClick()  ==================================
-->
Sub btnAtTransportOnClick()
	Call OpenMinorCd()
End Sub
<!--
'==================================  3.2.4 rdoAtDocAmt_OnPropertyChange()  ==============================
-->
Sub rdoAtDocAmt1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoAtDocAmt2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

<!--
'===============================  3.2.8 rdoAtTranshipment_OnPropertyChange()  ===========================
-->
Sub rdoAtTranshipment1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoAtTranshipment2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

<!--
'===============================  3.2.10 rdoAtPartialShip_OnPropertyChange()  ===========================
-->
Sub rdoAtPartialShip1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoAtPartialShip2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

<!--
'==================================  3.2.12 rdoAtTransfer_OnPropertyChange()  ===========================
-->
Sub rdoAtTransfer1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoAtTransfer2_OnPropertyChange()
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

	Call ggoOper.ClearField(Document, "2")						
	Call SetRadio()
	Call SetDefaultVal
	Call InitVariables											

	If Not chkField(Document, "1") Then							
		Exit Function
	End If

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
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ClickTab1()
	Call ggoOper.ClearField(Document, "A")	
	Call SetRadio()
	Call ggoOper.LockField(Document, "N")	
	Call SetDefaultVal
	Call InitVariables						
		
	frm1.txtLCAmdNo.focus
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
		
	If frm1.txtLCAmdNo.value  <> "" Then
		If lgBlnFlgChgValue = False Then			
			IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		    Exit Function
		End If
	End If
		
    If Not chkField(Document, "2") Then               
        If gPageNo > 0 Then
            gSelframeFlg = gPageNo
        End If
	        
        Exit Function
    End If
	    
   ' If UNICDbl(frm1.txtXchRt.Text) <= 0 Then
	'	Call DisplayMsgBox("200095", "X", "X", "X")
	'	Call ClickTab1()
	'	frm1.txtXchRt.focus
	'	Set gActiveElement = document.activeElement		
	'	Exit Function
'	End if
	    
    If ValidDateCheckLocal(frm1.txtAmendReqDt, frm1.txtAmendDt) = False Then Exit Function
    If ValidDateCheckLocal(frm1.txtAmendDt, frm1.txtAtLatestShipDt) = False Then Exit Function
    If ValidDateCheckLocal(frm1.txtAtLatestShipDt, frm1.txtAtExpiryDt) = False Then Exit Function
	    
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
'=============================================  5.2.1 DbLCQuery()  ========================================
-->
Function DbLCQuery()
	Dim strVal

	Err.Clear													

	DbLCQuery = False											

	strVal = BIZ_PGM_LCHQRY_ID & "?txtMode=" & Parent.UID_M0001		
	strVal = strVal & "&txtLCNo=" & Trim(ReadCookie("LCNo"))	

	Call RunMyBizASP(MyBizASP, strVal)							
	
	DbLCQuery = True											
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
	strVal = strVal & "&txtLCAmdNo=" & Trim(frm1.txtLCAmdNo.value)	

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
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
	End With
		
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
		
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
	strVal = strVal & "&txtLCAmdNo=" & Trim(frm1.txtLCAmdNo.value)	

	Call RunMyBizASP(MyBizASP, strVal)								

	DbDelete = True	
														
End Function
		
<!--
'=============================================  5.2.4 DbQueryOk()  ======================================
-->
Function DbQueryOk()												
		
	lgIntFlgMode = Parent.OPMD_UMODE										

	lgBlnFlgChgValue = False

	Call ggoOper.LockField(Document, "Q")							
	Call SetToolbar("11111000000111")
		
	Call ClickTab1()
	Call RefOk()
	frm1.txtLCAmdNo.focus 
	Set gActiveElement = document.activeElement
		
End Function
<!--
'=============================================  5.2.4 RefOk()  ======================================
-->
Function RefOk()
	Call SetToolbar("1111100000001111")	
	If	frm1.txtCurrency.value = Parent.gCurrency Then
		frm1.txtXchRt.value = 1
		ggoOper.SetReqAttr	frm1.txtXchRt , "Q"
	Else 
		ggoOper.SetReqAttr	frm1.txtXchRt , "N"
	End if	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C AMEND</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C AMEND 기타</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenLCRef">L/C참조</A></TD>
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
										<TD CLASS=TD5 NOWRAP>L/C AMEND 관리번호</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCAmdNo" SIZE=32 MAXLENGTH=18  TAG="12XXXU" ALT="L/C AMEND 관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCAmdNo" ALIGN=top TYPE="BUTTON" onclick="vbscript:btnLCAmdNoOnClick()"></TD>
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
									<TD CLASS=TD5 NOWRAP>L/C AMEND 관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCAmdNo1" SIZE=34 MAXLENGTH=18  TAG="25XXXU" ALT="L/C AMEND 관리번호"></TD>
									<TD CLASS=TD5>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>L/C번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LC번호" TYPE=TEXT SIZE=30 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24" ></TD>
									<TD CLASS=TD5 NOWRAP>수출자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  ALT="수출자" TAG="24XXXU">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>AMEND신청일</TD>
									<TD CLASS=TD6 NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=AMEND신청일 NAME="txtAmendReqDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD NOWRAP>
													&nbsp;AMEND일&nbsp;
												</TD>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=AMEND일 NAME="txtAmendDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD CLASS="TD6" WIDTH="*" NOWRAP>&nbsp;
												</TD>
											</TR>
										</Table>
									</TD>
									<TD CLASS=TD5 NOWRAP>개설일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=개설일 NAME="txtOpenDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>금액변경</TD>
									<TD CLASS=TD6 NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD CLASS="TD6" NOWRAP>
													<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtDocAmt" TAG="21X" VALUE="I" CHECKED ID="rdoAtDocAmt1"><LABEL FOR="rdoAtDocAmt1">INCREASE BY</LABEL>
													<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtDocAmt" VALUE="D" TAG="21X" ID="rdoAtDocAmt2"><LABEL FOR="rdoAtDocAmt2">DECREASE BY</LABEL>
												</TD>
												<TD CLASS="TD6" WIDTH="*" NOWRAP>&nbsp;
												</TD>
											</TR>
										</Table>
									</TD>
									<TD CLASS=TD5 NOWRAP>환율</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT NAME="txtXchRt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 235px" tag="22X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT NAME="txtAmendAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 250px" tag="21X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>AMEND금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD NOWRAP>
													<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10  MAXLENGTH=3 TAG="24XXXU">&nbsp;
												</TD>
												<TD NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT NAME="txtAtDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 165px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</Table>
									</TD>
									<TD CLASS=TD5 NOWRAP>개설금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT NAME="txtBeDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 235px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>유효기일연장</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=유효기일연장 NAME="txtAtExpiryDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>변경전유효기일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT NAME="txtBeExpiryDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>선적기일연장</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=선적기일연장 NAME="txtAtLatestShipDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>변경전선적기일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT NAME="txtBeLatestShipDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>환적여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtTranshipment" TAG="21X" VALUE="Y" CHECKED ID="rdoAtTranshipment1"><LABEL FOR="rdoAtTranshipment1">Y</LABEL><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtTranshipment" VALUE="N" TAG="21X" ID="rdoAtTranshipment2"><LABEL FOR="rdoAtTranshipment2">N</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>변경전환적여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeTranshipment" SIZE=10 TAG="24X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>분할선적여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtPartialShip" TAG="21X" VALUE="Y" CHECKED ID="rdoAtPartialShip1"><LABEL FOR="rdoAtPartialShip1">Y</LABEL><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtPartialShip" VALUE="N" TAG="21X" ID="rdoAtPartialShip2"><LABEL FOR="rdoAtPartialShip2">N</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>변경전분할선적</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBePartialShip" SIZE=10 TAG="24X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>양도여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtTransfer" TAG="21X" VALUE="Y" CHECKED ID="rdoAtTransfer1"><LABEL FOR="rdoAtTransfer1">Y</LABEL><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtTransfer" VALUE="N" TAG="21X" ID="rdoAtTransfer2"><LABEL FOR="rdoAtTransfer2">N</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>변경전양도여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeTransfer" SIZE=10 TAG="24X"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>운송방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAtTransport" SIZE=10  MAXLENGTH=5 TAG="22XXXU" ALT="운송방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAtTransport" align=top TYPE="BUTTON"  onclick="vbscript:btnAtTransportOnClick()">
														 <INPUT TYPE=TEXT NAME="txtAtTransportNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>변경전운송방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeTransport" SIZE=10 TAG="24">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtBeTransportNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>선적항</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAtLoadingPort" MAXLENGTH=5 SIZE=10 TAG="22XXXU" ALT="선적항"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAtLoadingPort" align=top TYPE="BUTTON" OnClick="OpenLoadingPort()">
														 <INPUT TYPE=TEXT NAME="txtAtLoadingPortNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>변경전선적항</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeLoadingPort" SIZE=10 TAG="24">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtBeLoadingPortNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>도착항</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAtDischgePort" MAXLENGTH=5 SIZE=10 TAG="22XXXU" ALT="도착항"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAtDischgePort" align=top TYPE="BUTTON" OnClick="OpenDischgePort()">
														 <INPUT TYPE=TEXT NAME="txtAtDischgePortNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>변경전도착항</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeDischgePort" SIZE=10 TAG="24">&nbsp;	
														 <INPUT TYPE=TEXT NAME="txtBeDischgePortNm" SIZE=20 TAG="24"></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(2)%>
							</TABLE>
						</DIV>
						<!-- 두번째 탭 내용 -->
						<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>통지은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdvBank" SIZE=10  TAG="24XXXU">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtAdvBankNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>개설은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10  TAG="24XXXU">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>구매그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10  TAG="24XXXU">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>수입자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10  TAG="24XXU">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>구매조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=10  TAG="24XXXU">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtPurOrgNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>기타참조</TD>
									<TD CLASS=TD6 Colspan=3 WIDTH=100% NOWRAP><INPUT NAME="txtRemark" ALT="기타참조" TYPE=TEXT MAXLENGTH=70 SIZE=90 TAG="21X"></TD>
								</TR><!--
								<TR>
									<TD CLASS=TD5 NOWRAP>통지번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAdvNo" ALT="통지번호" TYPE=TEXT MAXLENGTH=35 SIZE=34 STYLE="Text-Transform: uppercase" TAG="21X"></TD>
									<TD CLASS=TD5 NOWRAP>선통지참조사항</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPreAdvRef" ALT="선통지참조사항" TYPE=TEXT MAXLENGTH=35 SIZE=34 STYLE="Text-Transform: uppercase" TAG="21X"></TD>
								</TR>-->
								<%Call SubFillRemBodyTD5656(10)%>
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
				<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadLCAmendDtl()">AMEND내역등록</A><!-- | <A href="vbscript:LoadChargeHdr()">경비등록</A>--></TD>
				<TD WIDTH=10>&nbsp;</TD>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtLCNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPONo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHLCAmdNo" tag="24">
<!--<INPUT TYPE=HIDDEN NAME="txtBeTransport" tag="24">-->
<INPUT TYPE=HIDDEN NAME="txtIncAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtDecAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHExpiryDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHLatestShipDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHTranshipment" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPartialShip" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHTransfer" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHTransport" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHLoadingPort" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHDischgePort" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
