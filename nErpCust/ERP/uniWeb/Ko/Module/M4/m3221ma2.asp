<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3221ma2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Import Local L/C Amend 등록 ASP											*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/05/02																*
'*  8. Modified date(Last)  : 2003/05/21																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
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
	
	Const BIZ_PGM_QRY_ID = "m3221mb5.asp"	
	Const BIZ_PGM_SAVE_ID = "m3221mb5.asp"	
	Const BIZ_PGM_DEL_ID = "m3221mb5.asp"	
	Const BIZ_PGM_LCQRY_ID = "m3221mb8.asp"	
	Const LCAMEND_DETAIL_ENTRY_ID = "m3222ma2"
	Const CHARGE_HDR_ENTRY_ID = "m6111ma2"	
	Const BIZ_PGM_CAL_AMT_ID = "m3211mb9"

	Const TAB1 = 1

	Dim lgBlnFlgChgValue					
	Dim lgIntGrpCount				
	Dim lgIntFlgMode				
	Dim lgLCNo						
	
	Dim gSelframeFlg				
	Dim gblnWinEvent		
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
	frm1.txtAmendDt.text = EndDate
	frm1.txtAtExpireDt.text = EndDate
	frm1.txtAtLatestShipDt.text = EndDate
	frm1.txtAmendReqDt.text = EndDate
	frm1.txtAtDocAmt.text = UNIFormatNumber(CStr(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	Call SetToolbar("1110000000001111")
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
		
	iCalledAspName = AskPRAspName("M3221PA2")		
		
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3221PA2", "X")
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
'+	Name : OpenLCRef()																					+
'+	Description : L/C Reference Window Call																+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function OpenLCRef()
	Dim strRet,IntRetCD
	Dim iCalledAspName
		
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "X", "X", "X")
		Exit function
	End If
		
	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True

	iCalledAspName = AskPRAspName("M3211RA2")		
		
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3211RA2", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName,  Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	gblnWinEvent = False

	If strRet = "" Then
		frm1.txtLCAmdNo.focus
		Set gActiveElement = document.activeElement			
		Exit Function
	Else
		Call ggoOper.ClearField(Document, "A")	
		Call SetRadio()
		Call SetDefaultVal
		lgIntFlgMode = Parent.OPMD_CMODE
		Call SetLCRef(strRet)
	End If
End Function
<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetLCRef()  ++++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetLCRef()																					+
'+	Description : Set Return array from S/O Reference Window											+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Function SetLCRef(strRet)
	Dim strVal

	frm1.txtHLCNo.value = UCase(Trim(strRet))
		
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
		ggoOper.FormatFieldByObjectOfCur .txtBeDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtAtDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtAtXchRate, .txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		
	End With

End Sub
<!--
'=============================================  2.5.1 LoadLCAmendDtl()  =================================
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

	WriteCookie "txtLCAmdNo", UCase(Trim(frm1.txtLCAmdNo.value))

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

	WriteCookie "TmpNo", UCase(Trim(frm1.txtLCAmdNo.value))
	WriteCookie "LCAmdNo", UCase(Trim(frm1.txtLCAmdNo.value))		
	WriteCookie "BasNo", UCase(Trim(frm1.txtLCAmdNo.value))
	WriteCookie "Process_Step", "VO"
	WriteCookie "Po_No",	UCase(Trim(frm1.txtLCAmdNo.value))		
	WriteCookie "Pur_Grp",	UCase(Trim(frm1.txtPurGrp.value))	

	PgmJump(CHARGE_HDR_ENTRY_ID)

End Function
		
<!--
'============================================  2.5.1 OpenCookie()  ======================================
-->
Function OpenCookie()
		
	'frm1.txtLCNo.value = ReadCookie("txtLCNo")
	frm1.txtLCAmdNo.value = ReadCookie("txtLCAmdNo")

	If Trim(ReadCookie("txtLCNo")) <> "" Then		
		Call SetLCRef(Trim(ReadCookie("txtLCNo")))
	ElseIf frm1.txtLCAmdNo.value <> "" Then
		Call MainQuery()
	End If
		
	WriteCookie "txtLCNo", ""
	WriteCookie "txtLCAmdNo", ""
		
End Function

<!--
'==============================================  2.5.3 SetRadio()  ======================================
-->
Function SetRadio()
	frm1.rdoAtDocAmt1.checked = True
	frm1.rdoAtPartialShip1.checked = True
End Function
<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
Sub Form_Load()
	
		
	Call LoadInfTB19029	
		
	'Call AppendNumberRange("0","0","99999999")	
	'Call AppendNumberRange("1","0","999999999999")	
		
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")			
	Call SetDefaultVal
	Call InitVariables
		
	gSelframeFlg = TAB1
	Call OpenCookie()
		
End Sub
	
<!--
'=========================================   txtAtXchRate_OnBlur()  ===================================
-->	
Sub txtAtXchRate_OnBlur()
	Err.Clear																			
		
	If	frm1.txtCurrency.value = Parent.gCurrency Then
		frm1.txtAtXchRate.text = 1
	End If	
End Sub	
	
<!--
'=========================================  3.1.2 Form_QueryUnload()  ===================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
	  
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
		If frm1.chkAtTransport.checked = True Then
			Call OpenMinorCd()
		End If
	End Sub
	
<!--
'====================================  Ocx Event  ==================================
-->	
Sub rdoAtPartialShip1_OnPropertyChange()
		lgBlnFlgChgValue = True
End Sub

Sub rdoAtPartialShip2_OnPropertyChange()
		lgBlnFlgChgValue = True
End Sub

Sub rdoAtDocAmt1_OnPropertyChange()
		lgBlnFlgChgValue = True
End Sub

Sub rdoAtDocAmt2_OnPropertyChange()
		lgBlnFlgChgValue = True
End Sub

Sub txtAmendReqDt_Change()
	lgBlnFlgChgValue = True
End Sub	

Sub txtAmendDt_Change()
	lgBlnFlgChgValue = True
End Sub	

Sub txtAtXchRate_Change()
	lgBlnFlgChgValue = True
End Sub	

Sub txtAtDocAmt_Change()
	lgBlnFlgChgValue = True
End Sub	

Sub txtAtLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub	

Sub txtAtExpireDt_Change()
	lgBlnFlgChgValue = True
End Sub	

Sub txtatLatestShipDt_Change()
	lgBlnFlgChgValue = True
End Sub	

Sub txtAtExpireDt_DblClick(Button)
	if Button = 1 then
		frm1.txtAtExpireDt.Action = 7
		Call SetFocusToDocument("M")
   		frm1.txtAtExpireDt.focus
	End if
End Sub

Sub txtAmendAmt_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtAtLatestShipDt_DblClick(Button)
	if Button = 1 then
		frm1.txtAtLatestShipDt.Action = 7
		Call SetFocusToDocument("M")
   		frm1.txtAtLatestShipDt.focus
	End if
End Sub

Sub txtAmendReqDt_DblClick(Button)
	if Button = 1 then
		frm1.txtAmendReqDt.Action = 7
		Call SetFocusToDocument("M")
   		frm1.txtAmendReqDt.focus
	End if
End Sub

Sub txtAmendDt_DblClick(Button)
	if Button = 1 then
		frm1.txtAmendDt.Action = 7
		Call SetFocusToDocument("M")
   		frm1.txtAmendDt.focus
	End if
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

	Call ggoOper.ClearField(Document, "A")				
	Call SetRadio()
	Call ggoOper.LockField(Document, "N")				
		
	Call SetDefaultVal
	Call InitVariables
		
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
		
	If frm1.txtLCAmdNo.value  <> "" Then
		If lgBlnFlgChgValue = False Then		
			IntRetCD = DisplayMsgBox("900001", "X", "X", "X")		
		    Exit Function
		End If
	End If
	If	frm1.txtCurrency.value = Parent.gCurrency Then
		frm1.txtAtXchRate.text = 1
	End If
		
	If Not chkField(Document, "2") Then		
		Exit Function
	End If
		
	
	'If UNICDbl(frm1.txtAtXchRate.Text) <= 0 Then
	'	Call DisplayMsgBox("200095", "X", "X", "X")
	'	Exit Function
	'End if
		
	If UniConvDateToYYYYMMDD(frm1.txtAmendDt.Text,Parent.gDateFormat,"") < UniConvDateToYYYYMMDD(frm1.txtAmendReqDt.Text,Parent.gDateFormat,"") And (frm1.txtAmendDt.Text<>"" Or frm1.txtAmendDt.Text <> Null) Then
		Call DisplayMsgBox("970023", "X","AMEND일","AMEND신청일")			
		frm1.txtAmendDt.focus
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
	If	frm1.txtCurrency.value = Parent.gCurrency Then
		frm1.txtAtXchRate.value = 1
		ggoOper.SetReqAttr	frm1.txtAtXchRate , "Q"
	Else 
		ggoOper.SetReqAttr	frm1.txtAtXchRate , "N"
	End if							
	Call SetToolbar("1111100000011111")
End Function
<!--
'=============================================  5.2.4 DbQueryOk()  ======================================
-->
Function RefOk()													
	Call SetToolbar("1110100000001111")
	If	frm1.txtCurrency.value = Parent.gCurrency Then
		frm1.txtAtXchRate.value = 1
		ggoOper.SetReqAttr	frm1.txtAtXchRate , "Q"
	Else 
		ggoOper.SetReqAttr	frm1.txtAtXchRate , "N"
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>LOCAL L/C AMEND</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							</TR>
						</TABLE>
					</TD>					
					<TD WIDTH=* align=right><A href="vbscript:OpenLCRef">LOCAL L/C참조</A></TD>
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
										<TD CLASS=TD5 NOWRAP>LOCAL L/C AMEND관리번호</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCAmdNo" SIZE=29 MAXLENGTH=18 TAG="12XXXU" ALT="LOCAL L/C AMEND관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCAmdNo" ALIGN=top TYPE="BUTTON" onclick="vbscript:btnLCAmdNoOnClick()"></TD>
										<TD CLASS=TD6 NOWRAP></TD>
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
							<!-- 첫번째 탭 내용 -->
							<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>LOCAL L/C AMEND관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCAmdNo1" SIZE=32 MAXLENGTH=18 TAG="25XXXU" ALT="LOCAL L/C AMEND관리번호"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>LOCAL L/C번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LOCAL LC번호" TYPE=TEXT  SIZE=27 MAXLENGTH=35 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24" ></TD>
									<TD CLASS=TD5 NOWRAP>수혜자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  ALT="수혜자" MAXLENGTH=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>AMEND신청일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtAmendReqDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="AMEND신청일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>AMEND일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtAmendDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="AMEND일"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>금액변경</LABEL></TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtDocAmt" TAG="2X" VALUE="I" CHECKED ID="rdoAtDocAmt1">
										<LABEL FOR="rdoAtDocAmt1">INCREASE BY</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtDocAmt" VALUE="D" TAG="2X" ID="rdoAtDocAmt2">
										<LABEL FOR="rdoAtDocAmt2">DECREASE BY</LABEL>
									</TD>
									<TD CLASS=TD5 NOWRAP>환율</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtAtXchRate" style="HEIGHT: 20px; WIDTH: 230px" tag="22X5Z" ALT="환율" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtAmendAmt" style="HEIGHT: 20px; WIDTH: 235px" tag="21X2Z" ALT="개설금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>AMEND금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD >
													<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU">
												</TD>
												<TD >
													&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT NAME="txtAtDocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 150px" tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>																								
											</TR>
										</Table>										
									</TD>
									<TD CLASS=TD5 NOWRAP>개설금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtBeDocAmt" style="HEIGHT: 20px; WIDTH: 230px" tag="24X2Z" ALT="개설금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>	
											</TR>
										</TABLE>	
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>AMEND원화금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtAtLocAmt" style="HEIGHT: 20px; WIDTH: 235px" tag="24X2Z" ALT="AMEND원화금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>	
											</TR>
										</TABLE>	
									</TD>
									<TD CLASS=TD5 NOWRAP>개설원화금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtBeLocAmt" style="HEIGHT: 20px; WIDTH: 230px" tag="24X2Z" ALT="개설원화금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>	
											</TR>
										</TABLE>	
									</TD>									
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>유효기일연장</TD>
									<TD CLASS=TD6 NOWRAP>												
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=유효기일연장 NAME="txtAtExpireDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>												
									</TD>				 
									<TD CLASS=TD5 NOWRAP>변경전 유효기일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT NAME="txtBeExpireDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>인도기일연장</TD>
									<TD CLASS=TD6 NOWRAP>										
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=인도기일연장 NAME="txtAtLatestShipDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>변경전 인도기일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT NAME="txtBeLatestShipDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>분할인도여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtPartialShip" TAG="21X" VALUE="Y" CHECKED ID="rdoAtPartialShip1"><LABEL FOR="rdoAtPartialShip1">Y</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtPartialShip" VALUE="N" TAG="21X" ID="rdoAtPartialShip2"><LABEL FOR="rdoAtPartialShip2">N</LABEL>
									</TD>
									<TD CLASS=TD5 NOWRAP>변경전분할인도여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBePartialShip" SIZE=13 MAXLENGTH=10  STYLE="TEXT-ALIGN:left" TAG="24" ALT="변경전분할인도여부"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10 MAXLENGTH=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>통지은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdvBank" SIZE=10 MAXLENGTH=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtAdvBankNm" SIZE=20 TAG="24"></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>개설일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtOpenDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="개설일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>개설자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10  TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24" ALT="개설자"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>구매그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10 MAXLENGTH=4 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="24" ALT="구매그룹"></TD>
									<TD CLASS=TD5 NOWRAP>구매조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=10  MAXLENGTH=4 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtPurOrgNm" SIZE=20 TAG="24" ALT="구매조직"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>기타참조사항</TD>
									<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemark" ALT="기타참조사항" TYPE=TEXT MAXLENGTH=70 SIZE=90 TAG="21X"></TD>									
								</TR>
								<%Call SubFillRemBodyTD5656(5)%>
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
				<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadLCAmendDtl()">AMEND내역등록</A></TD>				
				<TD WIDTH=10>&nbsp;</TD>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex=-1></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHLCNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPONo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtLCAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHLCAmdNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtIncAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtDecAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtXchRate" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHExpiryDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHLatestShipDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPartialShip" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHTranshipment" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" tabindex=-1></IFRAME>
</DIV>
</BODY>
</HTML>
