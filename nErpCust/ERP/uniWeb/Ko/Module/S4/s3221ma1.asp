<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3221ma1.asp																*
'*  4. Program Name         : L/C Amend 등록															*
'*  5. Program Desc         : L/C Amend 등록															*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/31																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/31 : 화면 design												*
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

EndDate = UNIConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID				= "s3221mb1.asp"			
Const BIZ_PGM_LCQRY_ID			= "s3221mb3.asp"		
Const LCAMEND_DETAIL_ENTRY_ID	= "s3222ma1"	
Const EXPORT_CHARGE_ENTRY_ID	= "s6111ma1"	

Const TAB1 = 1
Const TAB2 = 2

Const gstrTransportMajor = "B9009"

Dim gSelframeFlg		
Dim gblnWinEvent				

'========================================================================================================
Function InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE							
	lgBlnFlgChgValue = False							
	lgIntGrpCount = 0									
		
	gblnWinEvent = False
End Function

'========================================================================================================
Sub SetDefaultVal()
	frm1.txtAmendDt.text = EndDate
	frm1.txtAmendAmt.text = UNIFormatNumber(CStr(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)

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
	frm1.txtLCAmdNo.focus
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
		
	Call changeTabs(TAB2)
	frm1.txtLCAmdNo.focus
		
	gSelframeFlg = TAB2
End Function

'========================================================================================================
Function OpenLCAmdNoPop()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Or UCase(frm1.txtLCAmdNo.className) = "PROTECTED" Then Exit Function
		
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("s3221pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3221pa1", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
		frm1.txtLCAmdNo.focus
		Exit Function
	Else
		Call SetLCAmdNo(strRet)
		
	End If	
End Function

'========================================================================================================
Function OpenLCRef()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If lgIntFlgMode = parent.OPMD_UMODE Then 
		Call DisplayMsgBox("200005", "x", "x", "x")
		Exit function
	End If 
		
	iCalledAspName = AskPRAspName("s3211ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3211ra1", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	If strRet = "" Then
		Exit Function
	Else
			
		Call ggoOper.ClearField(Document, "A")										
		Call SetRadio()
		Call InitVariables												
		Call SetDefaultVal
			
		Call SetLCRef(strRet)

	End If
End Function

'========================================================================================================
Function OpenMinorCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "운송방법"							
	arrParam(1) = "B_Minor"									
	arrParam(2) = Trim(frm1.txtAtTransport.Value)			
	arrParam(3) = ""										
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
		Exit Function
	Else
		Call SetMinorCd(arrRet)
	End If
End Function

'========================================================================================================
Function OpenPort(strMinorCD, strMinorNM, strPopNm, iwhere)
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
		
		arrParam(0) = strPopNm						
		arrParam(1) = "B_MINOR"						
		arrParam(2) = Trim(strMinorCD)				
		arrParam(3) = ""							
		arrParam(4) = "MAJOR_CD = " & FilterVar("B9092", "''", "S") & ""			
		arrParam(5) = strPopNm						
		
		arrField(0) = "Minor_CD"					
		arrField(1) = "Minor_NM"					
	    
		arrHeader(0) = strPopNm						
		arrHeader(1) = strPopNm & "명"			

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
Function SetLCAmdNo(strRet)
	frm1.txtLCAmdNo.value = strRet
	frm1.txtLCAmdNo.focus
End Function

'========================================================================================================
Function SetLCRef(strRet)
	Dim strVal

					
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If

	strVal = BIZ_PGM_LCQRY_ID & "?txtMode=" & parent.UID_M0001						
	strVal = strVal & "&txtLCNo=" & UCase(Trim(strRet))

	Call RunMyBizASP(MyBizASP, strVal)										

	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function SetMinorCd(arrRet)
	frm1.txtAtTransport.Value = arrRet(0)
	frm1.txtAtTransportNm.Value = arrRet(1)
	frm1.txtAtTransport.focus
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function SetOpenPort(iwhere, arrRet)
	
	Select Case iwhere				
		Case 0
			frm1.txtAtLoadingPort.Value = arrRet(0)
			frm1.txtAtLoadingPortNm.Value = arrRet(1)	
			frm1.txtAtLoadingPort.focus
		Case 1
			frm1.txtAtDischgePort.Value = arrRet(0)
			frm1.txtAtDischgePortNm.Value = arrRet(1)	
			frm1.txtAtDischgePort.focus
	End Select			
					
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function LoadLCAmendDtl()
End Function

'========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877					
	Dim strTemp, arrVal

	Select Case Kubun
		
		Case 1
			WriteCookie CookieSplit , frm1.txtLCAmdNo.value

		Case 0

			strTemp = ReadCookie(CookieSplit)
					
			If strTemp = "" then Exit Function
					
			frm1.txtLCAmdNo.value =  strTemp
				
			If Err.number <> 0 Then
				Err.Clear
				WriteCookie CookieSplit , ""
				Exit Function 
			End If
				
			Call MainQuery()
							
			WriteCookie CookieSplit , ""
				
		Case 2
			WriteCookie CookieSplit , "EL" & parent.gRowSep & frm1.txtSalesGroup.value & parent.gRowSep & frm1.txtSalesGroupNm.value & parent.gRowSep & frm1.txtLCAmdNo.value 
	End Select

End Function

'========================================================================================================
Function SetRadio()
	Dim blnOldFlag

	blnOldFlag = lgBlnFlgChgValue

	lgBlnFlgChgValue = blnOldFlag
End Function

'========================================================================================================
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
		Call PgmJump(LCAMEND_DETAIL_ENTRY_ID)
	Case 1
		Call CookiePage(2)
		Call PgmJump(EXPORT_CHARGE_ENTRY_ID)
	End Select
End Function

'========================================================================================================
Function ProtectXchRate()
	If frm1.txtAtCurrency.value = parent.gCurrency Then
		Call ggoOper.SetReqAttr(frm1.txtAtXchRate, "Q")
	End If	
End Function

'========================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		'개설금액 
		ggoOper.FormatFieldByObjectOfCur .txtAtDocAmt, .txtAtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		'변경전개설금액 
		ggoOper.FormatFieldByObjectOfCur .txtBeDocAmt, .txtBeCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		'변경금액 
		ggoOper.FormatFieldByObjectOfCur .txtAmendAmt, .txtBeCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		
		'변경환율 
		ggoOper.FormatFieldByObjectOfCur .txtAtXchRate, .txtAtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		
		'변경전환율 
		ggoOper.FormatFieldByObjectOfCur .txtBeXchRate, .txtBeCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
				
	End With
End Sub

'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029															
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")										
	Call SetDefaultVal
	<% '----------  Coding part  ------------------------------------------------------------- %>

	Call SetToolBar("11100000000011")											
	Call InitVariables
	Call CookiePage(0)
	Call changeTabs(TAB1)
		
	gSelframeFlg = TAB1
	frm1.txtLCAmdNo.focus
	Set gActiveElement = document.activeElement 
    gIsTab     = "Y" 
    gTabMaxCnt = 2   
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub
	

'========================================================================================================
Sub btnLCAmdNoOnClick()
	Call OpenLCAmdNoPop()
End Sub


'========================================================================================================
Sub btnAtTransportOnClick()
	Call OpenMinorCd()
End Sub
	

'========================================================================================================
Sub btnLoadingPortOnClick()
	Call OpenPort(frm1.txtAtLoadingPort.value, frm1.txtAtLoadingPortNm.value, "선적항", 0)
End Sub

'========================================================================================================
Sub btnDischgePortOnClick()
	Call OpenPort(frm1.txtAtDischgePort.value, frm1.txtAtDischgePortNm.value, "도착항", 1)
End Sub


'========================================================================================================
Sub txtAtDocAmt_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub rdoAtTranshipment1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoAtTranshipment2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub rdoAtPartialShip1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoAtPartialShip2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub rdoAtTransfer1_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoAtTransfer2_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub

Sub rdoAtDocAmt1_OnPropertyChange()
	lgBlnFlgChgValue = True
	Call txtAmendAmt_Change()
End Sub

Sub rdoAtDocAmt2_OnPropertyChange()
	lgBlnFlgChgValue = True
	Call txtAmendAmt_Change()
End Sub

'=======================================================================================================
Sub txtAmendDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAmendDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtAmendDt.Focus
    End If
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtAtExpireDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAtExpireDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtAtExpireDt.Focus
    End If
End Sub

'=======================================================================================================
Sub txtatLatestShipDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtatLatestShipDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtatLatestShipDt.Focus
    End If
End Sub
																
'========================================================================================================
Sub txtAtDocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtAtXchRate_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtAtExpireDt_Change()
	lgBlnFlgChgValue = True
End Sub


'=======================================================================================================
Sub txtAtLatestShipDt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtAmendDt_Change()
	lgBlnFlgChgValue = True
End Sub	

'=======================================================================================================
Sub txtAmendAmt_Change()
	lgBlnFlgChgValue = True
	
	Dim arrAmt
        
    arrAmt = UNICDbl(frm1.txtBeDocAmt.text)	
        
	If frm1.rdoAtDocAmt1.checked = True Then		
		frm1.txtAtDocAmt.text = UNIFormatNumberByCurrecny(arrAmt + UNICDbl(frm1.txtAmendAmt.text),frm1.txtAtCurrency.value, parent.ggAmtOfMoneyNo)
	Else		
		frm1.txtAtDocAmt.text = UNIFormatNumberByCurrecny(arrAmt - UNICDbl(frm1.txtAmendAmt.text),frm1.txtAtCurrency.value, parent.ggAmtOfMoneyNo)
	End If	
	
End Sub	

'========================================================================================================
Function FncQuery()
	Dim IntRetCD
	
	FncQuery = False											

	Err.Clear													

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "2")						
	Call InitVariables											

	If Not chkField(Document, "1") Then							
		Exit Function
	End If

	Call DbQuery()												

	FncQuery = True												
End Function
	

'========================================================================================================
Function FncNew()
	Dim IntRetCD 

	FncNew = False                                              

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "A")								
	Call ggoOper.LockField(Document, "N")						
	Call SetDefaultVal
	Call SetRadio()
	Call SetToolBar("11100000000011")							
	Call InitVariables											
		
	FncNew = True												
End Function
	

'========================================================================================================
Function FncDelete()
	Dim IntRetCD

	FncDelete = False											
		
	If lgIntFlgMode <> parent.OPMD_UMODE Then							
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "x", "x")

	If IntRetCD = vbNo Then
		Exit Function
	End If

	Call DbDelete												

	FncDelete = True											
End Function


'========================================================================================================
Function FncSave()
	Dim IntRetCD
		
	FncSave = False											
		
	Err.Clear												
		
	If lgBlnFlgChgValue = False Then						
	    IntRetCD = DisplayMsgBox("900001", "x", "x", "x")	
	    Exit Function
	End If
		
	If Not chkField(Document, "2") Then						
	    If gPageNo > 0 Then
	        gSelframeFlg = gPageNo
	    End If
		Exit Function
	End If

	If Len(Trim(frm1.txtAtLatestShipDt.Text)) Then
		If parent.UniConvDateToYYYYMMDD(frm1.txtAmendDt.Text, parent.gDateFormat, "-") > parent.UniConvDateToYYYYMMDD(frm1.txtAtLatestShipDt.Text, parent.gDateFormat, "-") Then
			Call DisplayMsgBox("970023", "x", frm1.txtAtLatestShipDt.Alt, frm1.txtAmendDt.Alt)
			Call changeTabs(TAB1)
			frm1.txtAtLatestShipDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If

	If Len(Trim(frm1.txtAtExpireDt.Text)) Then
		If parent.UniConvDateToYYYYMMDD(frm1.txtAtLatestShipDt.Text, parent.gDateFormat, "-") > parent.UniConvDateToYYYYMMDD(frm1.txtAtExpireDt.Text, parent.gDateFormat, "-") Then
			Call DisplayMsgBox("970023", "x", frm1.txtAtExpireDt.Alt, frm1.txtAtLatestShipDt.Alt)
			Call changeTabs(TAB1)
			frm1.txtAtExpireDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If
		
	If parent.UNICDbl(frm1.txtAtDocAmt.text) < 0 Then
		Call DisplayMsgBox("970023", "x", "개설금액","0")
		frm1.txtAtDocAmt.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If				
		
	If parent.UNICDbl(frm1.txtAtXchRate.text) <= 0 Then
		Call DisplayMsgBox("970023", "x", "환율","0")
		frm1.txtAtXchRate.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If			
		
	If frm1.rdoAtDocAmt1.checked = True Then
		frm1.txtRadio.value = "I"					
	ElseIf frm1.rdoAtDocAmt2.checked = True Then
		frm1.txtRadio.value = "D"  	 					
		If parent.UNICDbl(frm1.txtBeDocAmt.text) - parent.UNICDbl(frm1.txtAmendAmt.text) < 0 Then
			Call DisplayMsgBox("970023", "x", "개설금액","0")
			frm1.txtAmendAmt.focus 
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
			
	End If
		
	Call DbSave													
		
	FncSave = True												
End Function

'========================================================================================================
Function FncCopy()
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")		

		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	lgIntFlgMode = parent.OPMD_CMODE												

	Call ggoOper.ClearField(Document, "1")									
	Call ggoOper.LockField(Document, "N")									
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
    Call DisplayMsgBox("900002", "x", "x", "x")  '☜ 바뀐부분 
    Exit Function
End If
				
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If

frm1.txtPrevNext.value = "PREV"

strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						
strVal = strVal & "&txtLCAmdNo=" & Trim(frm1.txtLCAmdNo1.value)		
strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)	
         
Call RunMyBizASP(MyBizASP, strVal)
End Function


'========================================================================================
Function FncNext() 
Dim strVal
    
If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
    Call DisplayMsgBox("900002", "x", "x", "x")  '☜ 바뀐부분 
    Exit Function
End If

				
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If

frm1.txtPrevNext.value = "NEXT"

strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						
strVal = strVal & "&txtLCAmdNo=" & Trim(frm1.txtLCAmdNo1.value)		
strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)	
         
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
	Err.Clear														

	DbQuery = False													

	Dim strVal

					
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If



	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					
	strVal = strVal & "&txtLCAmdNo=" & Trim(frm1.txtLCAmdNo.value)	
	strVal = strVal & "&empty=empty" 

	Call RunMyBizASP(MyBizASP, strVal)								

	DbQuery = True													
End Function


'========================================================================================================
Function DbSave()
	Err.Clear													

	DbSave = False												

	Dim strVal

					
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If



	With frm1
		.txtMode.value = parent.UID_M0002								
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID


		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With




	DbSave = True												
End Function
	

'========================================================================================================
Function DbDelete()
	Err.Clear													

	DbDelete = False											

	Dim strVal

					
	If   LayerShowHide(1) = False Then
	         Exit Function 
	End If



	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003				
	strVal = strVal & "&txtLCAmdNo=" & Trim(frm1.txtLCAmdNo1.value)
	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)		

	Call RunMyBizASP(MyBizASP, strVal)								

	DbDelete = True													
End Function


'========================================================================================================
Function DbQueryOk()						
	lgIntFlgMode = parent.OPMD_UMODE										
	lgBlnFlgChgValue = False
		
	Call ggoOper.LockField(Document, "Q")							
	Call SetToolBar("11111000110111")
	frm1.txtLCAmdNo.focus
		
End Function

'========================================================================================================
Function LCQueryOk()												
	Call ProtectXchRate()
	Call SetToolBar("11101000000011")	
End Function
	

'========================================================================================================
Function DbSaveOk()													
	Call InitVariables
	Call MainQuery()
End Function
	

'========================================================================================================
Function DbDeleteOk()												
	Call MainNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>>&nbsp;</TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C Amend 정보</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C Amend 기타</font></td>
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
								<TD CLASS=TD5 NOWRAP>L/C AMEND관리번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCAmdNo" SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="L/C AMEND관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCAmdNo" ALIGN=top TYPE="BUTTON" ONCLICK ="vbscript:btnLCAmdNoOnClick()"></TD>
								<TD CLASS=TDT></TD>
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
				<TD WIDTH=100% VALIGN=TOP>
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
									<TABLE <%=LR_SPACE_TYPE_60%>>								
										<TR>
											<TD CLASS=TD5 NOWRAP>L/C AMEND관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCAmdNo1" SIZE=20 MAXLENGTH=18 TAG="25XXXU" ALT="L/C AMEND관리번호"></TD>
											<TD CLASS=TD5 NOWRAP>L/C관리번호</TD>
											<TD	CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCNo" SIZE=20 MAXLENGTH=18 TAG="24XXXU" ALT="L/C관리번호"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>AMEND일</TD>						
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtAmendDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="AMEND일"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>L/C번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LC번호" MAXLENGTH=35 TYPE=TEXT SIZE=30 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>수입자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 ALT="수입자" TAG="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>수출자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수출자">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>	
											<TD CLASS=TD5 NOWRAP>금액변경</TD>
											<TD CLASS=TD6 NOWRAP>
												<Table cellpadding=0 cellspacing=0>
													<TR>
														<TD CLASS=TD6 NOWRAP>
															<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtDocAmt" TAG="21X" VALUE="I" CHECKED ID="rdoAtDocAmt1"><LABEL FOR="rdoAtDocAmt1">INCREASE BY</LABEL>
															<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtDocAmt" VALUE="D" TAG="21X" ID="rdoAtDocAmt2"><LABEL FOR="rdoAtDocAmt2">DECREASE BY</LABEL>
														</TD>
													</TR>
												</Table>
											</TD>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP></TD>
										</TR>	
										<TR>	
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT NAME="txtAmendAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 style="HEIGHT: 20px; WIDTH: 150px" tag="21X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP></TD>
										</TR>	
										<TR>
											<TD CLASS=TD5 NOWRAP>개설금액</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtAtDocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2Z" ALT="개설금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>
														<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtAtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="통화"></TD>
													</TR>
												</TABLE>
											</TD>	
											<TD CLASS=TD5 NOWRAP>변경전개설금액</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtBeDocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2Z" ALT="변경전개설금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														</TD>
														<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtBeCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="통화"></TD>
													</TR>
												</TABLE>
											</TD>	
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>환율</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtAtXchRate" style="HEIGHT: 20px; WIDTH: 150px" tag="22X5Z" ALT="환율" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS=TD5 NOWRAP>변경전환율</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtBeXchRate" style="HEIGHT: 20px; WIDTH: 150px" tag="24X5Z" ALT="변경전환율" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>유효일</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtAtExpireDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="유효일"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>변경전 유효일</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtBeExpireDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="변경전 유효일"></OBJECT>');</SCRIPT></TD>
										</TR>	
										<TR>									
											<TD CLASS=TD5 NOWRAP>선적기일</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtAtLatestShipDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="선적일"></OBJECT>');</SCRIPT></TD>
											<TD CLASS=TD5 NOWRAP>변경전선적기일
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtBeLatestShipDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="변경전선적일"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>환적여부</TD>
											<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtTranshipment" TAG="21X" VALUE="Y" CHECKED ID="rdoAtTranshipment1"><LABEL FOR="rdoAtTranshipment1">Y</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtTranshipment" VALUE="N" TAG="21X" ID="rdoAtTranshipment2"><LABEL FOR="rdoAtTranshipment2">N</LABEL>
											</TD>
											<TD CLASS=TD5 NOWRAP>변경전환적여부</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeTranshipment" SIZE=20 MAXLENGTH=10  STYLE="TEXT-ALIGN: Reft" TAG="24" ALT="변경전환적여부"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>분할선적여부</TD>
											<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtPartialShip" TAG="21X" VALUE="Y" CHECKED ID="rdoAtPartialShip1"><LABEL FOR="rdoAtPartialShip1">Y</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtPartialShip" VALUE="N" TAG="21X" ID="rdoAtPartialShip2"><LABEL FOR="rdoAtPartialShip2">N</LABEL>
											</TD>
											<TD CLASS=TD5 NOWRAP>변경전분할선적여부</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBePartialShip" SIZE=20 MAXLENGTH=2  STYLE="TEXT-ALIGN: Reft" TAG="24" ALT="변경전분할선적여부"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>양도여부</TD>
											<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtTransfer" TAG="21X" VALUE="Y" CHECKED ID="rdoAtTransfer1"><LABEL FOR="rdoAtTransfer1">Y</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtTransfer" VALUE="N" TAG="21X" ID="rdoAtTransfer2"><LABEL FOR="rdoAtTransfer2">N</LABEL>
											</TD>
											<TD CLASS=TD5 NOWRAP>변경전양도여부</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeTransfer" SIZE=20 MAXLENGTH=2  STYLE="TEXT-ALIGN: Reft" TAG="24" ALT="변경전양도여부"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>운송방법</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAtTransport" SIZE=10 MAXLENGTH=5 TAG="22XXXU" ALT="운송방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAtTransport" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnAtTransportOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtAtTransportNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>변경전운송방법</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeTransport" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="변경전운송방법">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtBeTransportNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>선적항</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtatLoadingPort" ALT="선적항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingPort" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnLoadingPortOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtAtLoadingPortNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>변경전선적항</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBeLoadingPort" ALT="변경전선적항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtBeLoadingPortNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>도착항</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtatDischgePort" ALT="도착항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="22X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgePort" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnDischgePortOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtAtDischgePortNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>변경전도착항</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBeDischgePort" ALT="변경전도착항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24X">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtBeDischgePortNm" SIZE=20 TAG="24"></TD>													
										</TR>
										<%Call SubFillRemBodyTD5656(2)%>
									</TABLE>
					</DIV>
					<DIV ID="TabDiv" SCROLL=no>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>비고</TD>
											<TD CLASS=TD656 NOWRAP><INPUT NAME="txtDoc1" ALT="비고" TYPE=TEXT MAXLENGTH=120 SIZE=86 TAG="21X"></TD>
										</TR>	
										<TR>
											<TD CLASS=TD5 NOWRAP>통지은행</TD>
											<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtAdvBank" SIZE=10 TAG="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtAdvBankNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>영업그룹</TD>											
											<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 TAG="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>개설은행</TD>
											<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10 TAG="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>개설일</TD>						
											<TD CLASS=TD656 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtOpenDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="개설일"></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>제조자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="제조자">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>대행자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="대행자">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
										</TR>
										<%Call SubFillRemBodyTD656(13)%>
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
				<TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck(0)">AMEND내역등록</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(1)">판매경비등록</A></TD>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = -1></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHLCNo" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHLCAmdNo" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtPrevNext" TAG="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtRadio" TAG="24" TABINDEX = -1>
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
