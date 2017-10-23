<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3221ma2.asp																*
'*  4. Program Name         : Local L/C Amend 등록														*
'*  5. Program Desc         : Local L/C Amend 등록														*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/24																*
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

Const BIZ_PGM_ID				= "s3221mb2.asp"		
Const BIZ_PGM_LCQRY_ID			= "s3221mb4.asp"	
Const LCAMEND_DETAIL_ENTRY_ID	= "s3222ma2"	
Const EXPORT_CHARGE_ENTRY_ID	= "s6111ma1"	
Const BIZ_PGM_CAL_AMT_ID		= "s3211mb5.asp"

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
	frm1.txtAmendDt.text			= "<%=EndDate%>"
	frm1.txtAtExpireDt.text			= "<%=EndDate%>"
	frm1.txtAtLatestShipDt.text		= "<%=EndDate%>"
	frm1.txtAtLocCurrency1.value	= parent.gCurrency
	frm1.txtAtLocCurrency2.value	= parent.gCurrency
	lgBlnFlgChgValue = False
End Sub

'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %> 
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Function OpenLCAmdNoPop()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Or UCase(frm1.txtLCAmdNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("s3221pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3221pa2", "X")
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
		
	iCalledAspName = AskPRAspName("s3211ra2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3211ra2", "X")
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
			WriteCookie CookieSplit , "EO" & parent.gRowSep & frm1.txtSalesGroup.value & parent.gRowSep & frm1.txtSalesGroupNm.value & parent.gRowSep & frm1.txtLCAmdNo.value
			 			
	End Select

End Function

'========================================================================================================
Function SetRadio()
	Dim blnOldFlag

	blnOldFlag = lgBlnFlgChgValue
	frm1.rdoAtPartialShip1.checked = True

	lgBlnFlgChgValue = blnOldFlag
End Function

'========================================================================================================
Function JumpChgCheck(ByVal IWhere)

	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "x", "x")			
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
	If frm1.txtAtCurrency1.value = parent.gCurrency Then
		Call ggoOper.SetReqAttr(frm1.txtAtXchRate, "Q")
		Call ggoOper.SetReqAttr(frm1.txtAtLocAmt, "Q")
	End If	
End Function
	
'========================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		'금액변경 
		ggoOper.FormatFieldByObjectOfCur .txtAmendAmt, .txtAtCurrency1.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		'개설금액 
		ggoOper.FormatFieldByObjectOfCur .txtAtDocAmt, .txtAtCurrency1.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		'변경전개설금액 
		ggoOper.FormatFieldByObjectOfCur .txtBeDocAmt, .txtAtCurrency2.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
		
		'변경환율 
		ggoOper.FormatFieldByObjectOfCur .txtAtXchRate, .txtAtCurrency1.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		
		'변경전환율 
		ggoOper.FormatFieldByObjectOfCur .txtBeXchRate, .txtAtCurrency2.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
		
	End With
End Sub
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029													
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")								

	Call SetDefaultVal
	Call SetToolBar("11100000000011")											
	Call InitVariables
	Call CookiePage(0)	

	frm1.txtLCAmdNo.focus
	Set gActiveElement = document.activeElement 
End Sub
	
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub
	
'========================================================================================================
Sub txtAmendAmt_Change()
    Dim arrAmt
        
    arrAmt = UNICDbl(frm1.txtBeDocAmt.text)	
        
	If frm1.rdoAtDocAmt1.checked = True Then		
		frm1.txtAtDocAmt.text = UNIFormatNumberByCurrecny(arrAmt + UNICDbl(frm1.txtAmendAmt.text),frm1.txtAtCurrency1.value, parent.ggAmtOfMoneyNo)
	Else		
		frm1.txtAtDocAmt.text = UNIFormatNumberByCurrecny(arrAmt - UNICDbl(frm1.txtAmendAmt.text),frm1.txtAtCurrency1.value, parent.ggAmtOfMoneyNo)
	End If			
End Sub

'========================================================================================================
Sub txtAtXchRate_Change()
	Err.Clear																
	If frm1.txtAtCurrency1.value = parent.gCurrency Then
		frm1.txtAtXchRate.text = 1
		frm1.txtAtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtAtXchRate.text) * UNICDbl(frm1.txtAtDocAmt.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	Else 	
		If Len(frm1.txtAtCurrency1.value) And IsNumeric(frm1.txtAtXchRate.text) = True And IsNumeric(frm1.txtAtDocAmt.text) = True Then
			If frm1.txtExchRateOp.value = "*" then
				frm1.txtAtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtAtDocAmt.text) * UNICDbl(frm1.txtAtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
			ElseIf frm1.txtExchRateOp.value = "/" then
				frm1.txtAtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtAtDocAmt.text) / UNICDbl(frm1.txtAtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
			End If
		End If
	End If	
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtAtDocAmt_Change()
	If frm1.txtAtCurrency1.value = parent.gCurrency Then
		frm1.txtAtXchRate.text = 1
		frm1.txtAtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtAtXchRate.text) * UNICDbl(frm1.txtAtDocAmt.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	Else 			
		If Len(frm1.txtAtCurrency1.value) Then
			If IsNumeric(frm1.txtAtXchRate.text) = True And IsNumeric(frm1.txtAtDocAmt.text) = True Then
				If frm1.txtExchRateOp.value = "*" then
					frm1.txtAtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtAtDocAmt.text) * UNICDbl(frm1.txtAtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
				ElseIf frm1.txtExchRateOp.value = "/" then
					frm1.txtAtLocAmt.text = UNIFormatNumber(UNICDbl(frm1.txtAtDocAmt.text) / UNICDbl(frm1.txtAtXchRate.text), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
				End If 
			End If
		End If
	End If	
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub btnLCAmdNoOnClick()
	Call OpenLCAmdNoPop()
End Sub

'========================================================================================================
Sub chkAtDocAmt_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub rdoAtDocAmt2_OnPropertyChange()
	lgBlnFlgChgValue = True
		
	Call txtAmendAmt_Change
		
End Sub
	
'========================================================================================================
Sub rdoAtDocAmt1_OnPropertyChange()		
	Call txtAmendAmt_Change
End Sub
		
'========================================================================================================
Sub chkAtExpiryDt_OnClick()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub chkAtLatestShipDt_OnClick()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub chkAtPartialShip_OnClick()
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
Sub txtAmendDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAmendDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtAmendDt.Focus
    End If
End Sub

'========================================================================================================
Sub txtAtExpireDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAtExpireDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtAtExpireDt.Focus
    End If
End Sub

'========================================================================================================
Sub txtatLatestShipDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtatLatestShipDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtatLatestShipDt.Focus
    End If
End Sub
'========================================================================================================
Sub txtAtLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtAtExpireDt_Change()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtAtLatestShipDt_Change()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Sub txtAmendDt_Change()
	lgBlnFlgChgValue = True
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
	Call SetRadio()
	Call SetDefaultVal
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
	Call InitVariables												
	Call SetToolBar("11100000000011")								

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
		Exit Function
	End If

	If ValidDateCheck(frm1.txtAmendDt, frm1.txtAtLatestShipDt) = False Then Exit Function
	If ValidDateCheck(frm1.txtAtLatestShipDt, frm1.txtAtExpireDt) = False Then Exit Function

	If UNICDbl(frm1.txtAtDocAmt.text) < 0 Then
		Call DisplayMsgBox("970023", "x", "개설금액","0")
		frm1.txtAtDocAmt.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
		
	If UNICDbl(frm1.txtAtXchRate.text) <= 0 Then
		Call DisplayMsgBox("970023", "x", "환율","0")
		frm1.txtAtXchRate.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	If frm1.rdoAtDocAmt1.checked = True Then
		frm1.txtRadio.value = "I"
	ElseIf frm1.rdoAtDocAmt2.checked = True Then
		frm1.txtRadio.value = "D"  	 			

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
	    Call DisplayMsgBox("900002", "x", "x", "x")  
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

'========================================================================================================
Function FncNext() 
	Dim strVal
	    
	If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
	    Call DisplayMsgBox("900002", "x", "x", "x")  
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
	Call parent.FncFind(parent.C_SINGLE, False)
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

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003					
	strVal = strVal & "&txtLCAmdNo=" & Trim(frm1.txtLCAmdNo1.value)	
	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)		

	Call RunMyBizASP(MyBizASP, strVal)								

	DbDelete = True													
End Function

'========================================================================================================
Function LCQueryOk()												
	Call ProtectXchRate()
	Call txtAtDocAmt_Change()
	Call SetToolBar("11101000000011")	
End Function

'========================================================================================================
Function DbQueryOk()												
	lgIntFlgMode = parent.OPMD_UMODE										
	lgBlnFlgChgValue = False
		
	Call ggoOper.LockField(Document, "Q")							
	Call SetToolBar("11111000110111")
	Call ProtectXchRate()
	frm1.txtLCAmdNo.focus
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
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>Local L/C Amend 정보</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenLCRef">LOCAL L/C 참조</A></TD>
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
									<TD CLASS=TD5 NOWRAP>LOCAL L/C AMEND 관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCAmdNo" SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="L/C AMEND관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCAmdNo" ALIGN=top TYPE="BUTTON" ONCLICK ="vbscript:btnLCAmdNoOnClick()"></TD>
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
									<TD CLASS=TD5 NOWRAP>LOCAL L/C</TD>
									<TD CLASS=TD6 NOWRAP></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>AMEND 관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCAmdNo1" SIZE=20 MAXLENGTH=18 TAG="15XXXU" ALT="L/C AMEND관리번호"></TD>
									<TD CLASS=TD5 NOWRAP>LOCAL L/C관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCNo" ALT="LOCAL L/C관리번호" TYPE=TEXT MAXLENGTH=18 SIZE=20 TAG="24XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>AMEND일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtAmendDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="AMEND일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>LOCAL L/C번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LOCAL L/C번호" TYPE=TEXT MAXLENGTH=35 SIZE=30 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설신청인</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 ALT="개설신청인" TAG="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>수혜자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수혜자">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
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
												<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtAtCurrency1" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="통화"></TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>변경전개설금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtBeDocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2Z" ALT="변경전개설금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>
												<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtAtCurrency2" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="통화"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>환율</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtAtXchRate" style="HEIGHT: 20px; WIDTH: 150px" tag="22X5Z" ALT="환율" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
									<TD CLASS=TD5 NOWRAP>변경전환율</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtBeXchRate" style="HEIGHT: 20px; WIDTH: 150px" tag="24X5Z" ALT="환율" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설자국금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtAtLocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="22X2Z" ALT="개설자국금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>
												<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtAtLocCurrency1" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐"></TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>변경전개설자국금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtBeLocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2Z" ALT="변경전개설자국금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>
												<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtAtLocCurrency2" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>유효일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtAtExpireDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="유효일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>변경전 유효일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtBeExpireDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="변경전 유효일"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>인도기간</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtatLatestShipDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="인도일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>변경전 인도기간</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtBeLatestShipDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="변경전인도일"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>분할인도여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtPartialShip" TAG="21X" VALUE="Y" CHECKED ID="rdoAtPartialShip1"><LABEL FOR="rdoAtPartialShip1">Y</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAtPartialShip" VALUE="N" TAG="21X" ID="rdoAtPartialShip2"><LABEL FOR="rdoAtPartialShip2">N</LABEL>
									</TD>
									<TD CLASS=TD5 NOWRAP>변경전분할인도여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBePartialShip" SIZE=20 MAXLENGTH=2  STYLE="TEXT-ALIGN:left" TAG="24" ALT="변경전분할인도여부"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>통지번호</TD>
									<TD	CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdvNo" SIZE=35 MAXLENGTH=35 TAG="21" ALT="통지번호"></TD>
									<TD CLASS=TD5 NOWRAP>선통지참조사항</TD>
									<TD	CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPreAdvRef" SIZE=35 MAXLENGTH=120 TAG="21" ALT="선통지참조사항"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>기타 참조사항</TD>
									<TD CLASS=TD6 COLSPAN=3><INPUT TYPE=TEXT NAME="txtRef" SIZE=92 MAXLENGTH=120 TAG="21" ALT="기타참조사항"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>추심의뢰은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdvBank" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" TAG="24X" ALT="추심의뢰은행">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtAdvBankNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>개설은행</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10 MAXLENGTH=10 TAG="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설일</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime NAME="txtOpenDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT="개설일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=5 TAG="24XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
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
<INPUT TYPE=HIDDEN NAME="txtIncAmt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtDecAmt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtXchRate" TAG="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHExpiryDt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHLatestShipDt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHPartialShip" tag="24" TABINDEX = -1> 
<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24" TABINDEX = -1>	
<INPUT TYPE=HIDDEN NAME="txtExchRateOp" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtRadio" tag="24" TABINDEX = -1>
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
