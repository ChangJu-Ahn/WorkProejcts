<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1411MA1
'*  4. Program Name         : 여신관리그룹등록 
'*  5. Program Desc         : 여신관리그룹등록 
'*  6. Comproxy List        : PS1G111.dll, PS1G112.dll, PS1G113.dll
'*  7. Modified date(First) : 2000/08/05
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 2002/11/22 : Grid성능 적용, Kang Jun Gu
'*                            2002/12/10 : INCLUDE 다시 성능 적용, Kang Jun Gu
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>


<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

'=================  1.2.1 Global 상수 선언  ===============

Const BIZ_PGM_ID = "s1411mb1.asp"												'☆: Head Query 비지니스 로직 ASP명 

Const C_SHEETMAXROWS = 30

Dim C_BpCd
Dim C_BpNm
Dim C_CollAmt
Dim C_SOAmt
Dim C_DNREqAmt
Dim C_GIAmt
Dim C_BillAmt
Dim C_ARAmt
Dim C_NoteAmt
Dim C_PrRcptAmt
Dim C_ExpiredAmt
Dim C_OldExpiredAmt			' 변경전 만기일 경과 수금액 
Dim C_ChgFlg

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim IsOpenPop						' Popup

'========================================================================================================
Sub initSpreadPosVariables()  
	C_BpCd		= 1
	C_BpNm		= 2
	C_CollAmt	= 3
	C_SOAmt		= 4
	C_DNREqAmt	= 5
	C_GIAmt		= 6
	C_BillAmt	= 7
	C_ARAmt		= 8
	C_NoteAmt	= 9
	C_PrRcptAmt	= 10
	C_ExpiredAmt= 11
	C_OldExpiredAmt = 12	
	C_ChgFlg	= 13
End Sub

'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    With frm1
		.txtLocCurrency1.value = Parent.gCurrency
		.txtLocCurrency2.value = Parent.gCurrency 
		.txtLocCurrency3.value = Parent.gCurrency 
	End With
    
End Sub


'========================================================================================================
Sub SetDefaultVal()	
	With frm1
		.txtCreditGrp.focus

		.rdoGiChkFlag1.checked = True
		.rdoSoChkFlag1.checked = True
		
		.txtRadioGiChk.value = .rdoGiChkFlag1.value 
		.txtRadioSoChk.value = .rdoSoChkFlag1.value
		
		.txtCreditLmtAmt.text = "0"
		.txtAvailableAmtForSO.text = "0"
		.txtAvailableAmtForGI.text = "0"
		
		lgBlnFlgChgValue = False
	End With
End Sub


'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
						     
	    ggoSpread.Spreadinit "V20030603",,parent.gAllowDragDropSpread    

		.MaxCols = C_ChgFlg            '☜: 최대 Columns의 항상 1개 증가시킴 
	    .MaxRows = 0                                                                  ' ☜: Clear spreadsheet data 

	    Call GetSpreadColumnPos("A")
		.ReDraw = false
		ggoSpread.SSSetEdit		C_BpCd,      "고객", 15, 0
		ggoSpread.SSSetEdit		C_BpNm,      "고객명", 25, 0			
		ggoSpread.SSSetFloat    C_CollAmt,   "총담보금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat    C_SOAmt,     "수주금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat    C_DNREqAmt,  "출고요청금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat    C_GIAmt,     "출고금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat    C_BillAmt,   "매출채권금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat    C_ARAmt,     "외상매출금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat    C_NoteAmt,   "현금화대기금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat    C_PrRcptAmt, "선수금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

	    ggoSpread.SSSetFloat    C_ExpiredAmt, "만기일경과수금액",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    ggoSpread.SSSetFloat    C_OldExpiredAmt, "",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

		Call ggoSpread.MakePairsColumn(C_BpCd,C_BpNm)
		Call ggoSpread.SSSetColHidden(C_OldExpiredAmt,C_OldExpiredAmt,True)
		Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg,True)
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)   '☜: 공통콘트롤 사용 Hidden Column

		Call SetSpreadLock() 
		.ReDraw = true
	End With
End Sub

'========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLock C_BpCd, -1, C_PrRcptAmt
End Sub


'========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetRequired C_ExpiredAmt, pvStartRow, pvEndRow
End Sub


'========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_BpCd			= iCurColumnPos(1)
			C_BpNm			= iCurColumnPos(2)
			C_CollAmt		= iCurColumnPos(3)
			C_SOAmt			= iCurColumnPos(4)
			C_DNREqAmt		= iCurColumnPos(5)
			C_GIAmt			= iCurColumnPos(6)
			C_BillAmt		= iCurColumnPos(7)
			C_ARAmt			= iCurColumnPos(8)
			C_NoteAmt		= iCurColumnPos(9)
			C_PrRcptAmt		= iCurColumnPos(10)
			C_ExpiredAmt	= iCurColumnPos(11)
			C_OldExpiredAmt = iCurColumnPos(12)
			C_ChgFlg		= iCurColumnPos(13)

    End Select    
End Sub


'========================================================================================================
Function OpenPopupMenu()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "여신관리그룹"					<%' 팝업 명칭 %>
	arrParam(1) = "S_CREDIT_LIMIT"						<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtCreditGrp.value)			<%' Code Condition%>
	arrParam(3) = ""									<%' Name Cindition%>
	arrParam(4) = ""									<%' Where Condition%>
	arrParam(5) = "여신관리그룹"					<%' TextBox 명칭 %>
	
    arrField(0) = "CREDIT_GRP"							<%' Field명(0)%>
    arrField(1) = "CREDIT_GRP_NM"						<%' Field명(1)%>
    
    arrHeader(0) = "여신관리그룹"					<%' Header명(0)%>
    arrHeader(1) = "여신관리그룹명"					<%' Header명(1)%>
    
    frm1.txtCreditGrp.focus
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else 
		Call SetItemPopup(arrRet)
	End If	

End Function

'========================================================================================================
Function OpenContentPop(ByVal iPopupID)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iPopupID
	Case 0		<%' ---수주시여신체크방법 %>
		arrParam(0) = "여신체크"					<%' 팝업 명칭 %>
		arrParam(1) = "B_MINOR"							<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtCreditSoChkCd.value)	<%' Code Condition%>
		arrParam(3) = ""                               	<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD = " & FilterVar("S0007", "''", "S") & ""				<%' Where Condition%>
		arrParam(5) = "여신체크"					<%' TextBox 명칭 %>
		
	    arrField(0) = "MINOR_CD"						<%' Field명(0)%>
	    arrField(1) = "MINOR_NM"						<%' Field명(1)%>
	    
	    arrHeader(0) = "여신체크코드"				<%' Header명(0)%>
	    arrHeader(1) = "여신체크명"					<%' Header명(1)%>
	    
	    frm1.txtCreditSoChkCd.focus

	Case 1		<%' ---화폐 %>
		arrParam(0) = "화폐"						<%' 팝업 명칭 %>
		arrParam(1) = "B_CURRENCY"						<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtCurrency.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>
		arrParam(4) = ""								<%' Where Condition%>
		arrParam(5) = "화폐"						<%' TextBox 명칭 %>
		
	    arrField(0) = "CURRENCY"						<%' Field명(0)%>
	    arrField(1) = "CURRENCY_DESC"					<%' Field명(1)%>
	    
	    arrHeader(0) = "화폐코드"					<%' Header명(0)%>
	    arrHeader(1) = "화폐설명"					<%' Header명(1)%>

	Case 2		<%' ---출고시여신체크방법 %>
		arrParam(0) = "여신체크"					<%' 팝업 명칭 %>
		arrParam(1) = "B_MINOR"							<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtCreditChkCd.value)	<%' Code Condition%>
		arrParam(3) = ""                               	<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD = " & FilterVar("S0007", "''", "S") & ""				<%' Where Condition%>
		arrParam(5) = "여신체크"					<%' TextBox 명칭 %>
		
	    arrField(0) = "MINOR_CD"						<%' Field명(0)%>
	    arrField(1) = "MINOR_NM"						<%' Field명(1)%>
	    
	    arrHeader(0) = "여신체크코드"				<%' Header명(0)%>
	    arrHeader(1) = "여신체크명"					<%' Header명(1)%>

		frm1.txtCreditChkCd.focus
	End Select
    
    arrParam(3) = ""	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetContentPop(arrRet,iPopupID)
	End If	
End Function

'========================================================================================================
Function SetItemPopup(Byval arrRet)
	frm1.txtCreditGrp.value = arrRet(0)
	frm1.txtCreditGrpNm.value = arrRet(1)
End Function


'========================================================================================================
Function SetContentPop(Byval arrRet,ByVal iPopupID)

	If arrRet(0) <> "" Then 
		Select Case iPopupID
		Case 0												<%' 여신체크방법 %>
			frm1.txtCreditSoChkCd.value = arrRet(0)
			frm1.txtCreditSoChkNm.value = arrRet(1)
		Case 1												<%' 화폐 %>
			frm1.txtCurrency.value = arrRet(0)
		Case 2												<%' 여신체크방법 %>
			frm1.txtCreditChkCd.value = arrRet(0)
			frm1.txtCreditChkNm.value = arrRet(1)

		End Select
		lgBlnFlgChgValue = True
	End If
End Function

'========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877						<%'Cookie Split String : CookiePage Function Use%>
	Dim strTemp, arrVal

	Select Case Kubun
		
	Case 1
		WriteCookie CookieSplit , frm1.txtCreditGrp.value
	Case 0
		strTemp = ReadCookie(CookieSplit)
				
		If strTemp = "" then Exit Function
				
		frm1.txtCreditGrp.value =  strTemp
			
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If
			
		Call MainQuery()
						
		WriteCookie CookieSplit , ""
	Case Else
		Exit Function
	End Select 		
End Function

'========================================================================================================
Function RadioSettingValue()

	IF frm1.rdoSOChkFlag1.checked = True Then
		frm1.txtRadioSoChk.value = frm1.rdoSoChkFlag1.value
	ElseIf frm1.rdoSoChkFlag2.checked = True Then
		frm1.txtRadioSoChk.value = frm1.rdoSoChkFlag2.value
	ElseIf frm1.rdoSoChkFlag3.checked = True Then
		frm1.txtRadioSoChk.value = frm1.rdoSoChkFlag3.value
	End If


	IF frm1.rdoGiChkFlag1.checked = True Then
		frm1.txtRadioGiChk.value = frm1.rdoGiChkFlag1.value
	ElseIf frm1.rdoGiChkFlag2.checked = True Then
		frm1.txtRadioGiChk.value = frm1.rdoGiChkFlag2.value
	ElseIf frm1.rdoGiChkFlag3.checked = True Then
		frm1.txtRadioGiChk.value = frm1.rdoGiChkFlag3.value
	End If

End Function

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
 	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitSpreadSheet
    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call CookiePage (0)
    Call SetToolBar ("1110100000001111")										'⊙: 버튼 툴바 제어 
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"   

	Set gActiveSpdSheet = frm1.vspdData
	' Context 메뉴의 입력, 삭제, 데이터 입력, 취소 
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")
	    
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
	End If
	   	    
	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub     
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.vspdData.Row = Row
	'---frm1.vspdData.Col = C_MajorCd
		
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

   ggoSpread.UpdateRow Row
   
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

<%'==================================== 3.2.25 txtAdvDt_DblClick()  =====================================
'   Event Name : btnCurrency_OnClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================%>
	Sub btnCurrency_OnClick()
		If frm1.txtCurrency.readOnly <> True Then
			Call OpenContentPop (1)
		End If
	End Sub

'==========================================================================================
'   Event Name : Radio OnClick()
'   Event Desc : Radio Button Click시 lgBlnFlgChgValue 처리 / Value
'==========================================================================================
Sub rdoSoChkFlag1_OnClick()
	lgBlnFlgChgValue = True
	frm1.txtRadioSoChk.value = frm1.rdoSoChkFlag1.value
End Sub

Sub rdoSoChkFlag2_OnClick()
	lgBlnFlgChgValue = True
	frm1.txtRadioSoChk.value = frm1.rdoSoChkFlag2.value
End Sub

Sub rdoSoChkFlag3_OnClick()
	lgBlnFlgChgValue = True
	frm1.txtRadioSoChk.value = frm1.rdoSoChkFlag3.value
End Sub


Sub rdoGiChkFlag1_OnClick()
	lgBlnFlgChgValue = True
	frm1.txtRadioGiChk.value = frm1.rdoGiChkFlag1.value
End Sub

Sub rdoGiChkFlag2_OnClick()
	lgBlnFlgChgValue = True
	frm1.txtRadioGiChk.value = frm1.rdoGiChkFlag2.value
End Sub

Sub rdoGiChkFlag3_OnClick()
	lgBlnFlgChgValue = True
	frm1.txtRadioGiChk.value = frm1.rdoGiChkFlag3.value
End Sub

Sub chkBadCreditFlg_OnPropertyChange()
	lgBlnFlgChgValue = True
End Sub	

'========================================================================================================
Sub txtCreditLmtAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

<%    '-----------------------
    'Check previous data area
    '----------------------- %>
    
	ggoSpread.Source = frm1.vspdData
	
    If lgBlnFlgChgValue Or ggoSpread.SSCheckChange Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")										<%'⊙: Clear Contents  Field%>
    
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
    
<%  '-----------------------
    'Query function call area
    '----------------------- %>
	Call ggoOper.LockField(Document, "N")								<% '⊙: This function lock the suitable field %>

    Call InitVariables															<%'⊙: Initializes local global variables%>

    Call DbQuery																<%'☜: Query db data%>
       
    FncQuery = True																<%'⊙: Processing is OK%>
	       
End Function

'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          <%'⊙: Processing is NG%>
    
<%  '-----------------------
    'Check previous data area
    '-----------------------%>
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x") 
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
<%  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------%>
    Call ggoOper.ClearField(Document, "A")                                      <%'⊙: Clear Condition Field%>
    Call ggoOper.LockField(Document, "N")                                       <%'⊙: Lock  Suitable  Field%>
    Call SetDefaultVal
    Call InitVariables															<%'⊙: Initializes local global variables%>
    Call SetToolBar ("1110100000001111")										'⊙: 버튼 툴바 제어 

    FncNew = True																<%'⊙: Processing is OK%>
End Function

'========================================================================================================
Function FncDelete() 

    Dim IntRetCD
    
    FncDelete = False														<%'⊙: Processing is NG%>
    
<%  '-----------------------
    'Precheck area
    '-----------------------%>
    If lgIntFlgMode <> Parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"x","x")
    If IntRetCD = vbNo Then
        Exit Function
    End If
    
<%  '-----------------------
    'Delete function call area
    '-----------------------%>
    Call DbDelete															<%'☜: Delete db data%>
    
    FncDelete = True                                                        <%'⊙: Processing is OK%>
    
End Function


'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
<%  '-----------------------
    'Precheck area
    '-----------------------%>
	ggoSpread.Source = frm1.vspdData
	
    If Not (lgBlnFlgChgValue Or ggoSpread.SSCheckChange) Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If
    
<%  '-----------------------
    'Check content area
    '-----------------------%>
    If Not chkField(Document, "2") Then                             <%'⊙: Check contents area%>
       Exit Function
    End If

	Call RadioSettingValue

<%  '-----------------------
    'Save function call area
    '-----------------------%>
    Call DbSave				                                                <%'☜: Save db data%>
    
    FncSave = True                                                          <%'⊙: Processing is OK%>
    
End Function

'========================================================================================================
Function FncCopy() 
	Dim IntRetCD

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    lgIntFlgMode = Parent.OPMD_CMODE												<%'⊙: Indicates that current mode is Crate mode%>
    
    <% ' 조건부 필드를 삭제한다. %>
    Call ggoOper.ClearField(Document, "1")                                  <%'⊙: Clear Condition Field%>
    frm1.txtCreditLmtAmt.text = ""
    frm1.txtXchRate.text = ""
    frm1.txtAvailableAmtForSO.text = "0"
    frm1.txtAvailableAmtForGI.text = "0"
    frm1.chkBadCreditFlg.checked = False
    frm1.txtCreditGrpCd.value = ""
    frm1.txtCreditGrpName.value = ""
    Call ggoOper.LockField(Document, "N")									<%'⊙: This function lock the suitable field%>
    Call InitVariables														<%'⊙: Initializes local global variables%>
    Call InitSpreadSheet
    
End Function

'========================================================================================================
Function FncCancel() 
    On Error Resume Next                                                    <%'☜: Protect system from crashing%>
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  

    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement 
End Function
'========================================================================================================
Function FncPrint() 
    ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
End Function

'========================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLEMULTI)
End Function

'========================================================================================================
Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLEMULTI, TRUE)
End Function

'========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadColor(-1,-1)
End Sub

'========================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		  '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True
End Function

'========================================================================================================
Function DbDelete() 
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
    DbDelete = False														<%'⊙: Processing is NG%>
    
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If
    
    Dim iStrVal
    
    iStrVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003							<%'☜: 비지니스 처리 ASP의 상태 %>
    iStrVal = iStrVal & "&txtCreditGrp=" & Trim(frm1.txtCreditGrpCd.value)		<%'☜: 삭제 조건 데이타 %>
    
	Call RunMyBizASP(MyBizASP, iStrVal)										<%'☜: 비지니스 ASP 를 가동 %>
	
    DbDelete = True                                                         <%'⊙: Processing is NG%>

End Function

'========================================================================================================
Function DbDeleteOk()														<%'☆: 삭제 성공후 실행 로직 %>
	Call MainNew()
End Function

'========================================================================================================
Function DbQuery() 
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
    DbQuery = False                                                         <%'⊙: Processing is NG%>
    
    
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If
    
    Dim strVal
    
     If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001									<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtCreditGrp=" & Trim(frm1.txtHCreditGrp.value)						<%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtInsrtUserId=" & Trim(frm1.txtInsrtUserId.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001									<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtCreditGrp=" & Trim(frm1.txtCreditGrp.value)						<%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtInsrtUserId=" & Trim(frm1.txtInsrtUserId.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    End If
        
	Call RunMyBizASP(MyBizASP, strVal)										<%'☜: 비지니스 ASP 를 가동 %>
	
    DbQuery = True                                                          <%'⊙: Processing is NG%>

End Function

'========================================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE												<%'⊙: Indicates that current mode is Update mode%>
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")									<%'⊙: This function lock the suitable field%>

	If frm1.vspdData.maxrows <= 0 Then		
		Call SetToolBar("1111100000011111")
	Else
		Call SetToolBar("1111100100011111")
	End If
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus  
	Else
		frm1.txtCreditGrp.focus
	End If     
	
End Function

'========================================================================================================
Function DbSave() 
	on error resume next
	Err.Clear				

	Dim lRow
	Dim iStrVal

	<%'☜: Protect system from crashing%>
	If   LayerShowHide(1) = False Then
             Exit Function 
	End If
	
	DbSave = False															<%'⊙: Processing is NG%>

	
	With frm1
		.txtMode.value = Parent.UID_M0002											<%'☜: 비지니스 처리 ASP 의 상태 %>
		.txtFlgMode.value = lgIntFlgMode
		.txtInsrtUserId.value = Parent.gUsrID 
		.txtUpdtUserId.value = Parent.gUsrID
		
		' 부실 여신 그룹 여부 
		If .chkBadCreditFlg.checked Then
			.txtBadCreditFlg.value = "Y"
		Else
			.txtBadCreditFlg.value = "N"
		End If		 
    
		' 여신 그룹정보의 변경여부 
		If lgBlnFlgChgValue Then
			.txtCreditLimitChgFalg.value = "Y"
		Else
			.txtCreditLimitChgFalg.value = "N"
		End If

		ggoSpread.Source = .vspdData
				
		iStrVal = ""
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 0

			Select Case .vspdData.Text
				Case ggoSpread.UpdateFlag								<% '☜: 신규 %>
					iStrVal = iStrVal & CStr(lRow) & Parent.gColSep	<% '☜: U=Modify, Row위치 정보 %>

					' 고객 
					.vspdData.Col = C_BpCd
					iStrVal = iStrVal & Trim(.vspdData.Text) & Parent.gColSep

					' 만기일 경과 수금액 
					.vspdData.Col = C_ExpiredAmt
					iStrVal = iStrVal & CStr(UNIConvNum(.vspdData.Text, 0)) & Parent.gColSep
					
					' 변경전 만기일 경과 수금액 
					.vspdData.Col = C_OldExpiredAmt
					iStrVal = iStrVal & CStr(UNIConvNum(.vspdData.Text, 0)) & Parent.gRowSep
					
			End Select			
		Next

		.txtSpread.value = iStrVal
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
    DbSave = True                                                           <%'⊙: Processing is NG%>
End Function

'========================================================================================================
Function DbSaveOk()															<%'☆: 저장 성공후 실행 로직 %>

    frm1.txtCreditGrp.value = frm1.txtCreditGrpCd.value 
	frm1.vspdData.MaxRows = 0    
    Call InitVariables
    Call MainQuery()
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>여신관리그룹</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=500>&nbsp;</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
						<FIELDSET CLASS="CLSFLD"><TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>여신관리그룹</TD>
								<TD CLASS="TD6"><INPUT NAME="txtCreditGrp" ALT="여신관리그룹" TYPE="Text" MAXLENGTH="5" SiZE=10  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCreditGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopupMenu" >&nbsp;<INPUT NAME="txtCreditGrpNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="14"></TD>
								<TD CLASS="TDT"></TD>
								<TD CLASS="TD6"></TD>
							</TR>
						</TABLE></FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>여신관리그룹코드</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT NAME="txtCreditGrpCd" ALT="여신관리그룹코드" TYPE="Text" MAXLENGTH= "5" SIZE="10"  tag="23XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>여신관리그룹명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCreditGrpName" ALT="여신관리그룹명" TYPE="Text" MAXLENGTH="30" SIZE=30 tag="22XXX"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>수주시체크방법</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCreditSoChkCd" ALT="수주시체크방법" TYPE="Text" MAXLENGTH="4" SIZE=10  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCreditChkType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenContentPop 0">&nbsp;<INPUT NAME="txtCreditSoChkNm" TYPE="Text" MAXLENGTH="50" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>출고시체크방법</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCreditChkCd" ALT="출고시체크방법" TYPE="Text" MAXLENGTH="4" SIZE=10  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCreditDnChkType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenContentPop 2">&nbsp;<INPUT NAME="txtCreditChkNm" TYPE="Text" MAXLENGTH="50" SIZE=20 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>수주시여신체크</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSOChkFlag" TAG="21X" VALUE="N" CHECKED ID="rdoSOChkFlag1"><LABEL FOR="rdoSOChkFlag1">안함</LABEL>&nbsp;&nbsp;&nbsp;
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSOChkFlag" TAG="21X" VALUE="W" ID="rdoSOChkFlag2"><LABEL FOR="rdoSOChkFlag2">경고처리</LABEL>&nbsp;&nbsp;&nbsp;
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSOChkFlag" TAG="21X" VALUE="E" ID="rdoSOChkFlag3"><LABEL FOR="rdoSOChkFlag3">에러처리</LABEL>			
								</TD>
								<TD CLASS=TD5 NOWRAP>출고시여신체크</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoGIChkFlag" TAG="21XXX" VALUE="N" CHECKED ID="rdoGIChkFlag1"><LABEL FOR="rdoGIChkFlag1">안함</LABEL>&nbsp;&nbsp;&nbsp;
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoGIChkFlag" TAG="21XXX" VALUE="W" ID="rdoGIChkFlag2"><LABEL FOR="rdoGIChkFlag2">경고처리</LABEL>&nbsp;&nbsp;&nbsp;
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoGIChkFlag" TAG="21XXX" VALUE="E" ID="rdoGIChkFlag3"><LABEL FOR="rdoGIChkFlag3">에러처리</LABEL>			
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>사용가능액(수주)</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s1411ma1_fpDoubleSingle2_txtAvailableAmtForSO.js'></script></TD>
											<TD>&nbsp;&nbsp;&nbsp;<INPUT NAME="txtLocCurrency2" ALT="화폐" TYPE="Text" MAXLENGTH="3" SIZE="10"  tag="24"></TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>사용가능액(출고)</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s1411ma1_fpDoubleSingle3_txtAvailableAmtForGI.js'></script></TD>
											<TD>&nbsp;&nbsp;&nbsp;<INPUT NAME="txtLocCurrency3" ALT="화폐" TYPE="Text" MAXLENGTH="3" SIZE="10"  tag="24"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>여신한도금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0 STYLE="PADDING-BOTTOM:5px;PADDING-TOP:5px">
										<TR>
											<TD><script language =javascript src='./js/s1411ma1_fpDoubleSingle1_txtCreditLmtAmt.js'></script></TD>
											<TD>&nbsp;&nbsp;&nbsp;<INPUT NAME="txtLocCurrency1" ALT="화폐" TYPE="Text" MAXLENGTH="3" SIZE="10"  tag="24"></TD>
										</TR>
									</TABLE>
								</TD>		
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21XXX" VALUE="Y" NAME="chkBadCreditFlg" ID="chkBadCreditFlg">
									<LABEL FOR="chkBadCreditFlg">부실여신관리그룹</LABEL>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/s1411ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCreditLimitChgFalg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioSoChk" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioDnChk" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioGiChk" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHCreditGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtBadCreditFlg" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
