<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Basic Info. - Accounting
'*  2. Function Name        : Common Info.
'*  3. Program ID           : B1310MA1
'*  4. Program Name         : 은행정보등록 
'*  5. Program Desc         : Register of Bank Info.
'*  6. Component List       : PB2SA05
'*  7. Modified date(First) : 2000/03/22
'*  8. Modified date(Last)  : 2002/09/17
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit 

			'☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->			

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QUERY_ID = "b1310mb1_ko441.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID  = "b1310mb2_ko441.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID   = "b1310mb3_ko441.asp"

'==========================================================================================================
Dim C_ACCT_NO
Dim C_ACCT_TYPE_CD
Dim C_ACCT_TYPE_NM
Dim C_DPST_TYPE_CD
Dim C_DPST_TYPE_NM	
Dim C_ACCT_USE
Dim C_ACCT_USE_NM
Dim C_ACCT_PRNT		'>>AIR
Dim C_ACCT_PRNT_NM	'>>AIR
Dim C_BP_CD
Dim C_BP_POPUP
Dim C_BP_NM
Dim C_ACCT_LIMIT



 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey
'Dim lgLngCurRows
Dim strMode

'========================================================================================================= 
Dim IsOpenPop
Dim lgRetFlag
'Dim lgSortKey

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size

    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count

    lgSortKey = 1
End Sub


'========================================================================================
Sub initSpreadPosVariables()

	C_ACCT_NO		= 1
	C_ACCT_TYPE_CD	= 2
	C_ACCT_TYPE_NM	= 3
	C_DPST_TYPE_CD	= 4
	C_DPST_TYPE_NM	= 5
	C_ACCT_USE		= 6
	C_ACCT_USE_NM	= 7
	C_ACCT_PRNT		= 8		'>>AIR
	C_ACCT_PRNT_NM	= 9		'>>AIR
	C_BP_CD         = 10
	C_BP_POPUP		= 11
	C_BP_NM			= 12
	C_ACCT_LIMIT    = 13

End Sub

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
End Sub

'========================================================================================================= 
Sub  SetDefaultVal()

'	Dim NodX
	frm1.cboBankType.value		= "BK"	
	lgBlnFlgChgValue			= False
	frm1.txtzipcodechk.value	= "T"
End Sub

'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread

	With frm1.vspdData
		.MaxCols = C_ACCT_LIMIT + 1
        .MaxRows = 0

        .ReDraw = False 

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit   C_ACCT_NO,	  "계좌번호", 24, , , 30
		ggoSpread.SSSetCombo  C_ACCT_TYPE_CD, "예적금구분", 12
		ggoSpread.SSSetCombo  C_ACCT_TYPE_NM, "예적금구분", 12
		ggoSpread.SSSetCombo  C_DPST_TYPE_CD, "예적금유형", 12
		ggoSpread.SSSetCombo  C_DPST_TYPE_NM, "예적금유형", 12
		ggoSpread.SSSetCombo  C_ACCT_USE,     "계좌구분", 12
		ggoSpread.SSSetCombo  C_ACCT_USE_NM,  "계좌구분", 12
		ggoSpread.SSSetCombo  C_ACCT_PRNT,    "모계좌여부", 12		'>>AIR
		ggoSpread.SSSetCombo  C_ACCT_PRNT_NM, "모계좌여부", 12		'>>AIR	
		ggoSpread.SSSetEdit   C_BP_CD,        "거래처 코드", 20, , , 10, 2
		ggoSpread.SSSetButton C_BP_POPUP
		ggoSpread.SSSetEdit   C_BP_NM,        "거래처명", 30, , , 50
		ggoSpread.SSSetFloat  C_ACCT_LIMIT,   "한도금액",	17, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"

		.ReDraw = True

		Call ggoSpread.MakePairsColumn(C_BP_CD,C_BP_POPUP)

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_ACCT_TYPE_CD,C_ACCT_TYPE_CD,True)
		Call ggoSpread.SSSetColHidden(C_DPST_TYPE_CD,C_DPST_TYPE_CD,True)
		Call ggoSpread.SSSetColHidden(C_ACCT_USE,C_ACCT_USE,True)
		Call ggoSpread.SSSetColHidden(C_ACCT_PRNT,C_ACCT_PRNT,True)	'>>AIR

		Call SetSpreadLock 

	End With

End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData

    ggoSpread.SpreadLock C_ACCT_NO,-1, C_ACCT_NO
    ggoSpread.SpreadLock C_BP_NM,-1,C_BP_NM
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False

    ggoSpread.SSSetRequired	 C_ACCT_NO, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	 C_ACCT_TYPE_NM, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	 C_DPST_TYPE_NM, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BP_NM, pvStartRow, pvEndRow

    .vspdData.ReDraw = True

    End With
End Sub

'========================================================================================================= 
Sub InitComboBox()
	Dim IntRetCD1
	on error resume next

	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A1014", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	If IntRetCD1 <> False Then
		Call SetCombo2(frm1.cboBankType,lgF0,lgF1,chr(11))
	End If
End Sub

'========================================================================================
Sub InitGridComboBox()
	Dim IntRetCD1

	on error resume next

	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("F3011", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If intRetCD1 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ACCT_TYPE_CD
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ACCT_TYPE_NM
	End If

	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("F3012", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	If intRetCD <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DPST_TYPE_CD
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DPST_TYPE_NM
	End If

	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A3010", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	If intRetCD <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(vbTab & lgF0,Chr(11),vbTab), C_ACCT_USE
		ggoSpread.SetCombo Replace(vbTab & lgF1,Chr(11),vbTab), C_ACCT_USE_NM
	End If
	
	'모계좌여부 컬럼 셋팅 >>AIR
	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("F9000", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	If intRetCD <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(vbTab & lgF0,Chr(11),vbTab), C_ACCT_PRNT
		ggoSpread.SetCombo Replace(vbTab & lgF1,Chr(11),vbTab), C_ACCT_PRNT_NM
	End If	

End Sub


'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows

			.Row = intRow

			.Col = C_ACCT_TYPE_CD
			intIndex = .value
			.col = C_ACCT_TYPE_NM
			.value = intindex

			.Col = C_DPST_TYPE_CD
			intIndex = .value
			.col = C_DPST_TYPE_NM
			.value = intindex

			.Col = C_ACCT_USE
			intIndex = .value
			.col = C_ACCT_USE_NM
			.value = intindex
			
			'>>AIR
			.Col = C_ACCT_PRNT
			intIndex = .value
			.col = C_ACCT_PRNT_NM
			.value = intindex

			call subVspdSettingChange(C_BP_CD,intRow)

		Next
	End With
End Sub


'------------------------------------------  OpenPopup()  -------------------------------------------------
'	Name : OpenPopup()
'	Description : Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopup(Byval StrCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    Select Case iWhere

	Case 0
		arrParam(0) = "은행 팝업"				' 팝업 명칭 
		arrParam(1) = "B_BANK"						' TABLE 명칭 
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = ""							' Code Condition
		arrParam(5) = "은행"
	
		arrField(0) = "BANK_CD"						' Field명(0)
		arrField(1) = "BANK_NM"						' Field명(1)
    
		arrHeader(0) = "은행코드"				' Header명(0)
		arrHeader(1) = "은행명"					' Header명(1)
    
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False

		If arrRet(0) = "" then
			frm1.txtBankCd.focus
			Exit Function
		Else
			Call SetPopUp(arrRet, iWhere)
		End If

	Case 1

		iCalledAspName = AskPRAspName("ZipPopup")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZipPopup", "X")
			IsOpenPop = False
			Exit Function
		End If
		arrParam(0) = strCode
		arrParam(1) = ""
		If Trim(frm1.txtCountryCd.value) = "" Then
			arrParam(2) = Parent.gCountry
		Else
			arrParam(2) = frm1.txtCountryCd.value
		End If

		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False

		If arrRet(0) = "" then
			frm1.txtZipCd.focus
			Exit Function
		Else
			Call SetPopUp(arrRet, iWhere)
		End If

	Case 2

		arrParam(0) = "거래처 팝업"							' 팝업 명칭 
		arrParam(1) = "B_BIZ_PARTNER"							' TABLE 명칭 
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = ""										' Code Condition
		arrParam(5) = "거래처"

	    arrField(0) = "BP_CD"									' Field명(0)
		arrField(1) = "BP_NM"									' Field명(1)

	    arrHeader(0) = "거래처 코드"						' Header명(0)
		arrHeader(1) = "거래처 코드명"						' Header명(1)

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False

		If arrRet(0) = "" then
			Exit Function
		Else
			Call SetPopUp(arrRet, iWhere)
		End If	

	Case 3
		arrParam(0) = "국가 팝업"				<%' 팝업 명칭 %>
		arrParam(1) = "b_country"					<%' TABLE 명칭 %>
		arrParam(2) = frm1.txtCountryCd.value		<%' Code Condition%>
		arrParam(3) = ""							<%' Name Cindition%>
		arrParam(4) = ""							<%' Where Condition%>
		arrParam(5) = "국가"						<%' 조건필드의 라벨 명칭 %>

		arrField(0) = "country_cd"					<%' Field명(0)%>
		arrField(1) = "country_nm"					<%' Field명(1)%>

		arrHeader(0) = "국가코드"					<%' Header명(0)%>
		arrHeader(1) = "국가"						<%' Header명(1)%>

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False

		If arrRet(0) = "" then
			frm1.txtCountryCd.focus
			Exit Function
		Else
			Call SetPopUp(arrRet, iWhere)
		End If

	End Select
End Function
'========================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)'

	With frm1

	Select Case iWhere

		Case 0
			frm1.txtBankCd.focus
			.txtBankCd.value = arrRet(0)
			.txtBankNm.value = arrRet(1)

		Case 1
			.txtZipCd.focus
			.txtZipCd.value = arrRet(0)
			.txtAddr1.value  = arrRet(1)

			lgBlnFlgChgValue = True

		Case 2
			.vaSpread1.Col  = C_BP_CD
			.vaSpread1.Text = arrRet(0)
			.vaSpread1.Col  = C_BP_NM
			.vaSpread1.Text = arrRet(1)

			Call vspdData_Change(.vaSpread1.Col, .vaSpread1.Row)
		Case 3

			.txtCountryCd.focus
			.txtCountryCd.value = arrRet(0)
			.txtCountryNm.value = arrRet(1)

			lgBlnFlgChgValue = True

			Call txtCountryCd_OnChange() '
		End Select

	End With
End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function CookiePage(ByVal Kubun)

'	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp

	Select Case Kubun		
	Case "FORM_LOAD"
		strTemp = ReadCookie("BANK_CD")
		Call WriteCookie("BANK_CD", "")
		
		If strTemp = "" then Exit Function
					
		frm1.txtBankCd.value = strTemp
				
		If Err.number <> 0 Then
			Err.Clear
			Call WriteCookie("BANK_CD", "")
			Exit Function 
		End If
				
		Call FncQuery()
	
'	Case JUMP_PGM_ID_BANK_REP
'		Call WriteCookie("BANK_CD", frm1.txtBankCd.value)

	Case Else
		Exit Function
	End Select
End Function



'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029 
	Call ggoOper.LockField(Document, "N")
	'Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitSpreadSheet
	Call InitVariables
	'----------  Coding part  -------------------------------------------------------------
	Call InitComboBox
	Call InitGridComboBox
	Call SetToolbar("1110110100101111")
	Call CookiePage("FORM_LOAD")
	frm1.txtBankCd.focus 
End Sub

'==========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_ACCT_NO		= iCurColumnPos(1)
	        C_ACCT_TYPE_CD  = iCurColumnPos(2)
	        C_ACCT_TYPE_NM  = iCurColumnPos(3)
	        C_DPST_TYPE_CD  = iCurColumnPos(4)
	        C_DPST_TYPE_NM  = iCurColumnPos(5)
			C_ACCT_USE		= iCurColumnPos(6)
			C_ACCT_USE_NM	= iCurColumnPos(7)
			C_ACCT_PRNT		= iCurColumnPos(8)	'>>AIR
			C_ACCT_PRNT_NM	= iCurColumnPos(9)	'>>AIR		
	        C_BP_CD         = iCurColumnPos(10)
	        C_BP_POPUP      = iCurColumnPos(11)
	        C_BP_NM         = iCurColumnPos(12)
	        C_ACCT_LIMIT    = iCurColumnPos(13)

    End Select
End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If

	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If
	End If
    
End Sub


'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
	
	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
	
End Sub
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 '----------  Coding part  -------------------------------------------------------------   
	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
			Case  C_ACCT_TYPE_NM
				.Col = Col
				intIndex = .Value
				.Col = C_ACCT_TYPE_CD
				.Value = intIndex

			Case  C_DPST_TYPE_NM
				.Col = Col
				intIndex = .Value
				.Col = C_DPST_TYPE_CD
				.Value = intIndex
	
			Case  C_ACCT_USE_NM
				.Col = Col
				intIndex = .Value
				.Col = C_ACCT_USE
				.Value = intIndex
				
			'>>AIR	
			Case  C_ACCT_PRNT_NM
				.Col = Col
				intIndex = .Value
				.Col = C_ACCT_PRNT
				.Value = intIndex				
		End Select
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row

	If Col = C_BP_NM Then
		call subVspdSettingChange(C_BP_CD,Row)
	End If

End Sub

'========================================================================================
Sub subVspdSettingChange(ByVal Col , ByVal Row)	
	
	With frm1.vspdData			
		.Col = C_BP_CD
		.Row = Row    	    
				
		If Trim(.Text) <> "" Then
			ggoSpread.SpreadUnLock	C_ACCT_TYPE_NM, Row, C_ACCT_TYPE_NM, Row
			ggoSpread.SpreadUnLock	C_DPST_TYPE_NM, Row, C_DPST_TYPE_NM, Row
		Else
			ggoSpread.SSSetRequired	C_ACCT_TYPE_NM, Row, Row
			ggoSpread.SSSetRequired	C_DPST_TYPE_NM, Row, Row
		End If

	end with	
		
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

	With frm1.vspdData 

		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_BP_POPUP Then
		    .Col = Col
		    .Row = Row
			.Col = C_BP_CD
		    Call OpenPopup(.Text, 2)
			Call SetActiveCell(frm1.vspdData,Col -1,frm1.vspdData.ActiveRow ,"M","X","X")
		End If
    End With
End Sub

'========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 

		If Col <> NewCol Then

			
			If Col = C_BP_CD Then

				call subVspdSettingChange(Col,Row)

			End If
					
		End If


    End With
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			DbQuery
		End If
    End if
    
End Sub


'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False

    Err.Clear

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
    Call SetDefaultVal

    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data

    FncQuery = True																'⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False

    On Error Resume Next
    Err.Clear

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition Field
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call SetDefaultVal

    Call InitVariables

	Call SetToolbar("1110110100101111")

    FncNew = True

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False

    Err.Clear
    On Error Resume Next

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                  '☆:
        Exit Function
    End If

	If frm1.txtBankCd.value = frm1.hBankCd.value Then
    Else
        Call DisplayMsgBox("900009", "X", "X", "X")							'Check if key value is changed
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",Parent.VB_YES_NO, "X", "X")
    
    If IntRetCD = vbNo Then                                                '☜: Delete db data
       Exit Function                                                        '☜:
    End If

    Call DbDelete

    FncDelete = True
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False

   ' On Error Resume Next
    Err.Clear

	If (lgIntFlgMode = Parent.OPMD_UMODE) Then
		If frm1.txtBankCd1.value <> frm1.hBankCd.value Then
			Call DisplayMsgBox("900009","X","X","X")						'Check if key value is changed
			Exit Function
		End If
	End If
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData										'⊙: Preset spreadsheet pointer 
    
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then	'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")						'⊙: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
'   If Not chkField(Document, "1") Then
'      Exit Function
'   End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave
    FncSave = True
End Function

'========================================================================================
Function FncCopy()
	Dim IntRetCD 

	If frm1.vspdData.MaxRows < 1 Then Exit Function

	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    frm1.vspdData.Col = C_ACCT_NO
    frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True
End Function

'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
    Call InitData
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
    Dim imRow

    On Error Resume Next
    Err.Clear   

    FncInsertRow = False

    if IsNumeric(Trim(pvRowCnt)) Then
			imRow = CInt(pvRowCnt)
		else
			imRow = AskSpdSheetAddRowcount()

			If ImRow="" then
			Exit Function
			End If
	End If

	With frm1

		.vspdData.focus
		.vspdData.ReDraw = False

		ggoSpread.Source = .vspdData

		ggoSpread.InsertRow,imRow

		Call SetSpreadLock(.vspdData.ActiveRow)
		Call SetSpreadColor(.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1)

		.vspdData.ReDraw = True
	End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 
		lDelRows = ggoSpread.DeleteRow
    End With

End Function

'========================================================================================
Function FncPrev()
    On Error Resume Next
End Function

'========================================================================================
Function FncNext()
    On Error Resume Next
End Function

'========================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================
Function FncExcel()
    Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
Function FncFind()
    Call parent.FncFind(Parent.C_MULTI, False)
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitGridComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


'========================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

    ggoSpread.Source = frm1.vspdData

	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then

		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X","X")			'⊙: "Will you destory previous data"

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True

End Function

'========================================================================================
Function DbQuery() 
    Dim strVal

    DbQuery = False

    Call LayerShowHide(1)

   ' Err.Clear                                                               '☜: Protect system from crashing

	With frm1

    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QUERY_ID & "?txtMode=" & Parent.UID_M0001	
		strVal = strVal & "&txtBankCd=" & .hBankCd.value
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
	    strVal = BIZ_PGM_QUERY_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtBankCd=" & Trim(.txtBankCd.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If

    End With
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    DbQuery = True

End Function

'========================================================================================
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE
    Call InitData
    Call ggoOper.LockField(Document, "Q")
	Call SetSpreadLock

	Call txtCountryCd_Change()
	Call SetToolbar("1111111100111111")
	frm1.txtBankCd.focus 
	Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal,strDel

    DbSave = False

    On Error Resume Next

    Call LayerShowHide(1)                                                   '☜: Protect system from crashing

	lgRetFlag = False
	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode

	    strMode = .txtMode.value

	    '-----------------------
		'Data manipulate area
	    '-----------------------
		lGrpCnt = 1
	    strVal = ""
		strDel = ""

	    '-----------------------
		'Data manipulate area
	    '-----------------------
		For lRow = 1 To .vspdData.MaxRows

	        .vspdData.Row = lRow
		    .vspdData.Col = 0

			Select Case .vspdData.Text

            Case ggoSpread.InsertFlag											'☜: 신규 
				
				strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep  			'☜: C=Create, Row위치 정보 
                .vspdData.Col = C_ACCT_NO						'1
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_ACCT_TYPE_CD
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_DPST_TYPE_CD
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_ACCT_USE
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_BP_CD
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_ACCT_LIMIT
                strVal = strVal & Trim(UNIConvNum(.vspdData.Text,0)) & Parent.gColSep
				'모계좌여부 >>AIR
                .vspdData.Col = C_ACCT_PRNT
                strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep

                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.UpdateFlag

				strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep				'☜: U=Update, Row위치 정보 
                .vspdData.Col = C_ACCT_NO
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_ACCT_TYPE_CD
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_DPST_TYPE_CD
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_ACCT_USE
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_BP_CD
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_ACCT_LIMIT
                strVal = strVal & Trim(UNIConvNum(.vspdData.Text,0)) & Parent.gColSep
				'>>AIR
                .vspdData.Col = C_ACCT_PRNT
                strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep              
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag											'☜: 삭제 

				strDel = strDel & "D" & Parent.gColSep & lRow & Parent.gColSep				'☜: D=Delete, Row위치 정보 
                .vspdData.Col = C_ACCT_NO	'10
                strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep									
                               
                lGrpCnt = lGrpCnt + 1
                
                lgRetFlag = True
        End Select

    Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)											'☜: 비지니스 ASP 를 가동 

	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
Function DbSaveOk()
	ggoSpread.SSDeleteFlag 1
    Call InitVariables
	frm1.txtBankCd.value = frm1.txtBankCd1.value
	Call FncQuery
End Function

'========================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear

	DbDelete = False
	frm1.txtMode.value = Parent.UID_M0003

    strMode = frm1.txtMode.value
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & Parent.UID_M0003
    strVal = strVal & "&txtBankCd=" & Trim(frm1.txtBankCd.value)				'☜: 삭제 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)
    DbDelete = True
End Function

'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function

'========================================================================================
Sub cboBankType_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Sub txtCountryCd_OnChange()
'	If frm1.txtCountryCd.value <> "" Then
'		If Ucase(frm1.txtCountryCd.value) <> Ucase(Parent.gCountry) Then
'			Call ggoOper.SetReqAttr(frm1.txtZipCd,	"D")
'			Call ggoOper.SetReqAttr(frm1.txtAddr1,	"D")
'			Call ggoOper.SetReqAttr(frm1.txtEngAddr1,	"D")
'			frm1.txtzipcodechk.value = "F"
'		Else
'			Call ggoOper.SetReqAttr(frm1.txtZipCd,	"N")
'			Call ggoOper.SetReqAttr(frm1.txtAddr1,	"N")
'			Call ggoOper.SetReqAttr(frm1.txtEngAddr1,	"N")
'			frm1.txtzipcodechk.value = "T"
'		End If 
'		frm1.txtZipCd.value = ""
'		frm1.txtAddr1.value = "" 
'		frm1.txtEngAddr1.value = ""  
'	End If 
End Sub

'========================================================================================
Sub txtCountryCd_Change()
'	If frm1.txtCountryCd.value <> "" Then
'		If Ucase(frm1.txtCountryCd.value) <> Ucase(Parent.gCountry) Then
'			Call ggoOper.SetReqAttr(frm1.txtZipCd,	"D")
'			Call ggoOper.SetReqAttr(frm1.txtAddr1,	"D")
'			Call ggoOper.SetReqAttr(frm1.txtEngAddr1,	"D")
'			frm1.txtzipcodechk.value = "F"
'		Else
'			Call ggoOper.SetReqAttr(frm1.txtZipCd,	"N")
'			Call ggoOper.SetReqAttr(frm1.txtAddr1,	"N")
'			Call ggoOper.SetReqAttr(frm1.txtEngAddr1,	"N")
'			frm1.txtzipcodechk.value = "T"
'		End If
'	End If 
End Sub


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>은행정보등록(KO441)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>은행</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtBankCd" MAXLENGTH="10" SIZE=10 ALT ="은행코드" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(txtBankCd.value, 0)">
														   <INPUT NAME="txtBankNm" MAXLENGTH="30" STYLE="TEXT-ALIGN:left" ALT ="은행명" tag="14X"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
							    <TD CLASS=TD5 NOWRAP>은행코드</TD>
							    <TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtBankCd1"  ALT="은행코드" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag="23XXXU"> </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>은행약어명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBankShNm" ALT="은행약어명" MAXLENGTH="70" SIZE=35 STYLE="TEXT-ALIGN: left" tag ="22N"></TD>
								<TD CLASS=TD5 NOWRAP>은행명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBankFullNm" ALT="은행명" MAXLENGTH="70" SIZE=35 STYLE="TEXT-ALIGN: left" tag ="21N"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>은행영문명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBankEngNm" ALT="은행영문명" MAXLENGTH="70" SIZE=35 STYLE="TEXT-ALIGN:left" tag  ="21N"></TD>
								<TD CLASS=TD5 NOWRAP>금융기관 TYPE</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboBankType" ALT="금융기관 TYPE" STYLE="WIDTH: 120px"   tag="21"><OPTION VALUE=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>국가</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCountryCd" SIZE=10 MAXLENGTH=2 tag="21XXXU" ALT="국가코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCountryCd.value,3)">
													   <INPUT NAME="txtCountryNm" ALT="국가명" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X">
								</TD>
								<TD CLASS=TD5 NOWRAP>우편번호</TD>
								<TD CLASS=TD6 COLSPAN=3 NOWRAP><INPUT NAME="txtZipCd" ALT="우편번호" MAXLENGTH="12" SIZE = 10 STYLE="TEXT-ALIGN:left" tag ="21"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnZipCode" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtZipCd.value,1)"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>은행주소 1</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAddr1" ALT="은행주소 1" MAXLENGTH="70" SIZE=35 STYLE="TEXT-ALIGN:left" tag  ="21"></TD>
								<TD CLASS=TD5 NOWRAP>은행영문주소 1</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEngAddr1" ALT="은행영문주소 1" MAXLENGTH="30" SIZE=35 STYLE="TEXT-ALIGN:left" tag  ="21"></TD>
							</TR>
							<TR>	
								<TD CLASS=TD5 NOWRAP>은행주소 2</TD>
								<TD	CLASS=TD6 NOWRAP><INPUT NAME="txtAddr2" ALT="은행주소 2" MAXLENGTH="70" SIZE=35 STYLE="TEXT-ALIGN:left" tag="21">
								<TD CLASS=TD5 NOWRAP>은행영문주소 2</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEngAddr2" ALT="은행영문주소 2" MAXLENGTH="30" SIZE=35 STYLE="TEXT-ALIGN:left" tag="21"></TD>
							</TR>
							<TR>	
								<TD CLASS=TD5 NOWRAP>은행주소 3</TD>
								<TD	CLASS=TD6 NOWRAP><INPUT NAME="txtAddr3" ALT="은행주소 3" MAXLENGTH="70" SIZE=35 STYLE="TEXT-ALIGN:left" tag="21"></TD>
								<TD CLASS=TD5 NOWRAP>은행영문주소 3</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEngAddr3" ALT="은행영문주소 3" MAXLENGTH="30" SIZE=35 STYLE="TEXT-ALIGN:left" tag="21"></TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
									<script language =javascript src='./js/b1310ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</Table>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE><TEXTAREA class=hidden name=txtSpread tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hBankCd" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtzipcodechk" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

