<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Quality Management
'*  2. Function Name		:
'*  3. Program ID			: XI316QA1_ko119
'*  4. Program Name			: Q-FOCUS 전송현황
'*  5. Program Desc			:
'*  6. Component List		:
'*  7. Modified date(First)	: 2007/05/10
'*  8. Modified date(Last)	:
'*  9. Modifier (First)		: Kim Jin Tae
'* 10. Modifier (Last)		:
'* 11. Comment				:
'* 12. Common Coding Guide	: this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History				:
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit											'☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim strInspClass
Dim IsOpenPop
'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID		= "XI316QB1_ko119.asp"			'☆: 비지니스 로직 ASP명

Dim C_InvoiceNo
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_PalletNo
Dim C_TrayNo
Dim C_LotNo
Dim C_ProductionDt
Dim C_MatLotNo
Dim C_Part1
Dim C_Part2
Dim C_Part3
Dim C_Part4
Dim C_Part5
Dim C_Part6
Dim C_Part7
Dim C_Part8
Dim C_Part9
Dim C_Part10
Dim C_Part11
Dim C_Part12
Dim C_ReceiveDt
Dim C_SendYN
Dim C_SendDt
Dim C_CreateType
Dim C_PalletQty
Dim C_ProdtOrderNo

'--------------- 개발자 coding part(변수선언,End  )-----------------------------------------------------------

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
Dim CompanyYMD
CompanyYMD = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 시작 날짜
'--------------- 개발자 coding part(실행로직,End  )-----------------------------------------------------------

'==========================================  2.1.1 InitVariables()  ==========================================
'	Name		: InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=============================================================================================================
Sub InitVariables()
	lgBlnFlgChgValue = False
	IsOpenPop = False
    '###검사분류별 변경부분 Start###
    strInspClass = "R"
	'###검사분류별 변경부분 End  ###
End Sub

'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name		: SetDefaultVal()
'	Description :
'=============================================================================================================
Sub SetDefaultVal()
	Call CommonQueryRs(" UD_MINOR_CD, UD_MINOR_NM "," B_USER_DEFINED_MINOR "," UD_MAJOR_CD=" & FilterVar("QR001","''","S") & " ORDER BY UD_MINOR_CD",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	frm1.txtDtFr.text = CompanyYMD
	frm1.txtDtTo.text = CompanyYMD
End Sub

'============================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'============================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE", "QA") %>
End Sub

'------------------------------------------  OpenPlant()  ---------------------------------------------------
'	Name		: OpenPlant()
'	Description : Plant PopUp
'------------------------------------------------------------------------------------------------------------
Function OpenPlant()		' 공장에 대한 목록을 팝업으로 보여주는곳
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"							' 팝업 명칭
	arrParam(1) = "B_Plant"								' TABLE 명칭
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Condition
	arrParam(4) = ""
	arrParam(5) = "공장"								' TextBox 명칭

    arrField(0) = "Plant_Cd"							' Field명(0)
    arrField(1) = "Plant_NM"							' Field명(1)

    arrHeader(0) = "공장코드"							' Header명(0)
    arrHeader(1) = "공장명"								' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtPlantCd.Focus
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
	Else
		Exit Function
	End If
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'-------------------------------------------------------------------------------------------------------------
Function OpenItemInfo(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		IsOpenPop = False
		Exit Function
	End If

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분
	arrParam(3) = ""							' Default Value

	arrField(0) = 1 '"ITEM_CD"					' Field명(0)
	arrField(1) = 2 '"ITEM_NM"					' Field명(1)

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	

	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)		
End Function

'=======================================================================================================
'   Event Name : txtStartDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDtFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtFr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDtFr.Focus
	End If
End Sub

'=======================================================================================================
'   Event Name : txtEndDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDtTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDtTo.Focus
	End If
End Sub

'========================================= 2.6 InitSpreadSheet() =============================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=============================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20070515", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_ProdtOrderNo + 1
		.MaxRows = 0

 		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6", "15", "0")
		Call AppendNumberPlace("7", "13", "2")

		ggoSpread.SSSetEdit  C_InvoiceNo,	"INVOICE NO",	15
		ggoSpread.SSSetEdit  C_ItemCd,		"품목코드",		10
		ggoSpread.SSSetEdit  C_ItemNm,		"품목명",		10
		ggoSpread.SSSetEdit  C_Spec,		"규격",			25
		ggoSpread.SSSetEdit  C_PalletNo,	"PALLET NO",	15
		ggoSpread.SSSetEdit  C_TrayNo,		"TRAY NO",		15
		ggoSpread.SSSetEdit  C_LotNo,		"LOT NO",		15
		ggoSpread.SSSetDate  C_ProductionDt,"생산일",		10, 2, Parent.gDateFormat
		ggoSpread.SSSetEdit  C_MatLotNo,	"자재 LOT NO",	20
		ggoSpread.SSSetEdit  C_Part1,		"PART1",		10
		ggoSpread.SSSetEdit  C_Part2,		"PART2",		10
		ggoSpread.SSSetEdit  C_Part3,		"PART3",		10
		ggoSpread.SSSetEdit  C_Part4,		"PART4",		10
		ggoSpread.SSSetEdit  C_Part5,		"PART5",		10
		ggoSpread.SSSetEdit  C_Part6,		"PART6",		10
		ggoSpread.SSSetEdit  C_Part7,		"PART7",		10
		ggoSpread.SSSetEdit	 C_Part8,		"PART8",		10
		ggoSpread.SSSetEdit  C_Part9,		"PART9",		10
		ggoSpread.SSSetEdit  C_Part10,		"PART10",		10
		ggoSpread.SSSetEdit  C_Part11,		"PART11",		10
		ggoSpread.SSSetEdit  C_Part12,		"PART12",		15
		ggoSpread.SSSetDate  C_ReceiveDt,	"수신일시",		10, 2, Parent.gDateFormat
		ggoSpread.SSSetEdit  C_SendYN,		"전송여부",		10
		ggoSpread.SSSetEdit  C_SendDt,		"전송일시",		10
'		ggoSpread.SSSetDate  C_SendDt,		"전송일시",		10, 2, Parent.gDateFormat
		ggoSpread.SSSetEdit  C_CreateType,	"생성타입",		10
		ggoSpread.SSSetFloat C_PalletQTy,	"PALLET 수량",	10, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , , "Z"
		ggoSpread.SSSetEdit	 C_ProdtOrderNo,"제조오더번호", 10
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		ggoSpread.SpreadLockWithOddEvenRowColor()
	    ggoSpread.SSSetSplit2(3)

		.ReDraw = true
    End With
End Sub

'==========================================  2.6.1 InitSpreadPosVariables()  =================================
Sub InitSpreadPosVariables()
	C_InvoiceNo		= 1
	C_ItemCd		= 2
	C_ItemNm		= 3
	C_Spec			= 4
	C_PalletNo		= 5
	C_TrayNo		= 6
	C_LotNo			= 7
	C_ProductionDt	= 8
	C_MatLotNo		= 9
	C_Part1			= 10
	C_Part2			= 11
	C_Part3			= 12 
	C_Part4			= 13
	C_Part5			= 14
	C_Part6			= 15
	C_Part7			= 16
	C_Part8			= 17
	C_Part9			= 18
	C_Part10		= 19
	C_Part11		= 20
	C_Part12		= 21
	C_ReceiveDt		= 22
	C_SendYN		= 23
	C_SendDt		= 24
	C_CreateType	= 25
	C_PalletQty		= 26
	C_ProdtOrderNo	= 27
End Sub

'==========================================  2.6.2 GetSpreadColumnPos()  =====================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos

 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_InvoiceNo		= iCurColumnPos(1)
		C_ItemCd		= iCurColumnPos(2)
		C_ItemNm		= iCurColumnPos(3)
		C_Spec			= iCurColumnPos(4)
		C_PalletNo		= iCurColumnPos(5)
		C_TrayNo		= iCurColumnPos(6)
		C_LotNo			= iCurColumnPos(7)
		C_ProductionDt	= iCurColumnPos(8)
		C_MatLotNo		= iCurColumnPos(9)
		C_Part1			= iCurColumnPos(10)
		C_Part2			= iCurColumnPos(11)
		C_Part3			= iCurColumnPos(12)
		C_Part4			= iCurColumnPos(13)
		C_Part5			= iCurColumnPos(14)
		C_Part6			= iCurColumnPos(15)
		C_Part7			= iCurColumnPos(16)
		C_Part8			= iCurColumnPos(17)
		C_Part9			= iCurColumnPos(18)
		C_Part10		= iCurColumnPos(19)
		C_Part11		= iCurColumnPos(20)
		C_Part12		= iCurColumnPos(21)
		C_ReceiveDt		= iCurColumnPos(22)
		C_SendYN		= iCurColumnPos(23)
		C_SendDt		= iCurColumnPos(24)
		C_CreateType	= iCurColumnPos(25)
		C_PalletQty		= iCurColumnPos(26)
		C_ProdtOrderNo	= iCurColumnPos(27)
 	End Select
End Sub

'==========================================  3.1.1 Form_Load()  ======================================================
'	Name		: Form_Load()
'	Description :
'=====================================================================================================================
Sub Form_Load()
	Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")

	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어

'--------------- 개발자 coding part(실행로직,Start)------------------------------------------------------------------
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
	   	frm1.txtPlantNm.value = Parent.gPlantNm
	End If

	frm1.txtPlantCd.focus
'--------------- 개발자 coding part(실행로직,End  )------------------------------------------------------------------
End Sub

'====================================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'====================================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'====================================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'====================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then Exit Sub

 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey
 			lgSortKey = 1
 		End If
 	End If
End Sub

'====================================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc :
'====================================================================================================================
Sub vspdData_MouseDown(Button, Shift, x, y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
    End If
End Sub

'====================================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'====================================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'====================================================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정
'====================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'====================================================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경
'====================================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'====================================================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'====================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'====================================================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'====================================================================================================================
Sub PopRestoreSpreadColumnInf()	'###그리드 컨버전 주의부분###
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
 	'------ Developer Coding part (Start) ---------------------------------------------------------------------------
 	'------ Developer Coding part (End  ) ---------------------------------------------------------------------------
End Sub

'====================================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'====================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    If OldLeft <> NewLeft Then Exit Sub

	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then Exit Sub

			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If
End Sub

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************************
Function FncQuery()
    FncQuery = False                                                        '⊙: Processing is NG

    Err.Clear                                                               '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then Exit Function						'⊙: This function check indispensable field

    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData

	If Name_check("A") = False Then
		Set gActiveElement = document.activeElement
		Exit Function
	End If

    Call InitVariables

    '-----------------------
    'Query function call area
    '-----------------------
    lgStrPrevKey = ""
	If DbQuery = False then Exit Function

    FncQuery = True															'⊙: Processing is OK
End Function

'====================================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'====================================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'====================================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'====================================================================================================================
Function FncExcel()
	Call parent.FncExport(Parent.C_MULTI)
End Function

'====================================================================================================================
' Function Name : FncFind
' Function Desc : 폼에서 찾기를 하는 함수
'====================================================================================================================
Function FncFind()
    Call parent.FncFind(Parent.C_MULTI, False)								'☜:화면 유형, Tab 유무
End Function

'====================================================================================================================
' Function Name : FncExit
' Function Desc : 폼에서 나기기(종료)를 하는 함수
'====================================================================================================================
Function FncExit()
    FncExit = True
End Function

'====================================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'====================================================================================================================
Function DbQuery()
	Dim strVal
    DbQuery = False

    Err.Clear																'☜: Protect system from crashing
	Call LayerShowHide(1)


    With frm1
		If .rdoFlagYes.checked = True Then
			.txtrdoflag.value = "Y"
		ElseIf .rdoFlagNo.checked = True Then
			.txtrdoflag.value = "N"
		ElseIf .rdoFlagAll.checked = True Then
			.txtrdoflag.value = "A"
		Else
		End If

		strVal = BIZ_PGM_ID & "?txtPlantCd="	& Trim(.txtPlantCd.value) & _
							  "&txtrdoFlag="	& Trim(.txtrdoflag.value) & _
							  "&txtDtFr="		& Trim(.txtDtFr.Text) & _
							  "&txtDtTo="		& Trim(.txtDtTo.Text) & _
							  "&txtItemCd="		& Trim(.txtItemCd.value) & _
							  "&txtMaxRows="	& .vspdData.MaxRows & _
							  "&lgStrPrevKey="	& lgStrPrevKey				'☜: Next key tag
		Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동
	End With

	DbQuery = True
End Function

'====================================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김
'====================================================================================================================
Function DbQueryOk()
    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolbar("11000000000111")
	lgBlnFlgChgValue = False

	Set gActiveElement = document.activeElement
End Function

'====================================================================================================================
' Function Name : Name_Check
'====================================================================================================================
Function Name_Check(ByVal Check)
	Name_Check = False

	With frm1
		'-----------------------
		'Check Plant_Cd
		'-----------------------
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) = False Then

			.txtPlantNm.Value = ""
			Call DisplayMsgBox("125000", "X", "X", "X")
			.txtPlantCd.focus
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		.txtPlantNm.Value = lgF0(0)
	End With

	Name_Check = True
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR><TD <%=HEIGHT_TYPE_00%>></TD></TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif">
									<IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB">
									<FONT color=white>Q-FOCUS 전송현황</FONT></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right">
									<IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						   	</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR><TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
        							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;
															<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE="20" MAXLENGTH=40 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>수신일자</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/xi316qa1_ko119_fpDateTime5_txtDtFr.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/xi316qa1_ko119_fpDateTime6_txtDtTo.js'></script>
									</TD>
								</TR>
								<TR>
	     							<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;
															<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>전송여부</TD>
	        						<TD CLASS=TD6 NOWRAP><INPUT TYPE=radio CLASS=Radio NAME=rdoFlag ID=rdoFlagAll tag="12" VALUE = "A" CHECKED><LABEL FOR=rdoFlagAll>전체</LABEL>
												   &nbsp;<INPUT TYPE=radio CLASS=Radio NAME=rdoFlag ID=rdoFlagYes tag="12" VALUE = "Y"><LABEL FOR=rdoFlagYes>전송</LABEL>
												   &nbsp;<INPUT TYPE=radio CLASS=Radio NAME=rdoFlag ID=rdoFlagNo  tag="12" VALUE = "N"><LABEL FOR=rdoFlagNo>미전송</LABEL>
									</TD>
	     						</TR>
	     						<INPUT TYPE=hidden NAME="work_flag">
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>
						<TR>
							<TD HEIGHT=100% WIDTH=100% Colspan=2>
								<SCRIPT LANGUAGE =javascript src='./js/xi316qa1_I597258698_vspdData.js'></SCRIPT>
							</TD>
						</TR>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR><TD <%=HEIGHT_TYPE_01%>></TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtrdoflag" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
		<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
    </DIV>
</BODY>
</HTML>