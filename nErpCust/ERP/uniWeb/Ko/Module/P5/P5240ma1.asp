<%@	LANGUAGE="VBSCRIPT"	%>
<!--
'**********************************************************************************************
'*	1. Module	Name					:	Prucurement
'*	2. Function	Name				:
'*	3. Program ID						:	MC200MA1
'*	4. Program Name					:	납입지시조정 
'*	5. Program Desc					:	납입지시조정 
'*	6. Component List				:
'*	7. Modified	date(First)	:	2003-04-08
'*	8. Modified	date(Last)	:	2003/05/23
'*	9. Modifier	(First)			:	Ahn	Jung Je
'* 10. Modifier	(Last)			:	Kang Su	Hwan
'* 11. Comment							:
'* 12. Common	Coding Guide	:	this mark(☜)	means	that "Do not change"
'*														this mark(⊙)	Means	that "may	 change"
'*														this mark(☆)	Means	that "must change"
'* 13. History							:
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선	언 부 
'##########################################################################################################	-->
<!-- '******************************************	1.1	Inc	선언	 **********************************************
'	기능:	Inc. Include
'********************************************************************************************************* -->
<!-- #Include	file="../../inc/IncSvrCcm.inc" -->
<!-- #Include	file="../../inc/incSvrHTML.inc"	-->
<!--'==========================================	 1.1.1 Style Sheet	=======================================-->
<LINK	REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================	 1.1.2 공통	Include		=====================================-->

<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT	LANGUAGE=VBSCRIPT>
Option Explicit															'☜: indicates that	All	variables	must be	declared in	advance

'******************************************	 1.2 Global	변수/상수	선언	***********************************
'	1. Constant는	반드시 대문자	표기.
'**********************************************************************************************************
Const	BIZ_PGM_ID = "P5240mb1.asp"
'============================================	 1.2.1 Global	상수 선언	 ==================================
'========================================================================================================

<%'========================================================================================================%>

Dim	C_Plant_Cd
Dim	C_Plant_Nm
Dim	C_Item_Cd
Dim	C_Item_Nm
Dim	C_Spec
Dim	C_Unit
Dim	C_Rev_Inv
Dim	C_Curr_Inv

Const	C_SHEETMAXROWS = 50


'==========================================	 1.2.2 Global	변수 선언	 =====================================
'	1. 변수	표준에 따름. prefix로	g를	사용함.
'	2.Array인	경우는 ()를	반드시 사용하여	일반 변수와	구별해 됨 
'=========================================================================================================
<!-- #Include	file="../../inc/lgvariables.inc" -->

Dim	IsOpenPop
Dim	strDate
Dim	iDBSYSDate

'==========================================	 2.1.1 InitVariables()	======================================
'	Name : InitVariables()
'	Description	:	변수 초기화(Global 변수, 초기화가	필요한 변수	또는 Flag들을	Setting한다.)
'=========================================================================================================
<%'========================================================================================================%>

Sub	InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE
	lgPageNo			 = ""
		lgBlnFlgChgValue = False
		lgIntGrpCount	=	0
		lgStrPrevKey = ""
		lgLngCurRows = 0
		frm1.vspdData.MaxRows	=	0
End	Sub

'==========================================	 2.2.1 SetDefaultVal()	======================================
'	Name : SetDefaultVal()
'	Description	:	화면 초기화(수량 Field나 그	외 화면이	뜰 때	Default값을	정해줘야 하는	Field들	Setting)
'=========================================================================================================
Sub	SetDefaultVal()
    Dim strYear
    Dim strMonth
    Dim strDay

    Call ExtractDateFrom("<%= GetSvrDate %>",parent.gServerDateFormat , parent.gServerDateType      ,strYear,strMonth,strDay)

	frm1.txtYyyyMm.Year  = strYear
	frm1.txtYyyyMm.Month = strMonth

    frm1.txtItemCd.focus    

	Call SetToolbar("11000000000011")
End	Sub

'======================================================================================
'	Function Name	:	LoadInfTB19029
'	Function Desc	:	This method	loads	format inf
'======================================================================================
Sub	LoadInfTB19029()
	<!-- #Include	file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call	loadInfTB19029A("I", "*","NOCOOKIE","MA")	%>
	<% Call	LoadBNumericFormatA("I", "*","NOCOOKIE","MA")	%>
End	Sub

'============================= 2.2.3 InitSpreadSheet() ================================
'	Function Name	:	InitSpreadSheet
'	Function Desc	:	This method	initializes	spread sheet column	property
'======================================================================================
<%'========================================================================================================%>

Sub	InitSpreadSheet()
	Call InitSpreadPosVariables()

	'------------------------------------------
	'	Grid 1 - Operation Spread	Setting
	'------------------------------------------
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030226",	,Parent.gAllowDragDropSpread

		.ReDraw	=	false

		.MaxCols = C_Curr_Inv	+	1
		.MaxRows = 0

		Call GetSpreadColumnPos("A")



		ggoSpread.SSSetEdit	C_Plant_Cd,	"공장",	8,,,,2
		ggoSpread.SSSetEdit	C_Plant_Nm,	"공장명",	15
		ggoSpread.SSSetEdit	C_Item_Cd, "품목", 18
		ggoSpread.SSSetEdit	C_Item_Nm, "품목명", 30
		ggoSpread.SSSetEdit	C_Spec,	"규격",	8
		ggoSpread.SSSetEdit	C_Unit,	"단위",	8
		ggoSpread.SSSetFloat C_Rev_Inv,	"전월재고",	12,Parent.ggQtyNo	,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_Curr_Inv, "현재월재고", 12,Parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	Parent.gComNum1000,	Parent.gComNumDec



		Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)

		Call SetSpreadLock

		.ReDraw	=	true
		End	With
End	Sub

'==================================	2.2.4	SetSpreadLock()	==================================================
'	Function Name	:	SetSpreadLock
'	Function Desc	:	This method	set	color	and	protect	in spread	sheet	celles
'========================================================================================
Sub	SetSpreadLock()
		With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock		C_Plant_Cd,				-1,	C_Curr_Inv				,-1
			ggoSpread.SSSetProtected frm1.vspdData.MaxCols,	-1
		.vspdData.ReDraw = True
		End	With
End	Sub


'============================	 2.2.7 InitSpreadPosVariables()	===========================
'	Function Name	:	InitSpreadPosVariables
'	Function Desc	:	This method	Assigns	Sequential Number	to spread	sheet	column
'========================================================================================
Sub	InitSpreadPosVariables()
	C_Plant_Cd	 = 1
	C_Plant_Nm	 = 2
	C_Item_Cd		 = 3
	C_Item_Nm		 = 4
	C_Spec			 = 5
	C_Unit			 = 6
	C_Rev_Inv		 = 7
	C_Curr_Inv	 = 8
End	Sub

'============================	 2.2.8 GetSpreadColumnPos()	 ==============================
'	Function Name	:	GetSpreadColumnPos
'	Function Desc	:	This method	is used	to get specific	spreadsheet	column position	according	to the arguement
'========================================================================================
Sub	GetSpreadColumnPos(ByVal pvSpdNo)
		Dim	iCurColumnPos

		Select Case	UCase(pvSpdNo)
		Case "A"

 			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Plant_Cd	=	iCurColumnPos(1	)
			C_Plant_Nm	=	iCurColumnPos(2	)
			C_Item_Cd		=	iCurColumnPos(3	)
			C_Item_Nm		=	iCurColumnPos(4	)
			C_Spec			=	iCurColumnPos(5	)
			C_Unit			=	iCurColumnPos(6	)
			C_Rev_Inv		=	iCurColumnPos(7	)
			C_Curr_Inv	=	iCurColumnPos(8	)

		End	Select
End	Sub




'==========================================	 3.1.1 Form_Load()	======================================
'	Name : Form_Load()
'	Description	:	Window On	Load(공통	Include	파일에 선언)시 변수초기화	및 화면초기화를	하기 위해	함수를 Call하는	부분 
'=========================================================================================================
Sub	Form_Load()
	Call LoadInfTB19029																											'⊙: Load	table	,	B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")																		'⊙: Lock	 Suitable	 Field
	Call ggoOper.FormatDate(frm1.txtYyyyMm,	Parent.gDateFormat,	"2")

	Call InitSpreadSheet																										'⊙: Setup the Spread	sheet
	Call InitComboBox

	Call SetDefaultVal
	Call InitVariables																											'⊙: Initializes local global	variables


	Call SetToolBar("11000000000011")

End	Sub


'========================================================================================================
Sub	InitComboBox()
	'------	Developer	Coding part	(Start ) --------------------------------------------------------------
		Dim	iCodeArr
		Dim	iNameArr
	'------	Developer	Coding part	(End )	 --------------------------------------------------------------
End	Sub


'=======================================================================================================
'		Event	Name : txtYyyyMm_DblClick(Button)
'		Event	Desc : 달력을	호출한다.
'=======================================================================================================
Sub	txtYyyyMm_DblClick(Button)
		If Button	=	1	Then
				frm1.txtYyyyMm.Action	=	7
				Call SetFocusToDocument("M")
				frm1.txtYyyyMm.Focus
		End	If
End	Sub


Function txtYyyyMm_KeyPress(KeyAscii)
 If	KeyAscii = 13	Then
	Call MainQuery()
 End If
End	Function


'==========================================================================================
'		Event	Name : vspdData_GotFocus
'		Event	Desc : This	event	is spread	sheet	data changed
'==========================================================================================
Sub	vspdData_GotFocus()
		ggoSpread.Source = frm1.vspdData
End	Sub

'==========================================================================================
'		Event	Name : vspdData_TopLeftChange
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'==========================================================================================
Sub	vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop	,	ByVal	NewLeft	,	ByVal	NewTop )

		If OldLeft <>	NewLeft	Then
				Exit Sub
		End	If

	If CheckRunningBizProcess	=	True Then
		 Exit	Sub
	End	If

'		MSGBOX frm1.vspdData.MaxRows
'		MSGBOX VisibleRowCnt(frm1.vspdData,NewTop)
'		MSGBOX NEWTOP

		if frm1.vspdData.MaxRows < NewTop	+	VisibleRowCnt(frm1.vspdData,NewTop)	Then	'☜: 재쿼리	체크'
		If lgPageNo	<> ""	Then														'⊙: 다음	키 값이	없으면 더	이상 업무로직ASP를 호출하지	않음 
					 Call	DisableToolBar(Parent.TBC_QUERY)
					 Call	DbQuery
			End	If
	End	if

End	Sub


'==========================================================================================
'		Event	Name : vspdData_DblClick
'		Event	Desc : This	event	is spread	sheet	data changed
'==========================================================================================
Sub	vspdData_ButtonClicked(ByVal Col,	ByVal	Row, Byval ButtonDown)
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_BpCdPopup Then				'공급처 
				.Col = Col
				.Row = Row
				Call OpenBP()
		End	If
		End	With
End	Sub

'==========================================================================================
'		Event	Name : vspdData_Click
'		Event	Desc :
'==========================================================================================
Sub	vspdData_Click(ByVal Col , ByVal Row )
	Dim	IntRetCD

	If frm1.vspdData.MaxRows > 0 Then
		Call SetPopupMenuItemInf("0001111111")
	Else
		Call SetPopupMenuItemInf("0000111111")
	End	If
	'----------------------
	'Column	Split
	'----------------------
	gMouseClickStatus	=	"SPC"

	Set	gActiveSpdSheet	=	frm1.vspdData

		If frm1.vspdData.MaxRows = 0 Then	Exit Sub																									 'If there is	no data.

	 	frm1.vspdData.Row	=	frm1.vspdData.ActiveRow


	 	If Row <=	0	Then
				ggoSpread.Source = frm1.vspdData
				If lgSortKey = 1 Then
						ggoSpread.SSSort Col
						lgSortKey	=	2
				Else
						ggoSpread.SSSort Col,	lgSortKey
						lgSortKey	=	1
				End	If
		Else

		End	If
End	Sub

'==========================================================================================
'		Event	Name : vspdData_MouseDown(Button,Shift,x,y)
'		Event	Desc :
'==========================================================================================
Sub	vspdData_MouseDown(Button,Shift,x,y)
	If Button	=	2	And	gMouseClickStatus	=	"SPC"	Then
			 gMouseClickStatus = "SPCR"
		End	If
End	Sub

'========================================================================================================
'		Event	Name : vspdData_ColWidthChange
'		Event	Desc :
'========================================================================================================
Sub	vspdData_ColWidthChange(ByVal	pvCol1,	ByVal	pvCol2)
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End	Sub

'========================================================================================================
'		Event	Name : vspdData_ScriptDragDropBlock
'		Event	Desc :
'========================================================================================================
Sub	vspdData_ScriptDragDropBlock(	Col	,	 Row,	 Col2,	Row2,	 NewCol,	NewRow,	 NewCol2,	 NewRow2,	 Overwrite , Action	,	DataOnly , Cancel	)
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.SpreadDragDropBlock(Col , Row,	Col2,	Row2,	NewCol,	NewRow,	NewCol2, NewRow2,	Overwrite	,	Action , DataOnly	,	Cancel )
		Call GetSpreadColumnPos("A")
End	Sub

'==========================================================================================
'		Event	Name : vspdData_Change
'		Event	Desc :
'==========================================================================================
Sub	vspdData_Change(ByVal	Col	,	ByVal	Row	)
	ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow	Row

	lgBlnFlgChgValue = True

	Frm1.vspdData.Row	=	Row
	Frm1.vspdData.Col	=	Col

	Call CheckMinNumSpread(frm1.vspdData,	Col, Row)				 '	<------변경된	표준 라인 
End	Sub

'========================================================================================
'	Function Name	:	FncQuery
'	Function Desc	:	This function	is related to	Query	Button of	Main ToolBar
'========================================================================================
Function FncQuery()
		Dim	IntRetCD

		FncQuery = False														'⊙: Processing	is NG
		Err.Clear																'☜: Protect system	from crashing

	ggoSpread.Source = frm1.vspdData

		'-----------------------
		'Check previous	data area
		'-----------------------
		If ggoSpread.SSCheckChange = true	Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X",	"X")
		If IntRetCD	=	vbNo Then	Exit Function
		End	If

		'-----------------------
		'Check condition area
		'-----------------------
		If Not chkfield(Document,	"1") Then	Exit Function									'⊙: This	function check indispensable field
		If ChkKeyField()=False Then Exit Function 
		'-----------------------
		'Erase contents	area
		'-----------------------
		Call ggoOper.ClearField(Document,	"2")									'⊙: Clear Contents	 Field
		Call InitVariables														'⊙: Initializes local global	variables



		'-----------------------
		'Query function	call area
		'-----------------------
		If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function														'☜: Query db	data
	End	If

	Set	gActiveElement = document.activeElement
		FncQuery = True															'⊙: Processing	is OK
End	Function

'===========================================	5.1.2	FncNew()	===========================================
'=	Event	Name : FncNew																					=
'=	Event	Desc : This	function is	related	to New Button	of Main	ToolBar									=
'========================================================================================================
Function FncNew()

End	Function

'========================================================================================
'	Function Name	:	FncSave
'	Function Desc	:	This function	is related to	Delete Button	of Main	ToolBar
'========================================================================================
Function FncSave()
		Dim	IntRetCD
		Dim	intRow

		FncSave	=	False

		Err.Clear

		If CheckRunningBizProcess	=	True Then	Exit Function

	ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = False Then
				IntRetCD = DisplayMsgBox("900001", "X",	"X", "X")
				Exit Function
		End	If

		ggoSpread.Source = frm1.vspdData
		If Not ggoSpread.SSDefaultCheck	Then Exit	Function


		'-----------------------
		'Save	function call	area
		'-----------------------
		If DbSave	=	False	Then Exit	Function

	Set	gActiveElement = document.activeElement
		FncSave	=	True
End	Function

'========================================================================================
'	Function Name	:	FncCancel
'	Function Desc	:	This function	is related to	Cancel Button	of Main	ToolBar
'========================================================================================
Function FncCancel()
	if frm1.vspdData.Maxrows < 1	then exit	function
	ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo
	Set	gActiveElement = document.activeElement
End	Function

'========================================================================================
'	Function Name	:	FncInsertRow
'	Function Desc	:	This function	is related to	InsertRow	Button of	Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal	pvRowCnt)

End	Function

'========================================================================================
'	Function Name	:	FncPrint
'	Function Desc	:	This function	is related to	Print	Button of	Main ToolBar
'========================================================================================
Function FncPrint()
	Call parent.FncPrint()
	Set	gActiveElement = document.activeElement
End	Function

'========================================================================================
'	Function Name	:	FncExcel
'	Function Desc	:	This function	is related to	Excel
'========================================================================================
Function FncExcel()
		Call parent.FncExport(parent.C_MULTI)									'☜: Protect system	from crashing
	Set	gActiveElement = document.activeElement
End	Function

'========================================================================================
'	Function Name	:	FncFind
'	Function Desc	:
'========================================================================================
Function FncFind()
		Call parent.FncFind(parent.C_MULTI,	False)								'☜: Protect system	from crashing
	Set	gActiveElement = document.activeElement
End	Function

'========================================================================================
'	Function Name	:	FncSplitColumn
'	Function Desc	:
'========================================================================================
Sub	FncSplitColumn()
		If UCase(Trim(TypeName(gActiveSpdSheet)))	=	"EMPTY"	Then Exit	Sub

		ggoSpread.Source = gActiveSpdSheet
		ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
	Set	gActiveElement = document.activeElement
End	Sub

'========================================================================================
'	Function Name	:	FncExit
'	Function Desc	:
'========================================================================================
Function FncExit()
	Dim	IntRetCD

		On Error Resume	Next																													'☜: If	process	fails
		Err.Clear																																			'☜: Clear error status

	FncExit	=	False

		ggoSpread.Source = frm1.vspdData

	If lgBlnFlgChgValue	=	True Or	ggoSpread.SSCheckChange	=	True Then

		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")

		If IntRetCD	=	vbNo Then	Exit Function

		End	If

	Set	gActiveElement = document.activeElement
		FncExit	=	True
End	Function

'******************	 5.2 Fnc함수명에서 호출되는	개발 Function	 **************************
'	설명 :
'****************************************************************************************
'========================================================================================
'	Function Name	:	DbQuery
'	Function Desc	:	This function	is data	query	and	display
'========================================================================================
Function DbQuery()
	Dim	strVal
		Dim	strYear1
		Dim	strMonth1
		Dim	strDay1
		Dim	strDate1

		DbQuery	=	False

	Call LayerShowHide(1)


	Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
	strDate1 = strYear1	&	strMonth1

	With frm1
			strVal = BIZ_PGM_ID	&	"?txtMode="		&	parent.UID_M0001						'☜:
			strVal = strVal	&	"&txtYear="			&	Trim(strYear1)			'☆: 조회	조건 데이타 
			strVal = strVal	&	"&txtMonth="			&	Trim(strMonth1)			'☆: 조회	조건 데이타 
			strVal = strVal	&	"&txtItemCd="			&	Trim(.txtItemCd.value)
			strVal = strVal	&	"&lgStrPrevKey="		&	lgStrPrevKey
			strVal = strVal	&	"&lgIntFlgMode="		&	lgIntFlgMode
			strVal = strVal	&	"&txtMaxRows="		&	.vspdData.MaxRows
			strVal = strVal	&	"&lgPageNo="			&	lgPageNo													'☜: Next	key	tag
	End	With

		Call RunMyBizASP(MyBizASP, strVal)														'☜: 비지니스	ASP	를 가동 

		DbQuery	=	True
End	Function


'========================================================================================
'	Function Name	:	DbSave
'	Function Desc	:	This function	is data	save
'========================================================================================
Function DbSave()


End	Function
'========================================================================================
'	Function Name	:	DbQueryOk
'	Function Desc	:	DbQuery가	성공적일 경우	MyBizASP 에서	호출되는 Function, 현재	FncQuery에 있는것을	옮김 
'========================================================================================
Function DbQueryOk()
	Call SetToolBar("11000000000011")														'⊙: 버튼	툴바 제어 
	lgIntFlgMode = parent.OPMD_UMODE														'⊙: Indicates that	current	mode is	Update mode
		lgBlnFlgChgValue = False
End	Function


'========================================================================================
'	Function Name	:	DbSaveOk
'	Function Desc	:	DBSave가 성공적일	경우 MyBizASP	에서 호출되는	Function,	현재 FncSave에 있는것을	옮김 
'========================================================================================
Function DbSaveOk()
	Call InitVariables()
	Call MainQuery()
End	Function
'========================================================================================
'	Function Name	:	PopSaveSpreadColumnInf
'	Function Desc	:	그리드 현상태를	저장한다.
'========================================================================================
Sub	PopSaveSpreadColumnInf()
		ggoSpread.Source = gActiveSpdSheet
		Call ggoSpread.SaveSpreadColumnInf()
End	Sub

'========================================================================================
'	Function Name	:	PopRestoreSpreadColumnInf
'	Function Desc	:	그리드를 예전	상태로 복원한다.
'========================================================================================
Sub	PopRestoreSpreadColumnInf()
	Dim	LngRow

		ggoSpread.Source = gActiveSpdSheet
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet
		ggoSpread.Source = gActiveSpdSheet

	Call ggoSpread.ReOrderingSpreadData()
End	Sub

'########################################################################################
'########################################################################################
'# Area	Name	 : User-defined	Method Part
'# Description : This	part declares	user-defined method
'########################################################################################
'########################################################################################


'-----------------------	OpenItem()	-------------------------------------------------
Function OpenItem()
	If IsOpenPop = True	Then Exit	Function

	Dim	arrRet
	Dim	arrParam(5), arrField(6)
	Dim	iCalledAspName

	IsOpenPop	=	True



	arrParam(0)	=	"품목팝업"
	arrParam(1)	=	"B_Item_By_Plant,	B_Item"
	arrParam(2)	=	Trim(frm1.txtItemCd.value)
	arrParam(4)	=	"B_Item_By_Plant.Item_Cd = B_Item.Item_Cd	and	B_Item.phantom_flg = 'N' "

	arrParam(5)	=	"품목"

	arrField(0)	=	"B_Item_By_Plant.Item_Cd"
	arrField(1)	=	"B_Item.Item_NM"


	iCalledAspName = AskPRAspName("m1111pa1")

	If Trim(iCalledAspName)	=	"" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION,	"m1111pa1",	"X")
		IsOpenPop	=	False
		Exit Function
	End	If

	arrRet = window.showModalDialog(iCalledAspName,	Array(window.parent, arrParam,arrField), _
		"dialogWidth=760px;	dialogHeight=420px;	center:	Yes; help: No; resizable:	No;	status:	No;")

	IsOpenPop	=	False

	If arrRet(0) = ""	Then
		Exit Function
	Else
		frm1.txtItemCd.value = arrRet(0)
		frm1.txtItemNm.value = arrRet(1)
		frm1.txtItemCd.focus

	End	If
End	Function

'==========================================  2.2.6 ChkKeyField()  =======================================
'	Name : ChkKeyField()
'	Description : 
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       

	ChkKeyField = true		

	If Trim(frm1.txtItemCd.value) <> "" Then
		strWhere = " B_Item_By_Plant.Item_Cd = B_Item.Item_Cd	and	B_Item.phantom_flg = 'N'  and B_Item.item_Cd =  " & FilterVar(frm1.txtItemCd.value, "''", "S") & "  "

		Call CommonQueryRs(" B_Item_By_Plant.Item_Cd,B_Item.Item_NM "," B_Item_By_Plant,	 B_Item ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X","품목코드","X")
			frm1.txtItemCd.focus 
			frm1.txtItemNm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtItemNm.value = strDataNm(0)
	Else
		frm1.txtItemNM.value = ""
	End If
	
	

End Function


'☜: 아래	OBJECT Tag는 InterDev	사용자를 위한것으로	프로그램이 완성되면	아래 Include 코드로	대체되어야 한다 
</SCRIPT>
<!-- #Include	file="../../inc/UNI2KCM.inc" -->
</HEAD>

<BODY	TABINDEX="-1"	SCROLL="no">
<FORM	NAME=frm1	TARGET="MyBizASP"	METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD	<%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR	HEIGHT=23>
		<TD	WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD	WIDTH=10>&nbsp;</TD>
					<TD	CLASS="CLSMTABP">
						<TABLE ID="MyTab"	CELLSPACING=0	CELLPADDING=0>
							<TR>
								<td	background="../../../CShared/image/table/seltab_up_bg.gif"><img	src="../../../CShared/image/table/seltab_up_left.gif"	width="9"	height="23"></td>
								<td	background="../../../CShared/image/table/seltab_up_bg.gif" align="center"	CLASS="CLSMTAB"><font	color=white>설비부품재고현황</font></td>
								<td	background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img	src="../../../CShared/image/table/seltab_up_right.gif" width="10"	height="23"></td>
								</TR>
						</TABLE>
					</TD>
					<TD	WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR	HEIGHT=*>
		<TD	WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD	<%=HEIGHT_TYPE_02%>	WIDTH=100%></TD>
				</TR>
				<TR>
					<TD	HEIGHT=20	WIDTH=100%>
						<FIELDSET	CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
			 					<TR>
									<TD	CLASS="TD5"	NOWRAP>품목</TD>
									<TD	CLASS="TD6"	NOWRAP>
										<INPUT TYPE=TEXT NAME="txtItemCd"	SIZE="18"	MAXLENGTH="18" STYLE="Text-Transform:	uppercase" ALT="품목"	tag="11XXXU" ><IMG align=top height=20 name="btnItemCd"	onclick="vbscript:OpenItem()"	src="../../../CShared/image/btnPopup.gif"	width=16	TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtItemNm"	CLASS=protected	readonly=true	TABINDEX="-1"	SIZE="20"	tag="14" >
									</TD>
											<TD	CLASS="TD5"	NOWRAP>연월</TD>
											<TD	CLASS="TD6"	NOWRAP>
											<script language =javascript src='./js/p5240ma1_txtYyyyMm_txtYyyyMm.js'></script>
											</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD	<%=HEIGHT_TYPE_03%>	WIDTH=100%></TD>
				</TR>
				<TR>
				<TD	WIDTH=100% valign=top><TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD	HEIGHT="100%">
								<script language =javascript src='./js/p5240ma1_OBJECT3_vspdData.js'></script>
						</TD>
					</TR></TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
		<tr>
		<TD	<%=HEIGHT_TYPE_01%>></TD>
		</tr>
	<TR>
		<TD	WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100%	HEIGHT=<%=BizSize%>	FRAMEBORDER=0	SCROLLING=No noresize	framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode"	tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hAppToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hFacility_Accnt"	tag="24">
<INPUT TYPE=HIDDEN NAME="hFacility_Cd" tag="24">
</FORM>
<DIV ID="MousePT"	NAME="MousePT">
<iframe	name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO	noresize framespacing=0	width=220	height=41	src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

