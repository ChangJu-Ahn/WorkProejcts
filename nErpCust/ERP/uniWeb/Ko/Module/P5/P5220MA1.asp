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
Const	BIZ_PGM_ID = "P5220mb1.asp"
'============================================	 1.2.1 Global	상수 선언	 ==================================
'========================================================================================================

<%'========================================================================================================%>

Dim	C_Plan_Dt
Dim	C_Insp_Flag
Dim	C_Insp_FlagNm
Dim	C_Facility_CD
Dim	C_FacilityPop
Dim	C_Facility_Nm
Dim	C_Set_Plant
Dim	C_Set_PlantNm
Dim	C_Facility_Accnt_Nm
Dim	C_Plant_Sts
Dim	C_Chk_End_Dt
Dim	C_Rep_End_Dt

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
	Dim	LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtAppFrDt.text	=	UniConvDateAToB(UNIDateAdd ("D", -7, LocSvrDate, parent.gServerDateFormat),	parent.gServerDateFormat,	parent.gDateFormat)
	frm1.txtAppToDt.text	=	UniConvDateAToB(UNIDateAdd ("D", 7,	LocSvrDate,	parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	
	
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		Set gActiveElement = document.activeElement  
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
	End If
	
	Call SetToolbar("1110111100001111")
End	Sub

'======================================================================================
'	Function Name	:	LoadInfTB19029
'	Function Desc	:	This method	loads	format inf
'======================================================================================
Sub	LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("Q", "P", "NOCOOKIE", "MA") %>
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

		.MaxCols = C_Rep_End_Dt	+	1
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetDate		C_Plan_Dt, 				"계획일자",	12,2,gDateFormat
		ggoSpread.SSSetCombo 	C_Insp_Flag,			"점검여부",	 10, 0,	False
		ggoSpread.SSSetCombo 	C_Insp_FlagNm,			"점검여부",	 20, 0,	False
		ggoSpread.SSSetEdit		C_FACILITY_CD,			"설비코드",			15,,,18,2
		ggoSpread.SSSetButton	C_FacilityPop
		ggoSpread.SSSetEdit		C_FACILITY_NM,			"설비명",			18,,,40,2
		ggoSpread.SSSetEdit		C_SET_PLANT,			"설치공장",		15,,,4,2
		ggoSpread.SSSetEdit		C_SET_PLANTNm,			"설치공장",	15,,,40,2
		ggoSpread.SSSetEdit		C_FACILITY_ACCNT_NM,	"설비유형",			15,,,20,2
		ggoSpread.SSSetEdit		C_Plant_Sts,			"설비등급",			15,,,20,2
		ggoSpread.SSSetDate		C_Chk_End_Dt,			"최종점검일",	12,2,gDateFormat
		ggoSpread.SSSetDate		C_Rep_End_Dt,			"최종수리일",	12,2,gDateFormat


		Call ggoSpread.SSSetColHidden(C_Insp_FlagNm, C_Insp_FlagNm,	True)
		Call ggoSpread.SSSetColHidden(C_SET_PLANT, C_SET_PLANT,	True)
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
		ggoSpread.SSSetRequired		C_Insp_Flag		,	-1,	C_Insp_Flag

		ggoSpread.SpreadLock		C_Plan_Dt,				-1,	C_Plan_Dt				,-1
		ggoSpread.SpreadLock		C_FACILITY_CD,				-1,	C_Rep_End_Dt				,-1
			ggoSpread.SSSetProtected frm1.vspdData.MaxCols,	-1
		.vspdData.ReDraw = True
		End	With
End	Sub

Sub	SetSpreadColor(ByVal pvStarRow,	Byval	pvEndRow)
		ggoSpread.Source = frm1.vspdData
		With frm1
			.vspdData.ReDraw = False
		ggoSpread.SSSetRequired	 	C_Insp_Flag, pvStarRow,	pvEndRow
		ggoSpread.SSSetRequired	 	C_Plan_Dt, pvStarRow,	pvEndRow
		ggoSpread.SSSetRequired	 	C_FACILITY_CD, pvStarRow,	pvEndRow
 		ggoSpread.SpreadLock		C_FACILITY_NM,				-1,	C_Rep_End_Dt				,-1

		ggoSpread.SSSetProtected frm1.vspdData.MaxCols,	pvStarRow, pvEndRow
		.vspdData.ReDraw = True
		End	With
End	Sub

'============================	 2.2.7 InitSpreadPosVariables()	===========================
'	Function Name	:	InitSpreadPosVariables
'	Function Desc	:	This method	Assigns	Sequential Number	to spread	sheet	column
'========================================================================================
Sub	InitSpreadPosVariables()
	C_Plan_Dt								=	1
	C_Insp_Flag							=	2
	C_Insp_FlagNm						=	3
	C_Facility_CD						=	4
	C_FacilityPop			=	5
	C_Facility_Nm						=	6
	C_Set_Plant							=	7
	C_Set_PlantNm						=	8
	C_Facility_Accnt_Nm			=	9
	C_Plant_Sts							=	10
	C_Chk_End_Dt						=	11
	C_Rep_End_Dt						=	12
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

			C_Plan_Dt				=	iCurColumnPos(1	)
			C_Insp_Flag				=	iCurColumnPos(2	)
			C_Insp_FlagNm			=	iCurColumnPos(3	)
			C_Facility_CD			=	iCurColumnPos(4	)
			C_FacilityPop			=	iCurColumnPos(5	)
			C_Facility_Nm			=	iCurColumnPos(6	)
			C_Set_Plant				=	iCurColumnPos(7	)
			C_Set_PlantNm			=	iCurColumnPos(8	)
			C_Facility_Accnt_Nm		=	iCurColumnPos(9)
			C_Plant_Sts				=	iCurColumnPos(10)
			C_Chk_End_Dt			=	iCurColumnPos(11)
			C_Rep_End_Dt			=	iCurColumnPos(12)

		End	Select
End	Sub

'------------------------------------------	 OpenPlant()	-------------------------------------------------
'	Name : OpenPlant()
'	Description	:	Plant	PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim	arrRet
	Dim	arrParam(5), arrField(6),	arrHeader(6)

	If IsOpenPop = True	Then Exit	Function

	IsOpenPop	=	True

	arrParam(0)	=	"공장팝업"
	arrParam(1)	=	"B_PLANT"
	arrParam(2)	=	Trim(frm1.txtPlantCd.Value)
	arrParam(3)	=	""
	arrParam(4)	=	""
	arrParam(5)	=	"공장"

		arrField(0)	=	"PLANT_CD"
		arrField(1)	=	"PLANT_NM"

		arrHeader(0) = "공장"
		arrHeader(1) = "공장명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",	Array(arrParam,	arrField,	arrHeader),	_
		"dialogWidth=420px;	dialogHeight=450px;	center:	Yes; help: No; resizable:	No;	status:	No;")

	IsOpenPop	=	False

	If arrRet(0) = ""	Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value		 = arrRet(0)
		frm1.txtPlantNm.Value		 = arrRet(1)
		frm1.txtPlantCd.focus
	End	If
End	Function

'========================================================================================================
'	Name : OpenFacility_Popup()
'	Desc : developer describe	this line
'========================================================================================================
Function OpenFacility_Popup(Byval	iWhere)
	Dim	arrRet
	Dim	arrParam(5), arrField(6),	arrHeader(6)
	Dim strPlant, strWhere


	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
		Exit Function
	End If

	If frm1.txtPlantCd.value <> "" Then
		strPlant = frm1.txtPlantCd.value
	Else
		strPlant = "%"
	End if

	If IsOpenPop = True	 Then
		Exit Function
	End	If

	IsOpenPop	=	True


	strWhere = " SET_PLANT LIKE " & FilterVar(Trim(strPlant), "''", "S")

	Select Case	iWhere
		Case "1"
			arrParam(0)	=	"설비코드 팝업"
			arrParam(1)	=	"Y_FACILITY"
			arrParam(2)	=	frm1.txtFacility_Cd.value
			arrParam(3)	=	""												'	Name Cindition
			arrParam(4)	=	strWhere										'	Where	Condition
			arrParam(5)	=	"설비코드"									'	TextBox	명칭 

			arrField(0)	=	"Facility_cd"									'	Field명(0)
			arrField(1)	=	"Facility_Nm"									'	Field명(1)

			arrHeader(0) = "설비코드"									'	Header명(0)
			arrHeader(1) = "설비코드명"									'	Header명(1)
		Case "2"
			arrParam(0)	=	"설비코드	팝업"
			arrParam(1)	=	"Y_FACILITY"
			frm1.vspdData.Col	=	C_FACILITY_CD
			arrParam(2)	=	frm1.vspdData.text
			arrParam(3)	=	""												'	Name Cindition
			arrParam(4)	=	strWhere																		'	Where	Condition
			arrParam(5)	=	"설비코드"												'	TextBox	명칭 

			arrField(0)	=	"Facility_cd"											'	Field명(0)
			arrField(1)	=	"Facility_Nm"											'	Field명(1)

			arrHeader(0) = "설비코드"									'	Header명(0)
			arrHeader(1) = "설비코드명"									'	Header명(1)

	End	Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",	Array(arrParam,	arrField,	arrHeader),	_
	"dialogWidth=420px;	dialogHeight=450px;	center:	Yes; help: No; resizable:	No;	status:	No;")

	IsOpenPop	=	False
	If arrRet(0) = ""	Then
		 Exit	Function
	Else
		 Call	SetCondArea(arrRet,iWhere)
	End	If

End	Function

'======================================================================================================
'	Name : SetCondArea()
'	Description	:	Item Popup에서 Return되는	값 setting
'=======================================================================================================
Sub	SetCondArea(Byval	arrRet,	Byval	iWhere)
	With Frm1
		Select Case	iWhere
			Case "1"
					.txtFacility_Cd.value	=	arrRet(0)
					.txtFacility_Nm.value	=	arrRet(1)
					.txtFacility_Cd.focus
			Case "2"
				.vspdData.Col	=	C_FACILITY_CD
				.vspdData.text = arrRet(0)
				.vspdData.Col	=	C_FACILITY_NM
				.vspdData.text = arrRet(1)
				

				Call vspdData_Change(C_FACILITY_CD,	frm1.vspdData.ActiveRow)
		End	Select
	End	With
End	Sub



'==========================================	 3.1.1 Form_Load()	======================================
'	Name : Form_Load()
'	Description	:	Window On	Load(공통	Include	파일에 선언)시 변수초기화	및 화면초기화를	하기 위해	함수를 Call하는	부분 
'=========================================================================================================
Sub	Form_Load()
	Call LoadInfTB19029																											'⊙: Load	table	,	B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")																		'⊙: Lock	 Suitable	 Field
	Call InitSpreadSheet																										'⊙: Setup the Spread	sheet
	Call InitComboBox

	Call SetDefaultVal
	Call InitVariables																											'⊙: Initializes local global	variables


	Call SetToolBar("1110111100001111")

End	Sub


'========================================================================================================
Sub	InitComboBox()
	'------	Developer	Coding part	(Start ) --------------------------------------------------------------
		Dim	iCodeArr
		Dim	iNameArr

		Call CommonQueryRs(" MINOR_CD, MINOR_NM	","	B_MINOR	","	MAJOR_CD = 'Z410'	",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		Call SetCombo2(frm1.CboFacility_Accnt	,lgF0	 ,lgF1	,Chr(11))


	ggoSpread.SetCombo "Y" & vbtab & "N" , C_Insp_Flag
'		ggoSpread.SetCombo "완료"	&	vbtab	&	"미완료" , C_Insp_FlagNm

	'------	Developer	Coding part	(End )	 --------------------------------------------------------------
End	Sub

'=======================================================================================================
'		Event	Name : txtAppFrDt_DblClick(Button)
'		Event	Desc : 달력을	호출한다.
'=======================================================================================================
Sub	txtAppFrDt_DblClick(Button)
	If Button	=	1	Then
		frm1.txtAppFrDt.Action = 7
			Call SetFocusToDocument("M")
				frm1.txtAppFrDt.Focus
	End	If
End	Sub

'=======================================================================================================
'		Event	Name : txtAppToDt_DblClick(Button)
'		Event	Desc : 달력을	호출한다.
'=======================================================================================================
Sub	txtAppToDt_DblClick(Button)
	If Button	=	1	Then
		frm1.txtAppToDt.Action = 7
			Call SetFocusToDocument("M")
				frm1.txtAppToDt.Focus
	End	If
End	Sub

'=======================================================================================================
'		Event	Name : txtAppFrDt_KeyDown(keycode, shift)
'		Event	Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub	txtAppFrDt_KeyDown(keycode,	shift)
	If keycode = 13	Then
		Call MainQuery()
	End	If
End	Sub

'=======================================================================================================
'		Event	Name : txtAppToDt_KeyDown(keycode, shift)
'		Event	Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub	txtAppToDt_KeyDown(keycode,	shift)
	If keycode = 13	Then
		Call MainQuery()
	End	If
End	Sub

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

		If Row > 0 And Col = C_FacilityPop Then				'공급처 
				.Col = Col
				.Row = Row
				Call OpenFacility_Popup("2")
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
dim IntRetCd
	Select Case	Col
		Case	 C_Facility_CD
			If Trim(Frm1.vspdData.Text)	=	"" Then
					Frm1.vspdData.Col	=	C_Facility_Nm
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_SET_PLANT
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_SET_PLANTNm
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_FACILITY_ACCNT_NM
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_Plant_Sts
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_Chk_End_Dt
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_Rep_End_Dt
					Frm1.vspdData.Text	= ""
			Else
				Frm1.vspdData.Col	=	C_Facility_Cd

				IntRetCd =	CommonQueryRs("	Facility_nm, set_plant, Plant_Nm, dbo.ufn_GetCodeName('Z410', Facility_Accnt), dbo.ufn_GetCodeName('Z423', Plant_Sts), Chk_End_Dt, Rep_End_Dt"," Y_FACILITY, B_Plant b "," Y_FACility.set_plant = b.plant_cd and FACILITY_CD="	&	FilterVar(Trim(Frm1.vspdData.Text), "''", "S")	&	"	",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IntRetCd	=	false	then
					Call DisplayMsgBox("970000","X","설비","X")

					Frm1.vspdData.Col	=	C_Facility_Cd
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_Facility_Nm
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_SET_PLANT
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_SET_PLANTNm
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_FACILITY_ACCNT_NM
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_Plant_Sts
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_Chk_End_Dt
					Frm1.vspdData.Text	= ""
					Frm1.vspdData.Col	=	C_Rep_End_Dt
					Frm1.vspdData.Text	= ""
				Else
					Frm1.vspdData.Col	=	C_Facility_Nm
					Frm1.vspdData.Text	=	Trim(Replace(lgF0,Chr(11),""))
					Frm1.vspdData.Col	=	C_SET_PLANT
					Frm1.vspdData.Text	=	Trim(Replace(lgF1,Chr(11),""))
					Frm1.vspdData.Col	=	C_SET_PLANTNm
					Frm1.vspdData.Text	=	Trim(Replace(lgF2,Chr(11),""))
					Frm1.vspdData.Col	=	C_FACILITY_ACCNT_NM
					Frm1.vspdData.Text	=	Trim(Replace(lgF3,Chr(11),""))
					Frm1.vspdData.Col	=	C_Plant_Sts
					Frm1.vspdData.Text	=	Trim(Replace(lgF4,Chr(11),""))
					Frm1.vspdData.Col	=	C_Chk_End_Dt
					Frm1.vspdData.Text	=	Trim(Replace(lgF5,Chr(11),""))
					Frm1.vspdData.Col	=	C_Rep_End_Dt
					Frm1.vspdData.Text	=	Trim(Replace(lgF6,Chr(11),""))

				End	if
			End	if
	End	Select



	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow	Row


	lgBlnFlgChgValue = True

	Frm1.vspdData.Row	=	Row
	Frm1.vspdData.Col	=	Col

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


	If ValidDateCheck(frm1.txtAppFrDt, frm1.txtAppToDt)	=	False	Then Exit	Function

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
	Dim	IntRetCD

	FncNew = False

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True	Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "x",	"x")
		If IntRetCD	=	vbNo Then	Exit Function
	End	If

	Call ggoOper.ClearField(Document,	"1")
	Call ggoOper.ClearField(Document,	"2")
	Call ggoOper.LockField(Document, "N")
	Call SetDefaultVal
	Call SetToolBar("11100000000011")

	Call InitVariables


	Set	gActiveElement = document.activeElement
	FncNew = True
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
	Dim	IntRetCD
	Dim	imRow, iRow
	Dim	lgF0
	Dim	lgF1
	Dim	lgF2
	Dim	lgF3
	Dim	lgF4
	Dim	lgF5
	Dim	lgF6

	On Error Resume	Next																													'☜: If	process	fails
	Err.Clear																																			'☜: Clear error status

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow	=	CInt(pvRowCnt)
	Else
		imRow	=	AskSpdSheetAddRowCount()
	
		If imRow = ""	Then
			Exit Function
		End	if
	End	If

	With frm1
		If Not chkField(Document,	"2") Then
			Exit Function
		End	If

		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow, imRow
		SetSpreadColor .vspdData.ActiveRow,	.vspdData.ActiveRow	+	imRow	-1

		'.vspdData.Row=	.vspdData.ActiveRow
		For	iRow = .vspdData.ActiveRow to	.vspdData.ActiveRow	+	imRow	-1
'				.vspdData.Row	=	iRow
'				.vspdData.Col= C_OrderUnit
'				.vspdData.Text = lgF0(0)

			.vspdData.Row	=	iRow
			.vspdData.Col= C_Plan_Dt
			.vspdData.Text = UNIFormatDate("<%=CDate(GetSvrDate)+1%>")
		Next
		.vspdData.ReDraw = True
	End	With
	Set	gActiveElement = document.ActiveElement
	If Err.number	=	0	Then
		FncInsertRow = True																													 '☜:	Processing is	OK
	End	If
End	Function


'========================================================================================
'	Function Name	:	FncDeleteRow
'========================================================================================
Function FncDeleteRow()

	Dim	lDelRows
	Dim	lTempRows

	lDelRows = ggoSpread.DeleteRow
	lgLngCurRows = lDelRows	+	lgLngCurRows
	lTempRows	=	frm1.vspdData.MaxRows	-	lgLngCurRows
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

	With frm1
			strVal = BIZ_PGM_ID	&	"?txtMode="	&	parent.UID_M0001						'☜:
			strVal = strVal	&	"&txtPlantCd="		&	UCase(Trim(.txtPlantCd.value))		'☆: 조회	조건 데이타 
			strVal = strVal	&	"&txtAppFrDt="		&	Trim(.txtAppFrDt.text)			'☆: 조회	조건 데이타 
			strVal = strVal	&	"&txtAppToDt="		&	Trim(.txtAppToDt.text)			'☆: 조회	조건 데이타 
			strVal = strVal	&	"&CboFacility_Accnt="	&	Trim(.CboFacility_Accnt.value)
			strVal = strVal	&	"&txtFacility_Cd="	&	Trim(.TxtFacility_Cd.value)
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
	Dim	pP21011
	Dim	lRow
	Dim	lGrpCnt
	Dim	retVal
	Dim	boolCheck
	Dim	lStartRow
	Dim	lEndRow
	Dim	lRestGrpCnt
	Dim	strVal,	strDel

	DbSave = False

	If LayerShowHide(1)	=	False	Then
		Exit Function
	End	If

	strVal = ""
	strDel = ""
	lGrpCnt	=	1


	ggoSpread.Source = frm1.vspdData
		With Frm1

			 For lRow	=	1	To .vspdData.MaxRows
					 .vspdData.Row = lRow
					 .vspdData.Col = 0
					 Select	Case .vspdData.Text
							 Case	 ggoSpread.InsertFlag																			 '☜:	Create
																strVal = strVal	&	"C"	&	parent.gColSep
																strVal = strVal	&	lRow & parent.gColSep
										.vspdData.Col	=	C_FACILITY_CD			:	strVal = strVal	&	Trim(.vspdData.Text) & parent.gColSep
										.vspdData.Col	=	C_Plan_Dt				:	strVal = strVal	&	Trim(.vspdData.Text) & parent.gColSep
										.vspdData.Col	=	C_Insp_Flag				:	strVal = strVal	&	Trim(.vspdData.Text) & parent.gRowSep
										lGrpCnt	=	lGrpCnt	+	1
							 Case	 ggoSpread.UpdateFlag																			 '☜:	Update
																strVal = strVal	&	"U"	&	parent.gColSep
																strVal = strVal	&	lRow & parent.gColSep
										.vspdData.Col	=	C_FACILITY_CD			:	strVal = strVal	&	Trim(.vspdData.Text) & parent.gColSep
										.vspdData.Col	=	C_Plan_Dt				:	strVal = strVal	&	Trim(.vspdData.Text) & parent.gColSep
										.vspdData.Col	=	C_Insp_Flag				:	strVal = strVal	&	Trim(.vspdData.Text) & parent.gRowSep
										lGrpCnt	=	lGrpCnt	+	1
							 Case	 ggoSpread.DeleteFlag																			 '☜:	Delete

																strDel = strDel	&	"D"	&	parent.gColSep
																strDel = strDel	&	lRow & parent.gColSep
										.vspdData.Col	=	C_FACILITY_CD			:	strDel = strDel	&	Trim(.vspdData.Text) & parent.gColSep
										.vspdData.Col	=	C_Plan_Dt				:	strDel = strDel	&	Trim(.vspdData.Text) & parent.gColSep
										.vspdData.Col	=	C_Insp_Flag				:	strDel = strDel	&	Trim(.vspdData.Text) & parent.gRowSep
										lGrpCnt	=	lGrpCnt	+	1
					 End Select
			 Next
			 .txtMode.value				 =	parent.UID_M0002
			 .txtMaxRows.value		 = lGrpCnt-1
			 .txtSpread.value			 = strDel	&	strVal
		End	With
	Call ExecMyBizASP(frm1,	BIZ_PGM_ID)

		DbSave = True



End	Function
'========================================================================================
'	Function Name	:	DbQueryOk
'	Function Desc	:	DbQuery가	성공적일 경우	MyBizASP 에서	호출되는 Function, 현재	FncQuery에 있는것을	옮김 
'========================================================================================
Function DbQueryOk()
	Call SetToolBar("11101111000111")														'⊙: 버튼	툴바 제어 
	lgIntFlgMode = parent.OPMD_UMODE														'⊙: Indicates that	current	mode is	Update mode
		lgBlnFlgChgValue = False
End	Function

Sub	RemovedivTextArea()
	Dim	ii

	For	ii = 1 To	divTextArea.children.length
			divTextArea.removeChild(divTextArea.children(0))
	Next
End	Sub

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
	Call InitComboBox

	ggoSpread.Source = gActiveSpdSheet

	Call ggoSpread.ReOrderingSpreadData()
End	Sub

'########################################################################################
'########################################################################################
'# Area	Name	 : User-defined	Method Part
'# Description : This	part declares	user-defined method
'########################################################################################
'########################################################################################


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

	If Trim(frm1.txtPlantCd.value) <> "" Then
		strWhere = " Plant_Cd =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "  "

		Call CommonQueryRs(" plant_nm "," b_plant ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X","공장","X")
			frm1.txtPlantCd.focus 
			frm1.txtPlantNm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPlantNm.value = strDataNm(0)
	Else
		frm1.txtPlantNm.value = ""
	End If
	
	If Trim(frm1.txtFacility_Cd.value) <> "" Then
		strWhere = " Facility_cd =  " & FilterVar(frm1.txtFacility_Cd.value, "''", "S") & "  "

		Call CommonQueryRs(" Facility_Nm "," Y_FACILITY ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X","설비코드","X")
			frm1.txtFacility_Cd.focus 
			frm1.txtFacility_nm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtFacility_nm.value = strDataNm(0)
	Else
		frm1.txtFacility_nm.value = ""
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
								<td	background="../../../CShared/image/table/seltab_up_bg.gif" align="center"	CLASS="CLSMTAB"><font	color=white>설비점검계획수립및수정</font></td>
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
									<TD CLASS="TD5" NOWRAP>설치공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="설치공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=33 tag="14XXXU"></TD>
									<TD	CLASS=TD5	NOWRAP>계획기간</TD>
									<TD	CLASS="TD6">
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/p5220ma1_OBJECT1_txtAppFrDt.js'></script>
												</td>
												<td>&nbsp;~&nbsp;</td>
												<td>
													<script language =javascript src='./js/p5220ma1_OBJECT2_txtAppToDt.js'></script>
												</td>
											<tr>
										</table>
									</TD>
								</TR>
								<TR>
									<TD	CLASS="TD5"	NOWRAP>설비유형</TD>
									<TD	CLASS="TD6"	NOWRAP><SELECT NAME="CboFacility_Accnt"	ALT="설비유형" CLASS ="CboFacility_Accnt"	TAG="1XN"><OPTION	VALUE=""></OPTION></SELECT></TD>
									<TD	CLASS="TD5"	NOWRAP>설비코드</TD>
									<TD	CLASS="TD6"	NOWRAP><INPUT	ID=txtFacility_Cd	NAME="txtFacility_Cd"	ALT="설비코드" TYPE="Text" SiZE="18" MAXLENGTH="18"	tag="11XXXU"><IMG	SRC="../../../CShared/image/btnPopup.gif"	NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript:	OpenFacility_Popup('1')">
															<INPUT ID=txtFacility_Nm NAME="txtFacility_Nm" ALT="설비코드명"	TYPE="Text"	SiZE="25"	MAXLENGTH="40" tag="14XXXU"></TD>
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
								<script language =javascript src='./js/p5220ma1_OBJECT3_vspdData.js'></script>
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
<TEXTAREA	CLASS="hidden" NAME="txtSpread"	tag="24" TABINDEX	=	"-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"	tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hAppFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hAppToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hFacility_Accnt"	tag="24">
<INPUT TYPE=HIDDEN NAME="hFacility_Cd" tag="24">
</FORM>
<DIV ID="MousePT"	NAME="MousePT">
<iframe	name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO	noresize framespacing=0	width=220	height=41	src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>








