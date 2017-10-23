<%@	LANGUAGE="VBSCRIPT"	%>
<!--
======================================================================================================
*	 1.	Module Name					 : Facility	Resources
*	 2.	Function Name				 : 설비점검내용등록 
*	 3.	Program	ID					 : FA105MA1
*	 4.	Program	Name				 :
*	 5.	Program	Desc				 : 설비점검내용등록 
*	 6.	Comproxy List				 :
*	 7.	Modified date(First) : 2005/01/17
*	 8.	Modified date(Last)	 : 2005/01/17
*	 9.	Modifier (First)		 : Lee Chang-Je
*	10.	Modifier (Last)			 : Lee Chang-Je
*	11.	Comment							 : Who Let the dog out?
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include	file="../../inc/IncSvrCcm.inc" -->
<!-- #Include	file="../../inc/incSvrHTML.inc"	-->

<LINK	REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<Script	Language="VBScript">
Option Explicit
'========================================================================================================
'=											 4.2 Constant	variables
'========================================================================================================
Const	CookieSplit	=	1233
Const	BIZ_PGM_ID = "P5210mb1.asp"																			 'Biz	Logic	ASP
Const	C_SHEETMAXROWS		=	21																				'한	화면에 보여지는	최대갯수*1.5%>

'========================================================================================================
'=											 4.3 Common	variables
'========================================================================================================
<!-- #Include	file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=											 4.4 User-defind Variables
'========================================================================================================
<%'========================================================================================================%>
Dim	lsConcd
Dim	IsOpenPop

Dim	gSelframeFlg				 ' 현재	TAB의	위치를 나타내는	Flag
Dim	gCounts
Dim	isFirst		'첫화면이	열리는지 여부 
Dim	lgStrPrevKey1
Dim	lgStrPrevKey2
Dim	lgPageNo_A
Dim	lgPageNo_B
Dim	lgOldRow_A
Dim	lgOldRow_B


Dim	C_FACILITY_ACCNT_CD
Dim	C_FACILITY_ACCNT_NM
Dim	C_Facility_Lvl1
Dim	C_Facility_Lvl1_Pop
Dim	C_Facility_Lvl1Nm
Dim	C_Facility_Lvl2
Dim	C_Facility_Lvl2_Pop
Dim	C_Facility_Lvl2Nm
Dim  C_ChkF1


Dim	C_Seq
Dim	C_Zinsp_PartCd
Dim	C_Zinsp_PartNm
Dim	C_Insp_PartCd
Dim	C_Insp_PartNm
Dim	C_Insp_MethCd
Dim	C_Insp_MethNm
Dim	C_Insp_DecisionCd
Dim	C_Insp_DecisionNm
Dim	C_St_Go_GubunCd
Dim	C_St_Go_GubunNm
Dim  C_ChkF2

'==========================================	 1.2.2 Global	변수 선언	 =====================================
'	1. 변수	표준에 따름. prefix로	g를	사용함.
'	2.Array인	경우는 ()를	반드시 사용하여	일반 변수와	구별해 됨 
'=========================================================================================================
<%'========================================================================================================%>

Dim	iDBSYSDate
Dim	EndDate, StartDate

	'------	☆:	초기화면에 뿌려지는	마지막 날짜	------
	EndDate	=	"<%=GetSvrDate%>"
	'------	☆:	초기화면에 뿌려지는	시작 날짜	------
	StartDate	=	UNIDateAdd("m",	-1,	EndDate, Parent.gServerDateFormat)
	EndDate	=	UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
	StartDate	=	UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)

'========================================================================================================
'	Name : InitSpreadPosVariables()
'	Desc : Initialize	the	position
'========================================================================================================
Sub	initSpreadPosVariables(ByVal pvSpdNo)

	If  pvSpdNo = "A" or pvSpdNo= "" Then
		C_FACILITY_ACCNT_CD		=	1
		C_FACILITY_ACCNT_NM		=	2
		C_Facility_Lvl1			=	3
		C_Facility_Lvl1_Pop		=	4
		C_Facility_Lvl1Nm		=	5
		C_Facility_Lvl2			=	6
		C_Facility_Lvl2_Pop		=	7
		C_Facility_Lvl2Nm		=	8
		C_ChkF1				=9
	end if
	If pvSpdNo= "B" or pvSpdNo ="" then
		C_Seq					=	1
		C_Zinsp_PartCd		=	2
		C_Zinsp_PartNm		=	3
		C_Insp_PartCd		=	4
		C_Insp_PartNm		=	5
		C_Insp_MethCd		=	6
		C_Insp_MethNm		=	7
		C_Insp_DecisionCd	=	8
		C_Insp_DecisionNm	=	9
		C_St_Go_GubunCd		=	10
		C_St_Go_GubunNm		=	11
		C_ChkF2						=12

	End	If

End	Sub

'========================================================================================================
'	Name : InitVariables()
'	Desc : Initialize	value
'========================================================================================================
Sub	InitVariables()
	lgIntFlgMode			=	parent.OPMD_CMODE										'⊙: Indicates that	current	mode is	Create mode
	lgBlnFlgChgValue	=	False										'⊙: Indicates that	no value changed
	lgIntGrpCount			=	0										'⊙: Initializes Group View	Size
		lgStrPrevKey			=	""																			'⊙: initializes Previous	Key
		lgStrPrevKey1		=	""																			'⊙: initializes Previous	Key	Index
		lgStrPrevKey2		=	""																			'⊙: initializes Previous	Key	Index
		lgSortKey					=	1																				'⊙: initializes sort	direction
	lgOldRow_A = 0
	lgOldRow_B = 0
End	Sub

'========================================================================================================
'	Name : LoadInfTB19029()
'	Desc : Set System	Number format
'========================================================================================================
Sub	LoadInfTB19029()
	<!-- #Include	file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call	loadInfTB19029A("I", "H","NOCOOKIE","MA")	%>
End	Sub

'========================================================================================================
'	Name : CookiePage()
'	Description	:	Item Popup에서 Return되는	값 setting
'========================================================================================================
<%'========================================================================================================%>
Function CookiePage(ByVal	flgs)
End	Function

'========================================================================================================
'	Function Name	:	MakeKeyStream
'	Function Desc	:	This method	set	focus	to pos of	err
'========================================================================================================
Sub	MakeKeyStream(pRow)
	If lgCurrentSpd	=	"M"	Then
		 lgKeyStream = Frm1.CbohFacility_Accnt.Value & parent.gColSep																						'You Must	append one character(	parent.gColSep)
		 lgKeyStream = lgKeyStream & Frm1.txtItemGroupCd1.Value	&	parent.gColSep
		 lgKeyStream = lgKeyStream & Frm1.txtItemGroupCd2.Value	&	parent.gColSep
	Else
		frm1.vspdData1.Row = pRow
		frm1.vspdData1.Col = C_FACILITY_ACCNT_CD
		lgKeyStream	=	frm1.vspdData1.Text	&	parent.gColSep		 'You	Must append	one	character( parent.gColSep)
		frm1.vspdData1.Col = C_Facility_Lvl1
		lgKeyStream	=	frm1.vspdData1.Text	&	parent.gColSep		 'You	Must append	one	character( parent.gColSep)
		frm1.vspdData1.Col = C_Facility_Lvl2
		lgKeyStream	=	frm1.vspdData1.Text	&	parent.gColSep		 'You	Must append	one	character( parent.gColSep)
	End	If
End	Sub

'========================================================================================================
'	Name : InitComboBox()
'	Desc : Set ComboBox
'========================================================================================================
Sub	InitComboBox()
	Dim	iCodeArr
	Dim	iNameArr
	
	
	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z402' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Facility_Lvl1
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Facility_Lvl1Nm
	Call SetCombo2(frm1.txtItemGroupCd1	,lgF0	 ,lgF1	,Chr(11))



	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z403' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Facility_Lvl2
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Facility_Lvl2Nm
	Call SetCombo2(frm1.txtItemGroupCd2	,lgF0	 ,lgF1	,Chr(11))

	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z410' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_FACILITY_ACCNT_CD
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_FACILITY_ACCNT_NM

	Call SetCombo2(frm1.CbohFacility_Accnt ,lgF0	,lgF1	 ,Chr(11))


	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z425' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Zinsp_PartCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Zinsp_PartNm

	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z411' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Insp_PartCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Insp_PartNm

	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z412' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Insp_MethCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Insp_MethNm

	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z418' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Insp_DecisionCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Insp_DecisionNm

	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z419' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_St_Go_GubunCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_St_Go_GubunNm

End	Sub

'========================================================================================================
'	Name : InitData()
'	Desc : Reset Combox
'========================================================================================================
Sub	InitData()
	Dim	intRow
	Dim	intIndex

End	Sub

'========================================================================================================
'	Function Name	:	InitSpreadSheet
'	Function Desc	:	This method	initializes	spread sheet column	property
'========================================================================================================
Sub	InitSpreadSheet(ByVal	pvSpdNo)

	If pvSpdNo = ""	OR pvSpdNo = "A" Then

		Call initSpreadPosVariables("A")
		With frm1.vspdData1

					ggoSpread.Source = frm1.vspdData1
					ggoSpread.Spreadinit "V20030226",,parent.gAllowDragDropSpread

					.ReDraw	=	false

					.MaxCols = C_ChkF1 + 1																								<%'☜: 최대	Columns의	항상 1개 증가시킴 %>
					.Col = .MaxCols															<%'공통콘트롤	사용 Hidden	Column%>
					.MaxRows = 0


				Call AppendNumberPlace("6","2","0")
				Call GetSpreadColumnPos("A")
				
				ggoSpread.SSSetCombo 	C_FACILITY_ACCNT_CD,	"설비유형",	 10, 0,	False
				ggoSpread.SSSetCombo 	C_FACILITY_ACCNT_NM,	"설비유형",	 20, 0,	False

				ggoSpread.SSSetCombo		 C_Facility_Lvl1,					"대분류",	 10, 0,	False
				ggoSpread.SSSetButton		C_Facility_Lvl1_Pop
				ggoSpread.SSSetCombo		 C_Facility_Lvl1Nm,				 "대분류",	20,	0, False
				ggoSpread.SSSetCombo		 C_Facility_Lvl2,					"중분류",	10,	0, False
				ggoSpread.SSSetButton		C_Facility_Lvl2_Pop
				ggoSpread.SSSetCombo		 C_Facility_Lvl2Nm,					"중분류",	 20, 0,	False
				ggoSpread.SSSetEdit		C_ChkF1,					"chkFlag",2
				

				Call ggoSpread.SSSetColHidden(C_FACILITY_ACCNT_CD, C_FACILITY_ACCNT_CD,	True)
				Call ggoSpread.SSSetColHidden(C_Facility_Lvl1, C_Facility_Lvl1_Pop,	True)
				Call ggoSpread.SSSetColHidden(C_Facility_Lvl2, C_Facility_Lvl2_Pop,	True)
				Call ggoSpread.SSSetColHidden(C_ChkF1,	.MaxCols,	True)
			

				Call SetSpreadLock
				.ReDraw	=	true



		End	With

	End	if

		If pvSpdNo = ""	OR pvSpdNo = "B" Then
			Call initSpreadPosVariables("B")
			With frm1.vspdData2

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.Spreadinit "V20030226",,parent.gAllowDragDropSpread

				.ReDraw	=	false
				.MaxCols = C_ChkF2 + 1												<%'☜: 최대	Columns의	항상 1개 증가시킴 %>
				.Col = .MaxCols															<%'공통콘트롤	사용 Hidden	Column%>
	
				.MaxRows = 0

				Call AppendNumberPlace("6","2","0")
				Call GetSpreadColumnPos("B")

				ggoSpread.SSSetFloat	C_Seq,				"순서",    8, "7", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
				ggoSpread.SSSetCombo 	C_Zinsp_PartCd,		"부위",	 10, 0,	False
				ggoSpread.SSSetCombo	C_Zinsp_PartNm,		"부위",	 10, 0,	False
				ggoSpread.SSSetCombo	C_Insp_PartCd,		"점검항목",	 10, 0,	False
				ggoSpread.SSSetCombo	C_Insp_PartNm,		"점검항목",	 10, 0,	False
				ggoSpread.SSSetCombo	C_Insp_MethCd,		"점검방법",	 10, 0,	False
				ggoSpread.SSSetCombo	C_Insp_MethNm,		"점검방법",	 10, 0,	False
				ggoSpread.SSSetCombo	C_Insp_DecisionCd,	"판정기준",	 10, 0,	False
				ggoSpread.SSSetCombo	C_Insp_DecisionNm,	"판정기준",	 10, 0,	False
				ggoSpread.SSSetCombo	C_St_Go_GubunCd,	"운/휴구분",	14,	0, False
				ggoSpread.SSSetCombo	C_St_Go_GubunNm,	"운/휴구분",	14,	0, False

				ggoSpread.SSSetEdit	C_ChkF2,					"chkFlag",2
				
				Call ggoSpread.SSSetColHidden(C_Zinsp_PartCd,	C_Zinsp_PartCd,	True)
				Call ggoSpread.SSSetColHidden(C_Insp_PartCd, C_Insp_PartCd,	True)
				Call ggoSpread.SSSetColHidden(C_Insp_MethCd, C_Insp_MethCd,	True)
				Call ggoSpread.SSSetColHidden(C_Insp_DecisionCd, C_Insp_DecisionCd,	True)
				Call ggoSpread.SSSetColHidden(C_St_Go_GubunCd, C_St_Go_GubunCd,	True)
				Call ggoSpread.SSSetColHidden(C_ChkF2,	.MaxCols,	True)

				.ReDraw	=	true

			Call SetSpreadLock1

			End	With
		End	if


End	Sub

'======================================================================================================
'	Function Name	:	SetSpreadLock
'	Function Desc	:	This method	set	color	and	protect	in spread	sheet	celles
'======================================================================================================
Sub	SetSpreadLock()		
		With frm1
			.vspdData1.ReDraw = False
			ggoSpread.SpreadLock		C_FACILITY_ACCNT_CD,				-1,	C_Facility_Lvl2Nm				,-1
			ggoSpread.SSSetProtected frm1.vspdData1.MaxCols,	-1,-1
			.vspdData1.ReDraw = True
		End	With
End	Sub

Sub	SetSpreadLock1()

	With frm1.vspdData2
				ggoSpread.Source = frm1.vspdData2

				.ReDraw	=	False
				ggoSpread.SpreadLock 		C_Seq			,	-1,	C_Seq
				ggoSpread.SSSetRequired		C_Zinsp_PartNm		,	-1,	C_Zinsp_PartNm
				ggoSpread.SSSetRequired		C_Insp_PartNm		,	-1,	C_Insp_PartNm
				ggoSpread.SSSetRequired		C_Insp_MethNm		,	-1,	C_Insp_MethNm
				ggoSpread.SSSetRequired		C_Insp_DecisionNm	,	-1,	C_Insp_DecisionNm
				ggoSpread.SSSetRequired		C_St_Go_GubunNm		,	-1,	C_St_Go_GubunNm
				ggoSpread.SSSetProtected				.MaxCols,-1,-1
				.ReDraw	=	True
		End	With
End	Sub

'======================================================================================================
'	Function Name	:	SetSpreadColor
'	Function Desc	:	This method	set	color	and	protect	in spread	sheet	celles
'======================================================================================================
Sub	SetSpreadColor(ByVal pvStartRow,ByVal	pvEndRow)

	With frm1.vspdData1
		ggoSpread.Source = frm1.vspdData1

		.ReDraw	=	False
		ggoSpread.SSSetRequired		 C_FACILITY_ACCNT_NM		,	pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired		 C_Facility_Lvl1Nm			,	pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected	 C_Facility_Lvl1			,	pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected	 C_Facility_Lvl2			,	pvStartRow,	pvEndRow
		ggoSpread.SSSetProtected	.MaxCols,-1,-1
		.ReDraw	=	True
	End	With
End	Sub

Sub	SetSpreadColor1(ByVal	pvStartRow,ByVal pvEndRow)

	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2

		.ReDraw	=	False
		ggoSpread.SSSetRequired		C_Seq		,	pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired		C_Zinsp_PartNm		,	pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired		C_Insp_PartNm		,	pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired		C_Insp_MethNm		,	pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired		C_Insp_DecisionNm		,	pvStartRow,	pvEndRow
		ggoSpread.SSSetRequired		C_St_Go_GubunNm		,	pvStartRow,	pvEndRow

		ggoSpread.SSSetProtected	.MaxCols	,	pvStartRow,	pvEndRow
		.ReDraw	=	True
	End	With
End	Sub

'========================================================================================
'	Function Name	:	GetSpreadColumnPos
'	Description		:
'========================================================================================
Sub	GetSpreadColumnPos(ByVal pvSpdNo)
		Dim	iCurColumnPos

		Select Case	UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_FACILITY_ACCNT_CD		=	iCurColumnPos(1)
			C_FACILITY_ACCNT_NM		=	iCurColumnPos(2)
			C_Facility_Lvl1			=	iCurColumnPos(3)
			C_Facility_Lvl1_Pop		=	iCurColumnPos(4)
			C_Facility_Lvl1Nm		=	iCurColumnPos(5)
			C_Facility_Lvl2			=	iCurColumnPos(6)
			C_Facility_Lvl2_Pop		=	iCurColumnPos(7)
			C_Facility_Lvl2Nm		=	iCurColumnPos(8)
			C_ChkF1					=iCurColumnPos(9)
		Case "B"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Seq					=	iCurColumnPos(1	)
			C_Zinsp_PartCd			=	iCurColumnPos(2	)
			C_Zinsp_PartNm			=	iCurColumnPos(3	)
			C_Insp_PartCd			=	iCurColumnPos(4	)
			C_Insp_PartNm			=	iCurColumnPos(5	)
			C_Insp_MethCd			=	iCurColumnPos(6	)
			C_Insp_MethNm			=	iCurColumnPos(7	)
			C_Insp_DecisionCd		=	iCurColumnPos(8	)
			C_Insp_DecisionNm		=	iCurColumnPos(9	)
			C_St_Go_GubunCd			=	iCurColumnPos(10)
			C_St_Go_GubunNm			=	iCurColumnPos(11)
			C_ChkF2							= iCurColumnPos(12)

		End	Select
End	Sub

'========================================================================================================
'	Name : Form_Load
'	Desc : developer describe	this line	Called by	Window_OnLoad()	evnt
'========================================================================================================
Sub	Form_Load()

	Err.Clear																																				'☜: Clear err status
	Call LoadInfTB19029																															'☜: Load	table	,	B_numeric_format

	Call	ggoOper.FormatField(Document,	"1", ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gDateFormat,	parent.gComNum1000,	parent.gComNumDec)
	Call	ggoOper.FormatField(Document,	"2", ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gDateFormat,	parent.gComNum1000,	parent.gComNumDec)
	Call	ggoOper.LockField(Document,	"N")											'⊙: Lock	Field

	Call InitSpreadSheet("")																														'Setup the Spread	sheet
	Call InitVariables																															'Initializes local global	variables
	Call InitComboBox
	Call SetToolbar("1100111100011111")														'버튼	툴바 제어 
	gCounts	=	0
	isFirst	=	true
'	Call CookiePage	(0)																															'☜: Check Cookie
End	Sub

'========================================================================================================
'	Name : Form_QueryUnload
'	Desc : developer describe	this line	Called by	Window_OnUnLoad()	evnt
'========================================================================================================
Sub	Form_QueryUnload(Cancel, UnloadMode)
End	Sub

'========================================================================================================
'	Name : FncQuery
'	Desc : developer describe	this line	Called by	MainQuery	in Common.vbs
'========================================================================================================
Function FncQuery()

	Dim	IntRetCD
	Dim	ChgOK

	FncQuery = False															 '☜:	Processing is	NG
	Err.Clear																																		 '☜:	Clear	err	status

	ChgOK	=	false


	ggoSpread.Source = Frm1.vspdData2
	If	ggoSpread.SSCheckChange	=	True Then
	ChgOK	=	True
	End	If

	 ggoSpread.Source	=	Frm1.vspdData1
	If	ggoSpread.SSCheckChange	=	True Then
	ChgOK	=	True
	End	If

	If	ChgOK	Then
		IntRetCD =	DisplayMsgBox("900013",	 parent.VB_YES_NO,"x","x")		'☜: Data	is changed.	 Do	you	want to	display	it?

		If IntRetCD	=	vbNo Then
			Exit Function
		End	If
	End	If
	Call	ggoOper.ClearField(Document, "2")

	If Not chkField(Document,	"1") Then													 '☜:	This function	check	required field
		 Exit	Function
	End	If


	Call InitVariables																													 '⊙:	Initializes	local	global variables
	lgCurrentSpd = "M"
	Call MakeKeyStream("X")

		gCounts	=	0
		isFirst	=	true

		lgCurrentSpd = "M"	'	Master

	Call	DisableToolBar(	parent.TBC_QUERY)
	If DbQuery = False Then
		Call DbDtlQuery
		Call	RestoreToolBar()
				Exit Function
		End	If

		FncQuery = True																															 '☜:	Processing is	OK

End	Function

'========================================================================================================
'	Name : FncDelete
'	Desc : developer describe	this line	Called by	MainDelete in	Common.vbs
'========================================================================================================
Function FncDelete()
	Dim	intRetCD

	FncDelete	=	False																														 '☜:	Processing is	NG
	Err.Clear																																		 '☜:	Clear	err	status

	FncDelete	=	True																														 '☜:	Processing is	OK
End	Function

'========================================================================================================
'	Name : FncSave
'	Desc : developer describe	this line	Called by	MainSave in	Common.vbs
'========================================================================================================
Function FncSave()
    Dim	IntRetCD
    dim	lRow,iTemp,flagTxt
    DIM	strCD, strNm,iRow
    Dim strTmp1, strTmp2
    
    FncSave	=	False																															 '☜:	Processing is	NG
    Err.Clear
    
    frm1.ChgSave1.value	=	"F"
    frm1.ChgSave2.value	=	"F"

    ggoSpread.Source = frm1.vspdData1
    If	ggoSpread.SSCheckChange	=	True Then
        frm1.ChgSave1.value	=	"T"
    End	If

    ggoSpread.Source = Frm1.vspdData2
    If	ggoSpread.SSCheckChange	=	True Then
     frm1.ChgSave2.value	=	"T"
    End	If

	If frm1.ChgSave1.value = "F" and frm1.ChgSave2.value="F"	Then
		IntRetCD =	DisplayMsgBox("900001","x","x","x")														'☜:There	is no	changed	data.
		Exit Function
	End	If

	ggoSpread.Source = frm1.vspdData1
	If Not	ggoSpread.SSDefaultCheck Then																				'☜: Check contents	area
			 Exit	Function
	End	If

	ggoSpread.Source = frm1.vspdData2
	If Not	ggoSpread.SSDefaultCheck Then																				'☜: Check contents	area
			 Exit	Function
	End	If

	ggoSpread.Source = frm1.vspdData1
	With Frm1
	  For iRow =1 To .vspdData1.maxRows
		.vspdData1.Row = iRow
		.vspdData1.Col = 0
		flagTxt	=	.vspdData1.Text
		.vspdData1.Col = C_ChkF1 :strTmp1=.vspdDAta1.Text 
		.vspdData2.Row = .vspdData2.ActiveRow 
		.vspdDAta2.Col = C_ChkF2 : strTmp2 = .vspdData2.Text
		
		If flagTxt =	ggoSpread.InsertFlag  and strTmp1 <> strTmp2 Then
			frm1.vspdData2.Maxrows=0:frm1.vspdData2.focus	

			 If frm1.vspdData2.MaxRows	=<	1	then
				IntRetCD =	DisplayMsgBox("Y50060","x","x","x")

				ggoSpread.Source = .vspdData2				
				ggoSpread.InsertRow	0,	1
				SetSpreadColor1	0,	1
				.vspdData2.Row =1
				.vspdData2.Col=	C_Seq
				.vspdData2.Text	=	1
				.vspdData2.Focus		
				Exit Function
			End	If	
	  End If
	Next
  End With
  
	lgCurrentSpd = "M"
	Call	DisableToolBar(	parent.TBC_SAVE)
	
	If DbSave	=	False	Then
		Call	RestoreToolBar()
		Exit Function
	End	If

	FncSave	=	True																															'☜: Processing	is OK

End	Function

'========================================================================================================
'	Function Name	:	FncCopy
'	Function Desc	:	This function	is related to	Copy Button	of Main	ToolBar
'========================================================================================================
Function FncCopy()

    If Trim(lgActiveSpd) = ""	Then
    	 lgActiveSpd = "M"
    End	If
    
    Select Case	UCase(Trim(lgActiveSpd))
    	Case	"M"
    
    
    	Case	"S1"
    
    		If Frm1.vspdData2.MaxRows	<	1	Then
    				Exit Function
    		End	If
    
    		With Frm1.vspdData2
    
    			If .ActiveRow	>	0	Then
    				.ReDraw	=	False
    
    				ggoSpread.Source = frm1.vspdData2
    				ggoSpread.CopyRow
						SetSpreadColor1	.ActiveRow,	.ActiveRow

						.Col	=	C_Seq
            		.Row	=	.ActiveRow
            		.Text	=	""
    
    				.ReDraw	=	True
    				.focus
    			End	If
    		End	With
    End	Select
    
    Set	gActiveElement = document.ActiveElement

End	Function

'========================================================================================================
'	Function Name	:	FncCancel
'	Function Desc	:	This function	is related to	Cancel Button	of Main	ToolBar
'========================================================================================================
Function FncCancel()
	if lgCurrentSpd	=	"M"	then
		ggoSpread.Source = frm1.vspdData1
		If Not	ggoSpread.SSCheckChange Then																				'☜: Check contents	area
				 Exit	Function
		End	If

        Frm1.vspdData1.ReDraw	=	False
		ggoSpread.Source = Frm1.vspdData1
		ggoSpread.EditUndo
        Frm1.vspdData1.ReDraw	=	True
 		ggoSpread.Source = frm1.vspdData2
 		ggoSpread.ClearSpreadData

        If frm1.vspdData1.Row >1 Then Call DbDtlQuery() End If
	elseif lgCurrentSpd	=	"S1" then
		ggoSpread.Source = frm1.vspdData2
		If Not	ggoSpread.SSCheckChange Then																				'☜: Check contents	area
				 Exit	Function
		End	If
		ggoSpread.Source = Frm1.vspdData2
		ggoSpread.EditUndo
	end	if

End	Function

'========================================================================================================
'	Function Name	:	FncInsertRow
'	Function Desc	:	This function	is related to	InsertRow	Button of	Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal	pvRowCnt)

	Dim	imRow
	Dim	iRow
	Dim	IntRetCD
	Dim	iTemp
	
	On Error Resume	Next																													'☜: If	process	fails
	Err.Clear																																			'☜: Clear error status

		FncInsertRow = False																												 '☜:	Processing is	NG

		If IsNumeric(Trim(pvRowCnt)) Then
				imRow	=	CInt(pvRowCnt)
		Else
				imRow	=	AskSpdSheetAddRowCount()
				If imRow = ""	Then
						Exit Function
				End	If
		End	If

		If Trim(lgActiveSpd) = ""	Then
			 lgActiveSpd = "M"
		End	If

		Select Case	UCase(Trim(lgActiveSpd))
			Case	"M"
				ggoSpread.Source = Frm1.vspdData1
				If	ggoSpread.SSCheckChange	=	True Then
					Call	DisplayMsgBox("990027","X","X","X")																 '☆:
					Exit Function
				End	If
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.ClearSpreadData
	
				ggoSpread.Source = frm1.vspdData1
	
				With Frm1
					.vspdData1.Row = Row
					.vspdData1.Col = 0
					flagTxt	=	.vspdData1.Text
	
					.vspdData1.ReDraw	=	False
					.vspdData1.Focus
					ggoSpread.Source = .vspdData1
					ggoSpread.InsertRow	.vspdData1.ActiveRow,	1
					SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow	+	1	-	1
					Call SetToolbar("1100111100011111")														'버튼	툴바 제어 
					.vspdData1.ReDraw	=	True
				End	With

					With Frm1
						.vspdData2.ReDraw	=	False
						.vspdData2.Focus
						ggoSpread.Source = .vspdData2
						ggoSpread.InsertRow	.vspdData2.ActiveRow,	imRow
						SetSpreadColor1	.vspdData2.ActiveRow,	.vspdData2.ActiveRow + imRow - 1
						iTemp	=	0
						For	iRow =	.vspdData2.ActiveRow to	.vspdData2.ActiveRow + imRow - 1
							iTemp	=	iTemp	+	1
							.vspdData2.Row = iRow
							.vspdData2.Col=	C_Seq
							.vspdData2.Text	=	.vspdData2.Maxrows + iTemp - imRow
						Next
	
						.vspdData2.ReDraw	=	True
					End	With

			Case	"S1"

				if frm1.vspdData1.MaxRows	>	0	then
					With Frm1
						.vspdData2.ReDraw	=	False
						.vspdData2.Focus
						ggoSpread.Source = .vspdData2
						ggoSpread.InsertRow	.vspdData2.ActiveRow,	imRow
						SetSpreadColor1	.vspdData2.ActiveRow,	.vspdData2.ActiveRow + imRow - 1
						iTemp	=	0
						For	iRow =	.vspdData2.ActiveRow to	.vspdData2.ActiveRow + imRow - 1
							iTemp	=	iTemp	+	1
							.vspdData2.Row = iRow
							.vspdData2.Col=	C_Seq
							.vspdData2.Text	=	.vspdData2.Maxrows + iTemp - imRow
						Next
	
						.vspdData2.ReDraw	=	True
					End	With
				else
					Call	DisplayMsgBox("900025","X","X","X")																 '☆:
					Exit Function
				End	if
			End	Select

			If Err.number	=	0	Then
				 FncInsertRow	=	True																													'☜: Processing	is OK
			End	If

			Set	gActiveElement = document.ActiveElement

End	Function

'========================================================================================================
'	Function Name	:	FncDeleteRow
'	Function Desc	:	This function	is related to	DeleteRow	Button of	Main ToolBar
'========================================================================================================
Function FncDeleteRow()
		Dim	lDelRows

		if	lgCurrentSpd = "M" then
				If Frm1.vspdData1.MaxRows	<	1	then
					 Exit	function
			End	if
				With Frm1.vspdData1
					.focus
					 ggoSpread.Source	=	frm1.vspdData1
					lDelRows =	ggoSpread.DeleteRow
				End	With
		ELSEif lgCurrentSpd	=	"S1" then
				If Frm1.vspdData2.MaxRows	<	1	then
					 Exit	function
			End	if
				With Frm1.vspdData2
					.focus
					 ggoSpread.Source	=	frm1.vspdData2
					lDelRows =	ggoSpread.DeleteRow
				End	With

		END	IF
		Set	gActiveElement = document.ActiveElement
End	Function

'========================================================================================================
'	Function Name	:	FncPrint
'	Function Desc	:	This function	is related to	Print	Button of	Main ToolBar
'========================================================================================================
Function FncPrint()
		Call parent.FncPrint()
End	Function

'========================================================================================================
'	Function Name	:	FncExcel
'	Function Desc	:	This function	is related to	Excel
'========================================================================================================
Function FncExcel()
		Call parent.FncExport( parent.C_MULTI)																				 '☜:	화면 유형 
End	Function

'========================================================================================================
'	Function Name	:	FncFind
'	Function Desc	:
'========================================================================================================
Function FncFind()
		Call parent.FncFind( parent.C_MULTI, False)																		 '☜:화면	유형,	Tab	유무 
End	Function

'========================================================================================
'	Function Name	:	FncSplitColumn
'	Function Desc	:
'========================================================================================
Sub	FncSplitColumn()

		If UCase(Trim(TypeName(gActiveSpdSheet)))	=	"EMPTY"	Then
			 Exit	Sub
		End	If

		ggoSpread.Source = gActiveSpdSheet
		ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
	Set	gActiveElement = document.activeElement

End	Sub

'========================================================================================
'	Function Name	:	PopSaveSpreadColumnInf
'	Description		:
'========================================================================================
Sub	PopSaveSpreadColumnInf()
		ggoSpread.Source = gActiveSpdSheet
		Call ggoSpread.SaveSpreadColumnInf()
End	Sub
'========================================================================================
'	Function Name	:	PopRestoreSpreadColumnInf
'	Description		:
'========================================================================================
Sub	PopRestoreSpreadColumnInf()

	ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)

    Call InitComboBox
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.ReOrderingSpreadData
End	Sub
'========================================================================================================
'	Function Name	:	FncExit
'	Function Desc	:
'========================================================================================================
Function FncExit()
		Dim	IntRetCD

	FncExit	=	False

		 ggoSpread.Source	=	frm1.vspdData1
		If	ggoSpread.SSCheckChange	=	True Then
		IntRetCD =	DisplayMsgBox("900016",	 parent.VB_YES_NO,"x","x")			'⊙: Data	is changed.	 Do	you	want to	exit?
		If IntRetCD	=	vbNo Then
			Exit Function
		End	If
		End	If
		FncExit	=	True
End	Function

'========================================================================================================
'	Name : DbQuery
'	Desc : This	function is	called by	FncQuery
'========================================================================================================
Function DbQuery()

		DbQuery	=	False

		Err.Clear																																				 '☜:	Clear	err	status

	If LayerShowHide(1)	=	False	Then
		Exit Function
	End	If

	Dim	strVal



		With Frm1
		strVal = BIZ_PGM_ID	&	"?txtMode="						&	parent.UID_M0001
				strVal = strVal			&	"&lgCurrentSpd="			&	lgCurrentSpd											'☜: Next	key	tag
				strVal = strVal			&	"&txtKeyStream="			&	lgKeyStream												'☜: Query Key
			strVal = strVal			&	"&txtFacility_Accnt="	&	Frm1.CbohFacility_Accnt.value											'☜: Query Key
			strVal = strVal			&	"&txtItemGroupCd1="	&	Frm1.txtItemGroupCd1.value			'☜: Query Key
			strVal = strVal			&	"&txtItemGroupCd2="	&	Frm1.txtItemGroupCd2.value										 '☜:	Query	Key
				strVal = strVal			&	"&txtMaxRows="				&	.vspdData1.MaxRows
				strVal = strVal			&	"&lgStrPrevKey="		&	lgStrPrevKey								 '☜:	Next key tag
		strVal = strVal			&	"&lgPageNo_A="		&	lgPageNo_A													'☜: Next	key	tag
		strVal = strVal			&	"&txtType="			&	"A"													 '☜:	Next key tag
		End	With

	Call RunMyBizASP(MyBizASP, strVal)																							 '☜:	Run	Biz	Logic
		DbQuery	=	True
End	Function

'========================================================================================================
'	Name : DbDtlQuery
'	Desc : This	function is	called by	FncQuery
'========================================================================================================

Function DbDtlQuery()

		DbDtlQuery = False

		Err.Clear																																				 '☜:	Clear	err	status

	If LayerShowHide(1)	=	False	Then
		Exit Function
	End	If

	Dim	strVal

		With Frm1
			strVal = BIZ_PGM_ID	&	"?txtMode="						&	parent.UID_M0001
			strVal = strVal			&	"&lgCurrentSpd="			&	lgCurrentSpd											'☜: Next	key	tag
			strVal = strVal			&	"&txtKeyStream="			&	lgKeyStream												'☜: Query Key
			strVal = strVal			&	"&txtFacility_Accnt="	&	Frm1.hFacility_Accnt.value										 '☜:	Query	Key
			strVal = strVal			&	"&txtItemGroupCd1="		&	Frm1.hItemGroupCd1.value			'☜: Query Key
			strVal = strVal			&	"&txtItemGroupCd2="		&	Frm1.hItemGroupCd2.value										 '☜:	Query	Key
			strVal = strVal			&	"&txtMaxRows="				&	.vspdData2.MaxRows
			strVal = strVal			&	"&lgStrPrevKey=" 		&	lgStrPrevKey1									'☜: Next	key	tag
			strVal = strVal			&	"&lgPageNo_B="		&	lgPageNo_B													'☜: Next	key	tag
			strVal = strVal			&	"&txtType="			&	"B"													 '☜:	Next key tag
			strVal = strVal			&	"&hChkFlag="			&	frm1.hChkFlag.value													 '☜:	Next key tag
		End	With

	Call RunMyBizASP(MyBizASP, strVal)																							 '☜:	Run	Biz	Logic

		DbDtlQuery = True
End	Function


Function DbDtlQueryOk()														'☆: 조회	성공후 실행로직 
	Dim	i
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that	current	mode is	Update mode
	lgBlnFlgChgValue = False
End	Function


'========================================================================================================
'	Name : DbSave
'	Desc : This	function is	data query and display
'========================================================================================================
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

	if frm1.ChgSave1.value = "T"	then
		ggoSpread.Source	=	frm1.vspdData1
		With Frm1
			For lRow	=	1	To .vspdData1.MaxRows
				 .vspdData1.Row	=	lRow
				 .vspdData1.Col	=	0
				 Select	Case .vspdData1.Text
					 Case	 ggoSpread.DeleteFlag																			 '☜:	Delete

																	strDel = strDel	&	"D"	&	parent.gColSep
																	strDel = strDel	&	lRow & parent.gColSep
																	strDel = strDel	&	lgCurrentSpd & parent.gColSep
						.vspdData1.Col = C_FACILITY_ACCNT_CD	:	strDel = strDel	&	Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Facility_Lvl1		:	strDel = strDel	&	Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Facility_Lvl2		:	strDel = strDel	&	Trim(.vspdData1.Text)	&	parent.gRowSep
						lGrpCnt	=	lGrpCnt	+	1
				 End Select
			Next
		 .txtMode.value				 =	parent.UID_M0002
		 .txtMaxRows.value		 = lGrpCnt-1
		 .txtSpread.value			 = strDel	&	strVal
		End	With
	end	if



	if frm1.ChgSave2.value = "T"	then
        
		ggoSpread.Source = frm1.vspdData2
		With Frm1
            .vspdData1.Row	=	.vspdData1.ActiveRow
			 For lRow	=	1	To .vspdData2.MaxRows
				 .vspdData2.Row	=	lRow
				 .vspdData2.Col	=	0
				 Select	Case .vspdData2.Text
					 Case	 ggoSpread.InsertFlag																			 '☜:	Create
																	strVal = strVal	&	"C"	&	parent.gColSep
																	strVal = strVal	&	lRow & parent.gColSep
																	strVal = strVal	&	"S" & parent.gColSep
						.vspdData1.Col = C_FACILITY_ACCNT_CD	:	strVal = strVal	&	Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Facility_Lvl1		:	strVal = strVal	&	Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Facility_Lvl2		:	strVal = strVal	&	Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData2.Col = C_Seq					:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Zinsp_PartCd			:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Insp_PartCd			:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Insp_MethCd			:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Insp_DecisionCd		:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_St_Go_GubunCd		:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gRowSep
						lGrpCnt	=	lGrpCnt	+	1
					 Case	 ggoSpread.UpdateFlag																			 '☜:	Update
																	strVal = strVal	&	"U"	&	parent.gColSep
																	strVal = strVal	&	lRow & parent.gColSep
																	strVal = strVal	&	"S" & parent.gColSep
						.vspdData1.Col = C_FACILITY_ACCNT_CD	:	strVal = strVal	&	Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Facility_Lvl1		:	strVal = strVal	&	Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Facility_Lvl2		:	strVal = strVal	&	Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData2.Col = C_Seq					:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Zinsp_PartCd			:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Insp_PartCd			:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Insp_MethCd			:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Insp_DecisionCd		:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_St_Go_GubunCd		:	strVal = strVal	&	Trim(.vspdData2.Text)	&	parent.gRowSep
						lGrpCnt	=	lGrpCnt	+	1
					 Case	 ggoSpread.DeleteFlag																			 '☜:	Delete
																	strDel = strDel	&	"D"	&	parent.gColSep
																	strDel = strDel	&	lRow & parent.gColSep
																	strDel = strDel	&	"S" & parent.gColSep
						.vspdData1.Col = C_FACILITY_ACCNT_CD	:	strDel = strDel	&	Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Facility_Lvl1		:	strDel = strDel	&	Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Facility_Lvl2		:	strDel = strDel	&	Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData2.Col = C_Seq					:	strDel = strDel	&	Trim(.vspdData2.Text)	&	parent.gRowSep
						lGrpCnt	=	lGrpCnt	+	1
				 End Select
			Next
			.txtMode.value				 =	parent.UID_M0002
			.txtMaxRows.value		 = lGrpCnt-1
			.txtSpread.value			 = strDel	&	strVal
		End	With
	end	if

	Call ExecMyBizASP(frm1,	BIZ_PGM_ID)

	DbSave = True

End	Function

'========================================================================================================
'	Name : DbDelete
'	Desc : This	function is	called by	FncDelete
'========================================================================================================
Function DbDelete()
		Dim	IntRetCd

		FncDelete	=	False																											 '⊙:	Processing is	NG

		If lgIntFlgMode	<>	parent.OPMD_UMODE	Then																			'Check if	there	is retrived	data
				Call	DisplayMsgBox("900002","X","X","X")																 '☆:
				Exit Function
		End	If

		IntRetCD =	DisplayMsgBox("900003",	 parent.VB_YES_NO,"X","X")								'⊙: "Will you destory previous	data"
	If IntRetCD	=	vbNo Then													'------	Delete function	call area	------
		Exit Function
	End	If


		Call	DisableToolBar(	parent.TBC_DELETE)
	If DbDelete	=	False	Then
		Call	RestoreToolBar()
				Exit Function
		End	If

		FncDelete	=	True																												'⊙: Processing	is OK
End	Function

'========================================================================================================
'	Function Name	:	DbQueryOk
'	Function Desc	:	Called by	MB Area	when query operation is	successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode =	parent.OPMD_UMODE
	lgOldRow_A = 0
	Call	ggoOper.LockField(Document,	"Q")										'⊙: Lock	field
	Call InitData()
	Call SetToolbar("1100111100011111")

' 	if lgStrPrevKey1 <>	"" and isFirst = false then
' 		exit function
' 	end	if
' 	if lgStrPrevKey1 <>	"" or	isFirst	=	true then
' 		isFirst	=	false		'	첫화면이 열리고나서	오른쪽 그리드	세팅하기 위해 
		Call DisableToolBar(parent.TBC_QUERY)
		call vspdData1_click(1,frm1.vspdData1.activerow)
' 	end	if
	frm1.vspdData1.focus
End	Function

'========================================================================================================
'	Function Name	:	DbSaveOk
'	Function Desc	:	Called by	MB Area	when save	operation	is successful
'========================================================================================================
Function DbSaveOk()

	Call	ggoOper.ClearField(Document, "2")

	Call InitVariables															'⊙: Initializes local global	variables
	lgCurrentSpd = "M"
	Call MakeKeyStream("X")

	Call	DisableToolBar(	parent.TBC_QUERY)
	If DbQuery = False Then
		Call	RestoreToolBar()
		Exit Function
	End	If

End	Function

'========================================================================================================
'	Function Name	:	DbDeleteOk
'	Function Desc	:	Called by	MB Area	when delete	operation	is successful
'========================================================================================================
Function DbDeleteOk()

End	Function


'------------------------------------------	 OpenGroup()	---------------------------------------
'	Name : OpenGroup()
'	Description	:	OpenGroup	Popup에서	Return되는 값	setting
'---------------------------------------------------------------------------------------------------------
Function OpenGroup(Byval strCode,	iWhere)
	Dim	arrRet
	Dim	arrParam(5), arrField(6),	arrHeader(6)
	Dim	strTmp,	strvspdData
	If IsOpenPop = True	Then Exit	Function

	IsOpenPop	=	True

	if iWhere	=	1	or iWhere	=	3	then
		strTmp = ""
	elseif iWhere	=	2	then
		if Trim(UCase(frm1.txtItemGroupCd1.value)) = ""	then
			IsOpenPop	=	False
			Call DisplayMsgBox("127415", "X",	"X", "X")
			Exit Function
		else
			strTmp = " AND UPPER_ITEM_GROUP_CD = " & FilterVar(Trim(UCase(frm1.txtItemGroupCd1.value)),"''","S")
		end	if
	elseif iWhere	=	4	then
	 	ggoSpread.Source = frm1.vspdData1
		 	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_Facility_Lvl1
		strvspdData	=	frm1.vspdData1.text


		if Trim(UCase(strvspdData))	=	"" then
			IsOpenPop	=	False
			Call DisplayMsgBox("127415", "X",	"X", "X")
			Exit Function
		else
			strTmp = " AND UPPER_ITEM_GROUP_CD = " & FilterVar(Trim(UCase(strvspdData)),"''","S")
		end	if
	end	if

	arrParam(0)	=	"중분류팝업"
	arrParam(1)	=	"B_ITEM_GROUP"
	arrParam(2)	=	Trim(UCase(strCode))
	arrParam(3)	=	""
	arrParam(4)	=	"DEL_FLG = 'N'"	&	strTmp
	arrParam(5)	=	"중분류"

	arrField(0)	=	"ITEM_GROUP_CD"
	arrField(1)	=	"ITEM_GROUP_NM"

	arrHeader(0) = "중분류"
	arrHeader(1) = "중분류명"

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp",	Array(arrParam,	arrField,	arrHeader),	_
	"dialogWidth=420px;	dialogHeight=450px;	center:	Yes; help: No; resizable:	No;	status:	No;")

	IsOpenPop	=	False

	If arrRet(0) <>	"" Then
		Call SetItemGroup(arrRet,	iWhere)
	End	If
	Call SetFocusToDocument("M")
End	Function



'------------------------------------------	 SetItemGroup()	 -----------------------------------------
'	Name : SetItemGroup()
'	Description	:	ItemGroup	Popup에서	Return되는 값	setting
'---------------------------------------------------------------------------------------------------------
Function SetItemGroup(byval	arrRet,	iWhere)
	Select Case	iWhere
		Case 1
			frm1.txtItemGroupCd1.Value		=	arrRet(0)
			frm1.txtItemGroupNm1.Value		=	arrRet(1)
		Case 2
			frm1.txtItemGroupCd2.Value		=	arrRet(0)
			frm1.txtItemGroupNm2.Value		=	arrRet(1)
		Case 3
		 	ggoSpread.Source = frm1.vspdData1
		 	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
			frm1.vspdData1.Col = C_Facility_Lvl1
			frm1.vspdData1.text	=	arrRet(0)
			frm1.vspdData1.Col = C_Facility_Lvl1Nm
			frm1.vspdData1.text	=	arrRet(1)
		 	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
			frm1.vspdData1.Col = C_Facility_Lvl2
			frm1.vspdData1.text	=	""
			frm1.vspdData1.Col = C_Facility_Lvl2Nm
			frm1.vspdData1.text	=	""
		Case 4
		 	ggoSpread.Source = frm1.vspdData1
		 	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
			frm1.vspdData1.Col = C_Facility_Lvl2
			frm1.vspdData1.text	=	arrRet(0)
			frm1.vspdData1.Col = C_Facility_Lvl2Nm
			frm1.vspdData1.text	=	arrRet(1)
	End	Select

End	Function





'========================================================================================================
'		Event	Name : vspdData1_Change
'		Event	Desc :
'========================================================================================================
Sub	vspdData1_Change(ByVal Col , ByVal Row )
	Dim	iDx
	Dim	intIndex,	IntRetCd
	Dim	strName
	Dim	strDept_nm
	Dim	strRoll_pstn
	Dim	strPay_grd1
	Dim	strPay_grd2
	Dim	strEntr_dt
	Dim	strInternal_cd

 	Frm1.vspdData1.Row = Row
 	Frm1.vspdData1.Col = Col

	Select Case	Col
		Case C_FACILITY_ACCNT_NM
			Frm1.vspdData1.col = C_FACILITY_ACCNT_NM
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_FACILITY_ACCNT_CD
			Frm1.vspdData1.value = intindex
		Case C_Facility_Lvl1Nm
			Frm1.vspdData1.col = C_Facility_Lvl1Nm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Facility_Lvl1
			Frm1.vspdData1.value = intindex
		Case C_Facility_Lvl2Nm
			Frm1.vspdData1.col = C_Facility_Lvl2Nm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Facility_Lvl2
			Frm1.vspdData1.value = intindex
	End	Select

 	If Frm1.vspdData1.CellType =	parent.SS_CELL_TYPE_FLOAT	Then
		If	UNICDbl(Frm1.vspdData1.text) <	UNICDbl(Frm1.vspdData1.TypeFloatMin) Then
			 Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
		End	If
	End	If

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.UpdateRow	Row
	lgBlnFlgChgValue = TRUE

		lgCurrentSpd = "M"

End	Sub

Sub	vspdData2_Change(ByVal Col , ByVal Row )
	Dim	iDx, intIndex

 	Frm1.vspdData2.Row = Row
 	Frm1.vspdData2.Col = Col
	Select Case	Col
		Case	 C_Zinsp_PartNm
			Frm1.vspdData2.col = C_Zinsp_PartNm
			intIndex = Frm1.vspdData2.value
			Frm1.vspdData2.Col = C_Zinsp_PartCd
			Frm1.vspdData2.value = intindex
		Case	 C_Insp_PartNm
			Frm1.vspdData2.col = C_Insp_PartNm
			intIndex = Frm1.vspdData2.value
			Frm1.vspdData2.Col = C_Insp_PartCd
			Frm1.vspdData2.value = intindex
		Case	 C_Insp_MethNm
			Frm1.vspdData2.col = C_Insp_MethNm
			intIndex = Frm1.vspdData2.value
			Frm1.vspdData2.Col = C_Insp_MethCd
			Frm1.vspdData2.value = intindex
		Case	 C_Insp_DecisionNm
			Frm1.vspdData2.col = C_Insp_DecisionNm
			intIndex = Frm1.vspdData2.value
			Frm1.vspdData2.Col = C_Insp_DecisionCd
			Frm1.vspdData2.value = intindex
		Case	 C_St_Go_GubunNm
			Frm1.vspdData2.col = C_St_Go_GubunNm
			intIndex = Frm1.vspdData2.value
			Frm1.vspdData2.Col = C_St_Go_GubunCd
			Frm1.vspdData2.value = intindex
	End	Select

 	If Frm1.vspdData2.CellType =	parent.SS_CELL_TYPE_FLOAT	Then
		If	UNICDbl(Frm1.vspdData2.text) <	UNICDbl(Frm1.vspdData2.TypeFloatMin) Then
			 Frm1.vspdData2.text = Frm1.vspdData2.TypeFloatMin
		End	If
	End	If

	ggoSpread.Source	=	frm1.vspdData2
	ggoSpread.UpdateRow Row
	
	lgCurrentSpd = "S1"

End	Sub


'========================================================================================================
'		Event	Name : vspdData1_Click
'		Event	Desc : 컬럼을	클릭할 경우	발생 
'========================================================================================================
Sub	vspdData1_Click(ByVal	Col, ByVal Row)
	Dim	flagTxt,IntRetCD

	gMouseClickStatus	=	"SPC"
	Set	gActiveSpdSheet	=	frm1.vspdData1
	Call SetPopupMenuItemInf("0000111111")

	If frm1.vspdData1.MaxRows	<=	0	or Col <=0 or Row <=0 Then																										'If	there	is no	data.
		 Exit	Sub
 	End	If
 	
	ggoSpread.Source = frm1.vspdData1
	With Frm1
		.vspdData1.Row = Row
		.vspdData1.Col = 0
		flagTxt	=	.vspdData1.Text
		
		If flagTxt =	ggoSpread.InsertFlag or	flagTxt	=	 ggoSpread.UpdateFlag	or flagTxt =	ggoSpread.DeleteFlag Then
				Exit Sub
		End	If
	End	With


	If Row <=	0	Then
	'	 ggoSpread.Source	=	frm1.vspdData1

	'	 If	lgSortKey	=	1	Then
	'			 ggoSpread.SSSort	Col								'Sort	in ascending
	'			 lgSortKey = 2
	'	 Else
	'			 ggoSpread.SSSort	Col, lgSortKey		'Sort	in descending
	'			 lgSortKey = 1
	'	 End If
'
'		frm1.vspddata1.Row = frm1.vspdData1.ActiveRow
'		frm1.vspddata2.MaxRows = 0	
'		lgOldRow_A = frm1.vspddata1.Row
	Else
		If lgOldRow_A	<> Row  Then
				ggoSpread.Source = frm1.vspdDAta2
			If ggoSpread.SSCheckChange = True Then
				IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
				If IntRetCD = vbNo Then
					Exit Sub
				End If
			End If
			lgCurrentSpd = "S1"
			lgStrPrevKey1	=	""
			lgStrPrevKey2	=	""
			frm1.vspdData1.Row = Row
			frm1.vspdData1.Col = C_FACILITY_ACCNT_CD
			frm1.hFacility_Accnt.value = frm1.vspdData1.text
			frm1.vspdData1.Col = C_Facility_Lvl1
			frm1.hItemGroupCd1.value = frm1.vspdData1.text
			frm1.vspdData1.Col = C_Facility_Lvl2
			frm1.hItemGroupCd2.value = frm1.vspdData1.text
			frm1.vspdData1.Col = C_chkF1
			frm1.hChkFlag.value = frm1.vspdData1.text
			
			lgOldRow_A = Row
					
			Frm1.vspdData2.MaxRows = 0
			Call	DisableToolBar(	parent.TBC_QUERY)

			If DbDtlQuery	=	false	Then
				Call	RestoreToolBar()
				Exit Sub
			End	If			
		End	if
		lgCurrentSpd = "M"	
		'Set	gActiveSpdSheet	=	frm1.vspdData1
	End If
End	Sub

'========================================================================================================
'		Event	Name : vspdData2_Click
'		Event	Desc : 컬럼을	클릭할 경우	발생 
'========================================================================================================
Sub	vspdData2_Click(ByVal	Col, ByVal Row)

	Call SetPopupMenuItemInf("1101011111")

	IF lgBlnFlgChgValue	=	True and frm1.vspdData1.Maxrows	>	0	then
		Call SetToolbar("1100111100011111")
	End	if

	gMouseClickStatus	=	"SP1C"

	Set	gActiveSpdSheet	=	frm1.vspdData2

	If frm1.vspdData2.MaxRows	=	0	Then																										'If	there	is no	data.
		 Exit	Sub
 	End	If

	If Row <=	0	Then
		 ggoSpread.Source	=	frm1.vspdData2

		 If	lgSortKey	=	1	Then
				 ggoSpread.SSSort	Col								'Sort	in ascending
				 lgSortKey = 2
		 Else
				 ggoSpread.SSSort	Col, lgSortKey		'Sort	in descending
				 lgSortKey = 1
		 End If

		 Exit	Sub
	End	If
	lgCurrentSpd = "S1"
	Set	gActiveSpdSheet	=	frm1.vspdData2

End	Sub

'========================================================================================================
'		Event	Name : vspdData1_ColWidthChange
'		Event	Desc :
'========================================================================================================
Sub	vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End	Sub

'========================================================================================================
'		Event	Name : vspdData1_ColWidthChange
'		Event	Desc :
'========================================================================================================
Sub	vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End	Sub


'========================================================================================================
'		Event	Name : vspdData1_ScriptDragDropBlock
'		Event	Desc :
'========================================================================================================
Sub	vspdData1_ScriptDragDropBlock( Col ,	Row,	Col2,	 Row2,	NewCol,	 NewRow,	NewCol2,	NewRow2,	Overwrite	,	Action , DataOnly	,	Cancel )
	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SpreadDragDropBlock(Col , Row,	Col2,	Row2,	NewCol,	NewRow,	NewCol2, NewRow2,	Overwrite	,	Action , DataOnly	,	Cancel )
	Call GetSpreadColumnPos("A")
End	Sub

'========================================================================================================
'		Event	Name : vspdData1_ScriptDragDropBlock
'		Event	Desc :
'========================================================================================================
Sub	vspdData2_ScriptDragDropBlock( Col ,	Row,	Col2,	 Row2,	NewCol,	 NewRow,	NewCol2,	NewRow2,	Overwrite	,	Action , DataOnly	,	Cancel )
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.SpreadDragDropBlock(Col , Row,	Col2,	Row2,	NewCol,	NewRow,	NewCol2, NewRow2,	Overwrite	,	Action , DataOnly	,	Cancel )
	Call GetSpreadColumnPos("B")
End	Sub


Sub	vspdData1_MouseDown(Button , Shift , x , y)
	If	Button = 2 And	gMouseClickStatus	=	"SPC"	Then
		gMouseClickStatus = "SPCR"
	End	If
End	Sub
Sub	vspdData2_MouseDown(Button , Shift , x , y)
	If Button	=	2	And	gMouseClickStatus	=	"SP1C" Then
		gMouseClickStatus = "SP1CR"
	End If
End	Sub


'========================================================================================================
'		Event	Name : vspdData1_ButtonClicked
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData1_ButtonClicked(ByVal	Col, ByVal Row,	Byval	ButtonDown)
	dim	strTemp

 	frm1.vspdData1.Row = Row
	frm1.vspdData1.Col = Col
 	ggoSpread.Source = frm1.vspdData1

	Select Case	Col
		Case C_Facility_Lvl1_Pop
			frm1.vspdData1.Col = C_Facility_Lvl1
			strTemp	=	frm1.vspdData1.text
			Call OpenGroup(strTemp,	3)
		Case C_Facility_Lvl2_Pop
			frm1.vspdData1.Col = C_Facility_Lvl2
			strTemp	=	frm1.vspdData1.text
			Call OpenGroup(strTemp,	4)
	End	Select
End	Sub



'========================================================================================
'	Function Name	:	FncSplitColumn
'	Function Desc	:
'========================================================================================
Function FncSplitColumn()
	Dim	ACol
	Dim	ARow
	Dim	iRet
	Dim	iColumnLimit

	iColumnLimit	=	5

	If	gMouseClickStatus	=	"SPCR" Then
		 ACol	=	Frm1.vspdData1.ActiveCol
		 ARow	=	Frm1.vspdData1.ActiveRow

		 If	ACol > iColumnLimit	Then
				Frm1.vspdData1.Col = iColumnLimit	:	Frm1.vspdData1.Row = 0	:	iRet =	DisplayMsgBox("900030",	"X", Trim(frm1.vspdData1.Text),	"X")
				Exit Function
		 End If

		 Frm1.vspdData1.ScrollBars =	parent.SS_SCROLLBAR_NONE

			ggoSpread.Source = Frm1.vspdData1

			ggoSpread.SSSetSplit(ACol)

		 Frm1.vspdData1.Col	=	ACol
		 Frm1.vspdData1.Row	=	ARow

		 Frm1.vspdData1.Action = 0

		 Frm1.vspdData1.ScrollBars =	parent.SS_SCROLLBAR_BOTH
	End	If

	If	gMouseClickStatus	=	"SP1CR"	Then
		 ACol	=	Frm1.vspdData2.ActiveCol
		 ARow	=	Frm1.vspdData2.ActiveRow

		 If	ACol > iColumnLimit	Then
				Frm1.vspdData2.Col = iColumnLimit	:	Frm1.vspdData2.Row = 0	:	iRet =	DisplayMsgBox("900030",	"X", Trim(frm1.vspdData2.Text),	"X")
				Exit Function
		 End If

		 Frm1.vspdData2.ScrollBars =	parent.SS_SCROLLBAR_NONE

			ggoSpread.Source = Frm1.vspdData2

			ggoSpread.SSSetSplit(ACol)

		 Frm1.vspdData2.Col	=	ACol
		 Frm1.vspdData2.Row	=	ARow

		 Frm1.vspdData2.Action = 0

		 Frm1.vspdData2.ScrollBars =	parent.SS_SCROLLBAR_BOTH
	End	If

 End Function

'========================================================================================================
'		Event	Name : vspdData1_ScriptLeaveCell
'		Event	Desc : This	function is	called when	cursor leave cell
'========================================================================================================
Sub	vspdData1_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	If NewRow	<= 0 Or	Row	=	NewRow Then	Exit Sub

	ggoSpread.Source = frm1.vspdData1
	frm1.vspdData1.Row = NewRow
	frm1.vspdData1.Col = C_FACILITY_ACCNT_CD
	frm1.hFacility_Accnt.value = frm1.vspdData1.text
	frm1.vspdData1.Col = C_Facility_Lvl1
	frm1.hItemGroupCd1.value = frm1.vspdData1.text
	frm1.vspdData1.Col = C_Facility_Lvl2
	frm1.hItemGroupCd2.value = frm1.vspdData1.text

	frm1.vspdData1.Col = 0
	if frm1.vspdData1.text = ggoSpread.InsertFlag	Then
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		Exit sub
	end	if

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData

	If DbDtlQuery()	=	False	Then	Exit Sub
End	Sub

'========================================================================================================
'		Event	Name : vspdData1_OnFocus
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData1_OnFocus()
	lgActiveSpd			 = "M"
	lgCurrentSpd	="M"
End	Sub
'========================================================================================================
'		Event	Name : vspdData2_OnFocus
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData2_OnFocus()
		lgActiveSpd			 = "S1"
	lgCurrentSpd	="S1"
End	Sub

'========================================================================================================
'		Event	Name : vspdData1_TopLeftChange
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData1_TopLeftChange(ByVal	OldLeft	,	ByVal	OldTop , ByVal NewLeft , ByVal NewTop	)
	If OldLeft <>	NewLeft	Then
			Exit Sub
	End	If
	lgCurrentSpd	=	 "M"
	call MakeKeyStream("X")
	if frm1.vspdData1.MaxRows	<	NewTop +	VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey	<> ""	Then
			Call	 DisableToolBar( parent.TBC_QUERY)
			If	DbQuery	=	false	Then
				Call	RestoreToolBar()
				Exit Sub
			End If
		End	If
	End	if
End	Sub

'========================================================================================================
'		Event	Name : vspdData2_TopLeftChange
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData2_TopLeftChange(ByVal	OldLeft	,	ByVal	OldTop , ByVal NewLeft , ByVal NewTop	)

	If OldLeft <>	NewLeft	Then
			Exit Sub
	End	If
	lgCurrentSpd	="S1"
	If frm1.vspdData2.MaxRows	<	NewTop +	VisibleRowCnt(frm1.vspdData2,NewTop) Then

		If lgStrPrevKey1 <>	"" Then
			Call	 DisableToolBar( parent.TBC_QUERY)
			Call	MakeKeyStream(frm1.vspdData1.activeRow)
			If	DbDtlQuery = false Then
				Call	RestoreToolBar()
				Exit Sub
			End If
		End	If
	End	if
End	Sub


'========================================================================================================
'		Event	Name : vspdData1_GotFocus
'		Event	Desc : This	event	is spread	sheet	data changed
'========================================================================================================
Sub	vspdData1_GotFocus()
		ggoSpread.Source = Frm1.vspdData1
End	Sub
'========================================================================================================
'		Event	Name : vspdData1_GotFocus
'		Event	Desc : This	event	is spread	sheet	data changed
'========================================================================================================
Sub	vspdData2_GotFocus()
		ggoSpread.Source = Frm1.vspdData2
End	Sub





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
								<td	background="../../../CShared/image/table/seltab_up_bg.gif" align="center"	CLASS="CLSMTAB"><font	color=white>설비점검항목등록</font></td>
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
									<TD	CLASS="TD5"	NOWRAP>설비유형</TD>
									<TD	CLASS="TD6"	NOWRAP><SELECT NAME="CbohFacility_Accnt" ALT="설비유형"	CLASS	="CbohFacility_Accnt"	TAG="1XN"><OPTION	VALUE=""></OPTION></SELECT></TD>
									<TD	CLASS="TD5"	NOWRAP></TD>
									<TD	CLASS="TD6"	NOWRAP></TD>
								</TR>
								<TR>
									<TD	CLASS="TD5"	NOWRAP>대분류</TD>
									<TD	CLASS="TD6"	NOWRAP><SELECT NAME="txtItemGroupCd1"	ALT="대분류" CLASS ="txtItemGroupCd1"	TAG="1XN"><OPTION	VALUE=""></OPTION></SELECT></TD>
									<TD	CLASS="TD5"	NOWRAP>중분류</TD>
									<TD	CLASS="TD6"	NOWRAP><SELECT NAME="txtItemGroupCd2"	ALT="중분류" CLASS ="txtItemGroupCd2"	TAG="1XN"><OPTION	VALUE=""></OPTION></SELECT></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD	<%=HEIGHT_TYPE_03%>	WIDTH=100%></TD>
				</TR>
				<TR>
					<TD	WIDTH=100% HEIGHT=*	valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR	HEIGHT="50%">
								<TD	WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p5210ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR	HEIGHT="50%">
								<TD	WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p5210ma1_B_vspdData2.js'></script>
								</TD>
							</TR>
							<TR>
								<TD	HEIGHT=5 WIDTH=100%	colspan=4></TD>
							</TR>
						</Table>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD	<%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD	WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100%	HEIGHT=20	FRAMEBORDER=0	SCROLLING=NO noresize	framespacing=0 TABINDEX	=	"-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA	CLASS="hidden" NAME="txtSpread"	tag="24" TABINDEX	=	"-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"	tag="24"><INPUT	TYPE=HIDDEN	NAME="txtMaxRows"	tag="24" TABINDEX="-1"><INPUT	TYPE=HIDDEN	NAME="txtFlgMode"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hFacility_Accnt"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd1"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd2"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="ChgSave1" tag="24"	TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="ChgSave2" tag="24"	TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hChkFlag" tag="24" tabindex="-1">

</FORM>
<DIV ID="MousePT"	NAME="MousePT">
<iframe	name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO	noresize framespacing=0	width=220	height=41	src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
