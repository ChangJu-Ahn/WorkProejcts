<%@	LANGUAGE="VBSCRIPT"	%>
<!--
======================================================================================================
*	 1.	Module Name					 : Facility	Resources
*	 2.	Function Name				 : 설비수리내역등록 
*	 3.	Program	ID					 : FA105MA1
*	 4.	Program	Name				 :
*	 5.	Program	Desc				 : 설비수리내역등록 
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
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script	Language="VBScript">
Option Explicit
'========================================================================================================
'=											 4.2 Constant	variables
'========================================================================================================
Const	CookieSplit	=	1233
Const	BIZ_PGM_ID = "P5230mb1.asp"																			 'Biz	Logic	ASP
Const	C_SHEETMAXROWS		=	100																				'한	화면에 보여지는	최대갯수*1.5%>

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
Dim	lgPageNo_C
Dim	lgOldRow_A
Dim	lgOldRow_B
Dim	lgOldRow_C


Dim	C_Facility_Cd
Dim	C_FacilityPop
Dim	C_Facility_Nm
Dim	C_Set_Plant
Dim	C_Set_PlantNm
Dim	C_Facility_Accnt_Nm
Dim	C_Plan_Dt
Dim	C_Insp_Text
Dim	C_Insp_Hour
Dim	C_Insp_Min
Dim	C_Req_Dept
Dim	C_Req_Dept_POP
Dim	C_Req_Dept_Nm
Dim	C_Insp_Dept
Dim	C_Insp_Dept_POP
Dim	C_Insp_Dept_Nm
Dim	C_Insp_Emp_Qty
Dim	C_Payroll
Dim	C_Matl_Cost
Dim	C_Insp_Flag
Dim	C_Insp_FlagNm


Dim	C_Seq
Dim	C_Zinsp_PartCd
Dim	C_Zinsp_PartNm
Dim	C_Insp_PartCd
Dim	C_Insp_PartNm
Dim	C_Insp_MethCd
Dim	C_Insp_MethNm
Dim	C_Insp_DeciCd
Dim	C_Insp_DeciNm
Dim	C_St_GoGubunCd
Dim	C_St_GoGubunNm
Dim	C_Sury_Assy
Dim	C_Sury_Assy_Pop
Dim	C_Sury_Assy_Nm
Dim	C_S_Qty
Dim	C_Price
Dim	C_Sury_Amt
Dim C_Cur
Dim C_Cur_Popup
Dim	C_Sury_Type
Dim	C_Sury_TypeNm



'	Dim	C_Seq
Dim	C_Insp_Emp_Gb
Dim	C_Insp_Emp_GbNm
Dim	C_Insp_Emp_Cd
Dim	C_Insp_Emp_Pop
Dim	C_Insp_Emp_Nm
Dim	C_Cust_Cd
Dim	C_Cust_Pop
Dim	C_Cust_Nm
Dim	C_Insp_Hour2
Dim	C_Insp_Min2
Dim	C_Payroll2



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

	If pvSpdNo = "A" Then
		C_Facility_Cd			 	=	1
		C_FacilityPop				=	2
		C_Facility_Nm			 	=	3
		C_Set_Plant				 	=	4
		C_Set_PlantNm			 	=	5
		C_Facility_Accnt_Nm			=	6
		C_Plan_Dt					=	7
		C_Insp_Text				 	=	8
		C_Insp_Hour				 	=	9
		C_Insp_Min				 	=	10
		C_Req_Dept				 	=	11
		C_Req_Dept_POP		 		=	12
		C_Req_Dept_Nm			 	=	13
		C_Insp_Dept				 	=	14
		C_Insp_Dept_POP		 		=	15
		C_Insp_Dept_Nm		 		=	16
		C_Insp_Emp_Qty		 		=	17
		C_Payroll					=	18
		C_Matl_Cost				 	=	19
		C_Insp_Flag				 	=	20
		C_Insp_FlagNm			 	=	21
	ElseIf pvSpdNo = "B" Then
		C_Seq						=	1
		C_Zinsp_PartCd				=	2
		C_Zinsp_PartNm				=	3
		C_Insp_PartCd				=	4
		C_Insp_PartNm				=	5
		C_Insp_MethCd				=	6
		C_Insp_MethNm				=	7
		C_Insp_DeciCd				=	8
		C_Insp_DeciNm				=	9
		C_St_GoGubunCd				=	10
		C_St_GoGubunNm				=	11
		C_Sury_Assy					=	12
		C_Sury_Assy_Pop	 			=	13
		C_Sury_Assy_Nm	 			=	14
		C_S_Qty					 	=	15
		C_Price					 	=	16
		C_Sury_Amt			 		=	17
		C_Cur						=	18
		C_Cur_Popup					=	19
		C_Sury_Type			 		=	20
		C_Sury_TypeNm				=	21

	ElseIf pvSpdNo = "C" Then
		C_Seq						=	1
		C_Insp_Emp_Gb		 		=	2
		C_Insp_Emp_GbNm	 			=	3
		C_Insp_Emp_Cd		 		=	4
		C_Insp_Emp_Pop	 			=	5
		C_Insp_Emp_Nm		 		=	6
		C_Cust_Cd					=	7
		C_Cust_Pop			 		=	8
		C_Cust_Nm					=	9
		C_Insp_Hour2		 		=	10
		C_Insp_Min2			 		=	11
		C_Payroll2			 		=	12
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
	lgPageNo_A = 0
	lgPageNo_B = 0
	lgPageNo_C = 0
End	Sub

'========================================================================================================
'	Name : LoadInfTB19029()
'	Desc : Set System	Number format
'========================================================================================================
Sub	LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>  ' check
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
			 lgKeyStream = UNIConvDate(Trim(Frm1.txtWork_Dt.text)) & parent.gColSep
			 lgKeyStream = lgKeyStream & Frm1.txtPlantCd.value & parent.gColSep
			 lgKeyStream = lgKeyStream & Frm1.CboFacility_Accnt.value	&	parent.gColSep
			 lgKeyStream = lgKeyStream & Frm1.txtFacility_Cd.value & parent.gColSep
		Else

			frm1.vspdData.Row	=	pRow
		frm1.vspdData.Col	=	C_FACILITY_CD
				lgKeyStream	=	frm1.vspdData.Text & parent.gColSep			'You Must	append one character(	parent.gColSep)
		frm1.vspdData.Col	=	C_Plan_Dt
				lgKeyStream	=	lgKeyStream	&	UNIConvDate(Trim(frm1.vspdData.Text))	&	parent.gColSep		 'You	Must append	one	character( parent.gColSep)
		End	If
End	Sub

'========================================================================================================
'	Name : InitComboBox()
'	Desc : Set ComboBox
'========================================================================================================
Sub	InitComboBox()
	Dim	iCodeArr
	Dim	iNameArr


	Call CommonQueryRs(" MINOR_CD, MINOR_NM	","	B_MINOR	","	MAJOR_CD = 'Z410'	",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.CboFacility_Accnt	,lgF0	 ,lgF1	,Chr(11))

	ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo "Y" & vbtab & "N" , C_Insp_FlagNm




	ggoSpread.Source = frm1.vspdData1


	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z425' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Zinsp_PartCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Zinsp_PartNm

	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z411' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Insp_PartCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Insp_PartNm

	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z412' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Insp_MethCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Insp_MethNm

	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z418' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Insp_DeciCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Insp_DeciNm

	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z419' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_St_GoGubunCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_St_GoGubunNm

	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z420' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Sury_Type
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Sury_TypeNm


End	Sub
'========================================================================================================
'	Name : InitComboBox()
'	Desc : Set ComboBox
'========================================================================================================
Sub	InitComboBox2()
		Dim	iCodeArr
		Dim	iNameArr
		'	수당코드 
	Call	CommonQueryRs("	minor_cd,	minor_nm "," b_minor "," major_cd	=	'Z424'	",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iCodeArr = lgF0
		iNameArr = lgF1
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_Insp_Emp_GbNm
		ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_Insp_Emp_Gb
End	Sub
'========================================================================================================
'	Name : InitComboBox()
'	Desc : Set ComboBox
'========================================================================================================
Sub	InitComboBox3()
'			Dim	iCodeArr
'			Dim	iNameArr
'			'	수당코드 
'			Call	CommonQueryRs("	ALLOW_CD,ALLOW_NM	","	HDA010T	","	PAY_CD = " & FilterVar("*",	"''",	"S1")	&	"	 AND CODE_TYPE = " & FilterVar("1",	"''",	"S1")	&	"	 ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'			iCodeArr = lgF0
'			iNameArr = lgF1
'			ggoSpread.Source = frm1.vspdData1
'			ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab),	C_ALLOW_CD_NM
'			ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab),	C_ALLOW_CD
End	Sub

'========================================================================================================
'	Name : InitData()
'	Desc : Reset Combox
'========================================================================================================
Sub	InitData()
	Dim	intRow
	Dim	intIndex

			 ggoSpread.Source	=	frm1.vspdData1
			With frm1.vspdData1
				For	intRow = 1 To	.MaxRows
					.Row = intRow

'						.Col = C_ALLOW_CD					'	수당코드 
'						intIndex = .value
'						.col = C_ALLOW_CD_NM
'						.value = intindex

				Next
			End	With
End	Sub

'========================================================================================================
'	Function Name	:	InitSpreadSheet
'	Function Desc	:	This method	initializes	spread sheet column	property
'========================================================================================================
Sub	InitSpreadSheet(ByVal	pvSpdNo)

	If pvSpdNo = ""	OR pvSpdNo = "A" Then

		Call initSpreadPosVariables("A")
		With frm1.vspdData

				ggoSpread.Source = frm1.vspdData
				ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread

				.ReDraw	=	false

				.MaxCols = C_Insp_FlagNm + 1																								<%'☜: 최대	Columns의	항상 1개 증가시킴 %>
				.Col = .MaxCols															<%'공통콘트롤	사용 Hidden	Column%>
				.ColHidden = True

				.MaxRows = 0
				ggoSpread.ClearSpreadData

				Call AppendNumberPlace("6","2","0")
				Call GetSpreadColumnPos("A")

				ggoSpread.SSSetEdit		C_FACILITY_CD,			"설비코드",			15,,,18,2
				ggoSpread.SSSetButton	C_FacilityPop
				ggoSpread.SSSetEdit		C_FACILITY_NM,			"설비명",			20,,,40,2
				ggoSpread.SSSetEdit		C_SET_PLANT,				"설치공장",			15,,,20,2
				ggoSpread.SSSetEdit		C_SET_PLANTNm,			"설치공장",		15,,,40,2
				ggoSpread.SSSetEdit		C_FACILITY_ACCNT_NM,	"설비유형",			15,,,20,2
				ggoSpread.SSSetDate		C_Plan_Dt, 				"작업일자",	12,2,gDateFormat
				ggoSpread.SSSetEdit		C_Insp_Text,			"점검/수리내용",			15,,,25,1
				ggoSpread.SSSetFloat	C_Insp_Hour,			"소요시간",	11,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","23"
				ggoSpread.SSSetFloat	C_Insp_Min,				"소요분",11,"6",ggStrIntegeralPart,	ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","59"
				ggoSpread.SSSetEdit		C_Req_Dept,				"의뢰부서",			15,,,10,2
				ggoSpread.SSSetButton	C_Req_Dept_POP
				ggoSpread.SSSetEdit		C_Req_Dept_Nm,			"의뢰부서명",			15,,,40,2
				ggoSpread.SSSetEdit		C_Insp_Dept,			"수리부서",			15,,,10,2
				ggoSpread.SSSetButton	C_Insp_Dept_POP
				ggoSpread.SSSetEdit		C_Insp_Dept_Nm,			"수리부서명",			15,,,40,2
				ggoSpread.SSSetFloat	C_Insp_Emp_Qty,    		"수리인원", 		20, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
				ggoSpread.SSSetFloat	C_Payroll,				"인건비",			20,	 parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	 ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,,"0"
				ggoSpread.SSSetFloat	C_Matl_Cost,			"소모자재비",			20,	 parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	 ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,,"0"
				ggoSpread.SSSetCombo 	C_Insp_Flag,			"점검여부",	 10, 0,	False
				ggoSpread.SSSetCombo 	C_Insp_FlagNm,			"점검여부",	 10, 0,	False

				Call ggoSpread.MakePairsColumn(C_FACILITY_CD,	 C_FACILITY_NM)
				Call ggoSpread.SSSetColHidden(C_SET_PLANT,	C_SET_PLANT, True)
				Call ggoSpread.SSSetColHidden(C_Insp_Flag,	C_Insp_Flag, True)
				Call ggoSpread.SSSetColHidden(C_Insp_FlagNm,	C_Insp_FlagNm, True)

				.ReDraw	=	true

				Call SetSpreadLock

		End	With

	End	if

	If pvSpdNo = ""	OR pvSpdNo = "B" Then
		Call initSpreadPosVariables("B")
		With frm1.vspdData1

			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread

			.ReDraw	=	false
			.MaxCols = C_Sury_TypeNm + 1												<%'☜: 최대	Columns의	항상 1개 증가시킴 %>
			.Col = .MaxCols															<%'공통콘트롤	사용 Hidden	Column%>
			.ColHidden = True

			.MaxRows = 0

			Call AppendNumberPlace("6","2","0")
			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetFloat	C_Seq,				"순서",    8, "7", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"

			ggoSpread.SSSetCombo 	C_Zinsp_PartCd,		"부위",	 10, 0,	False
			ggoSpread.SSSetCombo	C_Zinsp_PartNm,		"부위",	 10, 0,	False
			ggoSpread.SSSetCombo	C_Insp_PartCd,		"수리항목",	 10, 0,	False
			ggoSpread.SSSetCombo	C_Insp_PartNm,		"수리항목",	 10, 0,	False
			ggoSpread.SSSetCombo	C_Insp_MethCd,		"수리방법",	 10, 0,	False
			ggoSpread.SSSetCombo	C_Insp_MethNm,		"수리방법",	 10, 0,	False
			ggoSpread.SSSetCombo	C_Insp_DeciCd,		"판정기준",	 10, 0,	False
			ggoSpread.SSSetCombo	C_Insp_DeciNm,		"판정기준",	 10, 0,	False
			ggoSpread.SSSetCombo	C_St_GoGubunCd,		"운/휴구분",	14,	0, False
			ggoSpread.SSSetCombo	C_St_GoGubunNm,		"운/휴구분",	14,	0, False
			ggoSpread.SSSetEdit		C_Sury_Assy,		"부품코드",			15,,,18,2
			ggoSpread.SSSetButton	C_Sury_Assy_Pop
			ggoSpread.SSSetEdit		C_Sury_Assy_Nm,		"부품명",			20,,,20,2
			ggoSpread.SSSetFloat	C_S_Qty,			"수량", 		19, Parent.ggQtyNo			, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0"
			ggoSpread.SSSetFloat	C_Price,			"단가",			15,		parent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
			ggoSpread.SSSetFloat	C_Sury_Amt,			"금액",			20,	 parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	 ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,,"0"
			ggoSpread.SSSetEdit     C_Cur,               "화폐", 10, 0,,3,2
			ggoSpread.SSSetButton   C_Cur_Popup 
			ggoSpread.SSSetCombo 	C_Sury_Type,		"조치유형",	14,	0, False
			ggoSpread.SSSetCombo 	C_Sury_TypeNm,		"조치유형",	 14, 0,	False

			Call ggoSpread.SSSetColHidden(C_Zinsp_PartCd,	C_Zinsp_PartCd,	True)
			Call ggoSpread.SSSetColHidden(C_Insp_PartCd, C_Insp_PartCd,	True)
			Call ggoSpread.SSSetColHidden(C_Insp_MethCd, C_Insp_MethCd,	True)
			Call ggoSpread.SSSetColHidden(C_Insp_DeciCd, C_Insp_DeciCd,	True)
			Call ggoSpread.SSSetColHidden(C_St_GoGubunCd,	C_St_GoGubunCd,	True)
			Call ggoSpread.SSSetColHidden(C_Sury_Type, C_Sury_Type,	True)

			.ReDraw	=	true

		Call SetSpreadLock1

		End	With
	End	if

	If pvSpdNo = ""	OR pvSpdNo = "C" Then

		Call initSpreadPosVariables("C")
		With frm1.vspdData2

			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread

			.ReDraw	=	false
			.MaxCols = C_Payroll2	+	1												<%'☜: 최대	Columns의	항상 1개 증가시킴 %>
			.Col = .MaxCols															<%'공통콘트롤	사용 Hidden	Column%>
			.ColHidden = True

			.MaxRows = 0

			Call AppendNumberPlace("6","2","0")
			Call GetSpreadColumnPos("C")

			ggoSpread.SSSetFloat	C_Seq,				"순서",    8, "7", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
			ggoSpread.SSSetCombo 	C_Insp_Emp_Gb,		"구분",	 10, 0,	False
			ggoSpread.SSSetCombo 	C_Insp_Emp_GbNm,	"구분",	 15, 0,	False
			ggoSpread.SSSetEdit		C_Insp_Emp_Cd,		"수리자",			10,,,13,2
			ggoSpread.SSSetButton		C_Insp_Emp_Pop
			ggoSpread.SSSetEdit		C_Insp_Emp_Nm,		"이름",			15,,,20,2
			ggoSpread.SSSetEdit		C_Cust_Cd,			"수리업체",			15,,,20,2
			ggoSpread.SSSetButton		C_Cust_Pop
			ggoSpread.SSSetEdit		C_Cust_Nm,			"업체명",			15,,,20,2
			ggoSpread.SSSetFloat	C_Insp_Hour2,		"소요시간",	11,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","23"
			ggoSpread.SSSetFloat	C_Insp_Min2,		"소요분",	11,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","59"
			ggoSpread.SSSetFloat	C_Payroll2,			"인건비",			20,	 parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	 ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,,"0"

			Call ggoSpread.SSSetColHidden(C_Insp_Emp_Gb,	C_Insp_Emp_Gb, True)

		.ReDraw	=	true

		Call SetSpreadLock2

		End	With
	End	if
End	Sub

'======================================================================================================
'	Function Name	:	SetSpreadLock
'	Function Desc	:	This method	set	color	and	protect	in spread	sheet	celles
'======================================================================================================
Sub	SetSpreadLock()

	With frm1.vspdData

				ggoSpread.Source = frm1.vspdData

				.ReDraw	=	False

			ggoSpread.SpreadLock 	C_FACILITY_CD			,	-1,	C_Plan_Dt
			ggoSpread.SpreadLock	C_Req_Dept_Nm			,	-1,	C_Req_Dept_Nm
			ggoSpread.SpreadLock	C_Insp_Dept_Nm			,	-1,	C_Insp_Dept_Nm
			ggoSpread.SpreadLock	C_Payroll				,	-1,	C_Payroll
			ggoSpread.SpreadLock	C_Insp_Hour				,	-1,	C_Insp_Min
			ggoSpread.SpreadLock	C_Insp_Flag				,	-1,	C_Insp_FlagNm
			ggoSpread.SpreadLock	C_SET_PLANT				,	-1,	C_SET_PLANT



			ggoSpread.SSSetProtected					.MaxCols,-1,-1
				.ReDraw	=	True

		End	With

End	Sub

Sub	SetSpreadLock1()

	With frm1.vspdData1

				ggoSpread.Source = frm1.vspdData1

				.ReDraw	=	False
						ggoSpread.SpreadLock C_Seq		,	-1,	C_Seq
						ggoSpread.SpreadLock C_Sury_Assy_Nm		,	-1,	C_Sury_Assy_Nm
						ggoSpread.SSSetRequired		C_Zinsp_PartNm		,	-1,	C_Zinsp_PartNm
						ggoSpread.SSSetRequired		C_Insp_PartNm		,	-1,	C_Insp_PartNm
						ggoSpread.SSSetRequired		C_Insp_MethNm		,	-1,	C_Insp_MethNm
						ggoSpread.SSSetRequired		C_Insp_DeciNm	,	-1,	C_Insp_DeciNm
						ggoSpread.SSSetRequired		C_St_GoGubunNm	,	-1,	C_St_GoGubunNm

						ggoSpread.SSSetProtected				.MaxCols,-1,-1
				.ReDraw	=	True

		End	With

End	Sub


Sub	SetSpreadLock2()

	With frm1.vspdData2

				ggoSpread.Source = frm1.vspdData2

				.ReDraw	=	False
					ggoSpread.SpreadLock C_Seq		,	-1,	C_Seq
					ggoSpread.SpreadLock C_Cust_Nm			,	-1,	C_Cust_Nm
					ggoSpread.SpreadLock C_Insp_Emp_Nm		,	-1,	C_Insp_Emp_Nm
					ggoSpread.SSSetProtected				.MaxCols,-1,-1
				.ReDraw	=	True

		End	With

End	Sub

'======================================================================================================
'	Function Name	:	SetSpreadColor
'	Function Desc	:	This method	set	color	and	protect	in spread	sheet	celles
'======================================================================================================
Sub	SetSpreadColor(ByVal pvStartRow,ByVal	pvEndRow)

	With frm1.vspdData

				ggoSpread.Source = frm1.vspdData

				.ReDraw	=	False


			ggoSpread.SSSetRequired		C_FACILITY_CD		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_Plan_Dt			,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_FACILITY_NM		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_SET_PLANT			,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_SET_PLANTnm		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_FACILITY_ACCNT_NM	,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_Req_Dept_Nm		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_Insp_Hour			,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_Insp_Min			,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_Insp_Dept_Nm		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_Payroll			,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_Insp_Flag			,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_Insp_FlagNm		,	pvStartRow,	pvEndRow



'				ggoSpread.SSSetRequired	C_Plan_Dt			,	pvStartRow,	pvEndRow

			ggoSpread.SSSetProtected						.MaxCols,-1,-1
				.ReDraw	=	True

		End	With

End	Sub

Sub	SetSpreadColor1(ByVal	pvStartRow,ByVal pvEndRow)

	With frm1.vspdData1

				ggoSpread.Source = frm1.vspdData1

				.ReDraw	=	False

			ggoSpread.SSSetProtected	C_Sury_Assy_Nm			,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_Seq		,	pvStartRow,	pvEndRow

			ggoSpread.SSSetRequired		C_Zinsp_PartCd		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_Zinsp_PartNm		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_Insp_PartCd		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_Insp_PartNm		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_Insp_MethCd		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_Insp_MethNm		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_Insp_DeciCd		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_Insp_DeciNm		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_St_GoGubunCd		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_St_GoGubunNm		,	pvStartRow,	pvEndRow

			ggoSpread.SSSetProtected						.MaxCols,	pvStartRow,	pvEndRow
				.ReDraw	=	True

		End	With

End	Sub

Sub	SetSpreadColor2(ByVal	pvStartRow,ByVal pvEndRow)

	With frm1.vspdData2

				ggoSpread.Source = frm1.vspdData2

				.ReDraw	=	False

			ggoSpread.SSSetProtected	C_Insp_Emp_Nm			,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected	C_Cust_Nm			,	pvStartRow,	pvEndRow
			ggoSpread.SSSetRequired		C_Seq		,	pvStartRow,	pvEndRow
			ggoSpread.SSSetProtected						.MaxCols,	pvStartRow,	pvEndRow
				.ReDraw	=	True

		End	With

End	Sub
'======================================================================================================
'	Function Name	:	SubSetErrPos
'	Function Desc	:	This method	set	focus	to pos of	err
'======================================================================================================
Sub	SubSetErrPos(iPosArr)
		Dim	iDx
		Dim	iRow
		iPosArr	=	Split(iPosArr, parent.gColSep)
		If IsNumeric(iPosArr(0)) Then
			 iRow	=	CInt(iPosArr(0))
			 For iDx = 1 To	 frm1.vspdData.MaxCols - 1
					 Frm1.vspdData.Col = iDx
					 Frm1.vspdData.Row = iRow
					 If	Frm1.vspdData.ColHidden	<> True	And	Frm1.vspdData.BackColor	<>	parent.UC_PROTECTED	Then
							Frm1.vspdData.Col	=	iDx
							Frm1.vspdData.Row	=	iRow
							Frm1.vspdData.Action = 0 ' go	to
							Exit For
					 End If

			 Next

		End	If
End	Sub

'========================================================================================
'	Function Name	:	GetSpreadColumnPos
'	Description		:
'========================================================================================
Sub	GetSpreadColumnPos(ByVal pvSpdNo)
		Dim	iCurColumnPos

		Select Case	UCase(pvSpdNo)
			 Case	"A"
				ggoSpread.Source = frm1.vspdData
				Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

				C_Facility_Cd			 	=	iCurColumnPos(1	)
				C_FacilityPop				=	iCurColumnPos(2	)
				C_Facility_Nm			 	=	iCurColumnPos(3	)
				C_Set_Plant				 	=	iCurColumnPos(4	)
				C_Set_PlantNm			 	=	iCurColumnPos(5	)
				C_Facility_Accnt_Nm	=	iCurColumnPos(6	)
				C_Plan_Dt					 	=	iCurColumnPos(7	)
				C_Insp_Text				 	=	iCurColumnPos(8	)
				C_Insp_Hour				 	=	iCurColumnPos(9	)
				C_Insp_Min				 	=	iCurColumnPos(10)
				C_Req_Dept				 	=	iCurColumnPos(11)
				C_Req_Dept_POP		 	=	iCurColumnPos(12)
				C_Req_Dept_Nm			 	=	iCurColumnPos(13)
				C_Insp_Dept				 	=	iCurColumnPos(14)
				C_Insp_Dept_POP		 	=	iCurColumnPos(15)
				C_Insp_Dept_Nm		 	=	iCurColumnPos(16)
				C_Insp_Emp_Qty		 	=	iCurColumnPos(17)
				C_Payroll					 	=	iCurColumnPos(18)
				C_Matl_Cost				 	=	iCurColumnPos(19)
				C_Insp_Flag				 	=	iCurColumnPos(20)
				C_Insp_FlagNm			 	=	iCurColumnPos(21)

			 Case	"B"
						ggoSpread.Source = frm1.vspdData1
						Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

				C_Seq						 	=	iCurColumnPos(1	)
				C_Zinsp_PartCd		=	iCurColumnPos(2	)
				C_Zinsp_PartNm		=	iCurColumnPos(3	)
				C_Insp_PartCd		=	iCurColumnPos(4	)
				C_Insp_PartNm		=	iCurColumnPos(5	)
				C_Insp_MethCd		=	iCurColumnPos(6	)
				C_Insp_MethNm		=	iCurColumnPos(7	)
				C_Insp_DeciCd		=	iCurColumnPos(8	)
				C_Insp_DeciNm		=	iCurColumnPos(9	)
				C_St_GoGubunCd		=	iCurColumnPos(10)
				C_St_GoGubunNm		=	iCurColumnPos(11)
				C_Sury_Assy			 	=	iCurColumnPos(12)
				C_Sury_Assy_Pop	 	=	iCurColumnPos(13)
				C_Sury_Assy_Nm	 	=	iCurColumnPos(14)
				C_S_Qty					 	=	iCurColumnPos(15)
				C_Price					 	=	iCurColumnPos(16)
				C_Sury_Amt			 	=	iCurColumnPos(17)
				C_Cur				=	iCurColumnPos(18)
				C_Cur_Popup			=	iCurColumnPos(19)
				C_Sury_Type			=	iCurColumnPos(20)
				C_Sury_TypeNM		=	iCurColumnPos(21)

			 Case	"C"
						ggoSpread.Source = frm1.vspdData2
						Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_Seq						 	=	iCurColumnPos(1	)
				C_Insp_Emp_Gb		 	=	iCurColumnPos(2	)
				C_Insp_Emp_GbNm	 	=	iCurColumnPos(3	)
				C_Insp_Emp_Cd		 	=	iCurColumnPos(4	)
				C_Insp_Emp_Pop	 	=	iCurColumnPos(5	)
				C_Insp_Emp_Nm		 	=	iCurColumnPos(6	)
				C_Cust_Cd				 	=	iCurColumnPos(7	)
				C_Cust_Pop			 	=	iCurColumnPos(8	)
				C_Cust_Nm				 	=	iCurColumnPos(9	)
				C_Insp_Hour2		 	=	iCurColumnPos(10)
				C_Insp_Min2			 	=	iCurColumnPos(11)
				C_Payroll2			 	=	iCurColumnPos(12)

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
	Call InitComboBox2
	Call SetToolbar("1100111100011111")														'버튼	툴바 제어 
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		Set gActiveElement = document.activeElement  
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
	End If
	frm1.txtWork_Dt.text = StartDate
	gCounts	=	0
	isFirst	=	true
	Call CookiePage	(0)																															'☜: Check Cookie
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

		ggoSpread.Source = Frm1.vspdData1
		If	ggoSpread.SSCheckChange	=	True Then
		ChgOK	=	True
		End	If

		ggoSpread.Source = Frm1.vspdData
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
	IF ChkKeyField()=False Then Exit Function 

	Call InitVariables																													 '⊙:	Initializes	local	global variables
	lgCurrentSpd = "M"

	Call MakeKeyStream("X")

		gCounts	=	0
		isFirst	=	true

		lgCurrentSpd = "M"	'	Master

	Call	DisableToolBar(	parent.TBC_QUERY)

	If DbQuery = False Then
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
		dim	lRow
		DIM	strCD, strNm

		FncSave	=	False																															 '☜:	Processing is	NG
		Err.Clear

	frm1.ChgSave1.value	=	"F"
	frm1.ChgSave2.value	=	"F"
	frm1.ChgSave3.value	=	"F"

		ggoSpread.Source = frm1.vspdData
		If	ggoSpread.SSCheckChange	=	True Then
		frm1.ChgSave1.value	=	"T"
		End	If

		ggoSpread.Source = Frm1.vspdData1
		If	ggoSpread.SSCheckChange	=	True Then
		frm1.ChgSave2.value	=	"T"
		End	If

		ggoSpread.Source = Frm1.vspdData2
		If	ggoSpread.SSCheckChange	=	True Then
		frm1.ChgSave3.value	=	"T"
		End	If

	If frm1.ChgSave1.value = "F" and frm1.ChgSave2.value="F" and frm1.ChgSave3.value="F" Then
				IntRetCD =	DisplayMsgBox("900001","x","x","x")														'☜:There	is no	changed	data.
		Exit Function
	End	If

		ggoSpread.Source = frm1.vspdData
	If Not	ggoSpread.SSDefaultCheck Then																					'☜: Check contents	area
			 Exit	Function
	End	If

		ggoSpread.Source = frm1.vspdData1
	If Not	ggoSpread.SSDefaultCheck Then																					'☜: Check contents	area
			 Exit	Function
	End	If

		ggoSpread.Source = frm1.vspdData2
	If Not	ggoSpread.SSDefaultCheck Then																					'☜: Check contents	area
			 Exit	Function
	End	If

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
						If Frm1.vspdData.MaxRows < 1 Then
								Exit Function
						End	If

					With Frm1.vspdData

						If .ActiveRow	>	0	Then
							.ReDraw	=	False

							ggoSpread.Source = frm1.vspdData
							ggoSpread.CopyRow
										SetSpreadColor .ActiveRow, .ActiveRow

										.Col	=	C_BAS_AMT_TYPE
					.Row	=	.ActiveRow
					.Text	=	""

					.Col	=	C_BAS_AMT_TYPE_NM
					.Row	=	.ActiveRow
					.Text	=	""

							.ReDraw	=	True
							.focus
						End	If
					End	With
			ggoSpread.Source = Frm1.vspdData1
			ggoSpread.ClearSpreadData
			ggoSpread.Source = Frm1.vspdData2
			ggoSpread.ClearSpreadData

				Case	"S1"

						If Frm1.vspdData1.MaxRows	<	1	Then
								Exit Function
						End	If

					With Frm1.vspdData1

						If .ActiveRow	>	0	Then
							.ReDraw	=	False

							ggoSpread.Source = frm1.vspdData1
							ggoSpread.CopyRow
										SetSpreadColor1	.ActiveRow,	.ActiveRow

'											.Col	=	C_ALLOW_CD
'						.Row	=	.ActiveRow
'						.Text	=	""

'						.Col	=	C_ALLOW_CD_NM
'						.Row	=	.ActiveRow
'						.Text	=	""

							.ReDraw	=	True
							.focus
						End	If
					End	With
				Case	Else

						If Frm1.vspdData2.MaxRows	<	1	Then
								Exit Function
						End	If

					With Frm1.vspdData2

						If .ActiveRow	>	0	Then
							.ReDraw	=	False

							ggoSpread.Source = frm1.vspdData2
							ggoSpread.CopyRow
										SetSpreadColor2	.ActiveRow,	.ActiveRow

'											.Col	=	C_ALLOW_CD
'						.Row	=	.ActiveRow
'						.Text	=	""

'						.Col	=	C_ALLOW_CD_NM
'						.Row	=	.ActiveRow
'						.Text	=	""

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
		ggoSpread.Source = Frm1.vspdData
		ggoSpread.EditUndo
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		ggoSpread.Source = Frm1.vspdData
		frm1.vspdData.Row	=	frm1.vspdData.ActiveRow
		frm1.vspdData.Col	=	C_Plan_Dt
		frm1.hWork_Dt.value	=	frm1.vspdData.text
		frm1.vspdData.Col	=	C_FACILITY_CD
		frm1.hFacility_Cd.value	=	frm1.vspdData.text
		Call DbDtlQuery1()

	elseif lgCurrentSpd	=	"S1" then
		ggoSpread.Source = Frm1.vspdData1
		ggoSpread.EditUndo
	else
		ggoSpread.Source = Frm1.vspdData2
		ggoSpread.EditUndo
	end	if
	Call initdata()
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



'			If lgIntFlgMode	<> parent.OPMD_UMODE Then																			 'Check	if there is	retrived data
'			Call DisplayMsgBox("900002", "X",	"X", "X")																	 '☆:
'					Exit Function
'			End	If


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
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.ClearSpreadData
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.ClearSpreadData
			ggoSpread.Source = frm1.vspdData
									With Frm1
						.vspdData.ReDraw = False
						.vspdData.Focus
						ggoSpread.Source = .vspdData
						ggoSpread.InsertRow	.vspdData.ActiveRow, imRow
						SetSpreadColor .vspdData.ActiveRow,	.vspdData.ActiveRow	+	imRow	-	1

						.vspdData.ReDraw = True
									End	With
				Case	"S1"
									With Frm1
						.vspdData1.ReDraw	=	False
						.vspdData1.Focus
						ggoSpread.Source = .vspdData1
						ggoSpread.InsertRow	.vspdData1.ActiveRow,	imRow
						SetSpreadColor1	.vspdData1.ActiveRow,	.vspdData1.ActiveRow + imRow - 1
						iTemp	=	0
						For	iRow =	.vspdData1.ActiveRow to	.vspdData1.ActiveRow + imRow - 1
							iTemp	=	iTemp	+	1
							.vspdData1.Row = iRow
							.vspdData1.Col=	C_Seq
							.vspdData1.Text	=	.vspdData1.Maxrows + iTemp - imRow
						Next
						.vspdData1.ReDraw	=	True
									End	With
				Case	Else
									With Frm1
						.vspdData2.ReDraw	=	False
						.vspdData2.Focus
						ggoSpread.Source = .vspdData2
						ggoSpread.InsertRow	.vspdData2.ActiveRow,	imRow
						SetSpreadColor2	.vspdData2.ActiveRow,	.vspdData2.ActiveRow + imRow - 1
						iTemp	=	0
						For	iRow =	.vspdData2.ActiveRow to	.vspdData2.ActiveRow + imRow - 1
							iTemp	=	iTemp	+	1
							.vspdData2.Row = iRow
							.vspdData2.Col=	C_Seq
							.vspdData2.Text	=	.vspdData2.Maxrows + iTemp - imRow
						Next
						.vspdData2.ReDraw	=	True
									End	With
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
				If Frm1.vspdData.MaxRows < 1 then
					 Exit	function
			End	if
				With Frm1.vspdData
					.focus
					 ggoSpread.Source	=	frm1.vspdData
					lDelRows =	ggoSpread.DeleteRow
				End	With
		ELSEif lgCurrentSpd	=	"S1" then
				If Frm1.vspdData1.MaxRows	<	1	then
					 Exit	function
			End	if
				With Frm1.vspdData1
					.focus
					 ggoSpread.Source	=	frm1.vspdData1
					lDelRows =	ggoSpread.DeleteRow
				End	With
		ELSE
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
		Select Case	gActiveSpdSheet.id
		Case "vaSpread"
			Call InitSpreadSheet("A")
		Case "vaSpread1"
			Call InitSpreadSheet("B")
			Call InitComboBox
		Case "vaSpread2"
			Call InitSpreadSheet("C")
			Call InitComboBox2
	End	Select
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End	Sub
'========================================================================================================
'	Function Name	:	FncExit
'	Function Desc	:
'========================================================================================================
Function FncExit()
		Dim	IntRetCD

	FncExit	=	False

		 ggoSpread.Source	=	frm1.vspdData
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
		strVal = BIZ_PGM_ID	&	"?txtMode="			&	parent.UID_M0001
				strVal = strVal			&	"&lgCurrentSpd="		&	lgCurrentSpd											'☜: Next	key	tag
				strVal = strVal			&	"&txtKeyStream="		&	lgKeyStream												'☜: Query Key
			strVal = strVal			&	"&txtWork_Dt="		&	Frm1.txtWork_Dt.text										 '☜:	Query	Key
			strVal = strVal			&	"&txtFacility_Cd="	&	Frm1.txtFacility_Cd.value			 '☜:	Query	Key
			strVal = strVal			&	"&txtPlantCd="		&	Frm1.txtPlantCd.value			 '☜:	Query	Key
			strVal = strVal			&	"&CboFacility_Accnt="	&	Frm1.CboFacility_Accnt.value			'☜: Query Key

				strVal = strVal			&	"&txtMaxRows="		&	.vspdData.MaxRows
				strVal = strVal			&	"&lgStrPrevKey="		&	lgStrPrevKey								 '☜:	Next key tag
				strVal = strVal			&	"&lgStrPrevKey="		&	lgStrPrevKey								 '☜:	Next key tag
		strVal = strVal			&	"&lgPageNo_A="		&	lgPageNo_A													'☜: Next	key	tag
		strVal = strVal			&	"&txtType="			&	"A"													 '☜:	Next key tag
		End	With

	Call RunMyBizASP(MyBizASP, strVal)																							 '☜:	Run	Biz	Logic
		DbQuery	=	True
End	Function

'========================================================================================================
'	Name : DbDtlQuery1
'	Desc : This	function is	called by	FncQuery
'========================================================================================================

Function DbDtlQuery1()

		DbDtlQuery1	=	False

		Err.Clear																																				 '☜:	Clear	err	status

	If LayerShowHide(1)	=	False	Then
		Exit Function
	End	If

	Dim	strVal

		With Frm1
		strVal = BIZ_PGM_ID	&	"?txtMode="			&	parent.UID_M0001
				strVal = strVal			&	"&lgCurrentSpd="		&	"S1"											'☜: Next	key	tag
				strVal = strVal			&	"&txtKeyStream="		&	lgKeyStream												'☜: Query Key
			strVal = strVal			&	"&txtWork_Dt="		&	Frm1.hWork_Dt.value											'☜: Query Key
			strVal = strVal			&	"&txtFacility_Cd="	&	Frm1.hFacility_Cd.value			 '☜:	Query	Key
			strVal = strVal			&	"&txtPlantCd="		&	""
			strVal = strVal			&	"&CboFacility_Accnt="	&	""
				strVal = strVal			&	"&txtMaxRows="		&	.vspdData1.MaxRows
				strVal = strVal			&	"&lgStrPrevKey=" 		&	lgStrPrevKey1									'☜: Next	key	tag
		strVal = strVal			&	"&lgPageNo_B="		&	lgPageNo_B													'☜: Next	key	tag
		strVal = strVal			&	"&txtType="			&	"B"													 '☜:	Next key tag
		End	With



	Call RunMyBizASP(MyBizASP, strVal)																							 '☜:	Run	Biz	Logic

		DbDtlQuery1	=	True
End	Function

'========================================================================================================
'	Name : DbDtlQuery2
'	Desc : This	function is	called by	FncQuery
'========================================================================================================

Function DbDtlQuery2()

		DbDtlQuery2	=	False

		Err.Clear																																				 '☜:	Clear	err	status

	If LayerShowHide(1)	=	False	Then
		Exit Function
	End	If

	Dim	strVal

		With Frm1
		strVal = BIZ_PGM_ID	&	"?txtMode="			&	parent.UID_M0001
				strVal = strVal			&	"&lgCurrentSpd="		&	"S2"											'☜: Next	key	tag
				strVal = strVal			&	"&txtKeyStream="		&	lgKeyStream												'☜: Query Key
			strVal = strVal			&	"&txtWork_Dt="		&	Frm1.hWork_Dt.value											'☜: Query Key
			strVal = strVal			&	"&txtFacility_Cd="	&	Frm1.hFacility_Cd.value			 '☜:	Query	Key
			strVal = strVal			&	"&txtPlantCd="		&	""
			strVal = strVal			&	"&CboFacility_Accnt="	&	""
				strVal = strVal			&	"&txtMaxRows="		&	.vspdData2.MaxRows
				strVal = strVal			&	"&lgStrPrevKey=" 		&	lgStrPrevKey2									'☜: Next	key	tag
		strVal = strVal			&	"&lgPageNo_C="		&	lgPageNo_C													'☜: Next	key	tag
		strVal = strVal			&	"&txtType="			&	"C"													 '☜:	Next key tag		End	With
		End	With

	Call RunMyBizASP(MyBizASP, strVal)																							 '☜:	Run	Biz	Logic

		DbDtlQuery2	=	True
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
		ggoSpread.Source = frm1.vspdData
			With Frm1
					 For lRow	=	1	To .vspdData.MaxRows
							 .vspdData.Row = lRow
							 .vspdData.Col = 0
							 Select	Case .vspdData.Text

									 Case	 ggoSpread.InsertFlag																			 '☜:	Update
																													strVal = strVal	&	"C"	&	parent.gColSep
																													strVal = strVal	&	lRow & parent.gColSep
																													strVal = strVal	&	"M"	&	parent.gColSep
												.vspdData.Col	=	C_FACILITY_CD		:	strVal = strVal	&				Trim(.vspdData.Text)	&	parent.gColSep
												.vspdData.Col	=	C_Plan_Dt			:	strVal = strVal	&	UNIConvDate(Trim(.vspdData.Text))		&	parent.gColSep
												.vspdData.Col	=	C_Insp_Text		:	strVal = strVal	&							Trim(.vspdData.Text) 	&	parent.gColSep
												.vspdData.Col	=	C_Insp_Hour		:	strVal = strVal	&	UNIConvNum (Trim(.vspdData.Text),0)	&	parent.gColSep
												.vspdData.Col	=	C_Insp_Min		:	strVal = strVal	&	UNIConvNum (Trim(.vspdData.Text),0)	&	parent.gColSep
												.vspdData.Col	=	C_Req_Dept		:	strVal = strVal	&				Trim(.vspdData.Text)	&	parent.gColSep
												.vspdData.Col	=	C_Insp_Dept		:	strVal = strVal	&				Trim(.vspdData.Text)	&	parent.gColSep
												.vspdData.Col	=	C_Insp_Emp_Qty	:	strVal = strVal	&	UNIConvNum (Trim(.vspdData.Text),0)	&	parent.gColSep
												.vspdData.Col	=	C_Payroll		:	strVal = strVal	&	UNIConvNum (Trim(.vspdData.Text),0)	&	parent.gColSep
												.vspdData.Col	=	C_Matl_Cost		:	strVal = strVal	&	UNIConvNum (Trim(.vspdData.Text),0)	&	parent.gColSep
												.vspdData.Col	=	C_Insp_FlagNm	:	strVal = strVal	&				Trim(.vspdData.Text) 	&	parent.gRowSep
												lGrpCnt	=	lGrpCnt	+	1

									 Case	 ggoSpread.UpdateFlag																			 '☜:	Update
																													strVal = strVal	&	"U"	&	parent.gColSep
																													strVal = strVal	&	lRow & parent.gColSep
																													strVal = strVal	&	"M"	&	parent.gColSep
												.vspdData.Col	=	C_FACILITY_CD		:	strVal = strVal	&				Trim(.vspdData.Text)	&	parent.gColSep
												.vspdData.Col	=	C_Plan_Dt			:	strVal = strVal	&	UNIConvDate(Trim(.vspdData.Text))		&	parent.gColSep
												.vspdData.Col	=	C_Insp_Text		:	strVal = strVal	&							Trim(.vspdData.Text) 	&	parent.gColSep
												.vspdData.Col	=	C_Insp_Hour		:	strVal = strVal	&	UNIConvNum (Trim(.vspdData.Text),0)	&	parent.gColSep
												.vspdData.Col	=	C_Insp_Min		:	strVal = strVal	&	UNIConvNum (Trim(.vspdData.Text),0)	&	parent.gColSep
												.vspdData.Col	=	C_Req_Dept		:	strVal = strVal	&				Trim(.vspdData.Text)	&	parent.gColSep
												.vspdData.Col	=	C_Insp_Dept		:	strVal = strVal	&				Trim(.vspdData.Text)	&	parent.gColSep
												.vspdData.Col	=	C_Insp_Emp_Qty	:	strVal = strVal	&	UNIConvNum (Trim(.vspdData.Text),0)	&	parent.gColSep
												.vspdData.Col	=	C_Payroll		:	strVal = strVal	&	UNIConvNum (Trim(.vspdData.Text),0)	&	parent.gColSep
												.vspdData.Col	=	C_Matl_Cost		:	strVal = strVal	&	UNIConvNum (Trim(.vspdData.Text),0)	&	parent.gColSep
												.vspdData.Col	=	C_Insp_FlagNm	:	strVal = strVal	&				Trim(.vspdData.Text) 	&	parent.gRowSep
												lGrpCnt	=	lGrpCnt	+	1
									 Case	 ggoSpread.DeleteFlag																			 '☜:	Delete

																													strDel = strDel	&	"D"	&	parent.gColSep
																													strDel = strDel	&	lRow & parent.gColSep
																													strDel = strDel	&	"M"	&	parent.gColSep
												.vspdData.Col	=	C_FACILITY_CD		:	strDel = strDel	&				Trim(.vspdData.Text)	&	parent.gColSep
												.vspdData.Col	=	C_Plan_Dt			:	strDel = strDel	&	UNIConvDate(Trim(.vspdData.Text))	&	parent.gRowSep
												lGrpCnt	=	lGrpCnt	+	1
							 End Select
					 Next

					 .txtMode.value				 =	parent.UID_M0002
				 .txtMaxRows.value		 = lGrpCnt-1
				 .txtSpread.value			 = strDel	&	strVal
			End	With
	end	if

		if	frm1.ChgSave2.value	=	"T"	then
		ggoSpread.Source = frm1.vspdData1
			With Frm1
					 For lRow	=	1	To .vspdData1.MaxRows
							 .vspdData1.Row	=	lRow
							 .vspdData1.Col	=	0
							 Select	Case .vspdData1.Text
									 Case	 ggoSpread.InsertFlag																			 '☜:	Create
																													strVal = strVal	&	"C"	&	parent.gColSep
																													strVal = strVal	&	lRow & parent.gColSep
																													strVal = strVal	&	"S1" & parent.gColSep
						frm1.vspdData.Row	=	frm1.vspdData.ActiveRow
												.vspdData.Col	=	C_FACILITY_CD		:	strVal = strVal	&	Trim(.vspdData.Text) & parent.gColSep
												.vspdData.Col	=	C_Plan_Dt			:	strVal = strVal	&	UNIConvDate(Trim(.vspdData.Text))	&	parent.gColSep
												.vspdData1.Col = C_Seq 			:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Zinsp_PartCd	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Insp_PartCd 	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Insp_MethCd 	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Insp_DeciCd 	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_St_GoGubunCd	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Sury_Assy	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_S_Qty		:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Price		:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Sury_Amt		:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Cur			:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Sury_Type	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gRowSep
												lGrpCnt	=	lGrpCnt	+	1
									 Case	 ggoSpread.UpdateFlag																			 '☜:	Update
																													strVal = strVal	&	"U"	&	parent.gColSep
																													strVal = strVal	&	lRow & parent.gColSep
																													strVal = strVal	&	"S1" & parent.gColSep
						frm1.vspdData.Row	=	frm1.vspdData.ActiveRow
												.vspdData.Col	=	C_FACILITY_CD		:	strVal = strVal	&	Trim(.vspdData.Text) & parent.gColSep
												.vspdData.Col	=	C_Plan_Dt			:	strVal = strVal	&	UNIConvDate(Trim(.vspdData.Text))	&	parent.gColSep
												.vspdData1.Col = C_Seq 			:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Zinsp_PartCd	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Insp_PartCd 	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Insp_MethCd 	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Insp_DeciCd 	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_St_GoGubunCd	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Sury_Assy	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_S_Qty		:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Price		:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Sury_Amt		:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Cur			:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gColSep
						.vspdData1.Col = C_Sury_Type	:	strVal = strVal	&				Trim(.vspdData1.Text)	&	parent.gRowSep
												lGrpCnt	=	lGrpCnt	+	1

									 Case	 ggoSpread.DeleteFlag																			 '☜:	Delete
																													strDel = strDel	&	"D"	&	parent.gColSep
																													strDel = strDel	&	lRow & parent.gColSep
																													strDel = strDel	&	"S1" & parent.gColSep
												frm1.vspdData.Row	=	frm1.vspdData.ActiveRow
												.vspdData.Col	=	C_FACILITY_CD		:	strDel = strDel	&	Trim(.vspdData.Text) & parent.gColSep
												.vspdData.Col	=	C_Plan_Dt			:	strDel = strDel	&	UNIConvDate(Trim(.vspdData.Text))	&	parent.gColSep
												.vspdData1.Col = C_Seq 			:	strDel = strDel	&	UNIConvNum(	Trim(.vspdData1.Text),0) & parent.gRowSep
												lGrpCnt	=	lGrpCnt	+	1
							 End Select
					 Next

					 .txtMode.value				 =	parent.UID_M0002
				 .txtMaxRows.value		 = lGrpCnt-1
				 .txtSpread.value			 = strDel	&	strVal
			End	With
		end	if

		if	frm1.ChgSave3.value	=	"T"	then
		ggoSpread.Source = frm1.vspdData2
			With Frm1
					 For lRow	=	1	To .vspdData2.MaxRows
							 .vspdData2.Row	=	lRow
							 .vspdData2.Col	=	0
							 Select	Case .vspdData2.Text
									 Case	 ggoSpread.InsertFlag																			 '☜:	Create
																													strVal = strVal	&	"C"	&	parent.gColSep
																													strVal = strVal	&	lRow & parent.gColSep
																													strVal = strVal	&	"S2" & parent.gColSep
						frm1.vspdData.Row	=	frm1.vspdData.ActiveRow
												.vspdData.Col	=	C_FACILITY_CD		:	strVal = strVal	&	Trim(.vspdData.Text) & parent.gColSep
												.vspdData.Col	=	C_Plan_Dt			:	strVal = strVal	&	UNIConvDate(Trim(.vspdData.Text))	&	parent.gColSep
												.vspdData2.Col = C_Seq 			:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData2.Text),0) & parent.gColSep
												.vspdData2.Col = C_Insp_Emp_Gb	:	strVal = strVal	&							Trim(.vspdData2.Text)		&	parent.gColSep
						.vspdData2.Col = C_Insp_Emp_Cd	:	strVal = strVal	&				Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Cust_Cd	 	:	strVal = strVal	&				Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Insp_Hour2	:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData2.Text),0) & parent.gColSep
						.vspdData2.Col = C_Insp_Min2	:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData2.Text),0) & parent.gColSep
						.vspdData2.Col = C_Payroll2		:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData2.Text),0) & parent.gRowSep


												lGrpCnt	=	lGrpCnt	+	1
									 Case	 ggoSpread.UpdateFlag																			 '☜:	Update
																													strVal = strVal	&	"U"	&	parent.gColSep
																													strVal = strVal	&	lRow & parent.gColSep
																													strVal = strVal	&	"S2" & parent.gColSep
						frm1.vspdData.Row	=	frm1.vspdData.ActiveRow
												.vspdData.Col	=	C_FACILITY_CD		:	strVal = strVal	&	Trim(.vspdData.Text) & parent.gColSep
												.vspdData.Col	=	C_Plan_Dt			:	strVal = strVal	&	UNIConvDate(Trim(.vspdData.Text))	&	parent.gColSep
												.vspdData2.Col = C_Seq 			:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData2.Text),0) & parent.gColSep
												.vspdData2.Col = C_Insp_Emp_Gb	:	strVal = strVal	&							Trim(.vspdData2.Text)		&	parent.gColSep
						.vspdData2.Col = C_Insp_Emp_Cd	:	strVal = strVal	&				Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Cust_Cd	 	:	strVal = strVal	&				Trim(.vspdData2.Text)	&	parent.gColSep
						.vspdData2.Col = C_Insp_Hour2	:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData2.Text),0) & parent.gColSep
						.vspdData2.Col = C_Insp_Min2	:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData2.Text),0) & parent.gColSep
						.vspdData2.Col = C_Payroll2		:	strVal = strVal	&	UNIConvNum(	Trim(.vspdData2.Text),0) & parent.gRowSep
												lGrpCnt	=	lGrpCnt	+	1

									 Case	 ggoSpread.DeleteFlag																			 '☜:	Delete
																													strDel = strDel	&	"D"	&	parent.gColSep
																													strDel = strDel	&	lRow & parent.gColSep
																													strDel = strDel	&	"S2" & parent.gColSep
												frm1.vspdData.Row	=	frm1.vspdData.ActiveRow
												.vspdData.Col	=	C_FACILITY_CD		:	strDel = strDel	&	Trim(.vspdData.Text) & parent.gColSep
												.vspdData.Col	=	C_Plan_Dt			:	strDel = strDel	&	UNIConvDate(Trim(.vspdData.Text))	&	parent.gColSep
												.vspdData2.Col = C_Seq 			:	strDel = strDel	&	UNIConvNum(	Trim(.vspdData2.Text),0) & parent.gColSep
												.vspdData2.Col = C_Insp_Emp_Gb	:	strDel = strDel	&							Trim(.vspdData2.Text)		&	parent.gRowSep
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
	lgOldRow_A = 0
	lgOldRow_B = 0
	lgOldRow_C = 0
		lgIntFlgMode =	parent.OPMD_UMODE
		Call	ggoOper.LockField(Document,	"Q")										'⊙: Lock	field
		Call InitData()

		Call SetToolbar("1100111100011111")

'		if lgStrPrevKey1 <>	"" and isFirst = false then
'			exit function
'		end	if

'		if lgStrPrevKey1 <>	"" or	isFirst	=	true then
		isFirst	=	false		'	첫화면이 열리고나서	오른쪽 그리드	세팅하기 위해 
		Call DisableToolBar(parent.TBC_QUERY)
		call vspdData_click(1,frm1.vspdData.activerow)
'		end	if
	frm1.vspdData.focus
End	Function

Function DbDtlQueryOk1()
		lgIntFlgMode =	parent.OPMD_UMODE

	Call InitData()
		Call SetToolbar("1100111100011111")

		Call	ggoOper.LockField(Document,	"Q")
		Set	gActiveElement = document.ActiveElement
'	frm1.vspdData1.focus
End	Function

Function DbDtlQueryOk2()
		lgIntFlgMode =	parent.OPMD_UMODE

	Call InitData()
		Call SetToolbar("1100111100011111")

		Call	ggoOper.LockField(Document,	"Q")
		Set	gActiveElement = document.ActiveElement
'	frm1.vspdData2.focus
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




'========================================================================================================
'	Name : OpenEmptName()
'	Desc : developer describe	this line
'========================================================================================================

Function OpenEmptName(iWhere)
	Dim	arrRet
	Dim	arrParam(2)

	If IsOpenPop = True	Then Exit	Function

	IsOpenPop	=	True

	If iWhere	=	1	Then 'TextBox(Condition)
		frm1.vspdData2.Row = frm1.vspdData2.activerow
			frm1.vspdData2.Col = C_Insp_Emp_Cd
		arrParam(0)	=	frm1.vspdData2.Text			'	Code Condition
				frm1.vspdData2.Col = C_Insp_Emp_Nm
			arrParam(1)	=	""'frm1.vspdData2.Text			'	Name Cindition
			arrParam(2)	=	lgUsrIntCd										'	자료권한 Condition
	End	If

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"),	Array(window.parent,arrParam), _
		"dialogWidth=760px;	dialogHeight=420px;	center:	Yes; help: No; resizable:	No;	status:	No;")

	IsOpenPop	=	False

	If arrRet(0) = ""	Then
		If iWhere	=	0	Then
			frm1.C_Insp_Emp_Cd.focus
		Else
			frm1.vspdData2.Col = C_Insp_Emp_Cd
			frm1.vspdData2.action	=0
		End	If
		Exit Function
	Else
		Call SubSetCondEmp(arrRet, iWhere)
	End	If

End	Function

'======================================================================================================
'	Name : SubSetCondEmp()
'	Description	:	Item Popup에서 Return되는	값 setting
'=======================================================================================================
Sub	SubSetCondEmp(Byval	arrRet,	Byval	iWhere)

	With frm1
		If iWhere	=	0	Then 'TextBox(Condition)

		Else 'spread
				ggoSpread.Source = Frm1.vspdData2
			frm1.vspdData2.row = frm1.vspdData2.activerow

			.vspdData2.Col = C_Insp_Emp_Cd
			.vspdData2.Text	=	arrRet(0)
			.vspdData2.Col = C_Insp_Emp_Nm
			.vspdData2.Text	=	arrRet(1)
			.vspdData2.action	=0

			ggoSpread.Source = frm1.vspdData2
			ggoSpread.UpdateRow	frm1.vspdData2.ActiveRow

		End	If
	End	With
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
	strWhere = " SET_PLANT LIKE " & FilterVar(Trim(strPlant), "''", "S")

	If IsOpenPop = True	 Then
		Exit Function
	End	If

	IsOpenPop	=	True
	Select Case	iWhere
		Case "1"
			arrParam(0)	=	"설비코드	팝업"
			arrParam(1)	=	"Y_FACILITY"
			arrParam(2)	=	frm1.txtFacility_Cd.value
			arrParam(3)	=	""												'	Name Cindition
			arrParam(4)	=	strWhere										'	Where	Condition
			arrParam(5)	=	"설비코드"												'	TextBox	명칭 

			arrField(0)	=	"Facility_cd"											'	Field명(0)
			arrField(1)	=	"Facility_Nm"											'	Field명(1)

			arrHeader(0) = "설비코드"									'	Header명(0)
			arrHeader(1) = "설비코드명"									'	Header명(1)
		Case "2"
			arrParam(0)	=	"설비코드	팝업"
			arrParam(1)	=	"Y_FACILITY"
			frm1.vspdData.Col	=	C_FACILITY_CD
			arrParam(2)	=	frm1.vspdData.text
			arrParam(3)	=	""												'	Name Cindition
			arrParam(4)	=	strWhere										'	Where	Condition
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
			 Frm1.txtFacility_Cd.focus
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
			Case "2"
				frm1.vspdData.Col	=	C_FACILITY_CD
				frm1.vspdData.text = arrRet(0)
				frm1.vspdData.Col	=	C_FACILITY_NM
				frm1.vspdData.text = arrRet(1)

				Call vspdData_Change(C_FACILITY_CD,	frm1.vspdData.ActiveRow)

		End	Select
	End	With
End	Sub


'------------------------------------------	 OpenBp()	 ---------------------------------------
'	Name : OpenBp()
'	Description	:	OpenAp Popup에서 Return되는	값 setting
'---------------------------------------------------------------------------------------------------------
Function OpenBp(Byval	strCode, byval iWhere)
	Dim	arrRet
	Dim	arrParam(5)

	If IsOpenPop = True	Then Exit	Function

	IsOpenPop	=	True
	arrParam(0)	=	strCode								'	 Code	Condition
	 	arrParam(1)	=	""							'	채권과 연계(거래처 유무)
	arrParam(2)	=	""								'FrDt
	arrParam(3)	=	""								'ToDt
	arrParam(4)	=	"T"							'B :매출 S:	매입 T:	전체 
	arrParam(5)	=	""									'SUP :공급처 PAYTO:	지급처 SOL:주문처	PAYER	:수금처	INV:세금계산 

	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam),	_
		"dialogWidth=780px;	dialogHeight=450px;	:	Yes; help: No; resizable:	No;	status:	No;")


	IsOpenPop	=	False

	If arrRet(0) = ""	Then

		Call EScCode(iwhere)
		Exit Function
	Else
			ggoSpread.Source = Frm1.vspdData2
		frm1.vspdData2.row = frm1.vspdData2.activerow

		Call SetBp(arrRet, iWhere)
		if iWhere	<> 0 then
			ggoSpread.UpdateRow	frm1.vspdData2.ActiveRow
				End	if
	End	If
End	Function

'========================================================================================================
'	Name : SetBp()
'	Description	:	Item Popup에서 Return되는	값 setting
'========================================================================================================
Function SetBp(Byval arrRet, Byval iWhere)

	With frm1
		Select Case	iWhere
				Case 3
					ggoSpread.Source = Frm1.vspdData2
				frm1.vspdData2.row = frm1.vspdData2.activerow

						.vspdData2.Col		=	C_Cust_Cd
					.vspdData2.text		=	arrRet(0)
					.vspdData2.Col		=	C_Cust_Nm
					.vspdData2.text		=	arrRet(1)
					Call SetActiveCell(.vspdData2,C_Cust_Cd,.vspdData2.ActiveRow ,"M","X","X")
				End	Select

	End	With

End	Function


'===========================================================================
Function  OpenCur(ByVal strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

    ggoSpread.Source = frm1.vspdData1                                   
 
	frm1.vspdData1.Col = 0
	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	 
	If frm1.vspdData1.Text <> ggoSpread.InsertFlag Then Exit Function 
	 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "화폐"      <%' 팝업 명칭 %>
	arrParam(1) = "B_CURRENCY"       <%' TABLE 명칭 %>
	arrParam(2) = Trim(strCode)       <%' Code Condition%>
	arrParam(3) = ""         <%' Name Cindition%>
	arrParam(4) = ""         <%' Where Condition%>
	arrParam(5) = "화폐"      <%' TextBox 명칭 %>

	arrField(0) = "CURRENCY"       <%' Field명(0)%>
	arrField(1) = "CURRENCY_DESC"      <%' Field명(1)%>

	arrHeader(0) = "화폐"      <%' Header명(0)%>
	arrHeader(1) = "화폐명"      <%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData1.Col = C_Cur
		frm1.vspdData1.Text = arrRet(0)
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.UpdateRow	frm1.vspdData1.ActiveRow
	End If

End Function

'------------------------------------------	 OpenSItem()	-------------------------------------------------
'	Name : OpenSItem()
'	Description	:	SpreadItem PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSItem(byval strCon)
	If IsOpenPop = True	Then Exit	Function

	Dim	arrRet
	Dim	arrParam(5), arrField(6)
	Dim	iCalledAspName
	Dim IntRetCd

	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
		Exit Function
	End If

	IsOpenPop	=	True

	With frm1.vspdData1

		.Row = .ActiveRow

		arrParam(0) = Trim(frm1.txtPlantCd.value)	' Item Code
		arrParam(1)	=	Trim(strCon)
		arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
		arrParam(3) = ""							' Default Value

	    arrField(0) = 1 							' Field명(0) : "ITEM_CD"
	    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
	    arrField(2) = 13 							' Field명(1) : "ITEM_NM"

	End	With

	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If


	arrRet = window.showModalDialog(iCalledAspName,	Array(window.parent, arrParam,arrField), _
		"dialogWidth=760px;	dialogHeight=420px;	center:	Yes; help: No; resizable:	No;	status:	No;")

	IsOpenPop	=	False

	If arrRet(0) = ""	Then
		Exit Function
	Else
		With frm1.vspdData1
			IntRetCd = CommonQueryRs(" STD_PRC "," I_MATERIAL_VALUATION "," ITEM_CD	=	"	&		FilterVar(Trim(arrRet(0)), "''", "S")	&	"	",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			If IntRetCd	=	false	then
				Frm1.vspdData1.Col	=	C_Price
				Frm1.vspdData1.value = 0
			ELSE
				Frm1.vspdData1.Col	=	C_Price
				Frm1.vspdData1.value = UNICDbl(Trim(Replace(lgF0,Chr(11),"")))
			END	IF
			Call vspdData1_Change(C_Price, frm1.vspdData1.ActiveRow)
			.Row = .ActiveRow
			.Col = C_Sury_Assy_Nm
			.text	=	arrRet(1)
			.Col = C_Sury_Assy
			.text	=	arrRet(0)
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.UpdateRow	frm1.vspdData1.ActiveRow


			.focus
		End	With
	End	If
End	Function



'========================================================================================================
'	Name : EScCode()
'	Description	:	Item Popup에서 Return되는	값 setting
'========================================================================================================
Function EScCode(Byval iWhere)

	With frm1

		Select Case	iWhere
				Case 3
					Call SetActiveCell(.vspdData2,C_Cust_Cd,.vspdData2.ActiveRow ,"M","X","X")
				End	Select

	End	With

End	Function
'======================================================================================================
'	Name : OpenCode()
'	Description	:
'=======================================================================================================

Function OpenCode(Byval	strCode, Byval iWhere, ByVal Row)
	Dim	arrRet
	Dim	arrParam(5), arrField(6),	arrHeader(6)

	If IsOpenPop = True	Then Exit	Function

	IsOpenPop	=	True

	Select Case	iWhere
			Case C_Insp_Dept_POP,	C_Req_Dept_POP
					arrParam(0)	=	"부서코드 팝업"							'	팝업 명칭 
				arrParam(1)	=	"H_CURRENT_DEPT"								'	TABLE	명칭 
				arrParam(2)	=	strCode																	'	Code Condition
				arrParam(3)	=	""									'	Name Cindition
				arrParam(4)	=	""														'	Where	Condition
				arrParam(5)	=	"부서코드" 									'	TextBox	명칭 

				arrField(0)	=	"dept_cd"									'	Field명(0)
				arrField(1)	=	"dept_nm"		 								'	Field명(1)

				arrHeader(0) = "부서코드"		 							'	Header명(0)
				arrHeader(1) = "부서코드명"									'	Header명(1)
	End	Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",	Array(arrParam,	arrField,	arrHeader),	_
		"dialogWidth=420px;	dialogHeight=450px;	center:	Yes; help: No; resizable:	No;	status:	No;")

	IsOpenPop	=	False

	If arrRet(0) = ""	Then
		 	frm1.vspdData.action=0
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow	Row
	End	If

End	Function

'======================================================================================================
'	Name : SetCode()
'	Description	:
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case	iWhere
				Case C_Insp_Dept_POP
					.vspdData.Col	=	C_Insp_Dept_NM
					.vspdData.text = arrRet(1)
						.vspdData.Col	=	C_Insp_Dept
					.vspdData.text = arrRet(0)
					.vspdData.action=0
				Case C_Req_Dept_POP
					.vspdData.Col	=	C_Req_Dept_NM
					.vspdData.text = arrRet(1)
						.vspdData.Col	=	C_Req_Dept
					.vspdData.text = arrRet(0)
					.vspdData.action=0
				End	Select
	End	With
End	Function

'========================================================================================================
'		Event	Name : vspdData_Change
'		Event	Desc :
'========================================================================================================
Sub	vspdData_Change(ByVal	Col	,	ByVal	Row	)
	Dim	intIndex
	Dim	IntRetCd,	iCodeArr,	iNameArr
	Dim	strFac

 	Frm1.vspdData.Row	=	Row
 	Frm1.vspdData.Col	=	Col

	Select Case	Col
		Case	 C_Facility_Cd
			strFac = Frm1.vspdData.text
			IntRetCd = CommonQueryRs(" facility_nm,	 dbo.ufn_GetCodeName('Z410', FACILITY_ACCNT) FACILITY_ACCNT_nm , Plant_nm "," Y_FACILITY LEFT OUTER JOIN B_PLANT ON Set_Plant = PLANT_CD "," facility_cd	=	"	&		FilterVar(Trim(UCase(strFac)), "''", "S")	&	"	",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			If IntRetCd	=	false	then
				Call DisplayMsgBox("970000","X","설비","X")
				Frm1.vspdData.text	= ""
				Frm1.vspdData.Col	=	C_Facility_Nm
				Frm1.vspdData.text	= ""
				Frm1.vspdData.Col	=	C_Facility_Accnt_Nm
				Frm1.vspdData.text	= ""
				Frm1.vspdData.Col	=	C_Set_PlantNm
				Frm1.vspdData.text	= ""
			ELSE
				Frm1.vspdData.Col	=	C_Facility_Nm
				Frm1.vspdData.text = Trim(Replace(lgF0,Chr(11),""))
				Frm1.vspdData.Col	=	C_Facility_Accnt_Nm
				Frm1.vspdData.text = Trim(Replace(lgF1,Chr(11),""))
				Frm1.vspdData.Col	=	C_Set_PlantNm
				Frm1.vspdData.text = Trim(Replace(lgF2,Chr(11),""))
			END	IF
	End	Select

 	If Frm1.vspdData.CellType	=	 parent.SS_CELL_TYPE_FLOAT Then
		If	UNICDbl(Frm1.vspdData.text)	<	 UNICDbl(Frm1.vspdData.TypeFloatMin) Then
			 Frm1.vspdData.text	=	Frm1.vspdData.TypeFloatMin
		End	If
	End	If

	ggoSpread.Source	=	frm1.vspdData
	ggoSpread.UpdateRow Row

	lgCurrentSpd = "M"

End	Sub

Sub	vspdData1_Change(ByVal Col , ByVal Row )
	Dim	intIndex
	Dim	Qty, Price,	DocAmt
	
	Frm1.vspdData1.Row = Row
	Frm1.vspdData1.Col = Col

	Select Case	Col
		Case	C_Zinsp_PartNm
			Frm1.vspdData1.col = C_Zinsp_PartNm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Zinsp_PartCd
			Frm1.vspdData1.value = intindex
		Case	C_Insp_PartNm
			Frm1.vspdData1.col = C_Insp_PartNm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Insp_PartCd
			Frm1.vspdData1.value = intindex
		Case	C_Insp_MethNm
			Frm1.vspdData1.col = C_Insp_MethNm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Insp_MethCd
			Frm1.vspdData1.value = intindex
		Case	C_Insp_DeciNm
			Frm1.vspdData1.col = C_Insp_DeciNm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Insp_DeciCd
			Frm1.vspdData1.value = intindex
		Case	C_St_GoGubunNm
			Frm1.vspdData1.col = C_St_GoGubunNm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_St_GoGubunCd
			Frm1.vspdData1.value = intindex
		Case	C_Sury_TypeNm
			Frm1.vspdData1.col = C_Sury_TypeNm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Sury_Type
			Frm1.vspdData1.value = intindex
		Case C_S_Qty,	C_Price
			frm1.vspdData1.Col = C_S_Qty
			Qty	=	UNICDbl(frm1.vspdData1.Text)
			frm1.vspdData1.Col = C_Price
			Price	=	UNICDbl(frm1.vspdData1.Text)

			DocAmt = Qty * Price
			frm1.vspdData1.Col = C_Sury_Amt
			frm1.vspdData1.Text	=	UNIConvNumPCToCompanyByCurrency(CStr(DocAmt),		parent.gCurrency,	parent.ggAmtOfMoneyNo, "X",	"X")
	End	Select


 	If Frm1.vspdData1.CellType =	parent.SS_CELL_TYPE_FLOAT	Then
		If	UNICDbl(Frm1.vspdData1.text) <	UNICDbl(Frm1.vspdData1.TypeFloatMin) Then
			 Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
		End	If
	End	If

	 ggoSpread.Source	=	frm1.vspdData1
	 ggoSpread.UpdateRow Row

	lgCurrentSpd = "S1"

End	Sub

Sub	vspdData2_Change(ByVal Col , ByVal Row )
	Dim	iDx

 	Frm1.vspdData2.Row = Row
 	Frm1.vspdData2.Col = Col

	Select Case	Col
			 Case	 C_Insp_Emp_GbNm
					iDx	=	Frm1.vspdData2.value
					Frm1.vspdData2.Col = C_Insp_Emp_Gb
					Frm1.vspdData2.value = iDx
			 Case	Else
	End	Select

 	If Frm1.vspdData2.CellType =	parent.SS_CELL_TYPE_FLOAT	Then
		If	UNICDbl(Frm1.vspdData2.text) <	UNICDbl(Frm1.vspdData2.TypeFloatMin) Then
			 Frm1.vspdData2.text = Frm1.vspdData2.TypeFloatMin
		End	If
	End	If

	 ggoSpread.Source	=	frm1.vspdData2
	 ggoSpread.UpdateRow Row

	lgCurrentSpd = "S2"

End	Sub
'========================================================================================================
'		Event	Name : vspdData_Click
'		Event	Desc : 컬럼을	클릭할 경우	발생 
'========================================================================================================
Sub	vspdData_Click(ByVal Col,	ByVal	Row)
	Dim	flagTxt
		Call SetPopupMenuItemInf("1101111111")

	IF lgBlnFlgChgValue	=	False	and	frm1.vspdData.Maxrows	=	0	then
		Call SetToolbar("1100110100011111")
	End	if

	gMouseClickStatus	=	"SPC"
	Set	gActiveSpdSheet	=	frm1.vspdData
	ggoSpread.Source = frm1.vspdData
	With Frm1
		.vspdData.Row	=	Row
		.vspdData.Col	=	0
		flagTxt	=	.vspdData.Text
		If flagTxt =	ggoSpread.InsertFlag or	flagTxt	=	 ggoSpread.UpdateFlag	or flagTxt =	ggoSpread.DeleteFlag Then
			Exit Sub
		End	If
	End	With

	gMouseClickStatus	=	"SPC"

	Set	gActiveSpdSheet	=	frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then																										 'If there is	no data.
		 Exit	Sub
 	End	If

	If Row <=	0	Then
		 ggoSpread.Source	=	frm1.vspdData

		 If	lgSortKey	=	1	Then
			 ggoSpread.SSSort	Col								'Sort	in ascending
			 lgSortKey = 2
		 Else
			 ggoSpread.SSSort	Col, lgSortKey		'Sort	in descending
			 lgSortKey = 1
		 End If

		 Exit	Sub
	End	If


	lgCurrentSpd = "M"
	Set	gActiveSpdSheet	=	frm1.vspdData
	lgStrPrevKey1	=	""
	lgStrPrevKey2	=	""


	If lgOldRow_A	<> Row Then
		frm1.vspdData.Col	=	C_Plan_Dt
		frm1.hWork_Dt.value	=	frm1.vspdData.text
		frm1.vspdData.Col	=	C_FACILITY_CD
		frm1.hFacility_Cd.value	=	frm1.vspdData.text
		lgOldRow_A = Row

		Call	DisableToolBar(	parent.TBC_QUERY)
			ggoSpread.Source			 = Frm1.vspdData1
			ggoSpread.ClearSpreadData
			ggoSpread.Source			 = Frm1.vspdData2
			ggoSpread.ClearSpreadData

		lgPageNo_B = 0
		lgPageNo_C = 0

		Call DbDtlQuery1

	End	if
End	Sub

'========================================================================================================
'		Event	Name : vspdData1_Click
'		Event	Desc : 컬럼을	클릭할 경우	발생 
'========================================================================================================
Sub	vspdData1_Click(ByVal	Col, ByVal Row)

	Call SetPopupMenuItemInf("1101111111")

	 gMouseClickStatus = "SP1C"

	Set	gActiveSpdSheet	=	frm1.vspdData1

	If frm1.vspdData1.MaxRows	=	0	Then																										'If	there	is no	data.
		 Exit	Sub
 	End	If

	If Row <=	0	Then
		 ggoSpread.Source	=	frm1.vspdData1

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
	Set	gActiveSpdSheet	=	frm1.vspdData1
End	Sub
'========================================================================================================
'		Event	Name : vspdData2_Click
'		Event	Desc : 컬럼을	클릭할 경우	발생 
'========================================================================================================
Sub	vspdData2_Click(ByVal	Col, ByVal Row)

	Call SetPopupMenuItemInf("1101111111")

	 gMouseClickStatus = "SP2C"

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
	lgCurrentSpd = "S2"
	Set	gActiveSpdSheet	=	frm1.vspdData2
End	Sub
'========================================================================================================
'		Event	Name : vspdData_TopLeftChange
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop	,	ByVal	NewLeft	,	ByVal	NewTop )
	If OldLeft <>	NewLeft	Then
			Exit Sub
	End	If

	If CheckRunningBizProcess	=	True Then
		 Exit	Sub
	End	If

	if frm1.vspdData.MaxRows < NewTop	+	VisibleRowCnt(frm1.vspdData,NewTop)	Then	'☜: 재쿼리	체크'
		If lgPageNo_A	<> ""	Then														'⊙: 다음	키 값이	없으면 더	이상 업무로직ASP를 호출하지	않음 
			 Call	DisableToolBar(Parent.TBC_QUERY)
			 Call	DbQuery
		End	If
	End	if
End	Sub
'========================================================================================================
'		Event	Name : vspdData1_TopLeftChange
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData1_TopLeftChange(ByVal	OldLeft	,	ByVal	OldTop , ByVal NewLeft , ByVal NewTop	)
	If OldLeft <>	NewLeft	Then
		Exit Sub
	End	If

	If CheckRunningBizProcess	=	True Then
		 Exit	Sub
	End	If

	if frm1.vspdData1.MaxRows	<	NewTop + VisibleRowCnt(frm1.vspdData1,NewTop)	Then	'☜: 재쿼리	체크'
		If lgPageNo_B	<> ""	Then														'⊙: 다음	키 값이	없으면 더	이상 업무로직ASP를 호출하지	않음 
			 Call	DisableToolBar(Parent.TBC_QUERY)
			 Call	DbDtlQuery1
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

	If CheckRunningBizProcess	=	True Then
		Exit	Sub
	End	If

	if frm1.vspdData1.MaxRows	<	NewTop + VisibleRowCnt(frm1.vspdData1,NewTop)	Then	'☜: 재쿼리	체크'
		If lgPageNo_C	<> ""	Then														'⊙: 다음	키 값이	없으면 더	이상 업무로직ASP를 호출하지	않음 
			 Call	DisableToolBar(Parent.TBC_QUERY)
			 Call	DbDtlQuery2
		End	If
	End	if
End	Sub
'========================================================================================================
'		Event	Name : vspdData_DblClick
'		Event	Desc :
'========================================================================================================
Sub	vspdData_DblClick(ByVal	Col, ByVal Row)
	Dim	iColumnName

	If Row <=	0	Then
		Exit Sub
	End	If

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End	If
End	Sub

'========================================================================================================
'		Event	Name : vspdData1_DblClick
'		Event	Desc :
'========================================================================================================
Sub	vspdData1_DblClick(ByVal Col,	ByVal	Row)
	Dim	iColumnName

	If Row <=	0	Then
		Exit Sub
	End	If

	If frm1.vspdData1.MaxRows	=	0	Then
		Exit Sub
	End	If

End	Sub

'========================================================================================================
'		Event	Name : vspdData2_DblClick
'		Event	Desc :
'========================================================================================================
Sub	vspdData2_DblClick(ByVal Col,	ByVal	Row)
	Dim	iColumnName

	If Row <=	0	Then
		Exit Sub
	End	If

	If frm1.vspdData2.MaxRows	=	0	Then
		Exit Sub
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
'		Event	Name : vspdData_ColWidthChange
'		Event	Desc :
'========================================================================================================
Sub	vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End	Sub
'========================================================================================================
'		Event	Name : vspdData_ColWidthChange
'		Event	Desc :
'========================================================================================================
Sub	vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData2
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

'========================================================================================================
'		Event	Name : vspdData_ScriptDragDropBlock
'		Event	Desc :
'========================================================================================================
Sub	vspdData1_ScriptDragDropBlock( Col ,	Row,	Col2,	 Row2,	NewCol,	 NewRow,	NewCol2,	NewRow2,	Overwrite	,	Action , DataOnly	,	Cancel )
	ggoSpread.Source = frm1.vspdData1
	Call ggoSpread.SpreadDragDropBlock(Col , Row,	Col2,	Row2,	NewCol,	NewRow,	NewCol2, NewRow2,	Overwrite	,	Action , DataOnly	,	Cancel )
	Call GetSpreadColumnPos("B")
End	Sub
'========================================================================================================
'		Event	Name : vspdData_ScriptDragDropBlock
'		Event	Desc :
'========================================================================================================
Sub	vspdData2_ScriptDragDropBlock( Col ,	Row,	Col2,	 Row2,	NewCol,	 NewRow,	NewCol2,	NewRow2,	Overwrite	,	Action , DataOnly	,	Cancel )
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.SpreadDragDropBlock(Col , Row,	Col2,	Row2,	NewCol,	NewRow,	NewCol2, NewRow2,	Overwrite	,	Action , DataOnly	,	Cancel )
	Call GetSpreadColumnPos("C")
End	Sub

Sub	vspdData_MouseDown(Button	,	Shift	,	x	,	y)
	If	Button = 2 And	gMouseClickStatus	=	"SPC"	Then
		gMouseClickStatus = "SPCR"
	End	If
End	Sub
Sub	vspdData1_MouseDown(Button , Shift , x , y)
	If Button	=	2	And	gMouseClickStatus	=	"SP1C" Then
		gMouseClickStatus = "SP1CR"
	End If
End	Sub
Sub	vspdData2_MouseDown(Button , Shift , x , y)
	If Button	=	2	And	gMouseClickStatus	=	"SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End	Sub

'========================================================================================================
'		Event	Name : vspdData_ButtonClicked
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData_ButtonClicked(ByVal Col,	ByVal	Row, Byval ButtonDown)
	With frm1.vspdData
		ggoSpread.Source	=	frm1.vspdData
		If Row > 0 Then
			Select Case	Col
				Case C_Req_Dept_POP
					.Col = Col - 1
					.Row = Row
					Call OpenCode(.text, C_Req_Dept_POP, Row)
				Case C_Insp_Dept_POP
					.Col = Col - 1
					.Row = Row
					Call OpenCode(.text, C_Insp_Dept_POP,	Row)
				Case C_FacilityPop
					.Col = Col - 1
					.Row = Row
					Call OpenFacility_Popup("2")
			End	Select
		End	If

	End	With
End	Sub


'========================================================================================================
'		Event	Name : vspdData1_ButtonClicked
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData1_ButtonClicked(ByVal	Col, ByVal Row,	Byval	ButtonDown)
	With frm1.vspdData1
		 ggoSpread.Source	=	frm1.vspdData1
		If Row > 0 Then
			Select Case	Col
				Case C_Sury_Assy_Pop
					.Col = C_Sury_Assy
					.Row = Row
					Call OpenSItem(.text)
				Case C_Cur_Popup
					.Col = Col - 1
					.Row = Row
					Call OpenCur (.Text)
			End	Select
		End	If
	End	With
End	Sub


'========================================================================================================
'		Event	Name : vspdData_ButtonClicked
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData2_ButtonClicked(ByVal	Col, ByVal Row,	Byval	ButtonDown)
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		frm1.vspdData2.row = frm1.vspdData2.ActiveRow
		If Row > 0 Then
			Select Case	Col
				Case C_Insp_Emp_Pop
					Call OpenEmptName("1")
				Case C_Cust_Pop
					frm1.vspdData2.Col	=	C_Cust_CD
					Call OpenBp(frm1.vspdData2.Text, 3)
			End	Select
		End	If

	End	With
End	Sub


'========================================================================================================
'		Event	Name : vspdData_ScriptLeaveCell
'		Event	Desc : This	function is	called when	cursor leave cell
'========================================================================================================
Sub	vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	Dim	iRet,	flagTxt

	If NewRow	<= 0 Or	Row	=	NewRow Then	Exit Sub



	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Row	=	NewRow
	frm1.vspdData.Col	=	C_Plan_Dt
	frm1.hWork_Dt.value	=	frm1.vspdData.text
	frm1.vspdData.Col	=	C_FACILITY_CD
	frm1.hFacility_Cd.value	=	frm1.vspdData.text
	lgOldRow_A = Row

	Call	DisableToolBar(	parent.TBC_QUERY)
		ggoSpread.Source			 = Frm1.vspdData1
		ggoSpread.ClearSpreadData
		ggoSpread.Source			 = Frm1.vspdData2
		ggoSpread.ClearSpreadData

	lgPageNo_B = 0
	lgPageNo_C = 0

	ggoSpread.Source = frm1.vspdData
	With Frm1
		.vspdData.Row	=	NewRow
		.vspdData.Col	=	0
		flagTxt	=	.vspdData.Text
		If flagTxt =	ggoSpread.InsertFlag or	flagTxt	=	 ggoSpread.UpdateFlag	or flagTxt =	ggoSpread.DeleteFlag Then
			Exit Sub
		End	If
	End	With

	If DbDtlQuery1() = False Then	Exit Sub

End	Sub

'========================================================================================================
'		Event	Name : vspdData_OnFocus
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData_OnFocus()
	lgActiveSpd			 = "M"
	lgCurrentSpd	="M"
End	Sub
'========================================================================================================
'		Event	Name : vspdData1_OnFocus
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData1_OnFocus()
	lgActiveSpd			 = "S1"
	lgCurrentSpd	="S1"
End	Sub

'========================================================================================================
'		Event	Name : vspdData2_OnFocus
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData2_OnFocus()
	lgActiveSpd			 = "S2"
	lgCurrentSpd	="S2"
End	Sub
'========================================================================================================
'		Event	Name : vspdData_TopLeftChange
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop	,	ByVal	NewLeft	,	ByVal	NewTop )

	If OldLeft <>	NewLeft	Then
		Exit Sub
	End	If

	If CheckRunningBizProcess	=	True Then
		Exit	Sub
	End	If

	if frm1.vspdData.MaxRows < NewTop	+	VisibleRowCnt(frm1.vspdData,NewTop)	Then	'☜: 재쿼리	체크'
		If lgPageNo_A	<> ""	Then														'⊙: 다음	키 값이	없으면 더	이상 업무로직ASP를 호출하지	않음 
			Call	DisableToolBar(Parent.TBC_QUERY)
			Call	DbQuery
		End	If
	End	if
End	Sub

'========================================================================================================
'		Event	Name : vspdData1_TopLeftChange
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub	vspdData1_TopLeftChange(ByVal	OldLeft	,	ByVal	OldTop , ByVal NewLeft , ByVal NewTop	)

	If OldLeft <>	NewLeft	Then
		Exit Sub
	End	If

	If CheckRunningBizProcess	=	True Then
		 Exit	Sub
	End	If

	if frm1.vspdData1.MaxRows	<	NewTop + VisibleRowCnt(frm1.vspdData1,NewTop)	Then	'☜: 재쿼리	체크'
		If lgPageNo_B	<> ""	Then														'⊙: 다음	키 값이	없으면 더	이상 업무로직ASP를 호출하지	않음 
			Call	DisableToolBar(Parent.TBC_QUERY)
			Call	DbDtlQuery1
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

	If CheckRunningBizProcess	=	True Then
		Exit	Sub
	End	If

	if frm1.vspdData2.MaxRows	<	NewTop + VisibleRowCnt(frm1.vspdData2,NewTop)	Then	'☜: 재쿼리	체크'
		If lgPageNo_C	<> ""	Then														'⊙: 다음	키 값이	없으면 더	이상 업무로직ASP를 호출하지	않음 
			Call	DisableToolBar(Parent.TBC_QUERY)
			Call	DbDtlQuery2
		End	If
	End	if
End	Sub


'========================================================================================================
'		Event	Name : txtOcpt_type_Onchange
'		Event	Desc :
'========================================================================================================
Function txtOcpt_type_Onchange()
	gCounts	=	0
End	Function
'========================================================================================================
'		Event	Name : vspdData_GotFocus
'		Event	Desc : This	event	is spread	sheet	data changed
'========================================================================================================
Sub	vspdData_GotFocus()
	ggoSpread.Source = Frm1.vspdData
End	Sub
'========================================================================================================
'		Event	Name : vspdData_GotFocus
'		Event	Desc : This	event	is spread	sheet	data changed
'========================================================================================================
Sub	vspdData1_GotFocus()
	ggoSpread.Source = Frm1.vspdData1
End	Sub
'========================================================================================================
'		Event	Name : vspdData_GotFocus
'		Event	Desc : This	event	is spread	sheet	data changed
'========================================================================================================
Sub	vspdData2_GotFocus()
	ggoSpread.Source = Frm1.vspdData2
End	Sub


'==========================================================================================
'		Event	Name : txtAppFrDt
'		Event	Desc :
'==========================================================================================

 Sub txtWork_Dt_DblClick(Button)
	if Button	=	1	then
		frm1.txtWork_Dt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtWork_Dt.Focus
	End	if
End	Sub

'==========================================================================================
'		Event	Name : OCX_KeyDown()
'		Event	Desc :
'==========================================================================================

Sub	txtWork_Dt_KeyDown(KeyCode,	Shift)
	If KeyCode = 13	Then Call	MainQuery()
End	Sub


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
		strWhere = " plant_cd= " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "  "

		Call CommonQueryRs(" plant_nm ","	 b_plant ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtPlantCd.alt,"X")
			frm1.txtPlantCd.focus 
			frm1.txtPlantnm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPlantNM.value = strDataNm(0)
	Else
		frm1.txtPlantNm.value = ""
	End If
	
	
	If Trim(frm1.txtFacility_Cd.value) <> "" Then
		strWhere = " set_plant = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "  "
		strwhere = strwhere & " and facility_cd =" &  FilterVar(frm1.txtFacility_Cd.value, "''", "S") & "  "

		Call CommonQueryRs(" Facility_Nm ","	 Y_FACILITY ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtFacility_Cd.alt,"X")
			frm1.txtFacility_Cd.focus 
			frm1.txtFacility_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtFacility_nm.value = strDataNm(0)
	Else
		frm1.txtFacility_nm.value = ""
	End If
End Function

</SCRIPT>
<!-- #Include	file="../../inc/uni2kcm.inc" -->
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
								<td	background="../../../CShared/image/table/seltab_up_bg.gif"><IMG	src="../../../CShared/image/table/seltab_up_left.gif"	width="9"	height="23"	></td>
								<td	background="../../../CShared/image/table/seltab_up_bg.gif" align="center"	CLASS="CLSMTABP"><font color=white>설비수리내역등록</font></td>
								<td	background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG	src="../../../CShared/image/table/seltab_up_right.gif" width="10"	height="23"	></td>
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
									<TD	CLASS="TD5"	NOWRAP>작업일자</TD>
									<TD	CLASS="TD6"	NOWRAP><script language =javascript src='./js/p5230ma1_txtWork_Dt_txtWork_Dt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>설치공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="설치공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=36 tag="14XXXU"></TD>
								</TR>
								<TR>
									<TD	CLASS="TD5"	NOWRAP>설비유형</TD>
									<TD	CLASS="TD6"	NOWRAP><SELECT NAME="CboFacility_Accnt"	ALT="근태구분" CLASS ="CboFacility_Accnt"	TAG="1XN"><OPTION	VALUE=""></OPTION></SELECT></TD>
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
					<TD	WIDTH=100% HEIGHT=*	valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR	HEIGHT="60%">
								<TD	WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p5230ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
							<TR	HEIGHT="20%">
								<TD	WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p5230ma1_vaSpread1_vspdData1.js'></script>
								</TD>
							</TR>
							<TR	HEIGHT="20%">
								<TD	WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p5230ma1_vaSpread2_vspdData2.js'></script>
								</TD>
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
<INPUT TYPE=HIDDEN NAME="txtMode"	tag="24"><INPUT	TYPE=HIDDEN	NAME="txtInsrtUserId"	tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="ChgSave1" tag="24">
<INPUT TYPE=HIDDEN NAME="ChgSave2" tag="24">
<INPUT TYPE=HIDDEN NAME="ChgSave3" tag="24">
<INPUT TYPE=HIDDEN NAME="hWork_Dt" tag="24">
<INPUT TYPE=HIDDEN NAME="hFacility_Cd" tag="24">

</FORM>
<DIV ID="MousePT"	NAME="MousePT">
<iframe	name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO	noresize framespacing=0	width=220	height=41	src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

