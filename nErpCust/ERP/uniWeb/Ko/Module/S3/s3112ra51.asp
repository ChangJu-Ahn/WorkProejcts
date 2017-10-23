<%@	LANGUAGE="VBSCRIPT"	%>
<!--
======================================================================================================
*	 1.	Module Name					: 
*	 2.	Function Name				: 
*	 3.	Program	ID					: J2070ra1
*	 4.	Program	Name				:
*	 5.	Program	Desc				: 
*	 6.	Comproxy List				:
*	 7.	Modified date(First) 		: 2005/05/11
*	 8.	Modified date(Last)	 		: 2005/05/11
*	 9.	Modifier (First)			: Lee Sang-Ho
*	10.	Modifier (Last)				: Lee Sang_Ho
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include	file="../../inc/IncSvrCcm.inc" -->
<!-- #Include	file="../../inc/incSvrHTML.inc"	-->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
-->
<LINK	REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--
'============================================  1.1.2 공통 Include  ======================================
'========================================================================================================
-->
<SCRIPT	LANGUAGE="VBScript"	  SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"	  SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"	  SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"	  SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script	Language="VBScript">
Option Explicit

'========================================================================================================
'=											 4.2 Constant	variables
'========================================================================================================
Const	BIZ_PGM_ID  = "s3112rb51.asp"					'Biz	Logic	ASP

'========================================================================================================
'=											 4.3 Common	variables
'========================================================================================================
<!-- #Include	file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=											 4.4 User-defind Variables
'========================================================================================================
<%'========================================================================================================%>

Dim	IsOpenPop

Dim	gSelframeFlg				 ' 현재	TAB의	위치를 나타내는	Flag

Dim	lgStrPrevKey1
Dim	lgStrPrevKey2
Dim	lgStrPrevKey3
Dim	lgPageNo_A
Dim	lgPageNo_B
Dim	lgPageNo_C
Dim	lgOldRow_A
Dim	lgOldRow_B
Dim	lgOldRow_C


Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_unit
Dim C_TrackingNo
Dim C_Qty
Dim C_Price
Dim C_PriceFlag
Dim C_PriceFlagNm
Dim C_NetAmt
Dim C_VatAmt
Dim C_PlantCd
Dim C_PlantNm
Dim C_Dlvydt
Dim C_BpCd
Dim C_BpNm
Dim C_VatType
Dim C_VatTypeNm
Dim C_VatRate
Dim C_VatIncFlag
Dim C_VatIncFlagNm

'==========================================	 1.2.2 Global	변수 선언	 =====================================
'	1. 변수	표준에 따름. prefix로	g를	사용함.
'	2.Array인	경우는 ()를	반드시 사용하여	일반 변수와	구별해 됨 
'=========================================================================================================
<%'========================================================================================================%>
														    'Window가 여러 개 뜨는 것을 방지하기 위해 
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim arrReturn												'☜: Return Parameter Group
Dim arrParam
Dim arrParent
					
arrParent = window.dialogArguments
Set popupParent = ArrParent(0)

arrparam = ArrParent(1)

top.document.title = popupParent.gActivePRAspName

Dim	iDBSYSDate
Dim	EndDate, StartDate

	'------	☆:	초기화면에 뿌려지는	마지막 날짜	------
	EndDate	=	"<%=GetSvrDate%>"
	'------	☆:	초기화면에 뿌려지는	시작 날짜	------
	StartDate	=	UNIDateAdd("S1",	-1,	EndDate, popupParent.gServerDateFormat)
	EndDate	=	UniConvDateAToB(EndDate, popupParent.gServerDateFormat, popupParent.gDateFormat)
	StartDate	=	UniConvDateAToB(StartDate, popupParent.gServerDateFormat, PopupParent.gDateFormat)

'========================================================================================================
'	Name : InitSpreadPosVariables()
'	Desc : Initialize	the	position
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)

	C_ItemCd		= 1
	C_ItemNm		= 2
	C_Spec			= 3
	C_unit			= 4
	C_TrackingNo	= 5
	C_Qty			= 6
	C_Price			= 7
	C_PriceFlag		= 8
	C_PriceFlagNm	= 9
	C_NetAmt		= 10
	C_VatAmt		= 11
	C_PlantCd		= 12
	C_PlantNm		= 13
	C_Dlvydt		= 14
	C_BpCd			= 15
	C_BpNm			= 16
	C_VatType		= 17
	C_VatTypeNm		= 18
	C_VatRate		= 19
	C_VatIncFlag	= 20
	C_VatIncFlagNm	= 21
	
End Sub

'========================================================================================================
'	Name : InitVariables()
'	Desc : Initialize	value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode			=	PopupParent.OPMD_CMODE										'⊙: Indicates that	current	mode is	Create mode
	lgBlnFlgChgValue	=	False										'⊙: Indicates that	no value changed
	lgIntGrpCount			=	0										'⊙: Initializes Group View	Size
	lgStrPrevKey			=	""																			'⊙: initializes Previous	Key
	lgStrPrevKey1		=	""																			'⊙: initializes Previous	Key	Index
	lgStrPrevKey2		=	""																			'⊙: initializes Previous	Key	Index
	lgStrPrevKey3		=	""																			'⊙: initializes Previous	Key	Index
	lgSortKey					=	1																				'⊙: initializes sort	direction
	lgOldRow_A = 0
	lgOldRow_B = 0
	lgOldRow_C = 0
	lgPageNo_A = 0
	lgPageNo_B = 0
	lgPageNo_C = 0

    ReDim arrReturn(0, 0)
    Self.Returnvalue = arrReturn

End Sub

'========================================================================================================
'	Name : LoadInfTB19029()
'	Desc : Set System	Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include	file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call	loadInfTB19029A("I", "H","NOCOOKIE","RA")	%>
End Sub


'========================================================================================================
'	Name : InitData()
'	Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim	intRow
	Dim	intIndex

End Sub

'========================================================================================================
'	Function Name	:	InitSpreadSheet
'	Function Desc	:	This method	initializes	spread sheet column	property
'========================================================================================================
Sub InitSpreadSheet(ByVal	pvSpdNo)

	Call initSpreadPosVariables(pvSpdNo)

	if pvSpdNo = "A" then
		ggoSpread.Source = frm1.vspddata
		With frm1.vspddata
			ggoSpread.Spreadinit "V20021129",,PopupParent.gAllowDragDropSpread

			.ReDraw	=	false

			.MaxCols = C_VatIncFlagNm +1																								<%'☜: 최대	Columns의	항상 1개 증가시킴 %>
			.Col = .MaxCols															<%'공통콘트롤	사용 Hidden	Column%>
			.ColHidden = True

			.MaxRows = 0
			ggoSpread.ClearSpreadData
			
			Call GetSpreadColumnPos(pvSpdNo)

			ggoSpread.SSSetEdit 	C_ItemCD, 		"품목코드"		, 12
			ggoSpread.SSSetEdit 	C_ItemNm, 		"품목명"		, 18
			ggoSpread.SSSetEdit 	C_Spec, 		"규격"			, 18
			ggoSpread.SSSetEdit 	C_Unit, 		"단위"			, 12
			ggoSpread.SSSetEdit 	C_TrackingNo, 	"Tracking No."	, 12
			ggoSpread.SSSetFloat	C_Qty,			"수량"			, 15,	PopupParent.ggQtyNo,	  ggStrIntegeralPart,	ggStrDeciPointPart,	PopupParent.gComNum1000,	 PopupParent.gComNumDec,	,	,	"Z"
			ggoSpread.SSSetFloat	C_Price,		"단가"			, 15,	PopupParent.ggUnitCostNo, ggStrIntegeralPart,	ggStrDeciPointPart,	PopupParent.gComNum1000,	 PopupParent.gComNumDec,	,	,	"Z"
			ggoSpread.SSSetEdit 	C_PriceFlag, 	"단가구분"		, 12
			ggoSpread.SSSetEdit 	C_PriceFlagNm, 	"단가구분명"	, 12
			ggoSpread.SSSetFloat	C_NetAmt,		"금액"			, 15,	PopupParent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	PopupParent.gComNum1000,	PopupParent.gComNumDec,	,	,	"Z"
			ggoSpread.SSSetFloat	C_VatAmt,		"금액"			, 15,	PopupParent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	PopupParent.gComNum1000,	PopupParent.gComNumDec,	,	,	"Z"
			ggoSpread.SSSetEdit 	C_PlantCd, 		"공장"			, 12
			ggoSpread.SSSetEdit 	C_PlantNm, 		"공장명"		, 12
			ggoSpread.SSSetDate		C_Dlvydt,		"납기일"		, 12,   2,PopupParent.gDateFormat
			ggoSpread.SSSetEdit 	C_BpCd, 		"공장"			, 12
			ggoSpread.SSSetEdit 	C_BpNm, 		"공장명"		, 12
			ggoSpread.SSSetEdit 	C_VatType, 		"VAT유형"		, 12
			ggoSpread.SSSetEdit 	C_VatTypeNm, 	"VAT유형명"		, 12
			ggoSpread.SSSetFloat	C_VatRate,		"VAT율"			, 10,	PopupParent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	PopupParent.gComNum1000, PopupParent.gComNumDec,	,	,	"Z"
			ggoSpread.SSSetEdit 	C_VatIncFlag, 	"VAT포함구분"   , 12
			ggoSpread.SSSetEdit 	C_VatIncFlagNm,	"VAT포함구분명" , 12
			
			Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)

			.ReDraw	=	true

			Call SetSpreadLock(pvSpdNo)

		End	With
	End If


End Sub

'======================================================================================================
'	Function Name	:	SetSpreadLock
'	Function Desc	:	This method	set	color	and	protect	in spread	sheet	celles
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

    With frm1
		If pvSpdNo = "A" Then
			ggoSpread.Source = frm1.vspddata
			.vspddata.ReDraw = False
 			ggoSpread.SpreadLock -1, -1, -1

			.vspddata.ReDraw = True
		End If

	End With

End Sub


'======================================================================================================
'	Function Name	:	SetSpreadColor
'	Function Desc	:	This method	set	color	and	protect	in spread	sheet	celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal	pvEndRow)

	With frm1.vspddata
		ggoSpread.Source = frm1.vspddata

		.ReDraw	=	False
		.ReDraw	=	True
	End	With

End Sub


'======================================================================================================
'	Function Name	:	SubSetErrPos
'	Function Desc	:	This method	set	focus	to pos of	err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
		Dim	iDx
		Dim	iRow
		iPosArr	=	Split(iPosArr, PopupParent.gColSep)
		If IsNumeric(iPosArr(0)) Then
			 iRow	=	CInt(iPosArr(0))
			 For iDx = 1 To	 frm1.vspddata.MaxCols	-	1
					 Frm1.vspddata.Col	=	iDx
					 Frm1.vspddata.Row	=	iRow
					 If	Frm1.vspddata.ColHidden <>	True And Frm1.vspddata.BackColor	<>	PopupParent.UC_PROTECTED	Then
							Frm1.vspddata.Col = iDx
							Frm1.vspddata.Row = iRow
							Frm1.vspddata.Action	=	0	'	go to
							Exit For
					 End If

			 Next

		End If
End Sub

'========================================================================================
'	Function Name	:	GetSpreadColumnPos
'	Description		:
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_Spec			= iCurColumnPos(3)
			C_unit			= iCurColumnPos(4)
			C_TrackingNo	= iCurColumnPos(5)
			C_Qty			= iCurColumnPos(6)
			C_Price			= iCurColumnPos(7)
			C_PriceFlag		= iCurColumnPos(8)
			C_PriceFlagNm	= iCurColumnPos(9)
			C_NetAmt		= iCurColumnPos(10)
			C_VatAmt		= iCurColumnPos(11)
			C_PlantCd		= iCurColumnPos(12)
			C_PlantNm		= iCurColumnPos(13)
			C_Dlvydt		= iCurColumnPos(14)
			C_BpCd			= iCurColumnPos(15)
			C_BpNm			= iCurColumnPos(16)
			C_VatType		= iCurColumnPos(17)
			C_VatTypeNm		= iCurColumnPos(18)
			C_VatRate		= iCurColumnPos(19)
			C_VatIncFlag	= iCurColumnPos(20)
			C_VatIncFlagNm	= iCurColumnPos(21)
			
	End Select
End Sub

'========================================================================================================
'	Name : Form_Load
'	Desc : developer describe	this line	Called by	Window_OnLoad()	evnt
'========================================================================================================
Sub Form_Load()

	Err.Clear																																				'☜: Clear err status
	Call LoadInfTB19029																															'☜: Load	table	,	B_numeric_format

	Call	ggoOper.FormatField(Document,	"1", ggStrIntegeralPart,	ggStrDeciPointPart,	PopupParent.gDateFormat,	PopupParent.gComNum1000,	PopupParent.gComNumDec)
	Call	ggoOper.FormatField(Document,	"2", ggStrIntegeralPart,	ggStrDeciPointPart,	PopupParent.gDateFormat,	PopupParent.gComNum1000,	PopupParent.gComNumDec)
	Call	ggoOper.LockField(Document,	"N")											'⊙: Lock	Field

	Call InitSpreadSheet("A")
	Call InitVariables
	Call SetDefaultVal()

	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")	
 	Call FncQuery()
	frm1.vspddata.focus	

End Sub

Sub SetDefaultVal()
	Dim arrParam
	
	arrParam = arrParent(1)

'     Call ggoOper.FormatDate(frm1.txtPoFrDt1, Parent.gDateFormat, 2)

' 	frm1.vspdData.OperationMode = 5
	
 	frm1.txtsono.value 	= Trim(arrParam(1))
 	frm1.txtsonm.value 	= Trim(arrParam(2))
 	
' 	frm1.txtPoFrDt1.Month	= right(Trim(arrParam(2)),2)


End Sub

'========================================================================================================
'	Name : Form_QueryUnload
'	Desc : developer describe	this line	Called by	Window_OnUnLoad()	evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
'	Name : FncQuery
'	Desc : developer describe	this line	Called by	FncQuery	in Common.vbs
'========================================================================================================
Function FncQuery()

	Dim	IntRetCD
	Dim	ChgOK

	FncQuery = False															 '☜:	Processing is	NG
	Err.Clear																																		 '☜:	Clear	err	status

	ChgOK	=	false


	ggoSpread.Source	=	Frm1.vspddata
	If	ggoSpread.SSCheckChange	=	True Then
		ChgOK	=	True
	End If

	If	ChgOK	Then
		IntRetCD =	DisplayMsgBox("900013",	 PopupParent.VB_YES_NO,"x","x")		'☜: Data	is changed.	 Do	you	want to	display	it?

		If IntRetCD	=	vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "2")
	Call InitVariables()

	If Not chkField(Document,	"1") Then													 '☜:	This function	check	required field
		 Exit	Function
	End If

	If DbDtlQuery1 = False Then
		Exit Function
	End If
		
	FncQuery = True																															 '☜:	Processing is	OK

End Function

'========================================================================================================
'	Name : FncDelete
'	Desc : developer describe	this line	Called by	MainDelete in	Common.vbs
'========================================================================================================
Function FncDelete()
		Dim	intRetCD

		FncDelete	=	False																														 '☜:	Processing is	NG
		Err.Clear																																		 '☜:	Clear	err	status

		FncDelete	=	True																														 '☜:	Processing is	OK
End Function

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

	If lgBlnFlgChgValue = False Then
		IntRetCD =	DisplayMsgBox("900001","x","x","x")
		Exit Function
	End If

	ggoSpread.Source = frm1.vspddata
	If Not	ggoSpread.SSDefaultCheck Then
		 Exit	Function
	End If

	lgCurrentSpd = "S1"

	If DbSave	=	False	Then

		Exit Function
	End If

	FncSave	=	True																															'☜: Processing	is OK

End Function

'========================================================================================================
'	Function Name	:	FncCopy
'	Function Desc	:	This function	is related to	Copy Button	of Main	ToolBar
'========================================================================================================
Function FncCopy()

End Function

'========================================================================================================
'	Function Name	:	FncCancel
'	Function Desc	:	This function	is related to	Cancel Button	of Main	ToolBar
'========================================================================================================
Function FncCancel()
	if lgCurrentSpd	=	"S1"	then
		ggoSpread.Source = Frm1.vspddata
		ggoSpread.EditUndo
	End If
End Function

'========================================================================================================
'	Function Name	:	FncInsertRow
'	Function Desc	:	This function	is related to	InsertRow	Button of	Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal	pvRowCnt)

End Function

'========================================================================================================
'	Function Name	:	FncDeleteRow
'	Function Desc	:	This function	is related to	DeleteRow	Button of	Main ToolBar
'========================================================================================================
Function FncDeleteRow()

End Function

'========================================================================================================
'	Function Name	:	FncPrint
'	Function Desc	:	This function	is related to	Print	Button of	Main ToolBar
'========================================================================================================
Function FncPrint()
		Call PopupParent.FncPrint()
End Function

'========================================================================================================
'	Function Name	:	FncExcel
'	Function Desc	:	This function	is related to	Excel
'========================================================================================================
Function FncExcel()
		Call PopupParent.FncExport( PopupParent.C_MULTI)																				 '☜:	화면 유형 
End Function

'========================================================================================================
'	Function Name	:	FncFind
'	Function Desc	:
'========================================================================================================
Function FncFind()
		Call PopupParent.FncFind( PopupParent.C_MULTI, False)																		 '☜:화면	유형,	Tab	유무 
End Function

'========================================================================================
'	Function Name	:	FncSplitColumn
'	Function Desc	:
'========================================================================================
Sub FncSplitColumn()

		If UCase(Trim(TypeName(gActiveSpdSheet)))	=	"EMPTY"	Then
			 Exit	Sub
		End If

		ggoSpread.Source = gActiveSpdSheet
		ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)

End Sub

'========================================================================================
'	Function Name	:	PopSaveSpreadColumnInf
'	Description		:
'========================================================================================
Sub PopSaveSpreadColumnInf()
		ggoSpread.Source = gActiveSpdSheet
		Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
'	Function Name	:	PopRestoreSpreadColumnInf
'	Description		:
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet(gActiveSpdSheet.Id)
	Call ggoSpread.ReOrderingSpreadData
End Sub


'========================================================================================================
'	Function Name	:	FncExit
'	Function Desc	:
'========================================================================================================
Function FncExit()
	Dim	IntRetCD

	FncExit	=	False

	ggoSpread.Source	=	frm1.vspddata
	If	ggoSpread.SSCheckChange	=	True Then
		IntRetCD =	DisplayMsgBox("900016",	 PopupParent.VB_YES_NO,"x","x")			'⊙: Data	is changed.	 Do	you	want to	exit?
		If IntRetCD	=	vbNo Then
			Exit Function
		End If
	End If

	FncExit	=	True
End Function


'========================================================================================================
'	Name : DbDtlQuery1
'	Desc : This	function is	called by	FncQuery
'========================================================================================================

Function DbDtlQuery1()

	DbDtlQuery1 = False

	Err.Clear

	If LayerShowHide(1)	=	False	Then
		Exit Function
	End If

	Dim	strVal

	With Frm1
		strVal = BIZ_PGM_ID	&	"?txtMode="			&	PopupParent.UID_M0001
		strVal = strVal		&	"&lgCurrentSpd="	&	lgCurrentSpd
		strVal = strVal		&	"&txtKeyStream="	&	lgKeyStream
		strVal = strVal		&	"&txtMaxRows="		&	.vspddata.MaxRows
		strVal = strVal		&	"&lgStrPrevKey=" 	&	lgStrPrevKey1
		strVal = strVal		&	"&lgPageNo_A="		&	lgPageNo_A
		strVal = strVal		&	"&txtSoNo="			&	.txtSoNo.value
	End	With

	Call RunMyBizASP(MyBizASP, strVal)

	DbDtlQuery1 = True
End Function

Function DbDtlQuery1Ok()														'☆: 조회	성공후 실행로직 
	Dim	i
	lgIntFlgMode = PopupParent.OPMD_UMODE												'⊙: Indicates that	current	mode is	Update mode
	lgBlnFlgChgValue = False
End Function

'========================================================================================================
'	Function Name	:	DbQueryOk
'	Function Desc	:	Called by	MB Area	when query operation is	successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode =	PopupParent.OPMD_UMODE
	lgOldRow_A = 0
	Call	ggoOper.LockField(Document,	"Q")										'⊙: Lock	field
	Call InitData()
' 	call vspddata_click(1,frm1.vspddata.activerow)
' 	frm1.vspddata.focus
End Function


'========================================================================================================
'		Event	Name : vspddata_Click
'		Event	Desc : 컬럼을	클릭할 경우	발생 
'========================================================================================================
Sub vspddata_Click(ByVal	Col, ByVal Row)
	Dim	flagTxt
	Call SetPopupMenuItemInf("1101111111")


	gMouseClickStatus	=	"SPC"
	Set	gActiveSpdSheet	=	frm1.vspddata
	ggoSpread.Source = frm1.vspddata


	gMouseClickStatus	=	"SPC"

	Set	gActiveSpdSheet	=	frm1.vspddata

	If frm1.vspddata.MaxRows	=	0	Then																										'If	there	is no	data.
		 Exit	Sub
 	End If

	If Row <=	0	Then
		 ggoSpread.Source	=	frm1.vspddata

		 If	lgSortKey	=	1	Then
				 ggoSpread.SSSort	Col
				 lgSortKey = 2
		 Else
				 ggoSpread.SSSort	Col, lgSortKey		'Sort	in descending
				 lgSortKey = 1
		 End If

		 Exit	Sub
	End If

	lgCurrentSpd = "S1"
	Set	gActiveSpdSheet	=	frm1.vspddata
End Sub


'========================================================================================================
'		Event	Name : vspddata_TopLeftChange
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub vspddata_TopLeftChange(ByVal	OldLeft	,	ByVal	OldTop , ByVal NewLeft , ByVal NewTop	)
	If OldLeft <>	NewLeft	Then
			Exit Sub
	End If

	If CheckRunningBizProcess	=	True Then
		 Exit	Sub
	End If

	if frm1.vspddata.MaxRows	<	NewTop + VisibleRowCnt(frm1.vspddata,NewTop)	Then	'☜: 재쿼리	체크'
		If lgPageNo_A	<> ""	Then														'⊙: 다음	키 값이	없으면 더	이상 업무로직ASP를 호출하지	않음 

			 Call	DbDtlQuery1
		End If
	End If
End Sub


'========================================================================================================
'		Event	Name : vspddata_ColWidthChange
'		Event	Desc :
'========================================================================================================
Sub vspddata_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspddata
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub


'========================================================================================================
'		Event	Name : vspddata_ScriptDragDropBlock
'		Event	Desc :
'========================================================================================================
Sub vspddata_ScriptDragDropBlock( Col ,	Row,	Col2,	 Row2,	NewCol,	 NewRow,	NewCol2,	NewRow2,	Overwrite	,	Action , DataOnly	,	Cancel )
	ggoSpread.Source = frm1.vspddata
	Call ggoSpread.SpreadDragDropBlock(Col , Row,	Col2,	Row2,	NewCol,	NewRow,	NewCol2, NewRow2,	Overwrite	,	Action , DataOnly	,	Cancel )
	Call GetSpreadColumnPos("A")
End Sub


Sub vspddata_MouseDown(Button , Shift , x , y)
	If	Button = 2 And	gMouseClickStatus	=	"SPC"	Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'========================================================================================================
'		Event	Name : vspddata_ScriptLeaveCell
'		Event	Desc : This	function is	called when	cursor leave cell
'========================================================================================================
Sub vspddata_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	If NewRow	<= 0 Or	Row	=	NewRow Then	Exit Sub
End Sub

'========================================================================================================
'		Event	Name : vspddata_OnFocus
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub vspddata_OnFocus()
	lgActiveSpd		= "S1"
	lgCurrentSpd	="S1"
End Sub

Function vspdData_DblClick(ByVal Col, ByVal Row)
 	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
 	     Exit Function
 	End If
 	With frm1.vspdData 
 		If .MaxRows > 0 Then
 			If .ActiveRow = Row Or .ActiveRow > 0 Then
 				Call OKClick
 			End If
 		End If
 	End With
End Function


Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow

	If frm1.vspdData.MaxRows = 0 then
		Exit Function
	end if

	with frm1
		intInsRow = 0
		Redim arrReturn(ggoSpread.Source.SelBlockRow2 - ggoSpread.Source.SelBlockRow , frm1.vspdData.MaxCols-2)
		For intRowCnt = ggoSpread.Source.SelBlockRow To ggoSpread.Source.SelBlockRow2
			frm1.vspdData.Row = intRowCnt
			For intColCnt = 0 To frm1.vspdData.MaxCols - 2
				frm1.vspdData.Col = intColCnt + 1
				arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
			Next
			intInsRow = intInsRow + 1
	    Next 
	end with
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
' msgbox "fsdfdsafdsafads"
' ' ' 	Dim arrReturn
' 	ReDim arrReturn(1,1)
' 	arrReturn(0,0) = ""
' 	Self.Returnvalue = arrReturn	
	Self.Close()
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<%
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
%>



<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>프로젝트번호</TD>
						<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtSoNo" SIZE=25 MAXLENGTH=25 tag="14xxxU" ALT="프로젝트번호">
												 <INPUT TYPE=TEXT NAME="txtSoNM" SIZE=30 MAXLENGTH=40 tag="14xxxU" >
						</TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% SRC="../../blank.htm" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>


<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"  TABINDEX=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


