<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Facility Resources
*  2. Function Name        : 설비제원정보등록 
*  3. Program ID           : p5111ma1
*  4. Program Name         : p5111ma1
*  5. Program Desc         : 설비제원정보등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2005/01/17
*  8. Modified date(Last)  : 2005/01/17
*  9. Modifier (First)     : Lee Chang-Je
* 10. Modifier (Last)      : Lee Chang-Je
* 11. Comment              : Who Let the dog out?
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js">			</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

<%'========================================================================================================%>

Const BIZ_PGM_ID    = "P5110mb1.asp"
Const C_SHEETMAXROWS = 100

Const C_PopPlantCd = 1

<%'========================================================================================================%>


dim C_FACILITY_ACCNT_CD
dim C_FACILITY_ACCNT_NM
dim C_FACILITY_CD
dim C_FACILITY_NM
dim C_SET_PLANT
dim C_SET_PLANTPopUp
dim C_SET_PLANTNm
dim C_PROD_CO
dim C_PUR_DT
dim C_PROD_AMT
dim C_LIFE_CYCLE
dim C_CHK_PRD1
dim C_PIC_FLAG


Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3


Dim lgBlnOpenedFlag

Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
Dim IsOpenPop
Dim IsOpenPopDept
Dim UserPrevNext
Dim gSelframeFlg
Dim lgOldRow

'========================================================================================================
<%
Dim StratDate
StratDate = GetSvrDate
%>



<%'========================================================================================================%>
Sub InitSpreadPosVariables()

	C_FACILITY_ACCNT_CD =		1
	C_FACILITY_ACCNT_NM =		2
	C_FACILITY_CD =     		3
	C_FACILITY_NM =     		4
	C_SET_PLANT =				5
	C_SET_PLANTPopUp = 			6
	C_SET_PLANTNm  = 			7
	C_PROD_CO = 				8
	C_PUR_DT =					9
	C_PROD_AMT =				10
	C_LIFE_CYCLE =				11
	C_CHK_PRD1 =				12
	C_PIC_FLAG =				13


End Sub

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode		= parent.OPMD_CMODE
	lgPageNo       = ""
	lgBlnFlgChgValue	= False
	lgIntGrpCount		= 0
	lgStrPrevKey		= ""
	lgStrPrevKeyIndex	= 0
	lgSortKey			= 1
	lgLngCurRows		= 0
	lgOldRow = 0
End Sub

'========================================================================================================
Sub SetDefaultVal()
	dim strYear,strMonth,strDay
	'------ Developer Coding part (Start ) --------------------------------------------------------------

'     Call ggoOper.FormatDate(frm1.txtToDt, parent.gDateFormat, 2)
	lgBlnFlgChgValue = False
'	Frm1.txtProd_Amt.text = 0
'	Frm1.txtProd_Amt.MaxValue = 9999999999999999999.99
'	Frm1.txtProd_Amt.MinValue = 0

	Frm1.txtChk_Prd1.MaxValue = 99999
	Frm1.txtChk_Prd1.MinValue = 0
	Frm1.txtChk_Prd2.MaxValue = 99999
	Frm1.txtChk_Prd2.MinValue = 0
	Frm1.txtLife_Cycle.MaxValue = 99999
	Frm1.txtLife_Cycle.MinValue = 0

'	Frm1.txtPress_Power.MaxValue = 99999999999.9999
'	Frm1.txtPress_Power.MinValue = 0
'	Frm1.txtMoter_Power.MaxValue = 99999999999.9999
'	Frm1.txtMoter_Power.MinValue = 0
'	Frm1.txtMoter_qty.MaxValue = 999999999
'	Frm1.txtMoter_qty.MinValue = 0
'	Frm1.txtMoter_Cir_Qty.MaxValue = 999999999.9999
'	Frm1.txtMoter_Cir_Qty.MinValue = 0
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>  ' check


End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
<%'========================================================================================================%>

Sub InitSpreadSheet()

	Call initSpreadPosVariables()

	With frm1.vspdData


		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread

		.ReDraw = false
		.MaxCols = C_PIC_FLAG + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>

		.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		.ColHidden = True

		.MaxRows = 0
		ggoSpread.Source = Frm1.vspdData
		ggoSpread.ClearSpreadData
		Call GetSpreadColumnPos("A")


		ggoSpread.SSSetCombo 	C_FACILITY_ACCNT_CD,	"설비유형",  10, 0, True
		ggoSpread.SSSetCombo 	C_FACILITY_ACCNT_NM,	"설비유형",  20, 0, True
		ggoSpread.SSSetEdit	C_FACILITY_CD,       	"설비코드",	    15,,,20,2
		ggoSpread.SSSetEdit	C_FACILITY_NM,		"설비명", 		15,,,20,0
		ggoSpread.SSSetEdit	C_SET_PLANT,  		"설치공장",   10,,,10,2
		ggoSpread.SSSetButton	C_SET_PLANTPopUp
		ggoSpread.SSSetEdit	C_SET_PLANTNm,		"공장명", 15,,,20,2
		ggoSpread.SSSetEdit	C_PROD_CO,			"제작회사", 15,,,20,2
		ggoSpread.SSSetDate	C_PUR_DT,			"구입일자",		12,2,parent.gDateFormat
		ggoSpread.SSSetFloat	C_PROD_AMT,    		"금액", 19, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_LIFE_CYCLE,    		"수명(년)", 19, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat	C_CHK_PRD1,    		"점검주기(주)", 19, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetCombo	C_PIC_FLAG,			"사진유무",		12,		2,		true


		Call ggoSpread.SSSetColHidden(C_FACILITY_ACCNT_CD, C_FACILITY_ACCNT_CD, True)
		call ggoSpread.MakePairsColumn(C_FACILITY_ACCNT_CD,C_FACILITY_ACCNT_NM)
		call ggoSpread.MakePairsColumn(C_FACILITY_CD,C_FACILITY_NM)
		call ggoSpread.MakePairsColumn(C_SET_PLANT,C_SET_PLANTNm)

		.ReDraw = true

	End With
	Call SetSpreadLock()

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()'byVal gird_fg, byVal lock_fg, byVal iRow)
    With frm1

		ggoSpread.Source = .vspddata
		.vspddata.ReDraw = False
		ggoSpread.SpreadLock		C_FACILITY_ACCNT_CD,	-1, C_PIC_FLAG		,-1

' 		ggoSpread.SpreadLock		C_FACILITY_ACCNT_CD,	-1, C_FACILITY_ACCNT_CD		,-1
' 		ggoSpread.SpreadLock		C_FACILITY_CD,			-1, C_FACILITY_CD			,-1
' 		ggoSpread.SpreadLock		C_SET_PLANTNm,			-1, C_SET_PLANTNm			,-1
' 		ggoSpread.SpreadLock		C_PIC_FLAG,				-1, C_PIC_FLAG				,-1


' 		ggoSpread.SpreadLock	  	C_FACILITY_ACCNT_NM, -1, -1

		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		.vspddata.ReDraw = True

   End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal iwhere, ByVal pvStartRow, ByVal pvEndRow)

	With frm1.vspddata

	ggoSpread.Source = frm1.vspddata

	.ReDraw = False
	Select Case iwhere
	Case "Q"
		ggoSpread.SSSetProtected	C_FACILITY_ACCNT_NM		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_FACILITY_CD			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_SET_PLANTNm			, pvStartRow, pvEndRow

	CASE "I"
		ggoSpread.SSSetRequired		C_FACILITY_ACCNT_NM		, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_FACILITY_CD			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_SET_PLANTNm			, pvStartRow, pvEndRow

	End Select
	.ReDraw = True

	End With

End Sub




'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	lgKeyStream = Frm1.CbohFacility_Accnt.Value						& parent.gColSep
	lgKeyStream = lgKeyStream & Trim(frm1.txthFacility_Cd.value)	& parent.gColSep
	lgKeyStream = lgKeyStream & Trim(frm1.CbohUse_Yn.value)			& parent.gColSep
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim iCodeArr
	Dim iNameArr

	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z410' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_FACILITY_ACCNT_CD
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_FACILITY_ACCNT_NM
	Call SetCombo2(frm1.CbohFacility_Accnt ,lgF0  ,lgF1  ,Chr(11))
	Call SetCombo2(frm1.CboFacility_Accnt ,lgF0  ,lgF1  ,Chr(11))

	ggoSpread.SetCombo "Y" & vbtab & "N" , C_PIC_FLAG


	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z402' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	Call SetCombo2(frm1.txtItemGroupCd1 ,lgF0  ,lgF1  ,Chr(11))

	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z403' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	Call SetCombo2(frm1.txtItemGroupCd2 ,lgF0  ,lgF1  ,Chr(11))


	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z415' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	Call SetCombo2(frm1.txtOil_Spec1 ,lgF0  ,lgF1  ,Chr(11))
	Call SetCombo2(frm1.txtOil_Spec2 ,lgF0  ,lgF1  ,Chr(11))
	Call SetCombo2(frm1.txtOil_Spec3 ,lgF0  ,lgF1  ,Chr(11))
	Call SetCombo2(frm1.txtOil_Spec4 ,lgF0  ,lgF1  ,Chr(11))
	Call SetCombo2(frm1.txtOil_Spec5 ,lgF0  ,lgF1  ,Chr(11))


	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z424' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	Call SetCombo2(frm1.txtProd_Flag ,lgF0  ,lgF1  ,Chr(11))


	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z413' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	Call SetCombo2(frm1.txtEmp_no ,lgF0  ,lgF1  ,Chr(11))


	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z423' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	Call SetCombo2(frm1.txtPlant_Sts ,lgF0  ,lgF1  ,Chr(11))


	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub


Sub vspddata_Change(ByVal Col, ByVal Row)
	Dim iDx
	Dim intIndex, IntRetCd
	Dim strName
	Dim strDept_nm
	Dim strRoll_pstn
	Dim strPay_grd1
	Dim strPay_grd2
	Dim strEntr_dt
	Dim strInternal_cd

		Frm1.vspdData.Row = Row
		Frm1.vspdData.Col = Col

	Select Case Col
		Case  C_FACILITY_ACCNT_NM
			Frm1.vspdData.col = C_FACILITY_ACCNT_NM
			intIndex = Frm1.vspdData.value
			Frm1.vspdData.Col = C_FACILITY_ACCNT_CD
			Frm1.vspdData.value = intindex

	End Select

	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
		If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
			Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
		End If
	End If

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = TRUE


End Sub


'
''===================================== CurFormatNumericOCX()  =======================================
''	Name : CurFormatNumericOCX()
''	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
''====================================================================================================
'Sub CurFormatNumericOCX()
'	With frm1
'	'해당되는 금액이 있는 Data 필드에 대하여 각각 처리 
'		'계약금액 
'		ggoOper.FormatFieldByObjectOfCur .txtProd_Amt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
'	End With
'End Sub
'
''===================================== CurFormatNumSprSheet()  ======================================
''	Name : CurFormatNumSprSheet()
''	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
''====================================================================================================
'Sub CurFormatNumSprSheet()
'' 	Dim ii
'' 	With frm1
'' 		'해당되는 금액이 있는 Grid에 대하여 각각 처리 
'' 		ggoSpread.Source = frm1.vspdData
'' 		For ii = 1 To .vspdData.MaxRows
'' 			Call FixDecimalPlaceByCurrency2(frm1.vspdData,ii,.txtDocCur.value,C_Amt,"A" ,"X","X")
''       	Next
''        Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,1,-1,.txtDocCur.value,C_Amt,"A" ,"I","X","X")
'
'' 		'ggoSpread.SSSetFloatByCellOfCur C_Amt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
'' 	End With
'
'End Sub


'==========================================================================================
Sub vspddata_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("1101111111")
	gMouseClickStatus = "SPC"	'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
		Exit Sub
	End If

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

	'------ Developer Coding part (Start)
	If lgOldRow <> Row Then

		frm1.vspdData.Col = 1
		frm1.vspdData.Row = row

		lgOldRow = Row


		Call InitData()
		Call LayerShowHide(1)
		frm1.hFacility_Accnt.value = ""
		frm1.vspdData.Col = C_FACILITY_CD
		frm1.hFacility_Cd.value = frm1.vspdData.text
		frm1.hUse_Yn.value = ""

		If DbDtlQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If

	End If
	'------ Developer Coding part (End)
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	Dim iColumnName

	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
	Call ClickTab2
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
' 	If Col <= C_SNm Or NewCol <= C_SNm Then
' 	   Cancel = True
' 	   Exit Sub
' 	End If
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	Dim iRet
	If NewRow <= 0 Or Row = NewRow Then Exit Sub

	Call InitData()
	Call LayerShowHide(1)

	frm1.vspdData.row = NewRow
	frm1.hFacility_Accnt.value = ""
	frm1.vspdData.Col = C_FACILITY_CD
	frm1.hFacility_Cd.value = frm1.vspdData.text
	frm1.hUse_Yn.value = ""

	If DbDtlQuery = False Then	Exit Sub
End Sub




'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================

Sub vspddata_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc :
'==========================================================================================

Sub vspddata_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	Dim strTemp
	Dim intPos1
	Dim strCard

	With frm1.vspddata
		If Row > 0 then
			if   Col = C_SET_PLANTPopUp Then
  				.Col = C_SET_PLANT
			end if
		End If
	End With
End Sub

Sub vspddata_KeyPress(index , KeyAscii )
	lgBlnFlgChgValue = True                                                 '⊙: Indicates that value changed
End Sub


'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(Parent.TBC_QUERY)
			Call DbQuery
		End If
	End if
End Sub
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()

	Dim intRow
	Dim intIndex

	With frm1
' 		.txtExRate.text	=1
' 		.txtPress_Power.text	=0
' 		.txtLocCntAmt.text	=0
' 		.txtAmt.text		=0
' 		.txtLocAmt.text	=0
' 		.cboPrivatePublic.value = "N"
	End With


	lgBlnFlgChgValue = False
End Sub
'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   :
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_FACILITY_ACCNT_CD =		iCurColumnPos(1 )
			C_FACILITY_ACCNT_NM =		iCurColumnPos(2 )
			C_FACILITY_CD =     		iCurColumnPos(3 )
			C_FACILITY_NM =     		iCurColumnPos(4 )
			C_SET_PLANT =				iCurColumnPos(5 )
			C_SET_PLANTPopUp = 			iCurColumnPos(6 )
			C_SET_PLANTNm  = 			iCurColumnPos(7 )
			C_PROD_CO = 				iCurColumnPos(8 )
			C_PUR_DT =					iCurColumnPos(9 )
			C_PROD_AMT =				iCurColumnPos(10)
			C_LIFE_CYCLE =				iCurColumnPos(11)
			C_CHK_PRD1 =				iCurColumnPos(12)
			C_PIC_FLAG =				iCurColumnPos(13)
	End Select
End Sub

'========================================================================================================
Sub Form_Load()
	Err.Clear                                                                        '☜: Clear err status

	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call AppendNumberPlace("6", "5", "0")                                   'Format Numeric Contents Field%>
'	Call AppendNumberPlace("7", "4", "0")                                   'Format Numeric Contents Field%>

' 	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

     Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)


	Call ggoOper.LockField(Document, "N")                                            '⊙: Lock  Suitable  Field

	Call SetDefaultVal
	Call InitData()
	Call InitVariables
	Call InitSpreadSheet()                                                               '⊙: Setup the Spread sheet
	Call InitComboBox()

	Call SetToolbar("1100000000011111")
	lgBlnOpenedflag = True

	'------ Developer Coding part (End )   --------------------------------------------------------------
	gSelframeFlg = TAB1
	Call ClickTab1
	frm1.CbohFacility_Accnt.Focus
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
	'--------- Developer Coding Part (Start) ----------------------------------------------------------
	'--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
Function FncQuery()
	Dim IntRetCD
	Dim var_m

	FncQuery = False															 '☜: Processing is NG
	Err.Clear                                                                    '☜: Clear err status

	If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData
	var_m = ggoSpread.SSCheckChange

	If lgBlnFlgChgValue = True Or var_m = True    Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X") '☜ "데이타가 변경되었습니다. 조회하시겠습니까?"

		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	call ClickTab1()

	Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
	Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field

	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData

	Call InitData()
	Call SetDefaultVal()

	'--------- Developer Coding Part (Start) ----------------------------------------------------------

	Call InitVariables()                                                         '⊙: Initializes local global variables
	Call MakeKeyStream("Q")

	' 설비코드 유효성 체크 
	If  frm1.txthFacility_Cd.value = "" Then
		frm1.txthFacility_Nm.value=""
	Else
		IntRetCD= CommonQueryRs(" FACILITY_NM "," Y_FACILITY ","    FACILITY_CD = " & FilterVar(Trim(frm1.txthFacility_Cd.value),"''","S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			If IntRetCD=False And Trim(frm1.txthFacility_Cd.value)<>"" Then
				Call DisplayMsgBox("970000","X",frm1.txthFacility_Cd.alt,"X")                         '☜ : 등록되지 않은 코드입니다.
				frm1.txthFacility_Cd.value=""
				frm1.txthFacility_Nm.value=""
				frm1.txthFacility_Cd.focus
				Set gActiveElement = document.activeElement
				Exit Function
			Else
				frm1.txthFacility_Nm.value=Trim(Replace(lgF0,Chr(11),""))
			End If
	End if
	
	'------ Developer Coding part (End )   --------------------------------------------------------------
	frm1.hFacility_Cd.value = frm1.txthFacility_Cd.value

	
	If DbQuery = False Then
		Exit Function
	End If                                                                 '☜: Query db data

	Set gActiveElement = document.ActiveElement
	FncQuery = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
	Dim IntRetCD

	FncNew = False																  '☜: Processing is NG
	Err.Clear                                                                     '☜: Clear err status


	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to make it new?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "A")                                        '☜: Clear Condition Field
	Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field

	Call pvLockField("N")

	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData

	' 	Call SetToolbar("1100000000011111")


	Call SetDefaultVal()
	Call InitData
	Call InitVariables                                                            '⊙: Initializes local global variables
	Call ClickTab2()

' 	call txtDocCur_OnChangeASP()   

	Set gActiveElement = document.ActiveElement
	FncNew = True															      '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
	Dim intRetCD

	FncDelete = False
	Err.Clear

	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002","x","x","x")
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")
	If IntRetCD = vbNo Then
		Exit Function
	End If

	Call MakeKeyStream("D")
	If DbDelete = False Then
		Exit Function
	End If

	Set gActiveElement = document.ActiveElement
	FncDelete = True
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
	Dim IntRetCD
	Dim pFromDt, pToDt, pDateTime
	Dim strYear, strMonth, strDay
	Dim FrDt
	Dim strSelect, strFrom, strWhere

	FncSave = False                                                              '☜: Processing is NG
	Err.Clear                                                                    '☜: Clear err status

	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data.
		Exit Function
	End If


	If Not chkField(Document, "2") Then                                          '☜: Check contents area
		Exit Function
	End If

    if Not chkContentArea() Then
		Exit Function
	End If



	Call MakeKeyStream("S")
	'------ Developer Coding part (End )   --------------------------------------------------------------
	If DbSave = False Then                                                       '☜: Query db data
		Exit Function
	End If
	Set gActiveElement = document.ActiveElement
	FncSave = True                                                               '☜: Processing is OK
End Function

Function chkContentArea() 
	dim bRet
	bRet = True


	Dim IntRetCd
	
	' 설치라인 
	If  frm1.txtSet_Place.value = "" Then
		frm1.txtConWcNm.value=""
    Else
        IntRetCD= CommonQueryRs(" WC_NM "," P_WORK_CENTER ","    WC_CD = " & FilterVar(Trim(frm1.txtSet_Place.value),"''","S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			If IntRetCD=False And Trim(frm1.txtSet_Place.value)<>"" Then
				Call changeTabs(TAB2)	 '~~~ 첫번째 Tab
				gSelframeFlg = TAB2
				Call DisplayMsgBox("970000","X",frm1.txtSet_Place.alt,"X")
				frm1.txtSet_Place.value=""
				frm1.txtConWcNm.value=""
				frm1.txtSet_Place.focus
				Set gActiveElement = document.activeElement
				bRet = False
			Else
				frm1.txtConWcNm.value=Trim(Replace(lgF0,Chr(11),""))
			End If
    End if	

	if bRet then
		' 통화 
		If  frm1.txtDocCur.value <> "" Then
			IntRetCD= CommonQueryRs(" CURRENCY "," B_CURRENCY ","    CURRENCY = " & FilterVar(Trim(frm1.txtDocCur.value),"''","S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			If IntRetCD=False Then
				Call changeTabs(TAB2)	 '~~~ 첫번째 Tab
				gSelframeFlg = TAB2
				Call DisplayMsgBox("970000","X",frm1.txtDocCur.alt,"X")
				frm1.txtDocCur.value=""
				frm1.txtDocCur.focus
				Set gActiveElement = document.activeElement
				bRet = False
			End If
		End if	
	End If

	
	' 설치업체 
	if bRet then
		If  frm1.txtSetCoCd.value = "" Then
			frm1.txtSetCoNm.value=""
		Else
			IntRetCD= CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ","    BP_CD = " & FilterVar(Trim(frm1.txtSetCoCd.value),"''","S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IntRetCD=False And Trim(frm1.txtSetCoCd.value)<>"" Then
					Call changeTabs(TAB2)	 '~~~ 첫번째 Tab
					gSelframeFlg = TAB2
					Call DisplayMsgBox("970000","X",frm1.txtSetCoCd.alt,"X")
					frm1.txtSetCoCd.value=""
					frm1.txtSetCoNm.value=""
					frm1.txtSetCoCd.focus
					Set gActiveElement = document.activeElement
					bRet = False
				Else
					frm1.txtSetCoNm.value=Trim(Replace(lgF0,Chr(11),""))
				End If
		End if
	End if

	' 구매업체 
	if bRet then
		If  frm1.txtPurCoCd.value = "" Then
			frm1.txtPurCoNm.value=""
		Else
			IntRetCD= CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ","    BP_CD = " & FilterVar(Trim(frm1.txtPurCoCd.value),"''","S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IntRetCD=False And Trim(frm1.txtPurCoCd.value)<>"" Then
					Call changeTabs(TAB2)	 '~~~ 첫번째 Tab
					gSelframeFlg = TAB2
					Call DisplayMsgBox("970000","X",frm1.txtPurCoCd.alt,"X")
					frm1.txtPurCoCd.value=""
					frm1.txtPurCoNm.value=""
					frm1.txtPurCoCd.focus
					Set gActiveElement = document.activeElement
					bRet = False
				Else
					frm1.txtPurCoNm.value=Trim(Replace(lgF0,Chr(11),""))
				End If
		End if
	End if

	' 제작업체 
	if bRet then
		If  frm1.txtProdCoCd.value = "" Then
			frm1.txtProdCoNm.value=""
		Else
			IntRetCD= CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ","    BP_CD = " & FilterVar(Trim(frm1.txtProdCoCd.value),"''","S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IntRetCD=False And Trim(frm1.txtProdCoCd.value)<>"" Then
					Call changeTabs(TAB2)	 '~~~ 첫번째 Tab
					gSelframeFlg = TAB2
					Call DisplayMsgBox("970000","X",frm1.txtProdCoCd.alt,"X")
					frm1.txtProdCoCd.value=""
					frm1.txtProdCoNm.value=""
					frm1.txtProdCoCd.focus
					Set gActiveElement = document.activeElement
					bRet = False
				Else
					frm1.txtProdCoNm.value=Trim(Replace(lgF0,Chr(11),""))
				End If
		End if
	End if

	' 자산번호1
	if bRet then
		If  frm1.txtCondAsstNo1.value = "" Then
			frm1.txtCondAsstNm1.value=""
		Else
			IntRetCD= CommonQueryRs(" ASST_NM "," A_ASSET_MASTER ","    ASST_NO = " & FilterVar(Trim(frm1.txtCondAsstNo1.value),"''","S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IntRetCD=False And Trim(frm1.txtCondAsstNo1.value)<>"" Then
					Call changeTabs(TAB2)	 '~~~ 첫번째 Tab
					gSelframeFlg = TAB2
					Call DisplayMsgBox("970000","X",frm1.txtCondAsstNo1.alt,"X")
					frm1.txtCondAsstNo1.value=""
					frm1.txtCondAsstNm1.value=""
					frm1.txtCondAsstNo1.focus
					Set gActiveElement = document.activeElement
					bRet = False
				Else
					frm1.txtCondAsstNm1.value=Trim(Replace(lgF0,Chr(11),""))
				End If
		End if
	End if

	' 자산번호2
	if bRet then
		If  frm1.txtCondAsstNo2.value = "" Then
			frm1.txtCondAsstNm2.value=""
		Else
			IntRetCD= CommonQueryRs(" ASST_NM "," A_ASSET_MASTER ","    ASST_NO = " & FilterVar(Trim(frm1.txtCondAsstNo2.value),"''","S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IntRetCD=False And Trim(frm1.txtCondAsstNo2.value)<>"" Then
					Call changeTabs(TAB2)	 '~~~ 첫번째 Tab
					gSelframeFlg = TAB2
					Call DisplayMsgBox("970000","X",frm1.txtCondAsstNo2.alt,"X")
					frm1.txtCondAsstNo2.value=""
					frm1.txtCondAsstNm2.value=""
					frm1.txtCondAsstNo2.focus
					Set gActiveElement = document.activeElement
					bRet = False
				Else
					frm1.txtCondAsstNm2.value=Trim(Replace(lgF0,Chr(11),""))
				End If
		End if
	End if



	chkContentArea = bRet 
End Function

'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status


    'Call ggoOper.LockField(Document, "N")									     '⊙: This lock the suitable field

    '--------- Developer Coding Part (Start) ----------------------------------------------------------

    If gSelframeFlg = TAB2 Then
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")				     '☜: Data is changed.  Do you want to continue?
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
		Call ggoOper.LockField(Document, "N")									     '⊙: This lock the suitable field

		lgIntFlgMode = Parent.OPMD_CMODE												     '⊙: Indicates that current mode is Crate mode

		ggoSpread.Source= frm1.vspdData
		ggoSpread.ClearSpreadData



		frm1.vspdData.ReDraw = True

	Elseif  gSelframeFlg = TAB1 Then

		If lgIntFlgMode <> Parent.OPMD_UMODE Then
			lgIntFlgMode = Parent.OPMD_CMODE
		End If

		frm1.vspddata.ReDraw = False

		if frm1.vspddata.MaxRows < 1 then Exit Function

		ggoSpread.Source = frm1.vspddata
		ggoSpread.CopyRow
		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,frm1.txtDocCur.value,C_Amt,   "A" ,"I","X","X")

		Call vspddata_Change(C_RcptType, frm1.vspddata.ActiveRow)
		Call SetSpreadColor("I", frm1.vspddata.ActiveRow, frm1.vspddata.ActiveRow)

    	frm1.vspddata.Col = C_RcptType
		'frm1.vspddata.Text = ""



		frm1.vspddata.ReDraw = True


	End if

    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncCopy = True                                                                '☜: Processing is OK
End Function

'========================================================================================================
Function FncCancel()
	Dim varData
	Dim intIndex
	Dim strVal
	Dim varFlag
	FncCancel = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	If gSelframeFlg = TAB1 Then  'Master단 
	With frm1.vspddata

		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo

	End With

	End If

	'------ Developer Coding part (End )   --------------------------------------------------------------
	Set gActiveElement = document.ActiveElement
	FncCancel = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
Function FncInsertRow(byval pvRowCnt)
	FncInsertRow = False														 '☜: Processing is NG
	Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	If   gSelframeFlg = TAB1 Then        '''' Acq Item

		with frm1

			.vspddata.focus

			.vspdData.Row = .vspddata.ActiveRow
			.vspdData.Col = 0
			if (.vspdData.Text = ggoSpread.InsertFlag) OR (.vspdData.Text = ggoSpread.UpdateFlag) Or pvRowCnt = "" then
				CALL  DisplayMsgBox("971012","x","저장","x")

				FncInsertRow = False                                                          '☜: Processing is OK
				Exit Function
			end if

			ggoSpread.Source = .vspddata
			.vspddata.ReDraw = False
			ggoSpread.InsertRow ,1
			.vspddata.ReDraw = True
			Call SetSpreadColor ("I", .vspdData.ActiveRow, .vspdData.ActiveRow)

		end with

	END if

	lgBlnFlgChgValue = True
'	Call ggoOper.LockField(Document, "Q")
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Set gActiveElement = document.ActiveElement
	FncInsertRow = True                                                          '☜: Processing is OK
End Function


'========================================================================================================
Function FncDeleteRow()
	Dim lDelRows

	FncDeleteRow = False														 '☜: Processing is NG
	Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	If gSelframeFlg = TAB1 Then
		frm1.vspdData.focus
    	ggoSpread.Source = frm1.vspdData
		if frm1.vspdData.MaxRows < 1 then Exit Function

		ggoSpread.DeleteRow
	End If

	lgBlnFlgChgValue = True
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Set gActiveElement = document.ActiveElement
	FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrint()
	Parent.fncPrint()

End Function



'========================================================================================================
Function FncExcel()
	FncExcel = False                                                             '☜: Processing is NG
	Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(parent.C_SINGLE)

	FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncFind()
	FncFind = False                                                              '☜: Processing is NG
	Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(parent.C_SINGLE, True)

	FncFind = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False                                                              '☜: Processing is NG
	Err.Clear                                                                    '☜: Clear err status

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")	                 '⊙: Data is changed.  Do you want to exit?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	FncExit = True                                                               '☜: Processing is OK
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
' Function Name : PopRestoreSpreadColumnInf
' Description   :
'========================================================================================
Sub PopRestoreSpreadColumnInf()
' 	Dim varData
' 	Dim intIndex

' 	ggoSpread.Source = gActiveSpdSheet
' 	Call ggoSpread.RestoreSpreadInf()
' 	Call InitSpreadSheet()
' 	Call InitComboBox()
' 	Call ggoSpread.ReOrderingSpreadData()
' 	Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData,-1 , -1 ,frm1.txtDocCur.value ,C_Amt ,   "A" ,"I","X","X")

End Sub

'========================================================================================================
Function DbQuery()
	Dim strVal

	Err.Clear                                                                    '☜: Clear err status
	DbQuery = False                                                              '☜: Processing is NG

	If   LayerShowHide(1) = False Then
		Exit Function
	End If

	strVal = BIZ_PGM_ID & "?txtMode="          		& parent.UID_M0001                       '☜: Query
	strVal = strVal     & "&txtPrevNext="      		& ""	                             '☜: Direction
	strVal = strVal     & "&txtFacility_Accnt="     & Frm1.CbohFacility_Accnt.value                     '☜: Query Key
	strVal = strVal     & "&txtFacility_Cd="   		& Frm1.hFacility_Cd.value      '☜: Query Key
	strVal = strVal     & "&txtUse_Yn="     		& Frm1.CbohUse_Yn.value                     '☜: Query Key
	strVal = strVal     & "&lgStrPrevKey="  		& lgStrPrevKey
	strVal = strVal     & "&lgStrPrevKeyIndex="  	& lgStrPrevKeyIndex
	strVal = strVal     & "&txtMaxRows="         	& Frm1.vspdData.MaxRows         '☜: Max fetched data
	strVal = strVal     & "&lgPageNo="				& lgPageNo                          '☜: Next key tag
	strVal = strVal     & "&txtType="				& "A"                          '☜: Next key tag


	'------ Developer Coding part (Start)  --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	DbQuery = True                                                               '☜: Processing is OK
	Set gActiveElement = document.ActiveElement
End Function



'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbDtlQuery()
	Dim strVal


	DbDtlQuery = False

'Call LayerShowHide(1)
' 	frm1.vspdData.row = frm1.vspdData.Activerow
' 	frm1.hFacility_Accnt.value = ""
' 	frm1.vspdData.Col = C_FACILITY_CD
' 	frm1.hFacility_Cd.value = frm1.vspdData.text
' 	frm1.hUse_Yn.value = ""

	Err.Clear                                                               '☜: Protect system from crashing

	With frm1

	strVal = BIZ_PGM_ID & "?txtMode="          		& parent.UID_M0001                       '☜: Query
	strVal = strVal     & "&txtPrevNext="      		& ""	                             '☜: Direction
	strVal = strVal     & "&txtFacility_Accnt="     & Frm1.hFacility_Accnt.value                     '☜: Query Key
	strVal = strVal     & "&txtFacility_Cd="   		& Frm1.hFacility_Cd.value      '☜: Query Key
	strVal = strVal     & "&txtUse_Yn="     		& Frm1.hUse_Yn.value                     '☜: Query Key
	strVal = strVal     & "&lgStrPrevKey="  		& lgStrPrevKey
	strVal = strVal     & "&lgStrPrevKeyIndex="  	& lgStrPrevKeyIndex
	strVal = strVal     & "&txtMaxRows="         	& Frm1.vspdData.MaxRows         '☜: Max fetched data
	strVal = strVal     & "&lgPageNo="				& lgPageNo                          '☜: Next key tag
	strVal = strVal     & "&txtType="				& "B"                          '☜: Next key tag


	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	End With
	DbDtlQuery = True

End Function

'========================================================================================================
Function DbDtlQueryOk()														'☆: 조회 성공후 실행로직 
	Dim i
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	lgBlnFlgChgValue = False

End Function

'========================================================================================================
Function DbSave()
'     Dim lRow
'     Dim lGrpCnt
' 	Dim strVal, strDel
' 	Dim C_txtSetDt, C_txtPurDt, C_txtChk_End_dt, C_txtRep_End_dt, C_txtJng_End_dt, C_txtPm_dt, strYear , strMonth , strDay
'     DbSave = False

' 	if LayerShowHide(1) = false then
' 		Exit Function
' 	end if

'     strVal = ""
'     strDel = ""
'     lGrpCnt = 1

' 	With Frm1
' 		For lRow = 1 To .vspdData.MaxRows

' 			.vspdData.Row = lRow
' 			.vspdData.Col = 0

' 			Select Case .vspdData.Text

' 				Case  ggoSpread.InsertFlag                                      '☜: Update추가 
' 																  strVal = strVal & "C"  &  parent.gColSep
' 																  strVal = strVal & lRow &  parent.gColSep
' 					.vspdData.Col = C_FACILITY_ACCNT_CD			: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_FACILITY_CD				: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_SET_PLANT					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PROD_CO					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PUR_DT					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PROD_AMT					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_LIFE_CYCLE				: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_CHK_PRD1					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PIC_FLAG					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					 lGrpCnt = lGrpCnt + 1
' 				Case  ggoSpread.UpdateFlag                                      '☜: Update
' 																  strVal = strVal & "U"  &  parent.gColSep
' 																  strVal = strVal & lRow &  parent.gColSep
' 					.vspdData.Col = C_FACILITY_ACCNT_CD			: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_FACILITY_CD				: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_SET_PLANT					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PROD_CO					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PUR_DT					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PROD_AMT					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_LIFE_CYCLE				: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_CHK_PRD1					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PIC_FLAG					: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
' 					lGrpCnt = lGrpCnt + 1

' 				Case  ggoSpread.DeleteFlag                                      '☜: Delete

' 					                              				  strDel = strDel & "D"  &  parent.gColSep
' 					                              				  strDel = strDel & lRow &  parent.gColSep
' 					.vspdData.Col = C_FACILITY_ACCNT_CD			: strDel = strDel & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_FACILITY_CD				: strDel = strDel & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_SET_PLANT					: strDel = strDel & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PROD_CO					: strDel = strDel & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PUR_DT					: strDel = strDel & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PROD_AMT					: strDel = strDel & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_LIFE_CYCLE				: strDel = strDel & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_CHK_PRD1					: strDel = strDel & Trim(.vspdData.Text) &  parent.gColSep
' 					.vspdData.Col = C_PIC_FLAG					: strDel = strDel & Trim(.vspdData.Text) &  parent.gRowSep
' 					lGrpCnt = lGrpCnt + 1
' 			End Select
' 		Next


' 		lRow = .vspdData.ActiveRow
' 		.vspdData.Row = lRow
' 		.vspdData.Col = 0




' 		if (.vspdData.Text = ggoSpread.InsertFlag) OR (.vspdData.Text = ggoSpread.UpdateFlag) then
' 			Call ExtractDateFrom(.txtSetDt.Text,parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)
' 			C_txtSetDt = strYear & strMonth & strDay
' 			Call ExtractDateFrom(.txtPurDt.Text,parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)
' 			C_txtPurDt = strYear & strMonth & strDay
' 			Call ExtractDateFrom(.txtChk_End_dt.Text,parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)
' 			C_txtChk_End_dt = strYear & strMonth & strDay
' 			Call ExtractDateFrom(.txtRep_End_dt.Text,parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)
' 			C_txtRep_End_dt = strYear & strMonth & strDay
' 			Call ExtractDateFrom(.txtJng_End_dt.Text,parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)
' 			C_txtJng_End_dt = strYear & strMonth & strDay
' 			Call ExtractDateFrom(.txtPm_dt.Text,parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)
' 			C_txtPm_dt = strYear & strMonth & strDay


' 			strVal = strVal & Trim(.txtItemGroupCd1.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(.txtItemGroupCd2.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(.txtModel_Sts.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtPress_Power.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(.txtPlant_Sts.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtProd_Amt.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtCondAsstNo1.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(.txtCondAsstNo2.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(.txtSet_Place.Value) 	&  parent.gColSep
' 			strVal = strVal & Trim(.CboUse_Yn.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtSetCoCd.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(C_txtSetDt)				&  parent.gColSep
' 			strVal = strVal & Trim(.txtPurCoCd.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(C_txtPurDt)				&  parent.gColSep
' 			strVal = strVal & Trim(.txtProdCoCd.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtEquip_Area.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(.txtProdNo.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtUseVolt.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtProd_Flag.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtUse_Amount.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(.txtLife_Cycle.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(.txtMoter_Type.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(C_txtChk_End_dt)			&  parent.gColSep
' 			strVal = strVal & Trim(.txtOil_Spec1.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtChk_Prd1.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtMoter_qty.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(C_txtRep_End_dt)			&  parent.gColSep
' 			strVal = strVal & Trim(.txtOil_Spec2.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtChk_Prd2.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtMoter_Power.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(C_txtJng_End_dt)			&  parent.gColSep
' 			strVal = strVal & Trim(.txtOil_Spec3.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtEmp_no.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtMoter_Cir_Qty.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(C_txtPm_dt)				&  parent.gColSep
' 			strVal = strVal & Trim(.txtOil_Spec4.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtMoter_Bearing.Value)	&  parent.gColSep
' 			strVal = strVal & Trim(.txtPm_Reason.Value)		&  parent.gColSep
' 			strVal = strVal & Trim(.txtOil_Spec5.Value)		&  parent.gRowSep
' 		End if


'        .txtMode.value        =  parent.UID_M0002
' 	   .txtMaxRows.value     = lGrpCnt-1
' 	   .txtSpread.value      = strDel & strVal

' 	End With

' 	MSGBOX strVal & "     :     " & strDel
' 	cALL LayerShowHide(0)
' 	exit Function


' 	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
'     DbSave = True


	Err.Clear

	DbSave = False

	LayerShowHide(1)


	With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	End With

	DbSave = True

End Function

'========================================================================================================
Function DbDelete()
	Err.Clear

	DbDelete = False

	LayerShowHide(1)


	With frm1
		.txtMode.value = parent.UID_M0003
		.txtFlgMode.value = lgIntFlgMode
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	End With

	DbDelete = True
End Function

'========================================================================================================
Sub DbQueryOk()
	Dim iRow,intIndex
	Dim varData
	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

	'------ Developer Coding part (Start)  --------------------------------------------------------------
	lgOldRow = 1
    Call pvLockField("Q")
	Call DbDtlQuery
	'------ Developer Coding part (End )   --------------------------------------------------------------

	On error Resume Next
	Frm1.vspdData.Focus
	On error goto 0

End Sub


Sub DbQueryNotOk()
	Dim iRow,intIndex
	Dim varData
	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

	'------ Developer Coding part (Start)  --------------------------------------------------------------
	lgOldRow = 1
    Call pvLockField("Q")
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub DbSaveOk()
	On error Resume next


	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData

'     frm1.txtInsuerCd.value = frm1.txtInsuer_Cd1.value
'     frm1.txtInsuerNm.value = frm1.txtInsuerNm1.value


	Set gActiveElement = document.ActiveElement

	Call MakeKeyStream("Q")
	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call DbQuery()

End Sub

'========================================================================================================
' Name : DbDeleteOk
' Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------

	Call InitVariables()
	Call SetToolbar("1100000000011111")

	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call FncQuery()
End Sub

'========================================================================================
' Function Name : pvLockField
' Function Desc : ggoOperLockField 대용 
'========================================================================================
Function pvLockField(byVal pvFlag)
	If pvFlag = "Q" Then
		Call LockHTMLField(frm1.CboFacility_Accnt,"P")
		Call LockHTMLField(frm1.txtFacility_Cd,"P")
		Call LockHTMLField(frm1.txtItemGroupCd1,"P")
	ElseIf pvFlag = "N" Then
		Call LockHTMLField(frm1.CboFacility_Accnt,"R")
		Call LockHTMLField(frm1.txtFacility_Cd,"R")
		Call LockHTMLField(frm1.txtItemGroupCd1,"R")
	End If
End Function

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************


'------------------------------------------  OpenCondPlant()  --------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						
	arrParam(1) = "B_PLANT"								
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			
	arrParam(3) = ""									
	arrParam(4) = ""									
	arrParam(5) = "공장"							
	
    arrField(0) = "PLANT_CD"							
    arrField(1) = "PLANT_NM"							
    arrField(2) = "CAL_TYPE"							
    
    arrHeader(0) = "공장"							
    arrHeader(1) = "공장명"							
    arrHeader(2) = "칼렌다 타입"						
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function
'------------------------------------------  OpenConWC()  ------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConWC()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"					
	arrParam(1) = "P_WORK_CENTER"					
	arrParam(2) = Trim(frm1.txtSet_Place.Value)		
	arrParam(3) = ""								
	arrParam(4) = "P_WORK_CENTER.PLANT_CD =" & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "작업장"						
	
    arrField(0) = "WC_CD"							
    arrField(1) = "WC_NM"							
    arrField(2) = "CASE WHEN INSIDE_FLG='Y' THEN '사내' ELSE '외주' END"
    
    arrHeader(0) = "작업장"						
    arrHeader(1) = "작업장명"					
    arrHeader(2) = "작업장구분"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConWC(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtSet_Place.focus
	
End Function



 '------------------------------------------  OpenMasterRef()  -------------------------------------------------
'	Name : OpenMasterRef()
'	Description : Asset Master Condition PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenMasterRef(iWhere)

	Dim arrRet
	Dim arrParam(7)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	iCalledAspName = AskPRAspName("a7103ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7103ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
		IsOpenPop = False
		Exit Function
	Else
		Call SetPoRef(arrRet, iWhere)
	End If

	IsOpenPop = False
	lgBlnFlgChgValue = True

		With frm1
			Select Case iWhere
				Case 1
					.txtCondAsstNo1.focus
				Case 2
					.txtCondAsstNo2.focus
			End Select
		End With

End Function
'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	lgBlnFlgChgValue = True		
End Function

'------------------------------------------  SetConWC()  --------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConWC(byval arrRet)
	frm1.txtSet_Place.Value    = arrRet(0)		
	frm1.txtConWcNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True

End Function
 '------------------------------------------  SetPoRef()  -------------------------------------------------
'	Name : SetPoRef()
'	Description :
'---------------------------------------------------------------------------------------------------------
Sub SetPoRef(strRet, iWhere)
	if iWhere = 1 then
		frm1.txtCondAsstNo1.value     = strRet(0)
		frm1.txtCondAsstNm1.value	 = strRet(1)
	else
		frm1.txtCondAsstNo2.value     = strRet(0)
		frm1.txtCondAsstNm2.value	 = strRet(1)
	end if
End Sub
'========================================================================================================
' Name : OpenFacility_Popup()
' Desc : developer describe this line
'========================================================================================================
Function OpenFacility_Popup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then
		Exit Function
	End If

	IsOpenPop = True
	Select Case iWhere
		Case "1"
			arrParam(0) = "설비코드 팝업"
			arrParam(1) = "Y_FACILITY"
			arrParam(2) = frm1.txthFacility_Cd.value
			arrParam(3) = ""						            ' Name Cindition
			arrParam(4) = ""                                    ' Where Condition
			arrParam(5) = "설비코드"                        ' TextBox 명칭 

			arrField(0) = "Facility_cd"                     ' Field명(0)
			arrField(1) = "Facility_Nm"                     ' Field명(1)

			arrHeader(0) = "설비코드"                 ' Header명(0)
			arrHeader(1) = "설비코드명"                 ' Header명(1)
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
		Frm1.txthFacility_Cd.focus
		Exit Function
	Else
		Call SetCondArea(arrRet,iWhere)
	End If

End Function

'======================================================================================================
' Name : SetCondArea()
' Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetCondArea(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
			Case "1"
				.txthFacility_Cd.value = arrRet(0)
				.txthFacility_Nm.value = arrRet(1)
				.CbohUse_Yn.focus
		End Select
	End With
End Sub


'===========================================================================
' Function Name : OpenCurrency()
' Function Desc : OpenCurrency Reference Popup
'===========================================================================
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg

	If IsOpenPop = True Or UCase(frm1.txtDocCur.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래통화 팝업"	
	arrParam(1) = "B_CURRENCY"
	arrParam(2) = Trim(frm1.txtDocCur.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "거래통화"

    arrField(0) = "CURRENCY"
    arrField(1) = "CURRENCY_DESC"

    arrHeader(0) = "거래통화"		
    arrHeader(1) = "거래통화명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtDocCur.Value    = arrRet(0)		
		lgBlnFlgChgValue = True
	End If
End Function

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 

	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")


	IsOpenPop = False

	If arrRet(0) = "" Then
		With frm1
			Select Case iWhere
				Case 3
					.txtSetCoCd.focus
				Case 4
					.txtPurCoCd.focus
				Case 5
					.txtProdCoCd.focus
			End Select
		End With
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
		lgBlnFlgChgValue = True
	End If
End Function
'=======================================================================================================



'======================================================================================================
'	Name : SetBizArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere

			Case 3
				.txtSetCoCd.focus
				.txtSetCoCd.value = arrRet(0)
				.txtSetCoNm.value = arrRet(1)
			Case 4
				.txtPurCoCd.focus
				.txtPurCoCd.value = arrRet(0)
				.txtPurCoNm.value = arrRet(1)
			Case 5
				.txtProdCoCd.focus
				.txtProdCoCd.value = arrRet(0)
				.txtProdCoNm.value = arrRet(1)
		End Select
	End With

	IF iWhere <> 0 Then
		lgBlnFlgChgValue = True
	End If
End Function




'========================================================================================================
' Name : txtSetCoCd_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtSetCoCd_Onchange()
	Dim IntRetCd

	If  frm1.txtSetCoCd.value = "" Then
		frm1.txtSetCoNm.value=""
    Else
        IntRetCD= CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ","    BP_CD = " & FilterVar(Trim(frm1.txtSetCoCd.value),"''","S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			If IntRetCD=False And Trim(frm1.txtSetCoCd.value)<>"" Then
				Call DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
				frm1.txtSetCoCd.value=""
				frm1.txtSetCoNm.value=""
				frm1.txtSetCoCd.focus
				Set gActiveElement = document.activeElement
			Else
				frm1.txtSetCoNm.value=Trim(Replace(lgF0,Chr(11),""))
			End If
    End if
	lgBlnFlgChgValue = True
End Function


'==========================================  2.3.1 Tab Click 처리  =================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'===================================================================================================================
 '----------------  ClickTab1(): Header Tab처리 부분 (Header Tab이 있는 경우만 사용)  ----------------------------
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab
	gSelframeFlg = TAB1

	Call SetToolbar("1100000000011111")

End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function

	Call changeTabs(TAB2)	 '~~~ 첫번째 Tab
	gSelframeFlg = TAB2
	Call SetToolbar("1111100000111111")                                                     '☆: Developer must customize

 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Function
Function ClickTab3()

	If gSelframeFlg = TAB3 Then Exit Function

	Call changeTabs(TAB3)	 '~~~ 첫번째 Tab
	gSelframeFlg = TAB3
	Call SetToolbar("1111100000111111")                                                     '☆: Developer must customize

 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Function


'==========================================  2.3.2 날짜와 숫자의 변화 처리  =================================================
'	기능: 날짜, 숫자 변동 체크 및 달력 더블클릭 
'	설명:
'===================================================================================================================


'========================================================================================================
' Name : txtPress_Power_onChange
' Desc : developer describe this line
'========================================================================================================

Function txtPress_Power_Change()
		lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function txtPlant_Sts_OnChange()
		lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function txtProd_Amt_Change()
  	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function txtLife_Cycle_Change()
  	lgBlnFlgChgValue = True
End Function



'========================================================================================================
Function txtChk_Prd1_Change()
  	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function txtMoter_qty_Change()
  	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function txtChk_Prd2_Change()
  	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function txtMoter_Power_Change()
  	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Function txtMoter_Cir_Qty_Change()
  	lgBlnFlgChgValue = True
End Function


'========================================================================================================
'========================================================================================================

Sub txtSetDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtSetDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtSetDt.Focus
	End If
End Sub

Sub txtSetDt_change()
	If CheckDateFormat(frm1.txtSetDt.Text,parent.gDateFormat) Then
		lgBlnFlgChgValue = True
	End If
End Sub


Sub txtPurDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPurDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtPurDt.Focus
	End If
End Sub

Sub txtPurDt_change()
	If CheckDateFormat(frm1.txtPurDt.Text,parent.gDateFormat) Then
		lgBlnFlgChgValue = True
	End If
End Sub

Sub txtChk_End_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtChk_End_dt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtChk_End_dt.Focus
	End If
End Sub

Sub txtChk_End_dt_change()
	If CheckDateFormat(frm1.txtChk_End_dt.Text,parent.gDateFormat) Then
		lgBlnFlgChgValue = True
	End If
End Sub

Sub txtRep_End_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtRep_End_dt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtRep_End_dt.Focus
	End If
End Sub

Sub txtRep_End_dt_change()
	If CheckDateFormat(frm1.txtRep_End_dt.Text,parent.gDateFormat) Then
		lgBlnFlgChgValue = True
	End If
End Sub

Sub txtJng_End_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtJng_End_dt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtJng_End_dt.Focus
	End If
End Sub

Sub txtJng_End_dt_change()
	If CheckDateFormat(frm1.txtJng_End_dt.Text,parent.gDateFormat) Then
		lgBlnFlgChgValue = True
	End If
End Sub


Sub txtPm_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPm_dt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtPm_dt.Focus
	End If
End Sub

Sub txtPm_dt_change()
	If CheckDateFormat(frm1.txtPm_dt.Text,parent.gDateFormat) Then
		lgBlnFlgChgValue = True
	End If
End Sub


Sub CboFacility_Accnt_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub txtItemGroupCd1_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub txtItemGroupCd2_Onchange()
	lgBlnFlgChgValue = True
End Sub

Sub txtSet_Place_Onchange()
	lgBlnFlgChgValue = True
End Sub


Sub CboUse_Yn_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub txtProd_Flag_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub txtOil_Spec1_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub txtOil_Spec2_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub txtOil_Spec3_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub txtOil_Spec4_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub txtOil_Spec5_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub txtEmp_no_Onchange()
	lgBlnFlgChgValue = True
End Sub

Sub txtPlantCd_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtPlantCd.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("PT", "''", "S") & "", C_PopPlantCd) Then
				.txtPlantCd.value = ""
				.txtPlantNm.value = ""
				.txtPlantCd.focus
			End If
		Else
			.txtPlantNm.value = ""
		End If
	End With
' 	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop Then Exit Function

	IsOpenPop = True
	
	Select Case pvIntWhere
	
	Case C_PopPlantCd												
		iArrParam(1) = "B_PLANT"						
		iArrParam(2) = Trim(frm1.txtPlantCd.value)		
		iArrParam(3) = ""								
		iArrParam(4) = ""								
		iArrParam(5) = "공장"						
		
		iArrField(0) = "ED15" & Parent.gColSep & "PLANT_CD"						
		iArrField(1) = "ED30" & Parent.gColSep & "PLANT_NM"					
    
	    iArrHeader(0) = "공장"						
	    iArrHeader(1) = "공장명"					

		frm1.txtPlantCd.focus 
	
	End Select
 
	iArrParam(0) = iArrParam(5)							 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function
'========================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		
		Case C_PopPlantCd
			.txtPlantCd.value = pvArrRet(0) 
			.txtPlantNm.value = pvArrRet(1)

		End Select
	End With
	
	SetConPopup = True

End Function
'========================================================================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(5), iArrTemp
	
	GetCodeName = False
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		iArrRs(5) = iArrTemp(3)				' 계획기간 순번 
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		If err.number = 0 Then
			If lgBlnOpenedFlag Then
				GetCodeName = OpenConPopup(pvIntWhere)
			End If
		Else
			MsgBox err.Description
		End If
	End if
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<%
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
%>



<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
			    <TR>
				    <TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>설비제원조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>설비제원상세등록1</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>설비제원상세등록2</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH="100%"> </TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
					    <FIELDSET CLASS="CLSFLD">
						  <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
							<TD CLASS="TD5" NOWRAP>설비유형</TD>
							<TD CLASS="TD6" NOWRAP><SELECT NAME="CbohFacility_Accnt" ALT="설비유형" CLASS ="CbohFacility_Accnt" TAG="1XN"><OPTION VALUE=""></OPTION></SELECT></TD>
							<TD CLASS="TD5" NOWRAP>설비코드</TD>
							<TD CLASS="TD6" NOWRAP><INPUT ID=txthFacility_Cd NAME="txthFacility_Cd" ALT="설비코드" TYPE="Text" SiZE="18" MAXLENGTH="18" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenFacility_Popup('1')">
											       <INPUT ID=txthFacility_Nm NAME="txthFacility_Nm" ALT="설비코드명" TYPE="Text" SiZE="25" MAXLENGTH="40" tag="14"></TD>
							</TR>
							<TR>
							<TD CLASS="TD5" NOWRAP>사용유무</TD>
							<TD CLASS="TD6" NOWRAP><SELECT NAME="CbohUse_Yn" ALT="사용유무" CLASS ="CbohUse_Yn" TAG="11XXXU"><OPTION VALUE=""></OPTION><OPTION VALUE="Y">Y</OPTION><OPTION VALUE="N">N</OPTION></SELECT></TD>
							<TD CLASS="TD5" NOWRAP></TD>
							<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%" ></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP >

					<!-- 첫번째 탭 내용 -->

					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%" SCROLL="no">
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/p5110ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>

						</TABLE>
					</DIV>

					<!-- 두번째 탭 내용 -->
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%" SCROLL="no">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>설비유형</TD>
								<TD CLASS="TD6" NOWRAP ><SELECT NAME="CboFacility_Accnt" ALT="설비유형" CLASS ="CboFacility_Accnt" TAG="22XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>설비코드</TD>
								<TD CLASS="TD6" NOWRAP ><INPUT TYPE=TEXT NAME="txtFacility_Cd" SIZE=18 MAXLENGTH=18 tag="22XXXU" ALT="설비코드">&nbsp;<INPUT TYPE=TEXT NAME="txtFacility_Nm" SIZE=30 MAXLENGTH=40 tag="23XXXX" ALT="설비명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>대분류</TD>
								<TD CLASS="TD6" NOWRAP ><SELECT NAME="txtItemGroupCd1" ALT="대분류" CLASS ="txtItemGroupCd1" TAG="22XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>중분류</TD>
								<TD CLASS="TD6" NOWRAP ><SELECT NAME="txtItemGroupCd2" ALT="중분류" CLASS ="txtItemGroupCd2" TAG="21XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>모델 및 형식</TD>
						    	<TD CLASS="TD6" NOWRAP><INPUT NAME="txtModel_Sts" ALT="모델 및 형식" TYPE="Text" SiZE=20 MAXLENGTH=20 tag="21XXX"></TD>
								<TD CLASS="TD5" NOWRAP>능력</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/p5110ma1_txtPress_Power_txtPress_Power.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>설비등급</TD>
								<TD CLASS="TD6" NOWRAP ><SELECT NAME="txtPlant_Sts" ALT="설비등급" CLASS ="txtPlant_Sts" TAG="22XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>자산번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCondAsstNo1" ALT="자산번호"  SIZE=18 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMasterRef(1)"> <INPUT TYPE="Text" NAME="txtCondAsstNm1" SIZE=24 MAXLENGTH=30 tag="24" ALT="자산명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>자산번호2</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCondAsstNo2" ALT="자산번호2" SIZE=18 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="21XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMasterRef(2)"> <INPUT TYPE="Text" NAME="txtCondAsstNm2" SIZE=24 MAXLENGTH=30 tag="24" ALT="자산명"></TD>
								<TD CLASS="TD5" NOWRAP>설치공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="22XXXU" ALT="설치공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>설치라인</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSet_Place" SIZE=8 MAXLENGTH=7 tag="21XXXU" ALT="설치라인"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenConWC()"> <INPUT TYPE=TEXT NAME="txtConWcNm" SIZE=25 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP>사용유무</TD>
					            <TD CLASS="TD6" NOWRAP><SELECT NAME="CboUse_Yn" ALT="사용유무" CLASS ="CboUse_Yn" TAG="2XN"><OPTION VALUE=""></OPTION><OPTION VALUE="Y">Y</OPTION><OPTION VALUE="N">N</OPTION></SELECT></TD>
							</TR>


							<TR>
								<TD CLASS="TD5" NOWRAP>금액</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/p5110ma1_txtProd_Amt_txtProd_Amt.js'></script>
							    </TD>
								<TD CLASS="TD5" NOWRAP>통화</TD>
								<TD CLASS="TD6"><INPUT NAME="txtDocCur" ALT="통화" TYPE="Text" MAXLENGTH=3 SiZE=10  tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCurrency()"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>설치업체</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtSetCoCd" TYPE=TEXT SIZE=10 MAXLENGTH="10" TAG="21XXXU" ALT="설치업체"><IMG SRC="../../image/btnPopup.gif" NAME="btnCustomCd" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenBp(frm1.txtSetCoCd.value,3)">&nbsp;<INPUT TYPE=TEXT NAME="txtSetCoNm"  SIZE="25" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS="TD5" NOWRAP>설치일자</TD>
						    	<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/p5110ma1_txtSetDt_txtSetDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>구매업체</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPurCoCd" TYPE=TEXT SIZE=10 MAXLENGTH="10" TAG="21XXXU" ALT="구매업체"><IMG SRC="../../image/btnPopup.gif" NAME="btnCustomCd" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenBp(frm1.txtPurCoCd.value,4)">&nbsp;<INPUT TYPE=TEXT NAME="txtPurCoNm"  SIZE="25" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS="TD5" NOWRAP>구매일자</TD>
						    	<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/p5110ma1_txtPurDt_txtPurDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>제작업체</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtProdCoCd" TYPE=TEXT SIZE=10 MAXLENGTH="10" TAG="21XXXU" ALT="제작업체"><IMG SRC="../../image/btnPopup.gif" NAME="btnProdCoCd" align=top TYPE="BUTTON" ONCLICK ="vbscript:Call OpenBp(frm1.txtProdCoCd.value,5)">&nbsp;<INPUT TYPE=TEXT NAME="txtProdCoNm"  SIZE="25" MAXLENGTH="50" TAG="24"></TD>
								<TD CLASS="TD5" NOWRAP>설비면적</TD>
						    	<TD CLASS="TD6" NOWRAP><INPUT NAME="txtEquip_Area" ALT="설비면적" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="21XXX"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>제작번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtProdNo" TYPE=TEXT SIZE=30 MAXLENGTH="30" TAG="21XXX" ALT="제작번호"></TD>
								<TD CLASS="TD5" NOWRAP>사용전압</TD>
						    	<TD CLASS="TD6" NOWRAP><INPUT NAME="txtUseVolt" ALT="사용전압" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="21X4Z"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>제작구분</TD>
								<TD CLASS="TD6" NOWRAP ><SELECT NAME="txtProd_Flag" ALT="제작구분" CLASS ="txtProd_Flag" TAG="21XXX"><OPTION VALUE=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>전기용량</TD>
						    	<TD CLASS="TD6" NOWRAP><INPUT NAME="txtUse_Amount" ALT="전기용량" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="21XXX"></TD>
							</TR>
						</TABLE>
					</DIV>

					<!-- 세번째 탭 내용 -->
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%" SCROLL="no">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>수명</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/p5110ma1_txtLife_Cycle_txtLife_Cycle.js'></script>&nbsp;
								</TD>
								<TD CLASS="TD5" NOWRAP>모타형식</TD>
								<TD CLASS="TD6" NOWRAP>
									<INPUT NAME="txtMoter_Type" ALT="모타형식" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="21XXX">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>최종점검일</TD>
						    	<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/p5110ma1_txtChk_End_dt_txtChk_End_dt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>오일규격1</TD>
								<TD CLASS="TD6" NOWRAP ><SELECT NAME="txtOil_Spec1" ALT="오일규격1" CLASS ="txtOil_Spec1" TAG="21XXX"><OPTION VALUE=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>점검주기(주)</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/p5110ma1_txtChk_Prd1_txtChk_Prd1.js'></script>&nbsp;
								</TD>
								<TD CLASS="TD5" NOWRAP>모타용량</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/p5110ma1_txtMoter_qty_txtMoter_qty.js'></script>&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>최종수리일</TD>
						    	<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/p5110ma1_txtRep_End_dt_txtRep_End_dt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>오일규격2</TD>
								<TD CLASS="TD6" NOWRAP ><SELECT NAME="txtOil_Spec2" ALT="오일규격2" CLASS ="txtOil_Spec2" TAG="21XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>정도주기(주)</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/p5110ma1_txtChk_Prd2_txtChk_Prd2.js'></script>&nbsp;
								</TD>
								<TD CLASS="TD5" NOWRAP>모타전압</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/p5110ma1_txtMoter_Power_txtMoter_Power.js'></script>&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>최종정도일</TD>
						    	<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/p5110ma1_txtJng_End_dt_txtJng_End_dt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>오일규격3</TD>
								<TD CLASS="TD6" NOWRAP ><SELECT NAME="txtOil_Spec3" ALT="오일규격3" CLASS ="txtOil_Spec3" TAG="21XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
							</TR>
							<TR>
				                <TD CLASS="TD5" NOWRAP>담당자</TD>
								<TD CLASS="TD6" NOWRAP ><SELECT NAME="txtEmp_no" ALT="담당자" CLASS ="txtEmp_no" TAG="21XXX"><OPTION VALUE=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>모타회전수</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/p5110ma1_txtMoter_Cir_Qty_txtMoter_Cir_Qty.js'></script>&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>폐기매각일</TD>
						    	<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/p5110ma1_txtPm_dt_txtPm_dt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>오일규격4</TD>
								<TD CLASS="TD6" NOWRAP ><SELECT NAME="txtOil_Spec4" ALT="오일규격4" CLASS ="txtOil_Spec4" TAG="21XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>베어링</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtMoter_Bearing" TYPE=TEXT SIZE=10 MAXLENGTH="6" TAG="22XXX" ALT="베어링"></TD>
								<TD CLASS="TD5" NOWRAP>폐기매각사유</TD>
						    	<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPm_Reason" ALT="폐기매각사유" TYPE="Text" SiZE=50 MAXLENGTH=100 tag="21XXXX"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>오일규격5</TD>
								<TD CLASS="TD6" NOWRAP ><SELECT NAME="txtOil_Spec5" ALT="오일규격5" CLASS ="txtOil_Spec5" TAG="21XXXU"><OPTION VALUE=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP></TD>
						    	<TD CLASS="TD6" NOWRAP></TD>
						</TR>
						</TABLE>
					</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"  TABINDEX=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hFacility_Accnt" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hFacility_Cd" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hUse_Yn" tag="24" TabIndex="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


