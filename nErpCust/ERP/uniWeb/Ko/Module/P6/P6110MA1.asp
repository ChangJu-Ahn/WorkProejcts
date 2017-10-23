<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Mold Resources
*  2. Function Name        : 금형제원정보등록 
*  3. Program ID           : P6110Ma1
*  4. Program Name         : P6110Ma1
*  5. Program Desc         : 금형제원정보등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2005/01/17
*  8. Modified date(Last)  : 2005/01/17
*  9. Modifier (First)     : Lee Sang Ho
* 10. Modifier (Last)      : Lee Sang Ho
* 11. Comment              :
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
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js">			</SCRIPT>

<Script Language="VBScript">
Option Explicit                                  '☜: indicates that All variables must be declared in advance

<%'========================================================================================================%>

Const BIZ_PGM_ID    = "p6110MB1.asp"
Const BIZ_PGM_SAVE_ID = "P6110MB2.asp"
Const BIZ_PGM_QUERY_ID = "P6110MB3.asp"
Const C_SHEETMAXROWS = 100

<%'========================================================================================================%>
dim C_CAST_CD			'금형코드 
dim C_CAST_NM			'금형명 
dim C_SET_PLANT_CD		'설치공장 
dim C_SET_PLANT_NM		'설치공장명 
dim C_CAR_KIND			'적용모델(Item_Group)
dim C_CAR_KIND_NM		'적용모델명 
dim C_MAKE_DT			'제작일자 
dim C_STR_TYPE			'구조 및 형식 
dim C_CHECK_END_DT		'최종점검일 
dim C_ITEM_CD			'생산품목 
dim C_ITEM_NM			'생산품목명 
dim C_CLOSE_DT			'폐기일자 
dim C_PIC_FLAG          '금형사진유무 

Const TAB1 = 1			'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3

Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>
Dim BaseDate
BaseDate = "<%=GetSvrDate%>"
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

Sub InitSpreadPosVariables()

	C_CAST_CD		= 1
	C_CAST_NM		= 2
	C_SET_PLANT_CD	= 3
	C_SET_PLANT_NM	= 4
	C_CAR_KIND		= 5
	C_CAR_KIND_NM   = 6
	C_MAKE_DT		= 7
	C_STR_TYPE		= 8
	C_CHECK_END_DT	= 9
	C_ITEM_CD		= 10
	C_ITEM_NM		= 11
	C_CLOSE_DT		= 12
	C_PIC_FLAG		= 13

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
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)

	frm1.txtFinAjDt.Year = strYear 		 '년월일 default value setting
	frm1.txtFinAjDt.Month = strMonth
	frm1.txtFinAjDt.Day = strDay

  lgBlnFlgChgValue = False

End Sub

'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================


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

		ggoSpread.SSSetEdit 	C_CAST_CD,				"금형코드",		18, 0, True
		ggoSpread.SSSetEdit 	C_CAST_NM,				"금형코드명",	20, 0, True
	    ggoSpread.SSSetEdit		C_SET_PLANT_CD,       	"설치공장",		10,,,20,2
		ggoSpread.SSSetEdit		C_SET_PLANT_NM,			"설치공장",		15,,,20,2
		ggoSpread.SSSetEdit		C_CAR_KIND,  			"적용모델",		10,,,10,2
		ggoSpread.SSSetEdit		C_CAR_KIND_NM,			"적용모델명",	15,,,20,2
		ggoSpread.SSSetDate		C_MAKE_DT,				"제작일자",		12,2,parent.gDateFormat
		ggoSpread.SSSetEdit		C_STR_TYPE,				"구조 및 형식", 15,,,20,1
		ggoSpread.SSSetDate		C_CHECK_END_DT,			"최종점검일",	12,2,parent.gDateFormat
		ggoSpread.SSSetEdit 	C_ITEM_CD,				"생산품목",		18, 0, True
		ggoSpread.SSSetEdit 	C_ITEM_NM,				"생산품목명",	20, 0, True
		ggoSpread.SSSetDate		C_CLOSE_DT,				"최종점검일",	12,2,parent.gDateFormat
		ggoSpread.SSSetEdit		C_PIC_FLAG,				"사진유무",		12,		2,		true


		Call ggoSpread.SSSetColHidden(C_SET_PLANT_CD, C_SET_PLANT_CD, True)
		Call ggoSpread.SSSetColHidden(C_CAR_KIND, C_CAR_KIND, True)

		.ReDraw = true

    End With
    Call SetSpreadLock()

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()'byVal gird_fg, byVal lock_fg, byVal iRow)
    With frm1

		ggoSpread.Source = .vspddata
		.vspddata.ReDraw = False
		ggoSpread.SpreadLock		C_CAST_CD,	-1, C_PIC_FLAG		,-1

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
		ggoSpread.SSSetProtected	C_CAST_CD			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_SET_PLANT_CD		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_SET_PLANT_NM		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_CAR_KIND			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_CAR_KIND_NM		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_MAKE_DT			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_STR_TYPE			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_CHECK_END_DT		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ITEM_CD			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ITEM_NM			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_CLOSE_DT			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PIC_FLAG			, pvStartRow, pvEndRow

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
	Select Case pOpt
		Case "Q"
			frm1.htxtCastCd.value = frm1.txtCastCd.value
			lgKeyStream = Trim(Frm1.txtSetPlantCd.Value)				& parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.txtCarKind.value)		& parent.gColSep
			lgKeyStream = lgKeyStream & Trim(frm1.htxtCastCd.value)		& parent.gColSep
		Case "S"
			lgKeyStream = Trim(Frm1.htxtCastCd.Value)					& parent.gColSep
		Case "D"
			lgKeyStream = Trim(Frm1.htxtCastCd.Value)					& parent.gColSep
	End Select
    '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : InitComboBox
' Desc : This method init for ComBox
'========================================================================================================
'========================================================================================================
Sub InitComboBox()

    Call CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Y6002' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.CboEmpCd ,lgF0  ,lgF1  ,Chr(11))

    Call CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Y6001' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboPrsSts ,lgF0  ,lgF1  ,Chr(11))

	lgF0 = "Y" & Chr(11) & "N" & Chr(11)
	Call SetCombo2(frm1.cboUseYn, lgF0, lgF0, Chr(11))
End Sub

'========================================================================================================
' Name : vspddata_Change
' Desc : This method vspddata_change Event Process
'========================================================================================================
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

'==========================================================================================
'	Name : vspddata_Click()
'	Description : This Method is vspddata_Click Event Process
'==========================================================================================
Sub vspddata_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"					'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then
       Exit Sub									'If there is no data.
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


	frm1.vspdData.Col = C_CAST_CD
	frm1.vspdData.Row = row

	lgOldRow = Row

	Call InitData()
	Call LayerShowHide(1)
	frm1.vspdData.Col = C_CAST_CD
	frm1.htxtCastCd.value = frm1.vspdData.text
	Call DbQuery2(frm1.htxtCastCd.value)

End Sub

'========================================================================================================
'    Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'    Name : vspdData_DblClick
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName

	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub

'========================================================================================================
'    Name : vspdData_ScriptDragDropBlock
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    'If Col <= C_SNm Or NewCol <= C_SNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
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

	if NewRow > 0 and Row > 0 and Row <> NewRow then
		With Frm1
			.vspdData.Row = Row
			.vspdData.Col = 0
			if (.vspdData.Text = ggoSpread.InsertFlag) OR (.vspdData.Text = ggoSpread.UpdateFlag)  then
				CALL  DisplayMsgBox("971012","x","저장","x")
				Exit Sub
			end if

		End With
		exit sub
	End if
	If Not (Row <> NewRow And NewRow > 0) Then
	   Exit Sub
	End If
	ggoSpread.Source = Frm1.vspdData
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
				Call OpenPLANT(frm1.vspdData.Text, 1, Row)
			end if
		End If
	End With
End Sub


Sub vspddata_KeyPress(index , KeyAscii )
    lgBlnFlgChgValue = True                                                 '⊙: Indicates that value changed
End Sub

'==========================================================================================
'   Event Name : OpenPLANT
'   Event Desc :
'==========================================================================================
Function OpenPLANT(strCode, iWhere, Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	IsOpenPop = True

	arrParam(0) = "공장팝업"
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(strCode)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "공장"

    arrField(0) = "PLANT_CD"
    arrField(1) = "PLANT_NM"

    arrHeader(0) = "공장"
    arrHeader(1) = "공장명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant(arrRet, iWhere, Row)
	End If

End Function


'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetRcpt()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(arrRet, iWhere, Row)

	With frm1
        .vspdData.Row = Row
		Select Case iWhere
		    Case 1
		        .vspdData.Col = C_SET_PLANT
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_SET_PLANTNm
		    	.vspdData.text = arrRet(1)
				Call vspdData_Change(C_SET_PLANT, .vspdData.Row)
				Call SetActiveCell(frm1.vspdData,C_SET_PLANT,frm1.vspdData.ActiveRow ,"M","X","X")

		End Select

		lgBlnFlgChgValue = True

	End With
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Function


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
		.cboEmpCd.value = ""
		.cboPrsSts.value = ""
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
			C_CAST_CD		=	iCurColumnPos(1 )
			C_CAST_NM		=	iCurColumnPos(2 )
			C_SET_PLANT_CD	=	iCurColumnPos(3 )
			C_SET_PLANT_NM	=	iCurColumnPos(4 )
			C_CAR_KIND		=	iCurColumnPos(5 )
			C_CAR_KIND_NM   =   ICUrColumnPos(6 )
			C_MAKE_DT		=	iCurColumnPos(7 )
			C_STR_TYPE		= 	iCurColumnPos(8 )
			C_CHECK_END_DT	= 	iCurColumnPos(9 )
			C_ITEM_CD		=	iCurColumnPos(10)
			C_ITEM_NM		=	iCurColumnPos(11)
			C_CLOSE_DT		=	iCurColumnPos(12)
			C_PIC_FLAG		=	iCurColumnPos(13)

    End Select


End Sub

'========================================================================================================
'   Name : Form_Load()
'========================================================================================================
Sub Form_Load()

	Dim strYear
	Dim strMonth
	Dim strDay

    Err.Clear                                                                        '☜: Clear err status

	Call LoadInfTB19029                                                            '☜: Load table , B_numeric_format

    Call AppendNumberPlace("6", "15", "4")                                   'Format Numeric Contents Field
    Call AppendNumberPlace("7", "5", "0")                                   'Format Numeric Contents Field
    Call AppendNumberPlace("8", "15", "2")                                   'Format Numeric Contents Field
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
     Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

	Call ggoOper.LockField(Document, "N")                                            '⊙: Lock  Suitable  Field

	Call SetDefaultVal

	Call InitComboBox()
	Call InitData()

	Call InitVariables
	Call InitSpreadSheet()                                                               '⊙: Setup the Spread sheet

	Call SetToolbar("1100000000011111")

	gSelframeFlg = TAB1
	Call ClickTab1
	If parent.gPlant <> "" Then
		frm1.txtSetPlantCd.value = parent.gPlant
		frm1.txtSetPlantNm.value = parent.gPlantNm
		frm1.txtCarKind.focus
		Set gActiveElement = document.activeElement
	Else
		frm1.txtSetPlantCd.focus
		Set gActiveElement = document.activeElement
	End If

End Sub

'========================================================================================================
'   Name : FncQuery()
'========================================================================================================
Function FncQuery()
	Dim IntRetCD
	Dim var_m

	FncQuery = False															 '☜: Processing is NG
	Err.Clear

	If lgBlnFlgChgValue = True  Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X") '☜ "데이타가 변경되었습니다. 조회하시겠습니까?"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
	Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field

	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData

	Call InitData()
	Call SetDefaultVal()

	'--------- Developer Coding Part (Start) ----------------------------------------------------------


	Call InitVariables()
	 	                                                     '⊙: Initializes local global variables
	Call MakeKeyStream("Q")

	IF frm1.txtSetPlantCd.value <> "" THEN
		Call  CommonQueryRs(" plant_nm "," b_plant "," plant_cd = '" & frm1.txtSetPlantCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			Call DisplayMsgBox("971012", "X", "공장코드", "X")
			frm1.txtSetPlantCd.focus
			Set gActiveElement = document.ActiveElement
			frm1.txtSetPlantNm.value = ""
			Exit Function
		ELSE
			frm1.txtSetPlantNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtSetPlantNm.value = ""
	END IF
	
	IF frm1.txtCarKind.value <> "" THEN
		Call  CommonQueryRs(" ITEM_GROUP_NM "," B_ITEM_GROUP "," ITEM_GROUP_CD = '" & frm1.txtCarKind.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			Call DisplayMsgBox("971012", "X", "적용모델", "X")
			frm1.txtcarKind.focus
			Set gActiveElement = document.ActiveElement
			frm1.txtCarKindNm.value = ""
			Exit Function
		ELSE
			frm1.txtCarKindNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtCarKindNm.value = ""
	END IF

	IF frm1.txtCastCd.value <> "" THEN
		Call  CommonQueryRs(" cast_nm "," y_cast "," SET_PLANT = '" & frm1.txtSetPlantCd.value & "' AND cast_cd = '" & frm1.txtCastCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			Call DisplayMsgBox("971012", "X", "금형코드", "X")
			frm1.txtCastCd.focus
			Set gActiveElement = document.ActiveElement
			frm1.txtCastNm.value = ""
			Exit Function
		ELSE
			frm1.txtCastNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtCastNm.value = ""
	END IF

	If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
		Exit Function
	End If

	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData

	Call clickTab1()
	If DbQuery = False Then
	Exit Function
	End If

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


	Call ggoOper.ClearField(Document, "2")                                        '☜: Clear Condition Field
	Call ggoOper.LockField(Document , "N")                                        '☜: Lock  Field

	ggoSpread.Source = frm1.vspddata
	ggoSpread.ClearSpreadData

	Call SetToolbar("1111100000101111")

	Call SetDefaultVal()

	Call InitData

	Call InitVariables                                                            '⊙: Initializes local global variables

	Call ClickTab2()
	frm1.txtCastCd1.focus
	Set gActiveElement = document.ActiveElement

	frm1.txtSetPlantCd.value = ""
	frm1.txtSetPlantNm.value = ""
	frm1.txtCarKind.value = ""
	frm1.txtCarKindNm.value = ""
	frm1.txtCastCd.value = ""
	frm1.txtCastNm.value = ""

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

    IF IsNull(frm1.txtPurCurCd.value) OR Trim(frm1.txtPurCurCd.value) = "" THEN
			frm1.txtPurCurCd.value = ""
    ELSE
			Call  CommonQueryRs(" CURRENCY "," B_CURRENCY "," CURRENCY = '" & frm1.txtPurCurCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
				Call DisplayMsgBox("971012", "X", "통화", "X")
				Call changeTabs(TAB2)
				frm1.txtPurCurCd.focus
				Set gActiveElement = document.ActiveElement
				Exit Function
			ELSE
			   frm1.txtPurCurCd.value = left(lgF0, len(lgF0) -1)
			END IF
		END IF


    Call MakeKeyStream("S")
	'------ Developer Coding part (End )   --------------------------------------------------------------

    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If
    Set gActiveElement = document.ActiveElement
    FncSave = True                                                               '☜: Processing is OK
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

		'Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
		frm1.txtCastCd.value = ""
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
		Call ReFormatSpreadCellByCellByCurrency2(Frm1.vspdData,frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow,frm1.txtTradeCur.value,C_Amt,   "A" ,"I","X","X")

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


'========================================================================================================
Function DbQuery()
    Dim strVal

    Err.Clear                                                                    '☜: Clear err status
    DbQuery = False                                                              '☜: Processing is NG

    If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	frm1.htxtCastCd.value = frm1.txtCastCd.value

    strVal = BIZ_PGM_ID & "?txtMode="          		& parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtPrevNext="      		& ""	                      '☜: Direction
    strVal = strVal     & "&txtSetPlantCd="         & Frm1.txtSetPlantCd.value                     '☜: Query Key
    strVal = strVal     & "&txtCarKind="   			& Frm1.txtCarKind.value			'☜: Query Key
    strVal = strVal     & "&txtCastCd="     		& Frm1.htxtCastCd.value                     '☜: Query Key
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

'========================================================================================================
Function DbQuery2(ByVal lCastCd)
    Dim strVal
    Dim StrCastCd


    Err.Clear                                                                    '☜: Clear err status
    DbQuery2 = False                                                              '☜: Processing is NG

    If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_QUERY_ID & "?txtMode="          		& parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtPrevNext="      		& ""							'☜: Direction
    strVal = strVal     & "&txtCastCd="     		& lCastCd			            '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey="  		& lgStrPrevKey
    strVal = strVal     & "&lgStrPrevKeyIndex="  	& lgStrPrevKeyIndex
    strVal = strVal     & "&txtMaxRows="         	& Frm1.vspdData.MaxRows         '☜: Max fetched data
	strVal = strVal     & "&lgPageNo="				& lgPageNo                          '☜: Next key tag
	strVal = strVal     & "&txtType="				& "A"                          '☜: Next key tag


	'------ Developer Coding part (Start)  --------------------------------------------------------------
   '------ Developer Coding part (End )   --------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery2 = True

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function DbSave()

    Err.Clear

	DbSave = False

	LayerShowHide(1)

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)

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
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)

	End With

    DbDelete = True
End Function

'========================================================================================================
Sub DbQueryOk()
	Dim iRow,intIndex
	Dim varData

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

' 	'------ Developer Coding part (Start)  --------------------------------------------------------------
	lgOldRow = 1
	frm1.vspdData.Row = 1
	frm1.vspdData.Col = C_CAST_CD
	frm1.htxtCastCd.value = frm1.vspdData.text
	Call DbQuery2(frm1.htxtCastCd.value)
' 	'------ Developer Coding part (End )   --------------------------------------------------------------


End Sub

Sub DbQueryOk2()

	Dim iRow,intIndex
	Dim varData
	Call  ggoOper.LockField(Document, "Q")
	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
' 	'------ Developer Coding part (Start)  --------------------------------------------------------------
    lgOldRow = 1

    lgBlnFlgChgValue = False

' 	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub DbSaveOk()
	On error Resume next

    ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData

	Call InitVariables()

    Set gActiveElement = document.ActiveElement

    Call MakeKeyStream("Q")
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Call FncQuery()

End Sub

'========================================================================================================
' Name : DbDeleteOk
' Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()

	'------ Developer Coding part (Start)  --------------------------------------------------------------
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData

	Call InitVariables()
	call ClickTab1()

	 Set gActiveElement = document.ActiveElement

    Call MakeKeyStream("Q")
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Call FncQuery()

	'------ Developer Coding part (End )   --------------------------------------------------------------

End Sub

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************
'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenSetPlant()
'	Description : Condition SetPlant PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenSetPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtSetPlantCd.Value)		' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)		' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "공장"							' TextBox 명칭 

    arrField(0) = "PLANT_CD"							' Field명(0)
    arrField(1) = "PLANT_NM"							' Field명(1)

    arrHeader(0) = "공장"							' Header명(0)
    arrHeader(1) = "공장명"							' Header명(1)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtSetPlantCd.focus

End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenSetPlant1()
'	Description : Condition SetPlant PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenSetPlant1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtSetPlantCd1.Value)		' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)		' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "공장"							' TextBox 명칭 

    arrField(0) = "PLANT_CD"							' Field명(0)
    arrField(1) = "PLANT_NM"							' Field명(1)

    arrHeader(0) = "공장"							' Header명(0)
    arrHeader(1) = "공장명"							' Header명(1)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant1(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtSetPlantCd1.focus

End Function

'------------------------------------------  OpenCast()  ------------------------------------------------
'	Name : OpenCast()
'	Description : Cast PopUp
'------------------------------------------------------------------------------------------------------------
Function OpenCast()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	IF frm1.txtSetPlantCd.value <> "" THEN
		Call  CommonQueryRs(" plant_nm "," b_plant "," plant_cd = '" & frm1.txtSetPlantCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			frm1.txtSetPlantNm.value = ""
			IsOpenPop = False
			Call DisplayMsgBox("971012", "X", "공장코드", "X")
			frm1.txtSetPlantCd.focus
			Set gActiveElement = document.ActiveElement
			Exit Function
		ELSE
			frm1.txtSetPlantNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtSetPlantNm.value = ""
		IsOpenPop = False
		Call DisplayMsgBox("971012", "X", "공장코드", "X")
		frm1.txtSetPlantCd.focus
		Set gActiveElement = document.ActiveElement
		Exit Function
	END IF 

		arrParam(0) = "금형코드"								' 팝업 명칭 
		arrParam(1) = "Y_CAST"											' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtCastCd.Value)		' Code Condition
		arrParam(3) = ""														' Name Cindition
		arrParam(4) = "SET_PLANT = " & FilterVar(frm1.txtSetPlantCd.value, "''", "S")								' Where Condition
		arrParam(5) = "금형코드"								' TextBox 명칭 

    arrField(0) = "ED15" & parent.gcolsep & "CAST_CD"							' Field명(0)
    arrField(1) = "ED15" & parent.gcolsep & "CAST_NM"							' Field명(1)
    arrField(2) = "ED20" & parent.gcolsep & "(SELECT ITEM_GROUP_NM FROM B_ITEM_GROUP WHERE ITEM_GROUP_CD = CAR_KIND )"						' Field명(2)
    arrField(3) = "ED20" & parent.gcolsep & "(SELECT ITEM_NM FROM B_ITEM WHERE ITEM_CD = ITEM_CD_1 )"						' Field명(3)
    arrField(4) = "F3"   & parent.gcolsep & "EXT1_QTY"						' Field명(4)

    arrHeader(0) = "금형코드"					' Header명(0)
    arrHeader(1) = "금형코드명"					' Header명(1)
    arrHeader(2) = "모델명"						' Header명(2)
    arrHeader(3) = "품목명"						' Header명(3)
    arrHeader(4) = "차수"						' Header명(4)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=800px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCast(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtCastCd.focus
End Function

'------------------------------------------  OpenCarKind()  -------------------------------------------------
'	Name : OpenCarKind()
'	Description : Condition CarKind PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenCarKind()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "적용모델"						' 팝업 명칭 
	arrParam(1) = "B_ITEM_GROUP"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCarKind.Value)			' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "적용모델"						' TextBox 명칭 

    arrField(0) = "ITEM_GROUP_CD"						' Field명(0)
    arrField(1) = "ITEM_GROUP_NM"						' Field명(1)

    arrHeader(0) = "적용모델"						' Header명(0)
    arrHeader(1) = "적용모델명"						' Header명(1)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCarKind(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtCarKind.focus
End Function

'------------------------------------------  OpenCarKind()  -------------------------------------------------
'	Name : OpenCarKind1()
'	Description : Condition CarKind PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenCarKind1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "적용모델"						' 팝업 명칭 
	arrParam(1) = "B_ITEM_GROUP"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCarKind1.Value)			' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "적용모델"						' TextBox 명칭 

    arrField(0) = "ITEM_GROUP_CD"						' Field명(0)
    arrField(1) = "ITEM_GROUP_NM"						' Field명(1)

    arrHeader(0) = "적용모델"						' Header명(0)
    arrHeader(1) = "적용모델명"						' Header명(1)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCarKind1(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtCarKind.focus

End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	If frm1.txtSetPlantCd1.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtSetPlantCd1.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If

	IsOpenPop = True

	Select Case iWhere
		Case "1"
			arrParam(0) = Trim(frm1.txtSetPlantCd1.value)	' Plant Code
			arrParam(1) = Trim(frm1.txtItemCd1.value)	' Item Code
	    Case "2"
			arrParam(0) = Trim(frm1.txtSetPlantCd1.value)	' Plant Code
			arrParam(1) = Trim(frm1.txtItemCd2.value)	' Item Code

	    Case "3"
			arrParam(0) = Trim(frm1.txtSetPlantCd1.value)	' Plant Code
			arrParam(1) = Trim(frm1.txtItemCd3.value)	' Item Code
	    Case "4"
			arrParam(0) = Trim(frm1.txtSetPlantCd1.value)	' Plant Code
			arrParam(1) = Trim(frm1.txtItemCd4.value)	' Item Code
	    Case "5"
			arrParam(0) = Trim(frm1.txtSetPlantCd1.value)	' Plant Code
			arrParam(1) = Trim(frm1.txtItemCd5.value)	' Item Code
	    Case "6"
			arrParam(0) = Trim(frm1.txtSetPlantCd1.value)	' Plant Code
			arrParam(1) = Trim(frm1.txtItemCd6.value)	' Item Code
	    Case "7"
			arrParam(0) = Trim(frm1.txtSetPlantCd1.value)	' Plant Code
			arrParam(1) = Trim(frm1.txtItemCd7.value)	' Item Code
	    Case "8"
			arrParam(0) = Trim(frm1.txtSetPlantCd1.value)	' Plant Code
			arrParam(1) = Trim(frm1.txtItemCd8.value)	' Item Code
	    Case "9"
			arrParam(0) = Trim(frm1.txtSetPlantCd1.value)	' Plant Code
			arrParam(1) = Trim(frm1.txtItemCd9.value)	' Item Code
	    Case "10"
			arrParam(0) = Trim(frm1.txtSetPlantCd1.value)	' Plant Code
			arrParam(1) = Trim(frm1.txtItemCd10.value)	' Item Code


	End Select

	arrParam(2) = "12!MO"						' Combo Set Data:"1029!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value

	arrField(0) = 1								'"ITEM_CD"
	arrField(1) = 2								'"ITEM_NM"

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemCd(iwhere, arrRet)
	End If

End Function

'------------------------------------------  OpenSetPlace()  -------------------------------------------------
'	Name : OpenSetPlace()
'	Description : SetPlacePopup
'---------------------------------------------------------------------------------------------------------
Function OpenSetPlace()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	If frm1.txtSetPlantCd1.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtSetPlantCd1.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	IsOpenPop = True

	arrParam(0) = "작업장팝업"
	arrParam(1) = "P_WORK_CENTER"
	arrParam(2) = frm1.txtSetPlace.value
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtSetPlantCd1.value, "''", "S") & " AND VALID_TO_DT >=  " & FilterVar(BaseDate , "''", "S") & ""
	arrParam(5) = "작업장"

    arrField(0) = "WC_CD"
    arrField(1) = "WC_NM"
    arrField(2) = "HH" & parent.gcolsep & "INSIDE_FLG"
    arrField(3) = "CASE WHEN INSIDE_FLG=" & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("사내", "''", "S") & " ELSE " & FilterVar("외주", "''", "S") & " END"
    arrField(4) = "dbo.ufn_GetCodeName(" & FilterVar("P1013", "''", "S") & ", WC_MGR)"

    arrHeader(0) = "작업장"
    arrHeader(1) = "작업장명"
    arrHeader(2) = "작업장구분"
    arrHeader(3) = "작업장구분"
    arrHeader(4) = "작업장담당자"

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlace(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtSetPlace.focus()

End Function


'------------------------------------------  OpenAsstCd()  -------------------------------------------------
'	Name : OpenAsstCd()
'	Description : Asset Master Condition PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenAsstCd(iWhere)

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
		Call SetAsstCd(arrRet, iWhere)
	End If

	IsOpenPop = False

		With frm1
			Select Case iWhere
				Case 1
					.txtCondAsstCd1.focus()
				Case 2
					.txtCondAsstCd2.focus()
			End Select
		End With

End Function

'------------------------------------------  OpenAsstCd()  -------------------------------------------------
'	Name : OpenAsstCd()
'	Description : Asset Master Condition PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenAsstCd(iWhere)

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
		Call SetAsstCd(arrRet, iWhere)
	End If

	IsOpenPop = False

		With frm1
			Select Case iWhere
				Case 1
					.txtAsstCd1.focus
				Case 2
					.txtAsstCd2.focus
			End Select
		End With

End Function


'------------------------------------------  OpenCurCd()  ---------------------------------------------
'	Name : OpenCurrency()
'	Description : Currency Popup
'---------------------------------------------------------------------------------------------------------
Function OpenCurCd()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "통화팝업"
	arrParam(1) = "B_CURRENCY"
	arrParam(2) = Ucase(Trim(frm1.txtPurCurCd.value))
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "통화"

    arrField(0) = "CURRENCY"
    arrField(1) = "CURRENCY_DESC"


    arrHeader(0) = "통화"
    arrHeader(1) = "통화명"


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCurrency(arrRet)
	End If

End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetCast()
'	Description : Cast POPUP에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Function SetCast(byval arrRet)
	frm1.txtCastCd.Value    = arrRet(0)
	frm1.txtCastNm.Value    = arrRet(1)

End Function

'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetCast1()
'	Description : Cast POPUP에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Function SetCast1(byval arrRet)
	frm1.txtCastCd1.Value    = arrRet(0)
	frm1.txtCastNm1.Value    = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetPlant()
'	Description : Condition SetPlant Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtSetPlantCd.Value    = arrRet(0)
	frm1.txtSetPlantNm.Value    = arrRet(1)

End Function

'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetPlant1()
'	Description : Condition SetPlant Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetPlant1(byval arrRet)
	frm1.txtSetPlantCd1.Value    = arrRet(0)
	frm1.txtSetPlantNm1.Value    = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetCarKind()
'	Description : Condition CarKind Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetCarKind(byval arrRet)
	frm1.txtCarKind.Value    = arrRet(0)
	frm1.txtCarKindNm.Value  = arrRet(1)

End Function

'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetCarKind1()
'	Description : Condition CarKind Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetCarKind1(byval arrRet)
	frm1.txtCarKind1.Value    = arrRet(0)
	frm1.txtCarKindNm1.Value  = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetSetPlace()  --------------------------------------------------
'	Name : SetPlace()
'	Description : SetPlace Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetPlace(Byval arrRet)
	frm1.txtSetPlace.value = arrRet(0)
	frm1.txtSetPlaceNm.value = arrRet(1)
	lgBlnFlgChgValue = True
End Function


'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(iWhere, ByRef arrRet)

    With frm1

    Select Case iWhere
		Case "1"
			.txtItemCd1.value = arrRet(0)
			.txtItemnm1.value = arrRet(1)
			.txtItemCd1.focus
		Case "2"
			.txtItemCd2.value = arrRet(0)
			.txtItemnm2.value = arrRet(1)
			.txtItemCd2.focus
		Case "3"
			.txtItemCd3.value = arrRet(0)
			.txtItemnm3.value = arrRet(1)
			.txtItemCd3.focus
		Case "4"
			.txtItemCd4.value = arrRet(0)
			.txtItemnm4.value = arrRet(1)
			.txtItemCd4.focus
		Case "5"
			.txtItemCd5.value = arrRet(0)
			.txtItemnm5.value = arrRet(1)
			.txtItemCd5.focus
		Case "6"
			.txtItemCd6.value = arrRet(0)
			.txtItemnm6.value = arrRet(1)
			.txtItemCd6.focus
		Case "7"
			.txtItemCd7.value = arrRet(0)
			.txtItemnm7.value = arrRet(1)
			.txtItemCd7.focus
		Case "8"
			.txtItemCd8.value = arrRet(0)
			.txtItemnm8.value = arrRet(1)
			.txtItemCd8.focus
		Case "9"
			.txtItemCd9.value = arrRet(0)
			.txtItemnm9.value = arrRet(1)
			.txtItemCd9.focus
		Case "10"
			.txtItemCd10.value = arrRet(0)
			.txtItemnm10.value = arrRet(1)
			.txtItemCd10.focus
	End Select

	lgBlnFlgChgValue = True
	Set gActiveElement = document.activeElement
    End With

End Function

 '------------------------------------------  SetAsstCd()  -------------------------------------------------
'	Name : SetAsstCd()
'	Description :
'---------------------------------------------------------------------------------------------------------
Sub SetAsstCd(strRet, iWhere)
    if iWhere = 1 then
		frm1.txtAsstCd1.value    = strRet(0)
		frm1.txtAsstNm1.value	 = strRet(1)
	else
		frm1.txtAsstCd2.value    = strRet(0)
		frm1.txtAsstNm2.value	 = strRet(1)
	end if
	lgBlnFlgChgValue = True
End Sub


'------------------------------------------  SetCurrency()  ----------------------------------------------
'	Name : SetCurrency()
'	Description : RoutingNo Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCurrency(Byval arrRet)
	With frm1

		.txtPurCurCd.value = UCase(arrRet(0))

		lgBlnFlgChgValue = True

	End With
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
    frm1.txtSetPlantCd.focus()
End Function
 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function ClickTab3()

	If gSelframeFlg = TAB3 Then Exit Function

	Call changeTabs(TAB3)	 '~~~ 첫번째 Tab
	gSelframeFlg = TAB3
	Call SetToolbar("1111100000111111")                                                     '☆: Developer must customize
	frm1.txtSetPlantCd.focus()
End Function

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
'==========================================  2.3.2 날짜와 숫자의 변화 처리  =================================================
'	기능: 날짜, 숫자 변동 체크 및 달력 더블클릭 
'	설명:
'===================================================================================================================

'=======================================================================================================
'   Event Name : Date_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtMakeDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtMakeDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtMakeDt.Focus
		lgBlnFlgChgValue = True
    End If
End Sub

Sub txtCloseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtCloseDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtCloseDt.Focus
		lgBlnFlgChgValue = True
    End If
End Sub

Sub txtCheckEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtCheckEndDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtCheckEndDt.Focus
		lgBlnFlgChgValue = True
    End If
End Sub

Sub txtRepEndDT_DblClick(Button)
    If Button = 1 Then
        frm1.txtRepEndDT.Action = 7
        SetFocusToDocument("M")
		Frm1.txtRepEndDT.Focus
		lgBlnFlgChgValue = True
    End If
End Sub

Sub txtFinAjDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFinAjDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtFinAjDt.Focus
		lgBlnFlgChgValue = True
    End If
End Sub

'=======================================================================================================
'   Event Name : OnChange & Change Event
'   Event Desc : 달력을 호출한다.
'=======================================================================================================

Sub txtMakeDt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCloseDt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCheckEndDt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtRepEndDT_Change()
	lgBlnFlgChgValue = True
 End Sub

Sub txtFinAjDt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtWeightT_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCushionPr_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCarKind1_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCstroke_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPurAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPurCurCd_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtLifeCycle_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtSHeight_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtDHeight_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPersonCount_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtInspPrid_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCurAccnt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtFinCurAccnt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPrsUnit_Change()
	lgBlnFlgChgValue = True
End Sub

Sub cboEmpCd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtLimitAccnt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub cboUseYn_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub cboPrsSts_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub CustomYn_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtlrFlag_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtSetPlace_Change()
	lgBlnFlgChgValue = True
End Sub



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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>금형제원조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>금형제원상세등록1</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>금형제원상세등록2</font></td>
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
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSetPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSetPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSetPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtSetPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>적용모델</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCarKind" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="적용모델"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCarKind" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCarKind()">&nbsp;<INPUT TYPE=TEXT NAME="txtCarKindNm" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>금형코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCastCd"  SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="금형코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCastCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCast()">&nbsp;<INPUT TYPE=TEXT NAME="txtCastNm" SIZE=20 tag="14" ALT="금형코드명"></TD>
									<TD CLASS=TD5 NOWRAP>
									<TD CLASS=TD6 NOWRAP>
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
									<script language =javascript src='./js/p6110ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>

						</TABLE>
					</DIV>

					<!-- 두번째 탭 내용 -->
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%" SCROLL="no">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>금형코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCastCd1"  SIZE=18 MAXLENGTH=18 tag="23XXXU" ALT="금형코드">&nbsp;<INPUT TYPE=TEXT NAME="txtCastNm1" SIZE=20 MAXLENGTH=40 tag="22xxx" ALT="금형코드명"></TD>
								<TD CLASS=TD5 NOWRAP>제작업체</TD>
						    	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMaker" ALT="제작업체" TYPE="Text" SiZE=30 MAXLENGTH=30 tag="21XXXX"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>설치공장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSetPlantCd1" SIZE=6 MAXLENGTH=4 tag="22xxxU" ALT="설치공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSetPlantCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSetPlant1()">&nbsp;<INPUT TYPE=TEXT NAME="txtSetPlantNm1" SIZE=25 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>고객지급여부</TD>
								<TD CLASS=TD6 NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="CustomYn"         ID="CustomYn1" VALUE="1" tag="12"><LABEL FOR="CustomYn1">유</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="CustomYn" CHECKED ID="CustomYn2" VALUE="2" tag="12"><LABEL FOR="CustomYn2">무</LABEL></SPAN>
								<TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>적용모델</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCarKind1" SIZE=10 MAXLENGTH=10 tag="22xxxU" ALT="적용모델"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCarKind1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCarKind1()">&nbsp;<INPUT TYPE=TEXT NAME="txtCarKindNm1" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>LR여부</TD>
								<TD CLASS=TD6 NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="LrFlag"         ID="LrFlag1" VALUE="1" tag="12"><LABEL FOR="LrFlag1">L</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="LrFlag" CHECKED ID="LrFlag2" VALUE="2" tag="12"><LABEL FOR="LrFlag2">R</LABEL></SPAN>
								<TD>
							<TR>
								<TD CLASS=TD5 NOWRAP>작업장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtSetPlace" ALT="작업장"  SIZE=10 MAXLENGTH=7 STYLE="TEXT-ALIGN: Left" tag="22xxxU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSetPlace" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenSetPlace()"> <INPUT TYPE="Text" NAME="txtSetPlaceNm" SIZE=24 MAXLENGTH=30 tag="24" ALT="작업장"></TD>
								<TD CLASS=TD5 NOWRAP>사용유무</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboUseYn" tag="22xxxU" STYLE="WIDTH: 150px;" ALT="사용유무" ><OPTION value=""></OPTION></SELECT></TD></TD>
							<TR>
								<TD CLASS=TD5 NOWRAP>수명</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/p6110ma1_txtLifeCycle_txtLifeCycle.js'></script>
											</TD>
											<TD valign=bottom>&nbsp;년</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>담당자</TD>
								<TD CLASS="TD6"><SELECT NAME="cboEmpCd" tag="21X" STYLE="WIDTH: 150px;"><OPTION value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>등급</TD>
								<TD CLASS="TD6"><SELECT NAME="cboPrsSts" tag="21X" STYLE="WIDTH: 150px;"><OPTION value=""></OPTION></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>구매금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/p6110ma1_txtPurAmt_txtPurAmt.js'></script>
											</TD>
											<TD>&nbsp;<INPUT NAME="txtPurCurCd" ALT="통화" TYPE="Text" SiZE=5 MAXLENGTH=3 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSetCurCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCurCd()">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>자산코드1</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAsstCd1" ALT="자산코드1" SIZE=18 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAsstCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenAsstCd(1)"> <INPUT TYPE="Text" NAME="txtAsstNm1" SIZE=24 MAXLENGTH=30 tag="24" ALT="자산명1"></TD>
								<TD CLASS="TD5" NOWRAP>자산코드2</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAsstCd2" ALT="자산코드2" SIZE=18 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="21XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAsstCd2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenAsstCd(2)"> <INPUT TYPE="Text" NAME="txtAsstNm2" SIZE=24 MAXLENGTH=30 tag="24" ALT="자산명2"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>생산품목1</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1"  SIZE=18 MAXLENGTH=18 tag="21XXXU" ALT="생산품목1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenitemCd(1)">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=20 tag="24" ALT="생산품목명1"></TD>
								<TD CLASS=TD5 NOWRAP>생산품목2</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd2"  SIZE=18 MAXLENGTH=18 tag="21XXXU" ALT="생산품목2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenitemCd(2)">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm2" SIZE=20 tag="24" ALT="생산품목명2"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>생산품목3</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd3"  SIZE=18 MAXLENGTH=18 tag="21XXXU" ALT="생산품목3"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd3" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenitemCd(3)">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm3" SIZE=20 tag="24" ALT="생산품목명3"></TD>
								<TD CLASS=TD5 NOWRAP>생산품목4</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd4"  SIZE=18 MAXLENGTH=18 tag="21XXXU" ALT="생산품목4"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd4" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenitemCd(4)">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm4" SIZE=20 tag="24" ALT="생산품목명4"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>생산품목5</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd5"  SIZE=18 MAXLENGTH=18 tag="21XXXU" ALT="생산품목5"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd5" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenitemCd(5)">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm5" SIZE=20 tag="24" ALT="생산품목명5"></TD>
								<TD CLASS=TD5 NOWRAP>생산품목6</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd6"  SIZE=18 MAXLENGTH=18 tag="21XXXU" ALT="생산품목6"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd6" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenitemCd(6)">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm6" SIZE=20 tag="24" ALT="생산품목명6"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>생산품목7</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd7"  SIZE=18 MAXLENGTH=18 tag="21xxxU" ALT="생산품목7"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd7" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenitemCd(7)">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm7" SIZE=20 tag="24" ALT="생산품목명7"></TD>
								<TD CLASS=TD5 NOWRAP>생산품목8</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd8"  SIZE=18 MAXLENGTH=18 tag="21xxxU" ALT="생산품목8"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd8" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenitemCd(8)">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm8" SIZE=20 tag="24" ALT="생산품목명8"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>생산품목9</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd9"  SIZE=18 MAXLENGTH=18 tag="21xxxU" ALT="생산품목9"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd9" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenitemCd(9)">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm9" SIZE=20 tag="24" ALT="생산품목명9"></TD>
								<TD CLASS=TD5 NOWRAP>생산품목10</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd10"  SIZE=18 MAXLENGTH=18 tag="21xxxU" ALT="생산품목10"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd10" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenitemCd(10)">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm10" SIZE=20 tag="24" ALT="생산품목명10"></TD>
							</TR>
						</TABLE>
					</DIV>

					<!-- 세번째 탭 내용 -->
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%" SCROLL="no">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>구조 및 형식</TD>
					           	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtStrType" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="구조 및 형식"></TD>
								<TD CLASS=TD5 NOWRAP>S/Height</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p6110ma1_txtSHeight_txtSHeight.js'></script>&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>재질</TD>
					           	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMatQ" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="재질"></TD>
								<TD CLASS=TD5 NOWRAP>D/Height</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p6110ma1_txtDHeight_txtDHeight.js'></script>&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>가공구분</TD>
					           	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtProcessType" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="가공구분"></TD>
								<TD CLASS=TD5 NOWRAP>성형력</TD>
					           	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFormingP" TYPE=TEXT SIZE=10 MAXLENGTH="10" TAG="21XXXX" ALT="성형력"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>중량(Kg)</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/p6110ma1_txtWeightT_txtWeightT.js'></script>
											</TD>
											<TD>
												&nbsp;Kg
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>CUSHION압</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/p6110ma1_txtCushionPr_txtCushionPr.js'></script>
											</TD>
											<TD>
												&nbsp;Kg/Cm2
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>규격</TD>
					           	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSpec" TYPE=TEXT SIZE=40 MAXLENGTH="40" TAG="21XXXX" ALT="규격"></TD>
								<TD CLASS=TD5 NOWRAP>C/STROKE</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/p6110ma1_txtCStroke_txtCStroke.js'></script>
											</TD>
											<TD>&nbsp;mm
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>사용기계</TD>
					           	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtUseMachine" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="사용기계"></TD>
								<TD CLASS=TD5 NOWRAP>LOCATE</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLocate" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="LOCATE"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>자동화</TD>
					           	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAutoMath" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="자동화"></TD>
								<TD CLASS=TD5 NOWRAP>LOAD'G</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoading" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="LOAD'G"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>인원</TD>
					           	<TD CLASS=TD6 NOWRAP>
					           		<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/p6110ma1_txtPersonCount_txtPersonCount.js'></script>
											</TD>
											<TD>
												&nbsp;명
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>UNLOAD'G</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtUnLoading" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="UNLOAD'G"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>가공방향</TD>
					           	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtModifyDire" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="가공방향"></TD>
								<TD CLASS=TD5 NOWRAP>SCRAP 처리</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtScrapProcess" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="SCRAP 처리"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>가이드방식</TD>
					           	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGuideMath" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="가이드방식"></TD>
								<TD CLASS=TD5 NOWRAP>보관장소</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCustodyArea" TYPE=TEXT SIZE=20 MAXLENGTH="20" TAG="21XXXX" ALT="보관장소"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>현타수적용일</TD>
						    	<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p6110ma1_txtFinAjDt_txtFinAjDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>현재타수</TD>
						    	<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p6110ma1_txtCurAccnt_txtCurAccnt.js'></script>&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>최종점검일</TD>
						    	<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p6110ma1_txtCheckEndDt_txtCheckEndDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>점검타수</TD>
						    	<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p6110ma1_txtInspPrid_txtInspPrid.js'></script>&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>최종수리일</TD>
						    	<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p6110ma1_txtRepEndDt_txtRepEndDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>최종점검타수</TD>
						    	<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p6110ma1_txtFinCurAccnt_txtFinCurAccnt.js'></script>&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>폐기매각일</TD>
						    	<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p6110ma1_txtCloseDt_txtCloseDt.js'></script></TD>

								<TD CLASS=TD5 NOWRAP>타수당생산수량</TD>
						    	<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p6110ma1_txtPrsUnit_txtPrsUnit.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>제작일자</TD>
						    	<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p6110ma1_txtMakeDt_txtMakeDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>한계타수수량</TD>
						    	<TD CLASS=TD6 NOWRAP>
						    		<script language =javascript src='./js/p6110ma1_txtLimitAccnt_txtLimitAccnt.js'></script>
						    	</TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"  TABINDEX=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="htxtCastCd" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


