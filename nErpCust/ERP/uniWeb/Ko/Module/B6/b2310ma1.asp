<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Basic Info. - Accounting
'*  2. Function Name        : Organization Info.
'*  3. Program ID           : B2310MA1
'*  4. Program Name         : Cost Center 등록 
'*  5. Program Desc         : Register of Cost Center
'*  6. Comproxy List        : B23011, B23018
'*  7. Modified date(First) : 2000/08/26
'*  8. Modified date(Last)  : 2001/02/03
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Song, Mun Gil / Cho Ig Sung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT> 

Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "b2310mb1.asp"												'☆: 비지니스 로직 ASP명 

'==========================================================================================================
Dim C_COST_CENTER_CD
Dim C_COST_CENTER_NM
Dim C_COST_CENTER_ENG_NM
Dim C_COST_CENTER_TYPE_CD
Dim C_COST_CENTER_TYPE_NM
Dim C_COST_CENTER_DI_FG_CD
Dim C_COST_CENTER_DI_FG_NM
Dim C_BIZ_AREA_CD
Dim C_BIZ_AREA_PB
Dim C_BIZ_AREA_NM
Dim C_BIZ_UNIT_CD
Dim C_BIZ_UNIT_PB
Dim C_BIZ_UNIT_NM
Dim C_PLANT_CD
Dim C_PLANT_PB
Dim C_PLANT_NM
Dim C_ORG_CHANGE_ID
Dim C_ORG_CHANGE_ID_PB
Dim C_INTERNAL_CD
Dim C_DEPT_CD
Dim C_DEPT_PB
Dim C_DEPT_NM
Dim C_CHKFLAG 

Const C_SHEETMAXROWS = 100

'========================================================================================================= 

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
'Dim lgIntFlgMode               ' Variable is for Operation Status

Dim igStrPrevKey
'Dim lgLngCurRows
'Dim lgPageNo

'======================================================================================== 
Dim IsOpenPop
Dim lgRetFlag
'Dim lgSortKey

'======================================================================================== 
 <!-- #Include file="../../inc/lgvariables.inc" -->

'======================================================================================== 
Sub initSpreadPosVariables()
	C_COST_CENTER_CD		= 1
	C_COST_CENTER_NM		= 2
	C_COST_CENTER_ENG_NM	= 3
	C_COST_CENTER_TYPE_CD	= 4
	C_COST_CENTER_TYPE_NM	= 5
	C_COST_CENTER_DI_FG_CD	= 6
	C_COST_CENTER_DI_FG_NM	= 7
	C_BIZ_AREA_CD			= 8
	C_BIZ_AREA_PB			= 9
	C_BIZ_AREA_NM			= 10
	C_BIZ_UNIT_CD			= 11
	C_BIZ_UNIT_PB			= 12
	C_BIZ_UNIT_NM			= 13
	C_PLANT_CD				= 14
	C_PLANT_PB				= 15
	C_PLANT_NM				= 16
	C_ORG_CHANGE_ID			= 17
	C_ORG_CHANGE_ID_PB      = 18
	C_INTERNAL_CD			= 19
	C_DEPT_CD				= 20
	C_DEPT_PB				= 21
	C_DEPT_NM				= 22
	C_CHKFLAG               = 23
End Sub


'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgSortKey = 1
End Sub

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'======================================================================================== 
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021203",,parent.gAllowDragDropSpread

	With frm1.vspdData

		.MaxCols = C_CHKFLAG + 1
		.MaxRows = 0

        .ReDraw = False

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit C_COST_CENTER_CD, "코스트센타코드", 15, , , 10, 2
		ggoSpread.SSSetEdit C_COST_CENTER_NM, "코스트센타명", 22, , , 20
		ggoSpread.SSSetEdit C_COST_CENTER_ENG_NM, "코스트센타 영문명", 22, , , 50
		ggoSpread.SSSetCombo C_COST_CENTER_TYPE_CD, "코스트센타 종류", 20, 2, True
		ggoSpread.SSSetCombo C_COST_CENTER_TYPE_NM, "코스트센타 종류", 20, 2, False
		ggoSpread.SSSetCombo C_COST_CENTER_DI_FG_CD, "직/간 구분", 15, 2, True
		ggoSpread.SSSetCombo C_COST_CENTER_DI_FG_NM, "직/간 구분", 15, 2, False
		ggoSpread.SSSetEdit C_BIZ_AREA_CD, "사업장코드", 12, , , 10, 2
		ggoSpread.SSSetButton C_BIZ_AREA_PB
		ggoSpread.SSSetEdit  C_BIZ_AREA_NM, "사업장명", 22, , , 50
		ggoSpread.SSSetEdit C_BIZ_UNIT_CD, "사업부코드", 12, , , 10, 2
		ggoSpread.SSSetButton C_BIZ_UNIT_PB
		ggoSpread.SSSetEdit  C_BIZ_UNIT_NM, "사업부명", 22, , , 50
		ggoSpread.SSSetEdit C_PLANT_CD, "공장코드", 10, , , 4, 2
		ggoSpread.SSSetButton C_PLANT_PB
		ggoSpread.SSSetEdit  C_PLANT_NM, "공장명", 22, , , 40
		ggoSpread.SSSetEdit C_ORG_CHANGE_ID, "조직변경ID", 10, , , 5, 2
		ggoSpread.SSSetButton C_ORG_CHANGE_ID_PB		
		ggoSpread.SSSetEdit C_INTERNAL_CD, "내부부서코드", 12, , , 30, 2
		ggoSpread.SSSetEdit C_DEPT_CD, "부서코드", 10, , , 12, 2
		ggoSpread.SSSetButton C_DEPT_PB
		ggoSpread.SSSetEdit  C_DEPT_NM, "부서명", 22, , , 40
		ggoSpread.SSSetCheck C_CHKFLAG, "사업장대표Flag", 15, ,"",true    

        Call ggoSpread.MakePairsColumn(C_BIZ_AREA_CD,C_BIZ_AREA_PB,"1")
        Call ggoSpread.MakePairsColumn(C_BIZ_UNIT_CD,C_BIZ_UNIT_PB,"1")
        Call ggoSpread.MakePairsColumn(C_PLANT_CD,C_PLANT_PB,"1")
        Call ggoSpread.MakePairsColumn(C_DEPT_CD,C_DEPT_PB,"1")
        'Call ggoSpread.SSSetColHidden(C_BizUnitCd,C_BizUnitCd,True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
        Call ggoSpread.SSSetColHidden(C_COST_CENTER_TYPE_CD,C_COST_CENTER_TYPE_CD,True)
        Call ggoSpread.SSSetColHidden(C_COST_CENTER_DI_FG_CD,C_COST_CENTER_DI_FG_CD,True)
'        Call ggoSpread.SSSetColHidden(C_ORG_CHANGE_ID,C_ORG_CHANGE_ID,True)
        Call ggoSpread.SSSetColHidden(C_INTERNAL_CD,C_INTERNAL_CD,True)

        Call InitComboBox
		.ReDraw = true
	End With

	Call SetSpreadLock 
End Sub


'========================================================================================
Sub SetSpreadLock()
    With frm1

	.vspdData.ReDraw = False
	ggoSpread.SpreadLock C_COST_CENTER_CD			, -1, C_COST_CENTER_CD
	ggoSpread.SSSetRequired C_COST_CENTER_NM		, -1, C_COST_CENTER_NM
	ggoSpread.SSSetRequired	C_COST_CENTER_TYPE_NM	, -1, C_COST_CENTER_TYPE_NM
	ggoSpread.SSSetRequired	C_COST_CENTER_DI_FG_NM	, -1, C_COST_CENTER_DI_FG_NM
	ggoSpread.SSSetRequired C_BIZ_AREA_CD			, -1, C_BIZ_AREA_CD
	ggoSpread.SpreadLock C_BIZ_AREA_NM				, -1, C_BIZ_AREA_NM
	ggoSpread.SSSetRequired C_BIZ_UNIT_CD			, -1, C_BIZ_UNIT_CD
	ggoSpread.SpreadLock C_BIZ_UNIT_NM				, -1, C_BIZ_UNIT_NM
	ggoSpread.SpreadLock C_PLANT_NM					, -1, C_PLANT_NM
	ggoSpread.SpreadLock C_DEPT_NM					, -1, C_DEPT_NM
	ggoSpread.SSSetProtected	.vspdData.MaxCols, -1,-1
	.vspdData.ReDraw = True

    End With
End Sub


'========================================================================================
Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
    With frm1

	    .vspdData.ReDraw = False

		ggoSpread.SSSetRequired C_COST_CENTER_CD, pvStarRow, pvEndRow
		ggoSpread.SSSetRequired	C_COST_CENTER_NM, pvStarRow, pvEndRow
		ggoSpread.SSSetRequired	C_COST_CENTER_TYPE_NM, pvStarRow, pvEndRow
		ggoSpread.SSSetRequired	C_COST_CENTER_DI_FG_NM, pvStarRow, pvEndRow
		ggoSpread.SSSetRequired C_BIZ_AREA_CD, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_BIZ_AREA_NM, pvStarRow, pvEndRow
		ggoSpread.SSSetRequired C_BIZ_UNIT_CD, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_BIZ_UNIT_NM, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_PLANT_NM, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_DEPT_NM, pvStarRow, pvEndRow

		.vspdData.ReDraw = True

    End With
End Sub


'========================================================================================================= 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_COST_CENTER_CD		= iCurColumnPos(1)
			C_COST_CENTER_NM		= iCurColumnPos(2)
			C_COST_CENTER_ENG_NM	= iCurColumnPos(3)
			C_COST_CENTER_TYPE_CD	= iCurColumnPos(4)
			C_COST_CENTER_TYPE_NM	= iCurColumnPos(5)
			C_COST_CENTER_DI_FG_CD	= iCurColumnPos(6)
			C_COST_CENTER_DI_FG_NM	= iCurColumnPos(7)
			C_BIZ_AREA_CD			= iCurColumnPos(8)
			C_BIZ_AREA_PB			= iCurColumnPos(9)
			C_BIZ_AREA_NM			= iCurColumnPos(10)
			C_BIZ_UNIT_CD			= iCurColumnPos(11)
			C_BIZ_UNIT_PB			= iCurColumnPos(12)
			C_BIZ_UNIT_NM			= iCurColumnPos(13)
			C_PLANT_CD				= iCurColumnPos(14)
			C_PLANT_PB				= iCurColumnPos(15)
			C_PLANT_NM				= iCurColumnPos(16)
			C_ORG_CHANGE_ID			= iCurColumnPos(17)
			C_ORG_CHANGE_ID_PB      = iCurColumnPos(18)
			C_INTERNAL_CD			= iCurColumnPos(19)
			C_DEPT_CD				= iCurColumnPos(20)
			C_DEPT_PB				= iCurColumnPos(21)
			C_DEPT_NM				= iCurColumnPos(22)
			C_CHKFLAG				= iCurColumnPos(23)
	End Select
End Sub

'========================================================================================================= 
Function OpenPopup(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim costcd,orgid
	Dim strBizAreaCd
	If IsOpenPop = True Then Exit Function

	With frm1
		Select Case iWhere
			Case 0
				arrParam(0) = "코스트센타 팝업"					' 팝업 명칭 
				arrParam(1) = "B_COST_CENTER"						' TABLE 명칭 
				arrParam(2) = strCode								' Code Condition
				arrParam(3) = ""									' Name Cindition
				arrParam(4) = ""									' Where Condition
				arrParam(5) = "코스트센타"

			    arrField(0) = "COST_CD"								' Field명(0)
				arrField(1) = "COST_NM"								' Field명(1)

				arrHeader(0) = "코스트센타코드"					' Header명(0)
				arrHeader(1) = "코스트센타명"						' Header명(1)
			Case 1
				arrParam(0) = "사업장 팝업"						' 팝업 명칭 
				arrParam(1) = "B_BIZ_AREA A"						' TABLE 명칭 
				arrParam(2) = strCode    							' Code Condition
				arrParam(3) = ""									' Name Cindition
				arrParam(4) = ""
				arrParam(5) = "사업장"

			    arrField(0) = "A.BIZ_AREA_CD"						' Field명(0)
				arrField(1) = "A.BIZ_AREA_NM"						' Field명(1)

			    arrHeader(0) = "사업장코드"						' Header명(0)
				arrHeader(1) = "사업장명"						    ' Header명(1)
			Case 2
			    ggoSpread.Source = .vspdData
			    .vspdData.Row = .vspdData.ActiveRow
				.vspdData.Col = C_BIZ_AREA_CD
				strBizAreaCd = Trim(.vspdData.text)

				arrParam(0) = "공장 팝업"							' 팝업 명칭 
				arrParam(1) = "b_plant a, b_biz_area b"							' TABLE 명칭 
				arrParam(2) = strCode    							' Code Condition
				arrParam(3) = ""									' Name Cindition
				arrParam(4) = " a.biz_area_cd= b.biz_area_cd and b.biz_area_cd = " & FilterVar(strBizAreaCd, "''", "S")
				arrParam(5) = "공장"

			    arrField(0) = "A.PLANT_CD"							' Field명(0)
				arrField(1) = "A.PLANT_NM"							' Field명(1)

			    arrHeader(0) = "공장코드"							' Header명(0)
				arrHeader(1) = "공장명"						    ' Header명(1)
			Case 3
				arrParam(0) = "사업부 팝업"						' 팝업 명칭 
				arrParam(1) = "B_BIZ_UNIT A"						' TABLE 명칭 
				arrParam(2) = strCode    							' Code Condition
				arrParam(3) = ""									' Name Cindition
				arrParam(4) = ""
				arrParam(5) = "사업부"

			    arrField(0) = "A.BIZ_UNIT_CD"						' Field명(0)
				arrField(1) = "A.BIZ_UNIT_NM"						' Field명(1)

			    arrHeader(0) = "사업부코드"						' Header명(0)
				arrHeader(1) = "사업부명"						' Header명(1)
			Case 4
				.vspdData.row = .vspdData.activerow
				.vspdData.Col = C_COST_CENTER_CD
				costcd = Trim(.vspdData.text)
				.vspdData.Col = C_ORG_CHANGE_ID		
				orgid = Trim(.vspdData.text)			

				arrParam(0) = "부서 팝업"						' 팝업 명칭 
				arrParam(1) = "B_ACCT_DEPT A"						' TABLE 명칭 
				arrParam(2) = strCode    							' Code Condition
				arrParam(3) = ""									' Name Cindition
				If orgid <> "" Then
					arrParam(4) = " A.ORG_CHANGE_ID =  " & FilterVar(orgid , "''", "S") & " "
				Else
					arrParam(4) = "  "		
				End If	
				arrParam(5) = "부서"

			    arrField(0) = "A.DEPT_CD"							' Field명(0)
				arrField(1) = "A.DEPT_NM"							' Field명(1)
				arrField(2) = "A.ORG_CHANGE_ID"						' Field명(2)
				arrField(3) = "A.INTERNAL_CD"						' Field명(3)

			    arrHeader(0) = "부서코드"						' Header명(0)
				arrHeader(1) = "부서명"						    ' Header명(1)
				arrHeader(2) = "조직변경아이디"					' Header명(2)
				arrHeader(3) = "내부부서코드"					' Header명(3)
			Case 5
				.vspdData.row = .vspdData.activerow
				.vspdData.Col = C_ORG_CHANGE_ID
				orgid = Trim(.vspdData.text)

				arrParam(0) = "부서개편ID팝업"					' 팝업 명칭 
				arrParam(1) = " Horg_abs "							' TABLE 명칭 
				arrParam(2) = strCode    							' Code Condition
				arrParam(3) = ""									' Name Cindition
				arrParam(4) = "  "
				arrParam(5) = "부서개편ID"

			    arrField(0) = "orgid"								' Field명(0)
				arrField(1) = "orgdt"								' Field명(1)
				arrField(2) = ""									' Field명(2)
				arrField(3) = ""									' Field명(3)

			    arrHeader(0) = "부서개편ID"						' Header명(0)
				arrHeader(1) = "부서개편일자"					' Header명(1)
				arrHeader(2) = ""									' Header명(2)
				arrHeader(3) = ""									' Header명(3)
			Case Else
				Exit Function
		End Select
	End With
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	Call GridsetFocus(iWhere)

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If
End Function

'=======================================================================================================
Function GridsetFocus(Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			frm1.txtCOST_CENTER_CD.focus
		End Select
	End With
End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
			Case 0
				.txtCOST_CENTER_CD.value = arrRet(0)
				.txtCOST_CENTER_NM.value = arrRet(1)
			Case 1
				.vspdData.Col = C_BIZ_AREA_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_BIZ_AREA_NM
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(.vspdData.Col, .vspdData.Row)
			Case 2
				.vspdData.Col = C_PLANT_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_PLANT_NM
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(.vspdData.Col, .vspdData.Row)
			Case 3
				.vspdData.Col = C_BIZ_UNIT_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_BIZ_UNIT_NM
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(.vspdData.Col, .vspdData.Row)
			Case 4
				.vspdData.Col = C_DEPT_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_DEPT_NM
				.vspdData.Text = arrRet(1)
				.vspdData.Col = C_ORG_CHANGE_ID
				.vspdData.Text = arrRet(2)
				.vspdData.Col = C_INTERNAL_CD
				.vspdData.Text = arrRet(3)
				Call vspdData_Change(.vspdData.Col, .vspdData.Row)
			Case 5
				.vspdData.Col = C_ORG_CHANGE_ID
				.vspdData.Text = arrRet(0)
				Call vspdData_Change(.vspdData.Col, .vspdData.Row)			
		End Select
	End With
End Function


'========================================================================================================= 
Sub InitComboBox()
	Dim iCodeArr,iNameArr

	Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("C2203", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_COST_CENTER_TYPE_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_COST_CENTER_TYPE_NM

    Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("C0002", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_COST_CENTER_DI_FG_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_COST_CENTER_DI_FG_NM
End Sub


'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow

			.Col = C_COST_CENTER_TYPE_CD
			intIndex = .value
			.col = C_COST_CENTER_TYPE_NM
			.value = intindex

			.Col = C_COST_CENTER_DI_FG_CD
			intIndex = .value
			.col = C_COST_CENTER_DI_FG_NM
			.value = intindex
		Next
	End With
End Sub


'========================================================================================================= 
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================= 
Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call InitComboBox

    Call SetToolBar("1100110100101111")
    frm1.txtCOST_CENTER_CD.focus
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================================= 
Sub vspdData_Click(ByVal Col, ByVal Row)
	If lgIntFlgMode = parent.OPMD_CMODE Then
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

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )'$$
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim IntRetCD
	Dim ii
	Dim arrVal1
	Dim arrVal2
	Dim jj
	Dim costcd,orgid

	With frm1	
		.vspdData.Row = Row
		.vspdData.Col = Col

		If .vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
			If CDbl(.vspdData.text) < CDbl(.vspdData.TypeFloatMin) Then
				.vspdData.text = .vspdData.TypeFloatMin
			End If
		End If

		Select Case Col
			Case  C_DEPT_CD
				.vspdData.Col = C_COST_CENTER_CD
				.vspdData.Row = Row
				costcd	= Trim(.vspdData.text)
				.vspdData.Col = C_DEPT_CD
				.vspdData.Row = Row

				If LTrim(RTrim(.vspdData.text)) <> ""  Then
					strSelect	=			 " dept_cd, dept_nm, org_change_id, internal_cd "
					strFrom		=			 " b_acct_dept(NOLOCK) "		
					strWhere	=			 " dept_Cd =  " & FilterVar(.vspdData.text, "''", "S") 
					frm1.vspdData.Col = C_ORG_CHANGE_ID
					strWhere	= strWhere & " and org_change_id =  " & FilterVar(.vspdData.text , "''", "S") & ""

					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
						IntRetCD = DisplayMsgBox("124600","X","X","X")
						.vspdData.Col = C_DEPT_CD
						.vspdData.text = ""
						.vspdData.Col = C_DEPT_NM
						.vspdData.text = ""
						.vspdData.Col = C_ORG_CHANGE_ID 
						.vspdData.text = ""
						.vspdData.Col = C_INTERNAL_CD 
						.vspdData.text = ""
					Else
						arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
						jj = Ubound(arrVal1,1)

						For ii = 0 to jj - 1
							arrVal2 = Split(arrVal1(ii), chr(11))
							.vspdData.Col = C_DEPT_CD
							.vspdData.text = Trim(arrVal2(1))
							.vspdData.Col = C_DEPT_NM
							.vspdData.text = Trim(arrVal2(2))
							.vspdData.Col = C_ORG_CHANGE_ID 
							.vspdData.text = Trim(arrVal2(3))
							.vspdData.Col = C_INTERNAL_CD 
							.vspdData.text = Trim(arrVal2(4))
						Next
					End If 
				Else
					.vspdData.Col = C_DEPT_CD
					.vspdData.text = ""
					.vspdData.Col = C_DEPT_NM
					.vspdData.text = ""
					.vspdData.Col = C_ORG_CHANGE_ID 
					.vspdData.text = ""
					.vspdData.Col = C_INTERNAL_CD 
					.vspdData.text = ""
				End If
			Case  C_ORG_CHANGE_ID
				.vspdData.Col = C_ORG_CHANGE_ID
				.vspdData.Row = Row
				orgid	= Trim(.vspdData.text)

				If LTrim(RTrim(.vspdData.text)) <> ""  Then
					strSelect	=			 " orgid "
					strFrom		=			 " horg_abs(NOLOCK) "		
					strWhere	=			 " orgid =  " & FilterVar(orgid, "''", "S") 

					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
						IntRetCD = DisplayMsgBox("124700","X","X","X")
						.vspdData.Col = C_DEPT_CD
						.vspdData.text = ""
						.vspdData.Col = C_DEPT_NM
						.vspdData.text = ""
						.vspdData.Col = C_ORG_CHANGE_ID 
						.vspdData.text = ""
						.vspdData.Col = C_INTERNAL_CD 
						.vspdData.text = ""
					End If 
				Else
					.vspdData.Col = C_DEPT_CD
					.vspdData.text = ""
					.vspdData.Col = C_DEPT_NM
					.vspdData.text = ""
					.vspdData.Col = C_ORG_CHANGE_ID 
					.vspdData.text = ""
					.vspdData.Col = C_INTERNAL_CD 
					.vspdData.text = ""
				End If
		End Select

		ggoSpread.Source = .vspdData
		ggoSpread.UpdateRow Row
	End With	
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
End Sub

'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 

		ggoSpread.Source = frm1.vspdData

		If Row > 0 Then
			If Col = C_BIZ_AREA_PB Then	
				.Col = Col
				.Row = Row

				.Col = C_BIZ_AREA_CD
				Call OpenPopup(.Text, 1)
			ElseIf Col = C_PLANT_PB Then
				.Col = Col
				.Row = Row

				.Col = C_PLANT_CD
				Call OpenPopup(.Text, 2)
			ElseIf Col = C_BIZ_UNIT_PB Then
				.Col = Col
				.Row = Row

				.Col = C_BIZ_UNIT_CD
				Call OpenPopup(.Text, 3)
			ElseIf Col = C_DEPT_PB Then
				.Col = Col
				.Row = Row

				.Col = C_DEPT_CD

				Call OpenPopup(.Text, 4)
			ElseIf Col = C_ORG_CHANGE_ID_PB Then
				.Col = Col
				.Row = Row

				.Col = C_ORG_CHANGE_ID

				Call OpenPopup(.Text, 5)				
			Else
				Exit Sub
			END If
			Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")
		End If

    End With

End Sub

'========================================================================================================= 
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub

'========================================================================================================= 
Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================= 
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
End Sub

'========================================================================================================= 
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex

	 '----------  Coding part  -------------------------------------------------------------   
	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
	
		.Row = Row
		Select Case Col
			Case  C_COST_CENTER_TYPE_NM
				.Col = Col
				intIndex = .Value
				.Col = C_COST_CENTER_TYPE_CD
				.Value = intIndex
			Case C_COST_CENTER_DI_FG_NM
				.Col = Col
				intIndex = .Value
				.Col = C_COST_CENTER_DI_FG_CD
				.Value = intIndex
			Case  C_COST_CENTER_TYPE_CD
				.Col = Col
				intIndex = .Value
				.Col = C_COST_CENTER_TYPE_NM
				.Value = intIndex
			Case C_COST_CENTER_DI_FG_CD
				.Col = Col
				intIndex = .Value
				.Col = C_COST_CENTER_DI_FG_NM
				.Value = intIndex
		End Select
	End With

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
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

    Call InitVariables
    Call InitComboBox

	If frm1.txtCOST_CENTER_CD.value = "" Then
		frm1.txtCOST_CENTER_NM.value = ""
	End If

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery

    FncQuery = True
End Function

'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function

'========================================================================================
Function FncDelete() 
	On Error Resume Next
End Function

'========================================================================================
Function FncSave() 
    Dim IntRetCD

    FncSave = False

    Err.Clear
	On Error Resume Next

    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False  Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
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
	If frm1.vspdData.MaxRows < 1 Then Exit Function

	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow

    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    frm1.vspdData.Col = C_COST_CENTER_CD
    frm1.vspdData.Text = ""
	frm1.vspdData.Col = C_DEPT_CD
	frm1.vspdData.text = ""
	frm1.vspdData.Col = C_DEPT_NM
	frm1.vspdData.text = ""
	frm1.vspdData.Col = C_ORG_CHANGE_ID
	frm1.vspdData.text = ""
	frm1.vspdData.Col = C_INTERNAL_CD
	frm1.vspdData.text = ""

    Call InitData 

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
Function FncInsertRow(Byval pvRowCnt)
	Dim imRow

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowcount()
		If ImRow="" then
			Exit Function
		End If
	End If

	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData

		.vspdData.ReDraw = False
		ggoSpread.InsertRow ,imRow

		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

		.vspdData.ReDraw = True
    End With
    If Err.number = 0 Then
       FncInsertRow = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

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
    Call parent.FncExport(parent.C_MULTI)												'☜: 화면 유형 
End Function

'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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
Function FncExit()
	Dim IntRetCD
	FncExit = False

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
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

    Err.Clear

    With frm1
		.txtCOST_CENTER_CD.value = UCase(Trim(.txtCOST_CENTER_CD.value))
		
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtCOST_CENTER_CD=" & UCase(Trim(.htxtCOST_CENTER_CD.value))	'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgPageNo=" & lgPageNo
 			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtCOST_CENTER_CD=" & UCase(Trim(.txtCOST_CENTER_CD.value))	'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgPageNo=" & lgPageNo
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

		End If

		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    End With

    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
    lgIntFlgMode = parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call InitData

   	Call SetToolBar("1100111100111111")
	frm1.vspdData.focus
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
Function DbSave()
	Dim lRow
	Dim lGrpCnt
	Dim strVal,strDel

    DbSave = False

    Call LayerShowHide(1)

    On Error Resume Next

	lgRetFlag = False
	With frm1
		.txtMode.value = parent.UID_M0002

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

				strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep				'☜: C=Create, Row위치 정보 
                .vspdData.Col = C_COST_CENTER_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_COST_CENTER_NM
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_COST_CENTER_ENG_NM
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_COST_CENTER_TYPE_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_COST_CENTER_DI_FG_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_BIZ_AREA_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_BIZ_UNIT_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_PLANT_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_INTERNAL_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep 
                .vspdData.Col = C_ORG_CHANGE_ID
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_DEPT_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_CHKFLAG   
                    IF .vspdData.Text = "1" Then
					  	strVal = strVal & "Y" & parent.gRowSep
					ELSE
					  	strVal = strVal & "N" & parent.gRowSep
					END IF

                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.UpdateFlag											'☜: 수정 

				strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep 				'☜: U=Update, Row위치 정보 
                .vspdData.Col = C_COST_CENTER_CD						'1
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_COST_CENTER_NM
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_COST_CENTER_ENG_NM
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_COST_CENTER_TYPE_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_COST_CENTER_DI_FG_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_BIZ_AREA_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_BIZ_UNIT_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_PLANT_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_INTERNAL_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_ORG_CHANGE_ID
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_DEPT_CD
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_CHKFLAG 
                    IF .vspdData.Text = "1" Then
					  	strVal = strVal & "Y" & parent.gRowSep
					ELSE
					  	strVal = strVal & "N" & parent.gRowSep
					END IF
                

                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.DeleteFlag											'☜: 삭제 

				strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep				'☜: D=Delete, Row위치 정보 
                .vspdData.Col = C_COST_CENTER_CD	'10
                strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep									

                lGrpCnt = lGrpCnt + 1
        End Select
    Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	End With

    DbSave = True
End Function

'========================================================================================
Function DbSaveOk()
	Call ggoOper.ClearField(Document, "2")
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData
	Call InitVariables
	Call Dbquery
End Function


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT="20" WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>코스트센타</TD>
									<TD CLASS="TD656" COLSPAN=3><INPUT NAME="txtCOST_CENTER_CD" MAXLENGTH="10" SIZE=10 ALT ="코스트센타 코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="Call OpenPopup(frm1.txtCOST_CENTER_CD.value, 0)"> <INPUT NAME="txtCOST_CENTER_NM" MAXLENGTH="20" SIZE=30 STYLE="TEXT-ALIGN:left" ALT ="코스트센타명" tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="htxtCOST_CENTER_CD" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

