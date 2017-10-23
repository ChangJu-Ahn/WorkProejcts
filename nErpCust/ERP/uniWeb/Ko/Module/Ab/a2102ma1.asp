<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNTING
'*  2. Function Name        : 
'*  3. Program ID		    : A2102MA1
'*  4. Program Name         : 관리항목 등록 
'*  5. Program Desc         : 관리항목 등록 수정 삭제 조회 
'*  6. Component List       : 
'*  7. ModIfied date(First) : 2000/09/07
'*  8. ModIfied date(Last)  : 2003/08/13
'*  9. ModIfier (First)     : Jong Hwan, Kim
'* 10. ModIfier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit	

Const BIZ_PGM_ID = "a2102mb1.asp"			'☆: 비지니스 로직 ASP명 
'==========================================================================================================
'⊙: Grid Columns
Dim C_CtrlCD
Dim C_CtrlNM
Dim C_SYSFG
Dim C_CtrlEngNM
Dim C_DataTypeCd
Dim C_DataTypeNm
Dim C_DataLen
Dim C_GL_CTRL_FLD
Dim C_GL_CTRL_NM
Dim C_TblID
Dim C_CtrlTblPopUp
Dim C_ColmID
Dim C_CtrlColmPopUp1
Dim C_ColmIDNM
Dim C_CtrlColmPopUp2
Dim C_MAJORCD
Dim C_MAJORPOPUP
Dim C_MAJORNM
Dim C_KeyColm1
Dim C_CtrlColmPopUp3
Dim C_DataTypeCd1
Dim C_DataTypeNm1
Dim C_KeyColm2	
Dim C_CtrlColmPopUp4
Dim C_DataTypeCd2
Dim C_DataTypeNm2
Dim C_KeyColm3
Dim C_CtrlColmPopUp5
Dim C_DataTypeCd3
Dim C_DataTypeNm3
Dim C_KeyColm4
Dim C_CtrlColmPopUp6
Dim C_DataTypeCd4
Dim C_DataTypeNm4
Dim C_KeyColm5
Dim C_CtrlColmPopUp7
Dim C_DataTypeCd5
Dim C_DataTypeNm5

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_CtrlCD			= 1
	C_CtrlNM			= 2
	C_SYSFG				= 3
	C_CtrlEngNM			= 4
	C_DataTypeCd		= 5
	C_DataTypeNm		= 6
	C_DataLen			= 7
	C_GL_CTRL_FLD       = 8
	C_GL_CTRL_NM        = 9	
	C_TblID				= 10
    C_CtrlTblPopUp		= 11
	C_ColmID			= 12
    C_CtrlColmPopUp1	= 13
	C_ColmIDNM			= 14
    C_CtrlColmPopUp2	= 15
	C_MAJORCD			= 16
	C_MAJORPOPUP		= 17
	C_MAJORNM			= 18
	C_KeyColm1			= 19
    C_CtrlColmPopUp3	= 20
	C_DataTypeCd1		= 21
	C_DataTypeNm1		= 22
	C_KeyColm2			= 23
    C_CtrlColmPopUp4	= 24
	C_DataTypeCd2		= 25
	C_DataTypeNm2		= 26
	C_KeyColm3			= 27
    C_CtrlColmPopUp5	= 28
	C_DataTypeCd3		= 29
	C_DataTypeNm3		= 30
	C_KeyColm4			= 31
    C_CtrlColmPopUp6	= 32
	C_DataTypeCd4		= 33
	C_DataTypeNm4		= 34
	C_KeyColm5			= 35
    C_CtrlColmPopUp7	= 36
	C_DataTypeCd5		= 37
	C_DataTypeNm5		= 38
End Sub

'========================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================
Dim  IsOpenPop

'========================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE
    lgIntGrpCount = 0

    lgStrPrevKey = ""
    lgLngCurRows = 0

    lgSortKey = 1
    lgPageNo = 0
End Sub

'========================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
End Sub

'========================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021212",,parent.gAllowDragDropSpread

    Dim		sList
	sList = "Y" & vbTab  & "N"

	With frm1.vspdData
		.MaxCols = C_DataTypeNm5 + 1
		.MaxRows = 0

		.ReDraw = False

        Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_CtrlCD        ,"관리항목코드"  , 20, , , 3, 2
		ggoSpread.SSSetEdit		C_CtrlNM        ,"관리항목명"    , 20, , , 30
		ggoSpread.SSSetCombo	C_SYSFG         ,"시스템구분"    , 12, True
		ggoSpread.SSSetEdit		C_CtrlEngNM     ,"관리항목영문명", 20, , , 50
		ggoSpread.SSSetCombo	C_DataTypeCd    ,"자료유형"      , 15
		ggoSpread.SSSetCombo	C_DataTypeNm    ,"자료유형"      , 15

'		Call AppendNumberPlace("6","3","0")
		ggoSpread.SSSetFloat	C_DataLen,"자료길이",15,"6"  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"1","999"

		ggoSpread.SSSetEdit		C_CtrlEngNM     ,"관리항목영문명", 20, , , 50
		
		ggoSpread.SSSetCombo	C_GL_CTRL_FLD   ,"전표관리항목"  , 15
		ggoSpread.SSSetCombo	C_GL_CTRL_NM    ,"전표관리항목"  , 15

		ggoSpread.SSSetEdit		C_TblID         ,"테이블ID"      , 20, , , 32, 2
        ggoSpread.SSSetButton   C_CtrlTblPopUp 'jsk add

		ggoSpread.SSSetEdit		C_ColmID        ,"컬럼ID"        , 20, , , 32, 2
        ggoSpread.SSSetButton   C_CtrlColmPopUp1 'jsk add

		ggoSpread.SSSetEdit		C_ColmIDNM      ,"컬럼명ID"      , 20, , , 32, 2
        ggoSpread.SSSetButton   C_CtrlColmPopUp2 'jsk add

		ggoSpread.SSSetEdit		C_MAJORCD       ,"종합코드"      ,  8, , , 5, 2
		ggoSpread.SSSetButton	C_MAJORPOPUP
		ggoSpread.SSSetEdit		C_MAJORNM       ,"종합코드명"    , 22, , , 40

		ggoSpread.SSSetEdit		C_KeyColm1      ,"KEY컬럼1"      , 20, , , 32, 2
        ggoSpread.SSSetButton   C_CtrlColmPopUp3 'jsk add
		ggoSpread.SSSetCombo	C_DataTypeCd1   ,"자료유형1"     , 15
		ggoSpread.SSSetCombo	C_DataTypeNm1   ,"자료유형1"     , 15

		ggoSpread.SSSetEdit		C_KeyColm2      ,"KEY컬럼2"      , 20, , , 32, 2
        ggoSpread.SSSetButton   C_CtrlColmPopUp4 'jsk add
		ggoSpread.SSSetCombo	C_DataTypeCd2   ,"자료유형2"     , 15
		ggoSpread.SSSetCombo	C_DataTypeNm2   ,"자료유형2"     , 15

		ggoSpread.SSSetEdit		C_KeyColm3      ,"KEY컬럼3"      , 20, , , 32, 2
        ggoSpread.SSSetButton   C_CtrlColmPopUp5 'jsk add
		ggoSpread.SSSetCombo	C_DataTypeCd3   ,"자료유형3"     , 15
		ggoSpread.SSSetCombo	C_DataTypeNm3   ,"자료유형3"     , 15
    
		ggoSpread.SSSetEdit		C_KeyColm4,      "Key컬럼4"      , 20, , , 32, 2
        ggoSpread.SSSetButton   C_CtrlColmPopUp6 'jsk add
		ggoSpread.SSSetCombo	C_DataTypeCd4   ,"자료유형4"     , 15
		ggoSpread.SSSetCombo	C_DataTypeNm4   ,"자료유형4"     , 15

		ggoSpread.SSSetEdit		C_KeyColm5      ,"Key컬럼5"      , 20, , , 32, 2
        ggoSpread.SSSetButton   C_CtrlColmPopUp7 'jsk add
		ggoSpread.SSSetCombo	C_DataTypeCd5   ,"자료유형5"     , 15
		ggoSpread.SSSetCombo	C_DataTypeNm5   ,"자료유형5"     , 15

		Call ggoSpread.MakePairsColumn(C_TblID,C_CtrlTblPopUp,"1")
		Call ggoSpread.MakePairsColumn(C_MAJORCD,C_MAJORPOPUP,"1")
		Call ggoSpread.MakePairsColumn(C_ColmIDNM,C_CtrlColmPopUp2,"1")
		Call ggoSpread.MakePairsColumn(C_KeyColm1,C_CtrlColmPopUp3,"1")
		Call ggoSpread.MakePairsColumn(C_KeyColm2,C_CtrlColmPopUp4,"1")
		Call ggoSpread.MakePairsColumn(C_KeyColm3,C_CtrlColmPopUp5,"1")
		Call ggoSpread.MakePairsColumn(C_KeyColm4,C_CtrlColmPopUp6,"1")
		Call ggoSpread.MakePairsColumn(C_KeyColm5,C_CtrlColmPopUp7,"1")

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_DataTypeCd,C_DataTypeCd,True)
		Call ggoSpread.SSSetColHidden(C_DataTypeCd1,C_DataTypeCd1,True)
		Call ggoSpread.SSSetColHidden(C_DataTypeCd2,C_DataTypeCd2,True)
		Call ggoSpread.SSSetColHidden(C_DataTypeCd3,C_DataTypeCd3,True)
		Call ggoSpread.SSSetColHidden(C_DataTypeCd4,C_DataTypeCd4,True)
		Call ggoSpread.SSSetColHidden(C_DataTypeCd5,C_DataTypeCd5,True)
		Call ggoSpread.SSSetColHidden(C_GL_CTRL_FLD,C_GL_CTRL_FLD,True)		
		
		.ReDraw = True
		Call SetSpreadLock
		Call initComboBox()

		ggoSpread.SetCombo sList, C_SYSFG
    End With
End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1.vspdData
		.ReDraw = False
		'SpreadLock(Col1, Row1, Optional Col2, Optional Row2)
		ggoSpread.SpreadLock C_CtrlCD   , -1, C_CtrlCD
		ggoSpread.SpreadLock C_CtrlNM   , -1, C_CtrlNM
		ggoSpread.SpreadLock C_SYSFG    , -1, C_SYSFG
		ggoSpread.SpreadLock C_CtrlEngNM, -1, C_CtrlEngNM
		ggoSpread.SpreadLock C_MAJORNM  , -1, C_MAJORNM
		ggoSpread.SSSetRequired  C_DataTypeNm,-1, -1	' 자료유형 
		ggoSpread.SSSetRequired  C_DataLen,   -1, -1	' 자료길이 
		ggoSpread.SSSetProtected	.MaxCols,-1,-1
		.ReDraw = True
    End With
End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
		' 필수 입력 항목으로 설정 
		' SSSetRequired(ByVal Col, ByVal Row, Optional Row2)
		ggoSpread.SSSetRequired  C_CtrlCD,    pvStartRow, pvEndRow	' 관리항목코드 
		ggoSpread.SSSetRequired  C_CtrlNM,    pvStartRow, pvEndRow	' 관리항목명 
		ggoSpread.SSSetProtected C_SYSFG,     pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_CtrlEngNM, pvStartRow, pvEndRow	' 관리항목영문명 
		ggoSpread.SSSetRequired  C_DataTypeNm,pvStartRow, pvEndRow	' 자료유형 
		ggoSpread.SSSetRequired  C_DataLen,   pvStartRow, pvEndRow	' 자료길이 
		ggoSpread.SSSetProtected C_MAJORNM,   pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    End With
End Sub

'========================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx

	'자료유형(Data Type)
	Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("A1018", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)    
	iCodeArr = vbTab & lgF0
    iNameArr = vbTab & lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DataTypeCd			'COLM_DATA_TYPE
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DataTypeNm

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DataTypeCd1			'KEY_DATA_TYPE_1
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DataTypeNm1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DataTypeCd2			'KEY_DATA_TYPE_2
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DataTypeNm2

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DataTypeCd3			'KEY_DATA_TYPE_3
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DataTypeNm3

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DataTypeCd4			'KEY_DATA_TYPE_4
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DataTypeNm4

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DataTypeCd5			'KEY_DATA_TYPE_5
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DataTypeNm5
    
	Call CommonQueryRs("GL_CTRL_FLD,ISNULL(GL_CTRL_NM,'')","A_SUBLEDGER_CTRL","GL_CTRL_NM IS NOT NULL",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	iCodeArr = vbTab & lgF0
    iNameArr = vbTab & lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_GL_CTRL_FLD			'COLM_DATA_TYPE
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_GL_CTRL_NM	
End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_CtrlCD			= iCurColumnPos(1) 
			C_CtrlNM			= iCurColumnPos(2) 
			C_SYSFG				= iCurColumnPos(3) 
			C_CtrlEngNM			= iCurColumnPos(4) 
			C_DataTypeCd		= iCurColumnPos(5) 
			C_DataTypeNm		= iCurColumnPos(6) 
			C_DataLen			= iCurColumnPos(7) 
			C_GL_CTRL_FLD       = iCurColumnPos(8)
			C_GL_CTRL_NM        = iCurColumnPos(9)			
			C_TblID				= iCurColumnPos(10) 
			C_CtrlTblPopUp		= iCurColumnPos(11) 
			C_ColmID			= iCurColumnPos(12)
			C_CtrlColmPopUp1	= iCurColumnPos(13)
			C_ColmIDNM			= iCurColumnPos(14)
			C_CtrlColmPopUp2	= iCurColumnPos(15)
			C_MAJORCD			= iCurColumnPos(16)
			C_MAJORPOPUP		= iCurColumnPos(17)
			C_MAJORNM			= iCurColumnPos(18)
			C_KeyColm1			= iCurColumnPos(19)
			C_CtrlColmPopUp3	= iCurColumnPos(20)
			C_DataTypeCd1		= iCurColumnPos(21)
			C_DataTypeNm1		= iCurColumnPos(22)
			C_KeyColm2			= iCurColumnPos(23)
			C_CtrlColmPopUp4	= iCurColumnPos(24)
			C_DataTypeCd2		= iCurColumnPos(25)
			C_DataTypeNm2		= iCurColumnPos(26)
			C_KeyColm3			= iCurColumnPos(27)
			C_CtrlColmPopUp5	= iCurColumnPos(28)
			C_DataTypeCd3		= iCurColumnPos(29)
			C_DataTypeNm3		= iCurColumnPos(30)
			C_KeyColm4			= iCurColumnPos(31)
			C_CtrlColmPopUp6	= iCurColumnPos(32)
			C_DataTypeCd4		= iCurColumnPos(33)
			C_DataTypeNm4		= iCurColumnPos(34)
			C_KeyColm5			= iCurColumnPos(35)
			C_CtrlColmPopUp7	= iCurColumnPos(36)
			C_DataTypeCd5		= iCurColumnPos(37)
			C_DataTypeNm5		= iCurColumnPos(38)
    End Select
End Sub

 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0
			arrParam(0) = "관리항목 팝업"			' 팝업 명칭 
			arrParam(1) = "A_CTRL_ITEM" 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "관리항목"				' 조건필드의 라벨 명칭 

			arrField(0) = "Ctrl_CD"						' Field명(0)
			arrField(1) = "Ctrl_NM"						' Field명(1)
			arrField(2) = "Ctrl_Eng_NM"					' Field명(2)

			arrHeader(0) = "관리항목코드"			' Header명(0)
			arrHeader(1) = "관리항목명"				' Header명(1)
			arrHeader(2) = "관리항목영문명"			' Header명(2)
		Case 1
			arrParam(0) = "종합코드 팝업"			' 팝업 명칭 
			arrParam(1) = "B_MAJOR" 					' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "종합코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "MAJOR_CD"					' Field명(0)
			arrField(1) = "MAJOR_NM"					' Field명(1)

			arrHeader(0) = "종합코드"				' Header명(0)
			arrHeader(1) = "종합코드명"				' Header명(1)
		Case 2
			frm1.vspdData.Col = C_TblID
			frm1.vspdData.Row = frm1.vspdData.ActiveRow

			arrParam(0) = "테이블명 팝업"
			arrParam(1) = "SYSOBJECTS"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "(XTYPE = " & FilterVar("U", "''", "S") & "  OR XTYPE = " & FilterVar("V", "''", "S") & " ) "
			arrParam(5) = "테이블"

			arrField(0) = "NAME"
			arrHeader(0) = "테이블명"
		Case 3
			arrParam(0) = "필드명 팝업"
			arrParam(1) = "SYSCOLUMNS A, SYSTYPES B"
			arrParam(2) = strCode
			arrParam(3) = ""

			frm1.vspdData.Col = C_TblID
			frm1.vspdData.row = frm1.vspdData.ActiveRow
			arrParam(4) = "A.XTYPE = B.XTYPE AND A.ID = (SELECT id FROM SYSOBJECTS  WHERE (XTYPE = " & FilterVar("U", "''", "S") & "  OR XTYPE = " & FilterVar("V", "''", "S") & " )  and name = " & FilterVar(UCase(frm1.vspdData.text), "''", "S") & ")"
			arrParam(5) = "필드명"

			arrField(0) = "UPPER(A.NAME)"
			arrField(1) = "UPPER(B.NAME)"
			arrField(2) = "UPPER(A.LENGTH)"

			arrHeader(0) = "필드명"
			arrHeader(1) = "필드타입"
		Case 3, 4, 6, 7, 8, 9, 10
			arrParam(0) = "필드명 팝업"
			arrParam(1) = "SYSCOLUMNS A, SYSTYPES B"
			arrParam(2) = strCode
			arrParam(3) = ""

			frm1.vspdData.Col = C_TblID
			frm1.vspdData.row = frm1.vspdData.ActiveRow
			arrParam(4) = "A.XTYPE = B.XTYPE AND A.ID = (SELECT id FROM SYSOBJECTS  WHERE (XTYPE = " & FilterVar("U", "''", "S") & "  OR XTYPE = " & FilterVar("V", "''", "S") & " )  and name = " & FilterVar(UCase(frm1.vspdData.text), "''", "S") & ")"   ' Where Condition
			arrParam(5) = "필드명"

			arrField(0) = "UPPER(A.NAME)"
			arrField(1) = "UPPER(B.NAME)"

			arrHeader(0) = "필드명"
			arrHeader(1) = "필드타입"
		Case Else
			Exit Function
	End Select

	IsOpenPop = True
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If arrRet(0) = "" Then
		Call GridSetFocus(iWhere)
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
			.txtCtrlCd.focus
		End Select
	End With
End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	Dim Row
	Dim colPos

	With frm1
		Select Case iWhere
			Case 0
				.txtCtrlCd.value = Trim(arrRet(0))
				.txtCtrlNm.value = arrRet(1)
			Case 1
				.vspdData.Col = C_MAJORCD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_MAJORNM
				.vspdData.Text = arrRet(1)

				Call vspdData_Change(.vspdData.Col, .vspdData.Row)
			Case 2 '테이블 ID
				.vspdData.Col = C_TblID
				.vspdData.Text = arrRet(0)

				Row = .vspdData.ActiveRow
				Call vspdData_Change(C_TblID, Row)
				lgBlnFlgChgValue = True
			Case 3 ' 컬럼ID
				colPos = .vspdData.ActiveCol
				.vspdData.Col = C_ColmID
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_DataLen
				.vspdData.Text = arrRet(2)
				.vspdData.Col = C_DataTypeCd
				'// 해당부분 찾아 선택하는 부분이 따로 필요함 
				If instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 Then
					.vspdData.value = "2"
					.vspdData.Col = C_DataTypeNm
					.vspdData.value = "2"
				ElseIf instr(1,UCase(arrRet(1)),"DATE")>0 Then
					.vspdData.value = "1"
					.vspdData.Col = C_DataTypeNm
					.vspdData.value = "1"  
				ElseIf instr(1,UCase(arrRet(1)),"CHAR")>0 Then
					.vspdData.value = "3"
					.vspdData.Col = C_DataTypeNm
					.vspdData.value = "3"
				End If
				
				Row = .vspdData.Row
				Call vspdData_Change(colPos, Row)
				lgBlnFlgChgValue = True
			Case 4													'컬럼명ID
				colPos = .vspdData.ActiveCol
				.vspdData.Col = C_ColmIDNM
				.vspdData.value = arrRet(0)
				lgBlnFlgChgValue = True
				Row = .vspdData.Row
				Call vspdData_Change(colPos, Row)
				lgBlnFlgChgValue = True
			Case 6
				colPos = .vspdData.ActiveCol
				.vspdData.Col = C_KeyColm1
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_DataTypeCd1
				'// 해당부분 찾아 선택하는 부분이 따로 필요함 
				If instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 Then
					 .vspdData.value = "2"
					 .vspdData.Col = C_DataTypeNm1
					 .vspdData.value = "2"
				ElseIf instr(1,UCase(arrRet(1)),"DATE")>0 Then
					 .vspdData.value = "1"
					 .vspdData.Col = C_DataTypeNm1
					 .vspdData.value = "1"  
				ElseIf instr(1,UCase(arrRet(1)),"CHAR")>0 Then
					 .vspdData.value = "3"
					 .vspdData.Col = C_DataTypeNm1
					 .vspdData.value = "3"  
				End If

				Row = .vspdData.Row
				Call vspdData_Change(colPos, Row)
				lgBlnFlgChgValue = True
			Case 7
				colPos = .vspdData.ActiveCol
				.vspdData.Col = C_KeyColm2
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_DataTypeCd2
				'// 해당부분 찾아 선택하는 부분이 따로 필요함 
				If instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 Then
					 .vspdData.value = "2"
					 .vspdData.Col = C_DataTypeNm2
					 .vspdData.value = "2"
				ElseIf instr(1,UCase(arrRet(1)),"DATE")>0 Then
					 .vspdData.value = "1"
					 .vspdData.Col = C_DataTypeNm2
					 .vspdData.value = "1"  
				ElseIf instr(1,UCase(arrRet(1)),"CHAR")>0 Then
					 .vspdData.value = "3"
					 .vspdData.Col = C_DataTypeNm2
					 .vspdData.value = "3"  
				End If

				Row = .vspdData.Row
				Call vspdData_Change(colPos, Row)
				lgBlnFlgChgValue = True
			Case 8
				colPos = .vspdData.ActiveCol
				.vspdData.Col = C_KeyColm3
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_DataTypeCd3
				'// 해당부분 찾아 선택하는 부분이 따로 필요함 
				If instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 Then
					 .vspdData.value = "2"
					 .vspdData.Col = C_DataTypeNm3
					 .vspdData.value = "2"
				ElseIf instr(1,UCase(arrRet(1)),"DATE")>0 Then
					 .vspdData.value = "1"
					 .vspdData.Col = C_DataTypeNm3
					 .vspdData.value = "1"  
				ElseIf instr(1,UCase(arrRet(1)),"CHAR")>0 Then
					 .vspdData.value = "3"
					 .vspdData.Col = C_DataTypeNm3
					 .vspdData.value = "3"  
				End If
				
				Row = .vspdData.Row
				Call vspdData_Change(colPos, Row)
				lgBlnFlgChgValue = True
			Case 9
				colPos = .vspdData.ActiveCol
				.vspdData.Col = C_KeyColm4
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_DataTypeCd4
				'// 해당부분 찾아 선택하는 부분이 따로 필요함 
				If instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 Then
					 .vspdData.value = "2"
					 .vspdData.Col = C_DataTypeNm4
					 .vspdData.value = "2"
				ElseIf instr(1,UCase(arrRet(1)),"DATE")>0 Then
					 .vspdData.value = "1"
					 .vspdData.Col = C_DataTypeNm4
					 .vspdData.value = "1"  
				ElseIf instr(1,UCase(arrRet(1)),"CHAR")>0 Then
					 .vspdData.value = "3"
					 .vspdData.Col = C_DataTypeNm4
					 .vspdData.value = "3"  
				End If
				
				Row = .vspdData.Row
				Call vspdData_Change(colPos, Row)
				lgBlnFlgChgValue = True
  			Case 10
				colPos = .vspdData.ActiveCol
				.vspdData.Col = C_KeyColm5
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_DataTypeCd5
				'// 해당부분 찾아 선택하는 부분이 따로 필요함 
				If instr(1,UCase(arrRet(1)),"INT")>0 or  instr(1,UCase(arrRet(1)),"NUMERIC") >0 Then
					 .vspdData.value = "2"
					 .vspdData.Col = C_DataTypeNm5
					 .vspdData.value = "2"
				ElseIf instr(1,UCase(arrRet(1)),"DATE")>0 Then
					 .vspdData.value = "1"
					 .vspdData.Col = C_DataTypeNm5
					 .vspdData.value = "1"  
				ElseIf instr(1,UCase(arrRet(1)),"CHAR")>0 Then
					 .vspdData.value = "3"
					 .vspdData.Col = C_DataTypeNm5
					 .vspdData.value = "3"  
				End If
				
				Row = .vspdData.Row
				Call vspdData_Change(colPos, Row)
				lgBlnFlgChgValue = True
		End Select
	End With
End Function



'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
    Call InitSpreadSheet                              '⊙: Setup the Spread Sheet
    Call InitVariables                            '⊙: Initializes local global Variables

    Call SetToolbar("110011010010111")										'⊙: 버튼 툴바 제어 
    frm1.txtCtrlCD.focus 
End Sub

'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    Dim strTemp
    Dim intPos1

    If Row <= 0 Then
        Exit Sub
    End If

    With frm1
        Select Case Col
            Case C_CtrlTblPopUp '테이블ID
                .vspdData.Col = Col-1
                .vspdData.Row = Row
                Call OpenPopUp(.vspdData.Text, 2)
            Case C_CtrlColmPopUp1 '컬럼ID
                .vspdData.Col = Col-1
                .vspdData.Row = Row
                Call OpenPopUp(.vspdData.Text, 3)
            Case C_CtrlColmPopUp2 '컬럼명ID
                .vspdData.Col = Col-1
                .vspdData.Row = Row
                Call OpenPopUp(.vspdData.Text, 4)
            Case C_MAJORPOPUP '종합코드 
                .vspdData.Col = Col-1
                .vspdData.Row = Row
                Call OpenPopUp(.vspdData.Text, 1)
            Case C_CtrlColmPopUp3 'Key컬럼1
                .vspdData.Col = Col-1
                .vspdData.Row = Row
                Call OpenPopUp(.vspdData.Text, 6)
            Case C_CtrlColmPopUp4 'Key컬럼2
                .vspdData.Col = Col-1
                .vspdData.Row = Row
                Call OpenPopUp(.vspdData.Text, 7)
            Case C_CtrlColmPopUp5 'Key컬럼3
                .vspdData.Col = Col-1
                .vspdData.Row = Row
                Call OpenPopUp(.vspdData.Text, 8)
            Case C_CtrlColmPopUp6 'Key컬럼4
                .vspdData.Col = Col-1
                .vspdData.Row = Row
                Call OpenPopUp(.vspdData.Text, 9)
            Case C_CtrlColmPopUp7 'Key컬럼5
                .vspdData.Col = Col-1
                .vspdData.Row = Row
                Call OpenPopUp(.vspdData.Text, 10)
        End Select

'		Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")
    End With
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case  C_GL_CTRL_NM
				.Col = Col
				intIndex = .Value
				.Col = C_GL_CTRL_FLD
				.Value = intIndex		
			Case  C_DataTypeNm
				.Col = Col
				intIndex = .Value
				.Col = C_DataTypeCd
				.Value = intIndex
			Case  C_DataTypeNm1
				.Col = Col
				intIndex = .Value
				.Col = C_DataTypeCd1
				.Value = intIndex
			Case  C_DataTypeNm2
				.Col = Col
				intIndex = .Value
				.Col = C_DataTypeCd2
				.Value = intIndex
			Case  C_DataTypeNm3
				.Col = Col
				intIndex = .Value
				.Col = C_DataTypeCd3
				.Value = intIndex
			Case  C_DataTypeNm4
				.Col = Col
				intIndex = .Value
				.Col = C_DataTypeCd4
				.Value = intIndex
			Case  C_DataTypeNm5
				.Col = Col
				intIndex = .Value
				.Col = C_DataTypeCd5
				.Value = intIndex
		End Select
	End With
End Sub

'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow
			
			.Col = C_GL_CTRL_FLD
			intIndex = .value
			.col = C_GL_CTRL_NM
			.value = intindex

			.Col = C_DataTypeCd
			intIndex = .value
			.col = C_DataTypeNm
			.value = intindex

			.Col = C_DataTypeCd1
			intIndex = .value
			.col = C_DataTypeNm1
			.value = intindex

			.Col = C_DataTypeCd2
			intIndex = .value
			.col = C_DataTypeNm2
			.value = intindex

			.Col = C_DataTypeCd3
			intIndex = .value
			.col = C_DataTypeNm3
			.value = intindex

			.Col = C_DataTypeCd4
			intIndex = .value
			.col = C_DataTypeNm4
			.value = intindex

			.Col = C_DataTypeCd5
			intIndex = .value
			.col = C_DataTypeNm5
			.value = intindex
		Next
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If

    gMouseClickStatus = "SPC"   
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
        Exit Sub
    End If
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================= 
Sub vspdData_MouseDown(Button , ShIft , x , y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================= 
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx,strText

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

    Select Case Col
        Case C_TblID 
            '//imsi
            strText = UCase(Trim(frm1.vspdData.text))
            With frm1 
                '//테이블이하 필드, 자료유형 을 모두 지워주기 
                .vspdData.Col = C_DataTypeCd
                .vspdData.Text = ""
                .vspdData.Col = C_DataTypeNm
                .vspdData.Text = ""
'                .vspdData.Col = C_DataLen
'                .vspdData.Text = ""
                .vspdData.Col = C_ColmID
                .vspdData.Text = ""
                .vspdData.Col = C_ColmIDNM
                .vspdData.Text = ""
                .vspdData.Col = C_KeyColm1
                .vspdData.Text = ""
                .vspdData.Col = C_DataTypeCd1
                .vspdData.Text = ""
                .vspdData.Col = C_DataTypeNm1
                .vspdData.Text = ""
                .vspdData.Col = C_KeyColm2
                .vspdData.Text = ""
                .vspdData.Col = C_DataTypeCd2
                .vspdData.Text = ""
                .vspdData.Col = C_DataTypeNm2
                .vspdData.Text = ""
                .vspdData.Col = C_KeyColm3
                .vspdData.Text = ""
                .vspdData.Col = C_DataTypeCd3
                .vspdData.Text = ""
                .vspdData.Col = C_DataTypeNm3
                .vspdData.Text = ""
                .vspdData.Col = C_KeyColm4
                .vspdData.Text = ""
                .vspdData.Col = C_DataTypeCd4
                .vspdData.Text = ""
                .vspdData.Col = C_DataTypeNm4
                .vspdData.Text = ""
                .vspdData.Col = C_KeyColm5
                .vspdData.Text = ""
                .vspdData.Col = C_DataTypeCd5
                .vspdData.Text = ""
                .vspdData.Col = C_DataTypeNm5
                .vspdData.Text = ""

				'//자동으로 키컬럼을 가져와 세팅해주기.
                .vspdData.Col = C_TblID
                If Trim(.vspdData.Text) <> "" Then
                    Call DBQUERY_SEVEN(Row)
                End If
            End With
        Case  C_DataTypeNm,C_DataTypeNm1,C_DataTypeNm2,C_DataTypeNm3,C_DataTypeNm4,C_DataTypeNm5
            strText = UCase(Trim(frm1.vspdData.text))
            If strText = "STRING" Then
                iDx = "3" 
            ElseIf strText = "NUMERIC" Then
                iDx = "2"
            ElseIf strText = "DATE" Then
                iDx = "1"
            End If
            Frm1.vspdData.value = iDx
		Case C_KeyColm1
			frm1.vspdData.col = C_KeyColm1
		    strText = UCase(Trim(frm1.vspdData.text))
           If strText = "" Then
			ENd If
		CAse C_KeyColm2
			frm1.vspdData.col = C_KeyColm2
		    strText = UCase(Trim(frm1.vspdData.text))
            If strText = "" Then
			ENd If
		Case C_KeyColm3
			frm1.vspdData.col = C_KeyColm3
		    strText = UCase(Trim(frm1.vspdData.text))
            If strText = "" Then
			ENd If
		Case C_KeyColm4
			frm1.vspdData.col = C_KeyColm4
		    strText = UCase(Trim(frm1.vspdData.text))
            If strText = "" Then
			ENd If
		Case C_KeyColm5
			frm1.vspdData.col = C_KeyColm5
		    strText = UCase(Trim(frm1.vspdData.text))
    End Select

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub 

'========================================================================================================= 
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
			Exit Sub
		End If
    End With
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
    	End If
    End If
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
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
	Call ggoOper.ClearField(Document, "2")
    Call InitVariables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then
		Exit Function
    End If

    Call DbQuery

    FncQuery = True
End Function


'========================================================================================
Function FncSave() 
	Dim IntRetCD 
    
    FncSave = False
    Err.Clear

    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = False Then   
        IntRetCD = DisplayMsgBox("900001","x","x","x")                          'No data changed!!
        Exit Function
    End If

    '-----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData

    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
		Exit Function
    End If

    Call DbSave

    FncSave = True
End Function

'========================================================================================
Function FncCopy()
	Dim IntRetCD

	frm1.vspdData.ReDraw = False

    If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    frm1.vspdData.Col = C_CtrlCd
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_SYSFG
    frm1.vspdData.Text = "N"

	frm1.vspdData.ReDraw = True
End Function

'========================================================================================
Function FncCancel() 
    If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo

	Call InitData()
End Function

Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    
    On Error Resume Next
    Err.Clear

    FncInsertRow = False

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
    imRow = AskSpdSheetAddRowCount()

    If imRow = "" Then
        Exit Function
		End If
    End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With

	If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
Function FncDeleteRow() 
	Dim lDeIRows
	Dim iDeIRowCnt, i
	Dim IntRetCD 

    If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 

		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_SYSFG
		If Trim(frm1.vspdData.Text) = "Y" Then
		    IntRetCD = DisplayMsgBox("183116","x","x","x")	'삭제할수 없습니다.
		    Exit Function
		End If

		lDeIRows = ggoSpread.DeleteRow
    End With
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
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()

    SetSpreadLock "I", 0, 1, ""	
End Sub

'========================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   '☜ 바뀐부분 
		 'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
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
    Err.Clear

	Call LayerShowHide(1)

    With frm1
	    If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtCtrlCd=" & Trim(.hCtrlCd.value)	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtCtrlCd=" & Trim(.txtCtrlCd.value)	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)		'☜: 비지니스 ASP 를 가동 
    End With

    DbQuery = True
End Function

'=======================================================================================================
' Function Name : DbQuery_Seven
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbQuery_Seven(ByVal Row)
    Dim strSelect
    Dim strFrom
    Dim strWhere, strWhere1  
    Dim strTableName
    Dim strField
    Dim strKeyField1
    Dim strKeyField2
    Dim strKeyField3
    Dim strKeyField4
    Dim strKeyField5
    Dim strDataType
    Dim strKeyType1
    Dim strKeyType2
    Dim strKeyType3
    Dim strKeyType4
    Dim strKeyType5
    Dim Rs0,Rs1
    Dim arrFldName
    Dim arrDataType
    Dim i,j
    Dim strFieldList
    '//입력, 수정한 테이블이름과 컬럼의 유효값 체크 
    Err.Clear

    With frm1
        .vspdData.Row = Row
        .vspdData.col = C_TblID
        strTableName = Trim(.vspdData.text)

        strSelect = " UPPER(c.name),UPPER(t.name) "
        strFrom  =  " sysindexes i, syscolumns c, sysobjects o, systypes t "
        strWhere =  " o.name = " & FilterVar(UCase(strTableName), "''", "S") 
        strWhere = strWhere & " and o.id = c.id "
        strWhere = strWhere & " and o.id = i.id "
        strWhere = strWhere & " and (i.status & 0x800) = 0x800 "
        strWhere = strWhere & " and c.xtype = t.xtype "
        strWhere = strWhere & " and  (o.XTYPE = " & FilterVar("U", "''", "S") & "  OR o.XTYPE = " & FilterVar("V", "''", "S") & " )  "
        strWhere = strWhere & " and (c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  1) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  2) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  3) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  4) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  5) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  6) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  7) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  8) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid,  9) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 10) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 11) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 12) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 13) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 14) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 15) or "
        strWhere = strWhere & "      c.name = index_col ( " & FilterVar(UCase(strTableName), "''", "S") & ", i.indid, 16)    "
        strWhere = strWhere & "      )                                              "
        strWhere = strWhere & "  order by c.colid                                   "

		If CommonQueryRs2by2(strSelect,strFrom,strWhere,lgF2By2) = False Then
		    Exit Function
		Else
		    Rs1 = split(lgF2By2,chr(11) & chr(12) )
		    i = 0
		    j = Ubound(rs1)
		    Redim arrFldName(j)
		    Redim arrDaType(j)

			Do While i< j
			    Rs0 = split(rs1(i), chr(11))
			    arrFldName(i) =Rs0(1) 
			    If instr(1,UCase(Rs0(2)),"INT")>0 or  instr(1,UCase(Rs0(2)),"NUMERIC") >0 Then
			        '//msgbox "int"
			        arrDaType(i) = "2"
			    ElseIf instr(1,UCase(Rs0(2)),"DATE")>0 Then
			        '//msgbox "date"
			        arrDaType(i) = "1"
			    ElseIf instr(1,UCase(Rs0(2)),"CHAR")>0 Then
			        '//msgbox "char"
			        arrDaType(i) = "3"
			    End If
			    i = i+1
			Loop

			'//field////////////////////////////
			If j>=1 Then
			    .vspdData.col      = C_KeyColm1
			    .vspdData.text     = arrFldName(0)
			    .vspdData.col      = C_DataTypeNm1
			    .vspdData.value    = arrDaType(0)
			    .vspdData.col      = C_DataTypeCd1
			    .vspdData.value    = arrDaType(0)
			End If

			If j>=2 Then
			    .vspdData.col      = C_KeyColm2
			    .vspdData.text     = arrFldName(1)
			    .vspdData.col      = C_DataTypeNm2
			    .vspdData.value    = arrDaType(1)
			    .vspdData.col      = C_DataTypeCd2
			    .vspdData.value    = arrDaType(1)
			End If

			If j>=3 Then
			    .vspdData.col      = C_KeyColm5
			    .vspdData.text     = arrFldName(2)
			    .vspdData.col      = C_DataTypeNm3
			    .vspdData.value    = arrDaType(2)
			    .vspdData.col      = C_DataTypeCd3
			    .vspdData.value    = arrDaType(2)
			End If

			If j>=4 Then
			    .vspdData.col      = C_KeyColm5
			    .vspdData.text     = arrFldName(3)
			    .vspdData.col      = C_DataTypeNm4
			    .vspdData.value    = arrDaType(3)
			    .vspdData.col      = C_DataTypeCd4
			    .vspdData.value    = arrDaType(3)
			End If

            If j>=5 Then
                .vspdData.col      = C_KeyColm5
                .vspdData.text     = arrFldName(4)
                .vspdData.col      = C_DataTypeNm5
                .vspdData.value    = arrDaType(4)
                .vspdData.col      = C_DataTypeCd5
                .vspdData.value    = arrDaType(4)
            End If
        End If
    End With
End Function

'========================================================================================
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")
    Call InitData
	Call SetToolbar("110011110011111")
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
Function DbSave() 
	Dim IRow
	Dim lGrpCnt
	Dim strVal, strDel

    DbSave = False
    On Error Resume Next
    Err.Clear 

	Call LayerShowHide(1)

	With frm1
		.txtMode.value = Parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""

		'-----------------------
		'Data manipulate area
		'-----------------------
		For IRow = 1 To .vspdData.MaxRows
		    .vspdData.Row = IRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag															'☜: 신규 
					strVal = strVal & "C" & Parent.gColSep & IRow & Parent.gColSep					'☜: C=Create
		            .vspdData.Col = C_CtrlCD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_CtrlNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_SYSFG
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_CtrlEngNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataLen
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_GL_CTRL_FLD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		            
		            .vspdData.Col = C_TblID
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_ColmID
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_ColmIDNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MAJORCD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KeyColm1
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd1
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KeyColm2
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd2
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KeyColm3
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd3
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KeyColm4
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd4
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KeyColm5
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd5
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
				Case ggoSpread.UpdateFlag															'☜: 수정 
					strVal = strVal & "U" & Parent.gColSep & IRow & Parent.gColSep					'☜: U=Update
		            .vspdData.Col = C_CtrlCD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_CtrlNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_SYSFG
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_CtrlEngNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataLen
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_GL_CTRL_FLD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		            		            
		            .vspdData.Col = C_TblID
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_ColmID
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_ColmIDNM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MAJORCD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KeyColm1
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd1
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KeyColm2
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd2
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KeyColm3
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd3
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KeyColm4
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd4
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_KeyColm5
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_DataTypeCd5
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
				    lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.DeleteFlag																'☜: 삭제 
					strDel = strDel & "D" & Parent.gColSep & IRow & Parent.gColSep
		            .vspdData.Col = C_CtrlCD
		            strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		    End Select
		Next
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)																'☜: 비지니스 ASP 를 가동 
	End With

    DbSave = True																						'⊙: Processing is NG
End Function

'========================================================================================
Function DbSaveOk()
	Call ggoOper.ClearField(Document, "2")
    Call InitVariables
	Call DbQuery
End Function

'========================================================================================
Sub SetGridFocus()	
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
End Sub

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gIf"><IMG src="../../../CShared/image/table/seltab_up_left.gIf" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="center" CLASS="CLSMTABP"><font color=white>관리항목등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gIf" width="10" height="23" ></td>
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
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100%  CELLSPACING=0>
								<TR>
									<TD CLASS="TD5" NOWRAP>관리항목</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtCtrlCD" MAXLENGTH="3" SIZE=10 ALT ="관리항목" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gIf" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtCtrlCD.Value, 0)">
													<INPUT NAME="txtCtrlNM" MAXLENGTH="30" SIZE=20 ALT ="관리항목명" tag="14X"></TD>
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
								<script language =javascript src='./js/a2102ma1_I818947350_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IfRAME NAME="MyBizASP" src="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IfRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hCtrlCd" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<Iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></Iframe>
</DIV>
</BODY>
</HTML>

