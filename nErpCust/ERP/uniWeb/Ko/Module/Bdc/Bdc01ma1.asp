<%@ LANGUAGE="VBSCRIPT" %>
<!--'==============================================================================================
'*  1. Module Name          : BDC
'*  2. Function Name        : 
'*  3. Program ID           : BDC04MA1
'*  4. Program Name         : BDC 업무등록 
'*  5. Program Desc         : BDC 마스터 데이타와 컴포넌트 정보 입력 
'*  6. Component List       : BDC001
'*  7. Modified date(First) : 2005/01/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kweon, SoonTae
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'===============================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
'☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->
'==================================================================================================
' 상수 및 변수 선언 
'--------------------------------------------------------------------------------------------------
Const BIZ_PGM_QUERY_ID = "BDC01MB1.ASP"								'☆: 조회 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID  = "BDC01MB2.ASP"								'☆: 저장 비지니스 로직 ASP명 

Const C_SP1_COMP_NAM = 1
Const C_SP1_METH_NAM = 2
Const C_SP1_TRET_DSC = 3

Const C_SP2_PARA_NAM = 1
Const C_SP2_PARA_TYP = 2
Const C_SP2_REQU_FLG = 3

Const C_SP3_FILD_NUM = 1
Const C_SP3_SHET_NUM = 2
Const C_SP3_FILD_SEQ = 3
Const C_SP3_FILD_NAM = 4
Const C_SP3_ATTH_CHR = 5

Const C_SP4_SHET_NUM = 1
Const C_SP4_FILD_SEQ = 2
Const C_SP4_FILD_NAM = 3
Const C_SP4_FILD_TYP = 4
Const C_SP4_FILD_FLG = 5
Const C_SP4_PANT_FLD = 6
Const C_SP4_FILD_SND = 7

'==================================================================================================
' 스프레드 시트의 통제를 위한 변수 
Dim nCurrentSpread
Dim nSpreadIndex1
Dim nSpreadIndex2
Dim nSpreadIndex3
Dim nSpreadIndex4

Dim strMode
Dim IsOpenPop
Dim lgRetFlag

' 동적으로 다르게 보여야할 각 스프레드의 값을 보관할 배열을 지정한다.
Dim arrParams(10)
Dim arrJoins(10, 20)

'==================================================================================================
' 페이지 로드가 완료되면 자동으로 호출되는 함수.
' 초기화 루틴을 이곳에 집중시켜 주어야 함.
' ../../inc/incCliMAMain.vbs 파일에 이 함수를 호출 하도록 하는 모듈이 있슴 
'--------------------------------------------------------------------------------------------------
Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")
	Call InitSpreadSheet
	Call InitVariables
	Call InitComboBox
	Call InitGridComboBox
	Call SetToolbar("1100111000000111")
	frm1.txtProcId.focus
	
	objCurSpread = frm1.vspdData
End Sub

'==================================================================================================
' 시스템에 설정된 화폐단위, 언어코드, 등등등의 설정값을 초기화 하는 함수.
' ../../inc/incCliVariables.vbs 과 ../../ComAsp/LoadInfTB19029.asp  파일에 종속적이다.
'--------------------------------------------------------------------------------------------------
Sub LoadInfTB19029()
<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
End Sub

'==================================================================================================
' 스프레드 초기화 함수 
' 프로그램에 따라 사용자들이 조정해 주어야 하는 부분 
'--------------------------------------------------------------------------------------------------
Sub InitSpreadSheet()
	With frm1.vspdData
        .ReDraw = False
		.MaxCols = C_SP1_TRET_DSC
        .MaxRows = 0

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit			'"V20021121",,parent.gAllowDragDropSpread
		ggoSpread.SSSetEdit   C_SP1_COMP_NAM,  "컴포넌트", 25, , , 80
		ggoSpread.SSSetEdit   C_SP1_METH_NAM,  "메 소 드", 25, , , 80
		ggoSpread.SSSetEdit   C_SP1_TRET_DSC,  "설    명", 44, , , 128
		.ReDraw = True
	End With
	
	With frm1.vspdData1
        .ReDraw = False
		.MaxCols = C_SP2_REQU_FLG
        .MaxRows = 0

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit			'"V20021121",,parent.gAllowDragDropSpread
		ggoSpread.SSSetEdit   C_SP2_PARA_NAM,  "인자명", 20, , , 80
		ggoSpread.SSSetCombo  C_SP2_PARA_TYP,  "형태", 8, 2, False
		ggoSpread.SSSetCombo  C_SP2_REQU_FLG,  "필수", 8, 2, False
		.ReDraw = True
	End With

	With frm1.vspdData2
        .ReDraw = False
		.MaxCols = C_SP3_ATTH_CHR
        .MaxRows = 0

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit			'"V20021121",,parent.gAllowDragDropSpread
		ggoSpread.SSSetEdit   C_SP3_FILD_NUM,  "", 2, , , 2
		ggoSpread.SSSetEdit   C_SP3_SHET_NUM,  "시트", 8, , , 1
		ggoSpread.SSSetEdit   C_SP3_FILD_SEQ,  "필드", 8, , , 2
		ggoSpread.SSSetEdit   C_SP3_FILD_NAM,  "필 드 명", 15, , , 40
		ggoSpread.SSSetCombo  C_SP3_ATTH_CHR,  "첨부", 8, 2, False
		Call ggoSpread.SSSetColHidden(C_SP3_FILD_NUM, C_SP3_FILD_NUM, True)

		.ReDraw = True
	End With

	ggoSpread.Source = frm1.vspdData3
	ggoSpread.Spreadinit				'"V20021121",,parent.gAllowDragDropSpread
	With frm1.vspdData3
        .ReDraw = False
		.MaxCols = C_SP4_FILD_SND
        .MaxRows = 0

		ggoSpread.SSSetEdit   C_SP4_SHET_NUM,  "시트", 8, , , 1
		ggoSpread.SSSetEdit   C_SP4_FILD_SEQ,  "필드", 8, , , 2
		ggoSpread.SSSetEdit   C_SP4_FILD_NAM,  "필 드 명", 25, , , 40
		ggoSpread.SSSetCombo  C_SP4_FILD_TYP,  "타입", 8, 2, False
		ggoSpread.SSSetCombo  C_SP4_FILD_FLG,  "필수", 8, 2, False
		ggoSpread.SSSetEdit   C_SP4_PANT_FLD,  "연결", 8, , , 2
		ggoSpread.SSSetButton C_SP4_FILD_SND   
		.ReDraw = True
	End With
	

	
End Sub

'==================================================================================================
' 광역 변수들을 초기화 시킨다.
'--------------------------------------------------------------------------------------------------
Sub InitVariables()
	Dim i, j

    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1

	nCurrentSpread = 0
	nSpreadIndex1 = 0
	nSpreadIndex2 = 0
	nSpreadIndex3 = 0
	nSpreadIndex4 = 0

    For i = 0 To 10
		arrParams(i) = ""
	Next

	For i = 0 To 10
		For j = 0 To 20
			arrJoins(i, j) = ""
		Next
	Next
End Sub

'==================================================================================================
' 스프레드시트 이외의 콤보박스들을 초기화 한다.
'--------------------------------------------------------------------------------------------------
Sub InitComboBox()
End Sub

'==================================================================================================
' 스프레드 시트의 콤보박스의 값을 초기화 한다.
'--------------------------------------------------------------------------------------------------
Sub InitGridComboBox()
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SetCombo "Var" & vbTab & "Str" & vbTab & "Num", C_SP2_PARA_TYP
    ggoSpread.SetCombo "Y" & vbTab & "N", C_SP2_REQU_FLG

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.SetCombo "N" & vbTab & "C" & vbTab & "R", C_SP3_ATTH_CHR

    ggoSpread.Source = frm1.vspdData3
    ggoSpread.SetCombo "Num" & vbTab & "Str" & vbTab & "Flo", C_SP4_FILD_TYP

    ggoSpread.Source = frm1.vspdData3
    ggoSpread.SetCombo "Y" & vbTab & "N", C_SP4_FILD_FLG
End Sub

'==================================================================================================
' 업무코드 참조 팝업 창을 생성시킨다.
'--------------------------------------------------------------------------------------------------
Function OpenPopup(Byval StrCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrParam(0) = "업무팝업"				' 팝업 명칭 
	arrParam(1) = "B_BDC_MASTER"			' TABLE 명칭 
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = " USE_FLAG= " & Filtervar("Y", "''", "S")				' Code Condition
	arrParam(5) = "업무"

	arrField(0) = "PROCESS_ID"				' Field명(0)
	arrField(1) = "PROCESS_NAME"			' Field명(1)

	arrHeader(0) = "업무코드"				' Header명(0)
	arrHeader(1) = "업무명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
	                                Array(arrParam, arrField, arrHeader), _
		                            "dialogWidth=420px; dialogHeight=450px; " & _
		                            "center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtProcId.focus
		Exit Function
	Else
		frm1.txtProcId.focus
		frm1.txtProcId.value = arrRet(0)
		frm1.txtProcNm.value = arrRet(1)
	End If
End Function

'==================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'--------------------------------------------------------------------------------------------------
Function  FncInsertRow(ByVal pvRowCnt)
    Dim i, j, k
    Dim arrTempRow, arrTempCol
    Dim iColSep, iRowSep

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep

    FncInsertRow = False

	Select Case (nCurrentSpread)
		Case 0
			'--------------------------------------------------------------------------------------
			' 두번째 스프레드와 세번째 스프레드의 값이 합당한지 검사 해 본다.
			Set gActiveSpdSheet = frm1.vspdData1
			ggoSpread.Source = frm1.vspdData1
			If Not ggoSpread.SSDefaultCheck Then
				Exit Function
			End If

			Set gActiveSpdSheet = frm1.vspdData2
			ggoSpread.Source = frm1.vspdData2
			If Not ggoSpread.SSDefaultCheck Then
				Exit Function
			End If

			If nSpreadIndex1 > 0 Then
				'--------------------------------------------------------------------------------------
				' 두번째 스프레드와 세번째 스프레드의 값을 배열로 읽어 들인다.
				arrParams(nSpreadIndex1-1) = ""
				For i = 1 To frm1.vspdData1.MaxRows
					arrParams(nSpreadIndex1-1) = arrParams(nSpreadIndex1-1) & _
								Chr(11) & GetSpreadText(frm1.vspdData1,C_SP2_PARA_NAM,i,"X","X") & _
								Chr(11) & GetSpreadText(frm1.vspdData1,C_SP2_PARA_TYP,i,"X","X") & _
								Chr(11) & GetSpreadText(frm1.vspdData1,C_SP2_REQU_FLG,i,"X","X") & _
								Chr(11) & Chr(12)
				Next

				If nSpreadIndex2 > 0 Then
					' 현재 vspdData2 의 자료를 저장한다.
					arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = ""
					For i = 1 To frm1.vspdData2.MaxRows
						arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) & _
									Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X") & _
									Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_SHET_NUM, i, "X", "X") & _
									Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_SEQ, i, "X", "X") & _
									Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NAM, i, "X", "X") & _
									Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_ATTH_CHR, i, "X", "X") & _
									Chr(11) & Chr(12)
					Next
				End If
			End If
			'--------------------------------------------------------------------------------------
			' 현재 행위치에 새로운 행을 생성시키고 현재행 이후를 하나씩 뒤로 밀어낸다.
			' 새로운 행이 컴포넌트 정보가 추가 되므로 
			' 현재 컴포넌트가 가리키고 있는 것 이후에 것의 컴포넌트 ID가 하나씩 밀려야 한다.
			' 또한 해당 컴포넌트가 가리키고 있는 파라메터 배열도 한칸씩 밀려야 한다.
			' 파라메터가 가리키고 있는 결합정보 배열도 한칸씩 밀려야 한다.
			' Dim arrParams(10)
			' Dim arrJoins(10, 20)
			For i = 10 To nSpreadIndex1 Step -1
				If arrParams(i) <> "" Then
					For j = 20 To 0 Step -1
						If arrJoins(i, j) <> "" Then
							arrJoins(i+1, j) = arrJoins(i, j)
							arrJoins(i, j) = ""
						End If
					Next
					arrParams(i+1) = arrParams(i)
					arrParams(i) = ""
				End If
			Next

			'--------------------------------------------------------------------------------------
			' 두번째 스프레드와 세번째 스프레드를 초기화 시켜준다.
			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.ClearSpreadData()
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ClearSpreadData()
			
			'--------------------------------------------------------------------------------------
			' 첫번째 스프레드에 새로운 행을 추가 시킨다.
			With frm1
				.vspdData.focus
				.vspdData.ReDraw = False
				ggoSpread.Source = .vspdData
				ggoSpread.InsertRow nSpreadIndex1, 1
				Call SetSpreadColor(0, .vspdData.ActiveRow, .vspdData.ActiveRow)
				nSpreadIndex1 = .vspdData.ActiveRow
				nSpreadIndex2 = 0
				nSpreadIndex3 = 0
				.vspdData.ReDraw = True
			End With

		Case 1
			'--------------------------------------------------------------------------------------
			' 세번째 스프레드의 값이 합당한지 검사 해 본다.
			Set gActiveSpdSheet = frm1.vspdData
			ggoSpread.Source = frm1.vspdData
			If Not ggoSpread.SSDefaultCheck Then
				Exit Function
			End If

			Set gActiveSpdSheet = frm1.vspdData2
			ggoSpread.Source = frm1.vspdData2
			If Not ggoSpread.SSDefaultCheck Then
				Exit Function
			End If
			If nSpreadIndex1 = 0 Or frm1.vspdData.MaxRows = 0 Then
				MsgBox "메소드가 선택되지 않았습니다."
				Exit Function
			End If

			If nSpreadIndex1 > 0 And nSpreadIndex2 > 0 And frm1.vspdData.MaxRows > 0 Then
				'--------------------------------------------------------------------------------------
				' 세번째 스프레드의 값을 배열로 읽어 들인다.
				arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = ""
				For i = 1 To frm1.vspdData2.MaxRows
					arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_SHET_NUM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_SEQ, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NAM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_ATTH_CHR, i, "X", "X") & _
								Chr(11) & Chr(12)
				Next
			End If
			'--------------------------------------------------------------------------------------
			' 현재 행위치에 새로운 행을 생성시키고 현재행 이후를 하나씩 뒤로 밀어낸다.
			' 새로운 행이 컴포넌트 정보가 추가 되므로 
			' 현재 컴포넌트가 가리키고 있는 것 이후에 것의 컴포넌트 ID가 하나씩 밀려야 한다.
			' 또한 해당 컴포넌트가 가리키고 있는 파라메터 배열도 한칸씩 밀려야 한다.
			' 파라메터가 가리키고 있는 결합정보 배열도 한칸씩 밀려야 한다.
			' Dim arrParams(10)
			' Dim arrJoins(10, 20)
			For i = 20 To nSpreadIndex2 Step -1
				If arrJoins(nSpreadIndex1-1, i) <> "" Then
					arrJoins(nSpreadIndex1-1, i+1) = arrJoins(nSpreadIndex1-1, i)
					arrJoins(nSpreadIndex1-1, i) = ""
				End If
			Next

			'--------------------------------------------------------------------------------------
			' 세번째 스프레드를 초기화 시켜준다.
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ClearSpreadData()

			'--------------------------------------------------------------------------------------
			' 첫번째 스프레드에 새로운 행을 추가 시킨다.
			With frm1
				.vspdData.focus
				.vspdData.ReDraw = False
				ggoSpread.Source = .vspdData1
				ggoSpread.InsertRow nSpreadIndex2, 1
				Call SetSpreadColor(1, .vspdData1.ActiveRow, .vspdData1.ActiveRow)
				.vspdData.ReDraw = True
				nSpreadIndex2 = .vspdData1.ActiveRow
				nSpreadIndex3 = 0
			End With

		Case 2
			Set gActiveSpdSheet = frm1.vspdData
			ggoSpread.Source = frm1.vspdData
			If Not ggoSpread.SSDefaultCheck Then
				Exit Function
			End If

			Set gActiveSpdSheet = frm1.vspdData1
			ggoSpread.Source = frm1.vspdData1
			If Not ggoSpread.SSDefaultCheck Then
				Exit Function
			End If

			If nSpreadIndex2 > 0 And frm1.vspdData1.MaxRows > 0 Then
				frm1.vspdData2.focus
				frm1.vspdData2.ReDraw = False
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.InsertRow, 1
				
				Call SetSpreadColor(2, frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow)
				frm1.vspdData2.ReDraw = True
				nSpreadIndex3 = frm1.vspdData2.ActiveRow
			Else
				MsgBox "파라메터가 선택되지 않았습니다."
			End If
		Case 3
			'--------------------------------------------------------------------------------------
			' 시트가 추가되었으므로 현재 추가되는 행 이후의 필드 아이디가 1 증가하므로 
			' 이후의 아이디를 가진 모든 배열값을 뒤져서 수정해 주어야 한다.
			' 완전히 생 노가다이다.

			If nSpreadIndex1 > 0 And nSpreadIndex2 > 0 Then
				'--------------------------------------------------------------------------------------
				' 세번째 스프레드의 값을 배열로 읽어 들인다.
				arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = ""
				For i = 1 To frm1.vspdData2.MaxRows
					arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_SHET_NUM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_SEQ, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NAM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_ATTH_CHR, i, "X", "X") & _
								Chr(11) & Chr(12)
					'--------------------------------------------------------------------------------------
					' 만약 해당 필드 아이디가 현재 nSpreadIndex4 보다 크면...
					If GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X") <> "" Then
						If CInt(GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X")) > nSpreadIndex4 Then
							Call SetSpreadValue(frm1.vspdData2, 1, i, _
												CInt(GetSpreadText( frm1.vspdData2, _
																	C_SP3_FILD_NUM, _
																	i, "X", "X")) + 1, "", "")
						End If
					End If
				Next
			End If

			'--------------------------------------------------------------------------------------
			' 이제 arrJoins(10,20) 배열을 모두 뒤져서 현재 필드 아이디 이후의 값이 존재하는지 본다.
			' 만약 존재하면 값들을 모두 변경해 준다.
			Dim szTempRow
			For i = 0 To 10
				If arrParams(i) <> "" Then
					For j = 0 To 20
						If arrJoins(i, j) <> "" Then
							'----------------------------------------------------------------------
							' 우선 문자열을 행별로 분할하여 배열에 넣는다.
'							szTempRow = arrJoins(i, j)
							arrTempRow = Split(arrJoins(i, j), iRowSep)
							szTempRow = ""
							For k = 0 To UBound(arrTempRow) -1
								arrTempCol = Split(arrTempRow(k), iColSep)
								If CInt(arrTempCol(1)) > nSpreadIndex4 Then
									szTempRow = szTempRow & _
												arrTempCol(1) + 1 & iColSep & _
												GetSpreadText(frm1.vspdData3, C_SP4_SHET_NUM, arrTempCol(1), "X", "X") & iColSep & _
												GetSpreadText(frm1.vspdData3, C_SP4_FILD_SEQ, arrTempCol(1), "X", "X") & iColSep & _
												GetSpreadText(frm1.vspdData3, C_SP4_FILD_NAM, arrTempCol(1), "X", "X") & iColSep & _
												arrTempCol(5) & iColSep & iRowSep
								Else
									szTempRow = szTempRow & arrTempRow(k) & iRowSep
								End If
							Next
							arrJoins(i, j) = szTempRow
						End If
					Next
				End If
			Next

			frm1.vspdData3.focus
			frm1.vspdData3.ReDraw = False
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.InsertRow, 1

			Call SetSpreadColor(3, frm1.vspdData3.ActiveRow, frm1.vspdData3.ActiveRow)
			frm1.vspdData3.ReDraw = True
			nSpreadIndex4 = frm1.vspdData3.ActiveRow
	End Select

	If Err.number = 0 Then
		FncInsertRow = True                                                  '☜: Processing is OK
	End If

    Set gActiveElement = document.ActiveElement
End Function

'==================================================================================================
' 현재 선택한 행을 삭제한다.
'--------------------------------------------------------------------------------------------------
Function FncDeleteRow()
	FncDeleteRow = false
	
	Dim szTemp, arrTempRow, arrTempCol
	Dim i, j, k
    Dim iColSep, iRowSep

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep

	Select Case (nCurrentSpread)
		Case 0
			If frm1.vspdData.MaxRows < 1 Or nSpreadIndex1 < 1 Then Exit Function
			
			'--------------------------------------------------------------------------------------
			' 현재선택된 행의 하위 정보들을 모두 지운후 현재행을 제외하고 첫번째 스프레드를 다시 
			' 그린후 2 번째 및 3번째 스프레드도 다시 그린다.
			If Confirm("하위 파라메터 및 파라메터 구성 정보도 지우시겠습니까?") Then
				'----------------------------------------------------------------------------------
				' 자 우선 자식들 정보부터 지우자 
				For i = nSpreadIndex1 - 1 To frm1.vspdData.MaxRows - 1
					arrParams(i) = arrParams(i+1)
					For j = 0 To 19
						arrJoins(i, j) = arrJoins(i+1, j)
					Next
				Next

				arrParams(10) = ""
				For j = 0 To 19
					arrJoins(10, j) = ""
				Next
				
				'----------------------------------------------------------------------------------
				' 세개의 스프레드를 지우고 다시 그린다.
				szTemp = ""
				For i = 1 To frm1.vspdData.MaxRows
					If i <> nSpreadIndex1 Then
					szTemp = szTemp & _
							iColSep & Trim(GetSpreadText(frm1.vspdData, C_SP1_COMP_NAM, i, "X", "X")) & _
							iColSep & Trim(GetSpreadText(frm1.vspdData, C_SP1_METH_NAM, i, "X", "X")) & _
							iColSep & Trim(GetSpreadText(frm1.vspdData, C_SP1_TRET_DSC, i, "X", "X")) & iRowSep
					End If
				Next

				ggoSpread.Source = frm1.vspdData
				Call ggoSpread.ClearSpreadData()
				ggoSpread.SSShowData szTemp
				Call SetSpreadColor(0, 1, frm1.vspdData0.MaxRows)
				frm1.vspdData1.ReDraw = True

				ggoSpread.Source = frm1.vspdData1
				Call ggoSpread.ClearSpreadData()
				ggoSpread.SSShowData arrParams(0)
				Call SetSpreadColor(1, 1, frm1.vspdData1.MaxRows)
				frm1.vspdData1.ReDraw = True

				' 현재 vspdData2 의 자료를 변경한다.
				ggoSpread.Source = frm1.vspdData2
				Call ggoSpread.ClearSpreadData()
				ggoSpread.SSShowData arrJoins(0, 0)
				Call SetSpreadColor(2, 1, frm1.vspdData2.MaxRows)
				frm1.vspdData2.ReDraw = True
				
				If frm1.vspdData.MaxRows > 0 Then 
					nSpreadIndex1 = 1
					If frm1.vspdData1.MaxRows > 0 Then
						nSpreadIndex2 = 1
						If frm1.vspdData2.MaxRows > 0 Then
							nSpreadIndex3 = 1
						Else
							nSpreadIndex3 = 0
						End If
					Else
						nSpreadIndex2 = 0
						nSpreadIndex3 = 0
					End If
				Else
					nSpreadIndex1 = 0
					nSpreadIndex2 = 0
					nSpreadIndex3 = 0
				End If
			Else
				Exit Function
			End If
		Case 1
			If frm1.vspdData1.MaxRows < 1 Or nSpreadIndex2 < 1 Then Exit Function
			
			'--------------------------------------------------------------------------------------
			' 현재선택된 행의 하위 정보들을 모두 지운후 현재행을 제외하고 첫번째 스프레드를 다시 
			' 그린후 2 번째 및 3번째 스프레드도 다시 그린다.
			If Confirm("하위 파라메터 구성 정보도 지우시겠습니까?") Then
				'----------------------------------------------------------------------------------
				' 자 우선 자식들 정보부터 지우자 
				For j = nSpreadIndex2 - 1 To frm1.vspdData1.MaxRows - 1
					arrJoins(nSpreadIndex1-1, j) = arrJoins(nSpreadIndex1-1, j+1)
				Next

				arrJoins(nSpreadIndex1-1, j) = ""

				'----------------------------------------------------------------------------------
				' 두개의 스프레드를 지우고 다시 그린다.
				arrParams(nSpreadIndex1-1) = ""
				For i = 1 To frm1.vspdData1.MaxRows
					If i <> nSpreadIndex2 Then
					arrParams(nSpreadIndex1-1) = arrParams(nSpreadIndex1-1) & _
							 iColSep & GetSpreadText(frm1.vspdData1,C_SP2_PARA_NAM,i,"X","X") & _
							 iColSep & GetSpreadText(frm1.vspdData1,C_SP2_PARA_TYP,i,"X","X") & _
							 iColSep & GetSpreadText(frm1.vspdData1,C_SP2_REQU_FLG,i,"X","X") & iRowSep
					End If
				Next

				ggoSpread.Source = frm1.vspdData1
				Call ggoSpread.ClearSpreadData()
				ggoSpread.SSShowData arrParams(nSpreadIndex1-1)
				Call SetSpreadColor(1, 1, frm1.vspdData1.MaxRows)
				frm1.vspdData1.ReDraw = True

				' 현재 vspdData2 의 자료를 변경한다.
				ggoSpread.Source = frm1.vspdData2
				Call ggoSpread.ClearSpreadData()
				ggoSpread.SSShowData arrJoins(nSpreadIndex1-1, 0)
				Call SetSpreadColor(2, 1, frm1.vspdData2.MaxRows)
				frm1.vspdData2.ReDraw = True

				If frm1.vspdData1.MaxRows > 0 Then
					nSpreadIndex2 = 1
					If frm1.vspdData2.MaxRows > 0 Then
						nSpreadIndex3 = 1
					Else
						nSpreadIndex3 = 0
					End If
				Else
					nSpreadIndex2 = 0
					nSpreadIndex3 = 0
				End If
			Else
				Exit Function
			End If
		Case 2
			If frm1.vspdData2.MaxRows < 1 Or nSpreadIndex3 < 1 Then Exit Function
			
			'--------------------------------------------------------------------------------------
			' 현재선택된 행의 하위 정보들을 모두 지운후 현재행을 제외하고 첫번째 스프레드를 다시 
			' 그린후 2 번째 및 3번째 스프레드도 다시 그린다.
			If Confirm("현재행을 지우시겠습니까?") Then
				'----------------------------------------------------------------------------------
				' 자신의 스프레드를 지우고 다시 그린다.
				arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = ""
				For i = 1 To frm1.vspdData2.MaxRows
					If i <> nSpreadIndex3 Then
					arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) & _
								iColSep & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X") & _
								iColSep & GetSpreadText(frm1.vspdData2, C_SP3_SHET_NUM, i, "X", "X") & _
								iColSep & GetSpreadText(frm1.vspdData2, C_SP3_FILD_SEQ, i, "X", "X") & _
								iColSep & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NAM, i, "X", "X") & _
								iColSep & GetSpreadText(frm1.vspdData2, C_SP3_ATTH_CHR, i, "X", "X") & iRowSep
					End If
				Next

				' 현재 vspdData2 의 자료를 변경한다.
				ggoSpread.Source = frm1.vspdData2
				Call ggoSpread.ClearSpreadData()
				ggoSpread.SSShowData arrJoins(nSpreadIndex1-1, nSpreadIndex2-1)
				Call SetSpreadColor(2, 1, frm1.vspdData2.MaxRows)
				frm1.vspdData2.ReDraw = True

				If frm1.vspdData2.MaxRows > 0 Then
					nSpreadIndex3 = 1
				Else
					nSpreadIndex3 = 0
				End If
			Else
				Exit Function
			End If
		Case 3
		
			If frm1.vspdData3.MaxRows < 1 Or nSpreadIndex4 < 1 Then Exit Function

			If Confirm("현재행을 지우시겠습니까?") Then
				'--------------------------------------------------------------------------------------
				' 현재 행이 사용중인지 검사한다.
				For i = 1 To frm1.vspdData2.MaxRows
					If GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i,"X","X") = nSpreadIndex4 Then
						MsgBox "삭제하려는 레코드가 사용중입니다."
						Exit Function
					End If
				Next

				For i = 0 To 10
					If arrParams(i) <> "" Then
						For j = 0 To 20
							If arrJoins(i, j) <> "" Then
								'----------------------------------------------------------------------
								' 우선 문자열을 행별로 분할하여 배열에 넣는다.
								arrTempRow = Split(arrJoins(i, j), iRowSep)
								For k = 0 To UBound(arrTempRow) -1
									arrTempCol = Split(arrTempRow(k), iColSep)
									If CInt(arrTempCol(1)) = nSpreadIndex4 Then
										MsgBox "삭제하려는 레코드가 사용중입니다."
										Exit Function
									End If
								Next
							End If
						Next
					End If
				Next
				'--------------------------------------------------------------------------------------
				' 현재 행에 대한 링크 정보가 있는지 검사한다.
				For i = 1 To frm1.vspdData3.MaxRows
					If GetSpreadText(frm1.vspdData3, C_SP4_PANT_FLD, i,"X","X") <> "" Then
						If CInt(GetSpreadText(frm1.vspdData3, C_SP4_PANT_FLD, i,"X","X")) = nSpreadIndex4 Then
							MsgBox "삭제하려는 행에 대한 시트연결정보가 존재합니다."
							Exit Function
						End If
					End If
				Next

				'----------------------------------------------------------------------------------
				' 자신의 스프레드를 지우고 다시 그린다.
				szTemp = ""
				For i = 1 To frm1.vspdData3.MaxRows
					If i <> nSpreadIndex4 Then
						szTemp = szTemp & _
							iColSep & Trim(GetSpreadText(frm1.vspdData3, C_SP4_SHET_NUM, i, "X", "X")) & _
							iColSep & Trim(GetSpreadText(frm1.vspdData3, C_SP4_FILD_SEQ, i, "X", "X")) & _
							iColSep & Trim(GetSpreadText(frm1.vspdData3, C_SP4_FILD_NAM, i, "X", "X")) & _
							iColSep & Trim(GetSpreadText(frm1.vspdData3, C_SP4_FILD_TYP, i, "X", "X")) & _
							iColSep & Trim(GetSpreadText(frm1.vspdData3, C_SP4_FILD_FLG, i, "X", "X")) & _
							iColSep & Trim(GetSpreadText(frm1.vspdData3, C_SP4_PANT_FLD, i, "X", "X")) & _
							iColSep & iRowSep
					End If
				Next
				
				' 현재 vspdData3 의 자료를 변경한다.
				ggoSpread.Source = frm1.vspdData3
				Call ggoSpread.ClearSpreadData()
				ggoSpread.SSShowData szTemp
				Call SetSpreadColor(3, 1, frm1.vspdData3.MaxRows)
				frm1.vspdData3.ReDraw = True

				If frm1.vspdData3.MaxRows > 0 Then
					nSpreadIndex4 = 1
				Else
					nSpreadIndex4 = 0
				End If
			Else
				Exit Function
			End If
	End Select
	
	FncDeleteRow = True
End Function

'==================================================================================================
' 현재 작업을 취소한다.
'--------------------------------------------------------------------------------------------------
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
    Call InitData
End Function

'==========================================================================================================
' 메뉴바의 조회 버튼을 눌렀을때 호출되는 메세지 핸들러이다.
' 전달인수:
'----------------------------------------------------------------------------------------------------------
Function FncQuery()
    Dim IntRetCD 
    FncQuery = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
          Exit Function
        End If
    End If
    
    If Not chkField(Document, "1") Then
		Exit Function
    End If
    
    Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.ClearSpreadData()
    Call InitVariables
    
    If DbQuery = False Then
       Exit Function
    End If
       
    FncQuery = True
End Function

'==========================================================================================================
' 메뉴바의 조회 버튼을 눌렀을때 호출되는 메세지 핸들러이다.
' 전달인수:
'----------------------------------------------------------------------------------------------------------
Function DbQuery() 
    Dim strVal    
    Dim IntRetCD

    DbQuery = False

    Call LayerShowHide(1)
    With frm1
        strVal = BIZ_PGM_QUERY_ID & _
                "?txtMode=" & Parent.UID_M0001 & _
                "&txtProcID=" & Trim(.txtProcID.value) & _
                "&txtMaxRows=" & .vspdData.MaxRows & _
                "&lgStrPrevKey=" & lgStrPrevKey
        Call RunMyBizASP(MyBizASP, strVal)
    End With

    DbQuery = True
End Function

'==========================================================================================================
' 조회 작업이 완료 되었을 때 자식 프레임에 의해 호출된다.
' 전달인수:
'----------------------------------------------------------------------------------------------------------
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE

    Call ggoOper.LockField(Document, "Q")
	Call SetSpreadColor(0, 1, frm1.vspdData.MaxRows)
	Call SetSpreadColor(1, 1, frm1.vspdData1.MaxRows)
	Call SetSpreadColor(2, 1, frm1.vspdData2.MaxRows)
	Call SetSpreadColor(3, 1, frm1.vspdData3.MaxRows)

	nSpreadIndex1 = 1
	nSpreadIndex2 = 1
	nSpreadIndex3 = 1
	nSpreadIndex4 = 1
    
    Call SetToolbar("11001110000111")
End Function

'==================================================================================================
' 메뉴바의 저장 버튼을 눌렀을때 호출되는 메세지 핸들러이다.
' 전달인수:
' 참고: 사용자가 입력한 값을 저장한다.
'--------------------------------------------------------------------------------------------------
Function FncSave() 
    Dim IntRetCD
    Dim i
    FncSave = False

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData3
    If Not ggoSpread.SSDefaultCheck Then
        Exit Function
    End If

	If nSpreadIndex1 > 0 Then
		' 현재 vspdData1 의 자료를 저장한다.
		arrParams(nSpreadIndex1-1) = ""
		For i = 1 To frm1.vspdData1.MaxRows
			arrParams(nSpreadIndex1-1) = arrParams(nSpreadIndex1-1) & _
								Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_PARA_NAM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_PARA_TYP, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_REQU_FLG, i, "X", "X") & _
								Chr(11) & Chr(12)
		Next

		If nSpreadIndex2 > 0 Then
			' 현재 vspdData2 의 자료를 저장한다.
			arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = ""
			For i = 1 To frm1.vspdData2.MaxRows
				arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_SHET_NUM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_SEQ, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NAM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_ATTH_CHR, i, "X", "X") & _
								Chr(11) & Chr(12)
			Next
		End If
	End If
	
    If DbSave = False Then
        Exit Function
    End If

    FncSave = True
End Function

'==================================================================================================
' FncSave 함수에 의해서 호출되는 함수로 사용자가 작성한 데이타를 가공하여 비지니스 로직이 있는 
' 프로그램에 전달해 준다.
' 전달인수:
'--------------------------------------------------------------------------------------------------
Function DbSave() 
    Dim i, j, k
	Dim strVal, strVal1, strVal2
    Dim iColSep, iRowSep
    Dim arrTempRow

    DbSave = False

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep

    On Error Resume Next

    Call LayerShowHide(1)                                        '☜: Protect system from crashing

	lgRetFlag = False
	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode

	    strMode = .txtMode.value

	    '-----------------------
		'Data manipulate area
        strVal  = ""
        strVal1 = ""
        strVal2 = ""

		'--------------------------------------------------------------------------------------
		' 이제 arrJoins(10,20) 배열을 모두 뒤져서 현재 필드 아이디 이후의 값이 존재하는지 본다.
		' 만약 존재하면 값들을 모두 변경해 준다.
        For i = 1 To .vspdData.MaxRows
            strVal = strVal & _
					 i & iColSep & _
					 Trim(GetSpreadText(.vspdData, C_SP1_COMP_NAM, i, "X", "X")) & iColSep & _
					 Trim(GetSpreadText(.vspdData, C_SP1_METH_NAM, i, "X", "X")) & iColSep & _
					 Trim(GetSpreadText(.vspdData, C_SP1_TRET_DSC, i, "X", "X")) & iRowSep

			If arrParams(i-1) <> "" Then
				arrTempRow = Split(arrParams(i - 1), iRowSep)
				For j = 0  To UBound(arrTempRow) - 1
					strVal1 = strVal1 & i & iColSep & j+1 & arrTempRow(j) & iRowSep
				Next

				For j = 0 To 20
					If arrJoins(i-1, j) <> "" Then
						'----------------------------------------------------------------------
						' 우선 문자열을 행별로 분할하여 배열에 넣는다.
						arrTempRow = Split(arrJoins(i-1, j), iRowSep)
						For k = 0 To UBound(arrTempRow) - 1
							strVal2 = strVal2 & i & iColSep & j + 1 & iColSep & k+1 & arrTempRow(k) & iRowSep
						Next
					End If
				Next
			End If
        Next

        .txtSpread.value  = strVal
        .txtSpread1.value = strVal1
        .txtSpread2.value = strVal2

        strVal  = ""
        For i = 1 To .vspdData3.MaxRows
            strVal = strVal & _
					 i & iColSep & _
					 Trim(GetSpreadText(.vspdData3, C_SP4_SHET_NUM, i, "X", "X")) & iColSep & _
					 Trim(GetSpreadText(.vspdData3, C_SP4_FILD_SEQ, i, "X", "X")) & iColSep & _
					 Trim(GetSpreadText(.vspdData3, C_SP4_FILD_NAM, i, "X", "X")) & iColSep & _
					 Trim(GetSpreadText(.vspdData3, C_SP4_FILD_TYP, i, "X", "X")) & iColSep & _
					 Trim(GetSpreadText(.vspdData3, C_SP4_FILD_FLG, i, "X", "X")) & iColSep & _
					 Trim(GetSpreadText(.vspdData3, C_SP4_PANT_FLD, i, "X", "X")) & iRowSep
        Next

        .txtSpread3.value = strVal

        Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
    End With

    DbSave = True                                                           '⊙: Processing is NG
End Function

'==================================================================================================
' 저장 작업이 완료한 차일드 페이지에서 호출하는 메소드 이다.
' 전달인수:
' 참    고: 사용자 변경 사항이 있을 경우 정말 떠날것인지를 물어본다.
'--------------------------------------------------------------------------------------------------
Function DbSaveOk()
    Call InitVariables
    frm1.vspdData.MaxRows = 0
    Call MainQuery()
End Function

'==================================================================================================
' 현재 작업 페이지를 떠날때 호출되는 메세지 핸들러이다.
' 전달인수:
' 참    고: 사용자 변경 사항이 있을 경우 정말 떠날것인지를 물어본다.
'--------------------------------------------------------------------------------------------------
Function FncExit()
    Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    ggoSpread.Source = frm1.vspdData1
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    ggoSpread.Source = frm1.vspdData3
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    FncExit = True
End Function

'==================================================================================================
' 스프레드의 필수입력 필드를 지정한다.
' 전달인수:
'  Flag			: 스프레드 구분자 
'  pvStartRow	: 시작행 
'  pvEndRow		: 종료행 
'--------------------------------------------------------------------------------------------------
Sub SetSpreadColor(ByVal Flag, ByVal pvStartRow,ByVal pvEndRow)
	Select Case Flag
		Case 0
			With frm1.vspdData
				.ReDraw = False
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SSSetRequired	 C_SP1_COMP_NAM, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP1_METH_NAM, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP1_TRET_DSC, pvStartRow, pvEndRow
				.ReDraw = True
			End With
		Case 1
			With frm1.vspdData1
				.ReDraw = False
				ggoSpread.Source = frm1.vspdData1
				ggoSpread.SSSetRequired	 C_SP2_PARA_NAM, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP2_PARA_TYP, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP2_REQU_FLG, pvStartRow, pvEndRow
				.ReDraw = True
			End With
		Case 2
			With frm1.vspdData2
				.ReDraw = False
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.SSSetRequired	 C_SP3_FILD_NUM, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP3_SHET_NUM, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP3_FILD_SEQ, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP3_FILD_NAM, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP3_ATTH_CHR, pvStartRow, pvEndRow
		        ggoSpread.SpreadLock     C_SP3_FILD_NUM, -1, C_SP3_FILD_NAM, -1
				.ReDraw = True
			End With
		Case 3
			With frm1.vspdData3
				.ReDraw = False
				ggoSpread.Source = frm1.vspdData3
				ggoSpread.SSSetRequired	 C_SP4_SHET_NUM, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP4_FILD_SEQ, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP4_FILD_NAM, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP4_FILD_TYP, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP4_FILD_FLG, pvStartRow, pvEndRow
				.ReDraw = True
			End With
		End Select
End Sub

'==================================================================================================
' 첫번째 스프레드의 셀이 클릭되었을때의 메세지 핸들러이다.
' 전달인수:
'
'  참고: 저장된 행과 현재 클릭된 행이 일치 하지 않을 경우 
'		 vspdData1 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'		 vspdData2 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'--------------------------------------------------------------------------------------------------
Sub vspdData_MouseDown(Button, Shift, x, y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    '------ Developer Coding part (Start ) --------------------------------------------------------
    nCurrentSpread = 0
    '------ Developer Coding part (End   ) --------------------------------------------------------
End Sub

'==================================================================================================
' 첫번째 스프레드의 셀이 클릭되었을때의 메세지 핸들러이다.
' 전달인수:
'
'  참고: 저장된 행과 현재 클릭된 행이 일치 하지 않을 경우 
'		 vspdData1 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'		 vspdData2 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'--------------------------------------------------------------------------------------------------
Sub vspdData1_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    '------ Developer Coding part (Start ) --------------------------------------------------------
    nCurrentSpread = 1
    '------ Developer Coding part (End   ) --------------------------------------------------------
End Sub

'==================================================================================================
' 첫번째 스프레드의 셀이 클릭되었을때의 메세지 핸들러이다.
' 전달인수:
'  Col		  : 클릭된 열 
'  Row		  : 클릭된 행 
'
'  참고: 저장된 행과 현재 클릭된 행이 일치 하지 않을 경우 
'		 vspdData1 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'		 vspdData2 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'--------------------------------------------------------------------------------------------------
Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    '------ Developer Coding part (Start ) --------------------------------------------------------
    nCurrentSpread = 2
    '------ Developer Coding part (End   ) --------------------------------------------------------
End Sub

'==================================================================================================
' 첫번째 스프레드의 셀이 클릭되었을때의 메세지 핸들러이다.
' 전달인수:
'  Col		  : 클릭된 열 
'  Row		  : 클릭된 행 
'
'  참고: 저장된 행과 현재 클릭된 행이 일치 하지 않을 경우 
'		 vspdData1 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'		 vspdData2 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'--------------------------------------------------------------------------------------------------
Sub vspdData3_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    '------ Developer Coding part (Start ) --------------------------------------------------------
    nCurrentSpread = 3
    '------ Developer Coding part (End   ) --------------------------------------------------------
End Sub

'==================================================================================================
' 첫번째 스프레드의 셀이 클릭되었을때의 메세지 핸들러이다.
' 전달인수:
'  Col		  : 클릭된 열 
'  Row		  : 클릭된 행 
'
'  참고: 저장된 행과 현재 클릭된 행이 일치 하지 않을 경우 
'		 vspdData1 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'		 vspdData2 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'--------------------------------------------------------------------------------------------------
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim i, j
	If nSpreadIndex1 <> Row And nSpreadIndex1 > 0 Then
		' 현재 vspdData1 의 자료를 저장한다.
		arrParams(nSpreadIndex1-1) = ""
		For i = 1 To frm1.vspdData1.MaxRows
            arrParams(nSpreadIndex1-1) = arrParams(nSpreadIndex1-1) & _
							Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_PARA_NAM, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_PARA_TYP, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_REQU_FLG, i, "X", "X") & _
							Chr(11) & Chr(12)
		Next

		If nSpreadIndex2 > 0 Then
			' 현재 vspdData2 의 자료를 저장한다.
			arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = ""
			For i = 1 To frm1.vspdData2.MaxRows
				arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_SHET_NUM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_SEQ, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NAM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_ATTH_CHR, i, "X", "X") & _
								Chr(11) & Chr(12)
			Next
		End If

		If Row > 0 Then
			' 현재 vspdData1 의 자료를 변경한다.
			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.ClearSpreadData()
			ggoSpread.SSShowData arrParams(Row - 1)
			Call SetSpreadColor(1, 1, frm1.vspdData1.MaxRows)
			frm1.vspdData1.ReDraw = True

			' 현재 vspdData2 의 자료를 변경한다.
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ClearSpreadData()
			ggoSpread.SSShowData arrJoins(Row - 1, 0)
			Call SetSpreadColor(2, 1, frm1.vspdData2.MaxRows)
			frm1.vspdData2.ReDraw = True
			nSpreadIndex2 = 1
			nSpreadIndex3 = 1
		End If
	End If

	nSpreadIndex1 = Row
End Sub

'==================================================================================================
' 첫번째 스프레드의 셀이 클릭되었을때의 메세지 핸들러이다.
' 전달인수:
'  Col		  : 클릭된 열 
'  Row		  : 클릭된 행 
'
'  참고: 저장된 행과 현재 클릭된 행이 일치 하지 않을 경우 
'		 vspdData1 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'		 vspdData2 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'--------------------------------------------------------------------------------------------------
Sub vspdData1_Click(ByVal Col, ByVal Row)
	Dim i, j
	
	If nSpreadIndex2 <> Row And nSpreadIndex1 > 0 And nSpreadIndex2 > 0 Then
		' 현재 vspdData1 의 자료를 저장한다.
		arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = ""
		For i = 1 To frm1.vspdData2.MaxRows
            arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) & _
							Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_SHET_NUM, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_SEQ, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NAM, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_ATTH_CHR, i, "X", "X") & _
							Chr(11) & Chr(12)
		Next

		If Row > 0 Then
			' 현재 vspdData1 의 자료를 변경한다.
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ClearSpreadData()
			ggoSpread.SSShowData arrJoins(nSpreadIndex1-1, Row - 1)
			Call SetSpreadColor(2, 1, frm1.vspdData2.MaxRows)
			frm1.vspdData2.ReDraw = True

			nSpreadIndex3 = 1
		End If
	End If

	nSpreadIndex2 = Row
End Sub

'==================================================================================================
' 첫번째 스프레드의 셀이 클릭되었을때의 메세지 핸들러이다.
' 전달인수:
'  Col		  : 클릭된 열 
'  Row		  : 클릭된 행 
'
'  참고: 저장된 행과 현재 클릭된 행이 일치 하지 않을 경우 
'		 vspdData1 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'		 vspdData2 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'--------------------------------------------------------------------------------------------------
Sub vspdData2_Click(ByVal Col, ByVal Row)
	Dim i
	If nSpreadIndex3 <> Row And nSpreadIndex1 > 0 And nSpreadIndex2 > 0 Then
		' 현재 vspdData1 의 자료를 저장한다.
		arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = ""
		For i = 1 To frm1.vspdData2.MaxRows
            arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) & _
							Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_SHET_NUM, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_SEQ, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NAM, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_ATTH_CHR, i, "X", "X") & _
							Chr(11) & Chr(12)
		Next
	End If
	
	nSpreadIndex3 = Row
End Sub

'==================================================================================================
' 첫번째 스프레드의 셀이 클릭되었을때의 메세지 핸들러이다.
' 전달인수:
'  Col		  : 클릭된 열 
'  Row		  : 클릭된 행 
'
'  참고: 저장된 행과 현재 클릭된 행이 일치 하지 않을 경우 
'		 vspdData1 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'		 vspdData2 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'--------------------------------------------------------------------------------------------------
Sub vspdData3_Click(ByVal Col, ByVal Row)
	nSpreadIndex4 = Row
End Sub

'==========================================================================================================
' 네번째 스프레드의 버튼셀이 클릭되었을때의 메세지 핸들러이다.
' 전달인수:
'  Col		  : 클릭된 열 
'  Row		  : 클릭된 행 
'  ButtonDown : 항상 0 임(무시해도 됨)
'
'  참고: vspdData2의 현제 행에 필드정보를 보낸다.
'----------------------------------------------------------------------------------------------------------
Sub vspdData3_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	'If nSpreadIndex3 = 0 Or frm1.vspdData2.MaxRows = 0 Then
	'	MsgBox "@@결합할 행을 선택하십시요"
	'	Exit Sub
	'End If
		
     dim jRow , sATTH_CHR
     
   
     sATTH_CHR= GetSpreadText(frm1.vspdData1, C_SP2_PARA_TYP, frm1.vspdData1.ActiveRow, "X", "X")
	 ggoSpread.source= frm1.vspdData2
	 frm1.vspdData2.Row =  frm1.vspdData2.MaxRows 
	 ggoSpread.InsertRow ,1
	 jRow =  frm1.vspdData2.ActiveRow
	 
	Call SetSpreadColor(2, jRow, jRow)
	
'
 
	Call SetSpreadValue(frm1.vspdData2, 1, frm1.vspdData2.ActiveRow, _
						nSpreadIndex4, "", "")
	Call SetSpreadValue(frm1.vspdData2, 2, frm1.vspdData2.ActiveRow, _
						GetSpreadText(frm1.vspdData3, C_SP4_SHET_NUM, nSpreadIndex4, "X", "X"), "", "")
	Call SetSpreadValue(frm1.vspdData2, 3, frm1.vspdData2.ActiveRow, _
						GetSpreadText(frm1.vspdData3, C_SP4_FILD_SEQ, nSpreadIndex4, "X", "X"), "", "")
	Call SetSpreadValue(frm1.vspdData2, 4, frm1.vspdData2.ActiveRow, _
						GetSpreadText(frm1.vspdData3, C_SP4_FILD_NAM, nSpreadIndex4, "X", "X"), "", "")
				
    if uCase(sATTH_CHR) ="VAR" then
		Call SetSpreadValue(frm1.vspdData2, C_SP3_ATTH_CHR, frm1.vspdData2.ActiveRow, _
						"N", "", "")
						
    else

		Call SetSpreadValue(frm1.vspdData2, C_SP3_ATTH_CHR, frm1.vspdData2.ActiveRow, _
						"C", "", "")
    end if						
End Sub

'==========================================================================================================
' 네번째 스프레드의 셀이 변경되었을때 메세지 핸들러이다.
' 전달인수:
'  Col		  : 클릭된 열 
'  Row		  : 클릭된 행 
'
'  참고: 최초 생성시 자동으로 Field_ID 를 생성시켜야 한다.
'        현재 레코드에 존재하는 Field_ID의 최대값에 1 을 더한값을 새로운 Field_Id로 한다.
'----------------------------------------------------------------------------------------------------------
Sub vspdData3_Change(ByVal Col, ByVal Row)
	Dim i, j, k
	Exit Sub
	'----------------------------------------------------------------------------------------------
	' 시트가 추가되었으므로 현재 추가되는 행 이후의 필드 아이디가 1 증가하므로 
	' 이후의 아이디를 가진 모든 배열값을 뒤져서 수정해 주어야 한다.
	' 완전히 생 노가다이다.

	If nSpreadIndex1 > 0 And nSpreadIndex2 > 0 Then
		'------------------------------------------------------------------------------------------
		' 세번째 스프레드의 값을 배열로 읽어 들인다.
		arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = ""
		For i = 1 To frm1.vspdData2.MaxRows
			arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) = arrJoins(nSpreadIndex1-1, nSpreadIndex2-1) & _
						i & _
						Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X") & _
						Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_SHET_NUM, i, "X", "X") & _
						Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_SEQ, i, "X", "X") & _
						Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_FILD_NAM, i, "X", "X") & _
						Chr(11) & GetSpreadText(frm1.vspdData2, C_SP3_ATTH_CHR, i, "X", "X") & _
						Chr(11) & Chr(12)
			'--------------------------------------------------------------------------------------
			' 만약 해당 필드 아이디가 현재 nSpreadIndex4 보다 크면...
			If CInt(GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X")) > nSpreadIndex4 Then
				Call SetSpreadValue(frm1.vspdData2, 1, i, _
									CInt(GetSpreadText(frm1.vspdData2, _
													   C_SP3_FILD_NUM, i, "X", "X")) + 1, "", "")
			End If
		Next
	End If

	'--------------------------------------------------------------------------------------
	' 이제 arrJoins(10,20) 배열을 모두 뒤져서 현재 필드 아이디 이후의 값이 존재하는지 본다.
	' 만약 존재하면 값들을 모두 변경해 준다.
	Dim szTempRow
	For i = 0 To 10
		If arrParams(i) <> "" Then
			For j = 0 To 20
				If arrJoins(i, j) <> "" Then
					'----------------------------------------------------------------------
					' 우선 문자열을 행별로 분할하여 배열에 넣는다.
'							szTempRow = arrJoins(i, j)
					arrTempRow = Split(arrJoins(i, j), iRowSep)
					szTempRow = ""
					For k = 0 To UBound(arrTempRow) -1
						arrTempCol = Split(arrTempRow(k), iColSep)
							MsgBox arrTempCol(1)
						If CInt(arrTempCol(1)) > nSpreadIndex4 Then
							szTempRow = szTempRow & _
										arrTempCol(1) + 1 & iColSep & _
										GetSpreadText(frm1.vspdData3, C_SP4_SHET_NUM, arrTempCol(1), "X", "X") & iColSep & _
										GetSpreadText(frm1.vspdData3, C_SP4_FILD_SEQ, arrTempCol(1), "X", "X") & iColSep & _
										GetSpreadText(frm1.vspdData3, C_SP4_FILD_NAM, arrTempCol(1), "X", "X") & iColSep & _
										arrTempCol(5) & iColSep & iRowSep
						Else
							szTempRow = szTempRow & arrTempRow(k) & iRowSep
						End If
					Next
					arrJoins(i, j) = szTempRow
				End If
			Next
		End If
	Next
End Sub


'==================================================================================================
' 라디오 버튼 1 이벤트 핸들러 
'--------------------------------------------------------------------------------------------------
Function Radio1_OnClick()
	frm1.hUseFlag.value = frm1.Radio1.value
End Function

'==================================================================================================
' 라디오 버튼 2 이벤트 핸들러 
'--------------------------------------------------------------------------------------------------
Function Radio2_OnClick()
	frm1.hUseFlag.value = frm1.Radio2.value
End Function

'==================================================================================================
' 라디오 버튼 3 이벤트 핸들러 
'--------------------------------------------------------------------------------------------------
Function Radio3_OnClick()
	frm1.hJoinMethod.value = frm1.Radio3.value
End Function

'==================================================================================================
' 라디오 버튼 4 이벤트 핸들러 
'--------------------------------------------------------------------------------------------------
Function Radio4_OnClick()
	frm1.hJoinMethod.value = frm1.Radio4.value
End Function

'==================================================================================================
' 라디오 버튼 6 이벤트 핸들러 
'--------------------------------------------------------------------------------------------------
Function Radio6_OnClick()
	frm1.hTranFlag.value = frm1.Radio6.value
End Function

'==================================================================================================
' 라디오 버튼 7 이벤트 핸들러 
'--------------------------------------------------------------------------------------------------
Function Radio7_OnClick()
	frm1.hTranFlag.value = frm1.Radio7.value
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>업무등록</font></td>
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
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>업무</TD>
									<TD CLASS="TD656" NOWRAP>
									    <INPUT NAME="txtProcID" MAXLENGTH="20" SIZE=20 ALT ="업무코드" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcID" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(txtProcID.value, 0)">
										<INPUT NAME="txtProcNm" MAXLENGTH="80" SIZE=40 ALT ="업 무 명" tag="14X"  STYLE="TEXT-ALIGN:left"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=60 WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>업무명</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtNProcNm" ALT="업무명" MAXLENGTH="80" SIZE=40 STYLE="TEXT-ALIGN: left" tag ="22N">
								</TD>
                                <TD CLASS="TD5">사용여부</TD>
                                <TD CLASS="TD6">
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoUseFlag" ID="Radio1" VALUE="Y" TAG = "11X" CHECKED>
										<LABEL FOR="Radio1">사용</LABEL>&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoUseFlag" ID="Radio2" VALUE="N" TAG = "11X">
										<LABEL FOR="Radio2">미사용</LABEL>&nbsp;&nbsp;
                                </TD>
							</TR>
							<TR>
                                <TD CLASS="TD5">연결방법</TD>
                                <TD CLASS="TD6">
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoJoinFlag" ID="Radio3" VALUE="N" TAG = "11X" CHECKED>
										<LABEL FOR="Radio3">단독</LABEL>&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoJoinFlag" ID="Radio4" VALUE="S" TAG = "11X">
										<LABEL FOR="Radio4">시트</LABEL>&nbsp;&nbsp;
                                </TD>
                                <TD CLASS="TD5">트랜젝션</TD>
                                <TD CLASS="TD6">
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoTranFlag" ID="Radio6" VALUE="Y" TAG = "11X" CHECKED>
										<LABEL FOR="Radio6">사용</LABEL>&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoTranFlag" ID="Radio7" VALUE="N" TAG = "11X">
										<LABEL FOR="Radio7">미사용</LABEL>&nbsp;&nbsp;
                                </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>시 작 행</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtStartRow" ALT="시 작 행" MAXLENGTH="2" SIZE=10 STYLE="TEXT-ALIGN: left" tag ="22N">
								</TD>
								<TD CLASS=TD5 NOWRAP>실행시간</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtRunTime" ALT="실행시간" MAXLENGTH="5" SIZE=10 STYLE="TEXT-ALIGN: left" tag ="22N">
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD WIDTH=100% HEIGHT=10%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>

								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD WIDTH=50% HEIGHT=40%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData1 NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>

								</TD>
								<TD WIDTH=50% HEIGHT=40%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData2 NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>

								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD WIDTH=100% HEIGHT=30%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData3 NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>

								</TD>
							</TR>
						</Table>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hProcId" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hUseFlag" VALUE="Y" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hJoinMethod" VALUE="N" TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hTranFlag" VALUE="Y" TAG="24" TABINDEX="-1">
<TEXTAREA CLASS=hidden NAME=txtSpread tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA CLASS=hidden NAME=txtSpread1 tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA CLASS=hidden NAME=txtSpread2 tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA CLASS=hidden NAME=txtSpread3 tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" TAG="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

