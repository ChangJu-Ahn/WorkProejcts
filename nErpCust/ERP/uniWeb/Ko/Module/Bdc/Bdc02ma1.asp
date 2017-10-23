<%@ LANGUAGE="VBSCRIPT" %>
<!--'==============================================================================================
'*  1. Module Name          : BDC
'*  2. Function Name        : 
'*  3. Program ID           : BDC02MA1
'*  4. Program Name         : BDC 검증로직 등록 
'*  5. Program Desc         : BDC SQL과 Mapping 정보 입력 
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
Const BIZ_PGM_QUERY_ID = "BDC02MB1.ASP"								'☆: 조회 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID  = "BDC02MB2.ASP"								'☆: 저장 비지니스 로직 ASP명 

Const C_SP1_DESC = 1
Const C_SP1_RSLT = 2

Const C_SP2_NUM = 1
Const C_SP2_NAM = 2
Const C_SP2_TYP = 3
Const C_SP2_LNG = 4
Const C_SP2_FID = 5
Const C_SP2_FNM = 6

Const C_SP3_NUM = 1
Const C_SP3_SEQ = 2
Const C_SP3_NAM = 3
Const C_SP3_SND = 4

'==================================================================================================
' 동적으로 다르게 보여야할 각 스프레드의 값을 보관할 배열을 지정한다.
Dim arrQuery(10, 1)

' 스프레드 시트의 통제를 위한 변수 
Dim nCurrentSpread
Dim nSpreadIndex1
Dim nSpreadIndex2

Dim strMode
Dim IsOpenPop
Dim lgRetFlag


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
	With frm1.vspdData1
        .ReDraw = False
		.MaxCols = C_SP1_RSLT
        .MaxRows = 0

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit			'"V20050321",,parent.gAllowDragDropSpread
		ggoSpread.SSSetEdit  C_SP1_DESC,  "질의문명", 50, , , 40
		ggoSpread.SSSetEdit  C_SP1_RSLT,  "결과값", 30, , , 40
		.ReDraw = True
	End With
	
	With frm1.vspdData2
        .ReDraw = False
		.MaxCols = C_SP2_FNM
        .MaxRows = 0

		ggoSpread.Source =	frm1.vspdData2
		ggoSpread.Spreadinit			'"V20021121",,parent.gAllowDragDropSpread
		ggoSpread.SSSetEdit  C_SP2_NUM,  "번호", 3, , , 20
		ggoSpread.SSSetEdit  C_SP2_NAM,  "이름", 12, , , 20
		ggoSpread.SSSetCombo C_SP2_TYP,  "유형", 8, 2, False
		ggoSpread.SSSetEdit  C_SP2_LNG,  "길이", 8, , , 2
		ggoSpread.SSSetEdit  C_SP2_FID,  "필드ID", 8, , , 2
		ggoSpread.SSSetEdit  C_SP2_FNM,  "필드명", 12, , ,40
		Call ggoSpread.SSSetColHidden(C_SP2_NUM, C_SP2_NUM, True)
		.ReDraw = True
	End With

	ggoSpread.Source = frm1.vspdData3
	ggoSpread.Spreadinit			'"V20021121",,parent.gAllowDragDropSpread
	With frm1.vspdData3
        .ReDraw = False
		.MaxCols = C_SP3_SND
        .MaxRows = 0

		ggoSpread.SSSetEdit   C_SP3_NUM,  "시트", 6, , , 1
		ggoSpread.SSSetEdit   C_SP3_SEQ,  "필드", 6, , , 2
		ggoSpread.SSSetEdit   C_SP3_NAM,  "필드명", 15, , , 40
		ggoSpread.SSSetButton C_SP3_SND
		.ReDraw = True
	End With
End Sub

'==================================================================================================
' 광역 변수들을 초기화 시킨다.
'--------------------------------------------------------------------------------------------------
Sub InitVariables()
	Dim i

    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1

	nCurrentSpread = 1
	nSpreadIndex1 = 0
	nSpreadIndex2 = 0
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
	ggoSpread.Source = frm1.vspdData2
    ggoSpread.SetCombo "VarChar" & vbTab & "NVarChar" & vbTab & "Integer", C_SP2_TYP
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

	arrParam(0) = "업무팝업"			' 팝업 명칭 
	arrParam(1) = "B_BDC_MASTER"			' TABLE 명칭 
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = " USE_FLAG= " & Filtervar("Y", "''", "S")						' Code Condition
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
    Dim i
    Dim iColSep, iRowSep

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep

    FncInsertRow = False
	Select Case (nCurrentSpread)
		Case 1
			'--------------------------------------------------------------------------------------
			' 두번째 스프레드와 세번째 스프레드의 값이 합당한지 검사 해 본다.
			ggoSpread.Source = frm1.vspdData1
			If Not ggoSpread.SSDefaultCheck Then
				Exit Function
			End If

			ggoSpread.Source = frm1.vspdData2
			If Not ggoSpread.SSDefaultCheck Then
				Exit Function
			End If

			If nSpreadIndex1 > 0 Then
			    arrQuery(nSpreadIndex1 - 1, 0) = frm1.txtSQLStatements.value
			    
				' 현재 vspdData2 의 자료를 저장한다.
				arrQuery(nSpreadIndex1 - 1, 1) = ""
				For i = 1 To frm1.vspdData2.MaxRows
					arrQuery(nSpreadIndex1 - 1, 1) = arrQuery(nSpreadIndex1 - 1, 1) & i & iColSep & _
								GetSpreadText(frm1.vspdData2, C_SP2_NUM, i, "X", "X") & iColSep & _
								GetSpreadText(frm1.vspdData2, C_SP2_NAM, i, "X", "X") & iColSep & _
								GetSpreadText(frm1.vspdData2, C_SP2_TYP, i, "X", "X") & iColSep & _
								GetSpreadText(frm1.vspdData2, C_SP2_LNG, i, "X", "X") & iColSep & _
								GetSpreadText(frm1.vspdData2, C_SP2_FID, i, "X", "X") & iColSep & _
								GetSpreadText(frm1.vspdData2, C_SP2_FNM, i, "X", "X") & iColSep & iRowSep
				Next

				For i = 9 To nSpreadIndex1 Step -1
					arrQuery(i + 1, 0) = arrQuery(i, 0)
					arrQuery(i + 1, 1) = arrQuery(i, 1)
					arrQuery(i, 0) = ""
					arrQuery(i, 1) = ""
				Next
			End If
			'--------------------------------------------------------------------------------------
			' 현재 행위치에 새로운 행을 생성시키고 현재행 이후를 하나씩 뒤로 밀어낸다.
			' 새로운 행이 컴포넌트 정보가 추가 되므로 
			' 현재 컴포넌트가 가리키고 있는 것 이후에 것의 컴포넌트 ID가 하나씩 밀려야 한다.
			' 또한 해당 컴포넌트가 가리키고 있는 파라메터 배열도 한칸씩 밀려야 한다.
			' 파라메터가 가리키고 있는 결합정보 배열도 한칸씩 밀려야 한다.
			
			frm1.txtSQLStatements.value = ""
			'--------------------------------------------------------------------------------------
			' 두번째 스프레드를 초기화 시켜준다.
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ClearSpreadData()
			
			'--------------------------------------------------------------------------------------
			' 첫번째 스프레드에 새로운 행을 추가 시킨다.
			With frm1
				.vspdData1.focus
				.vspdData1.ReDraw = False
				ggoSpread.Source = .vspdData1
				ggoSpread.InsertRow nSpreadIndex1, 1
				Call SetSpreadColor(1, .vspdData1.ActiveRow, .vspdData1.ActiveRow)
				nSpreadIndex1 = .vspdData1.ActiveRow
				nSpreadIndex2 = 0
				.vspdData1.ReDraw = True
			End With
		Case 2
			With frm1
				.vspdData2.focus
				.vspdData2.ReDraw = False
				ggoSpread.Source = .vspdData2
				ggoSpread.InsertRow nSpreadIndex2, 1
				Call SetSpreadColor(2, .vspdData2.ActiveRow, .vspdData2.ActiveRow)
				nSpreadIndex2 = .vspdData2.ActiveRow
	            Call SetSpreadValue(.vspdData2, C_SP2_NUM, .vspdData2.ActiveRow, _
						            nSpreadIndex1, "", "")
				.vspdData2.ReDraw = True
			End With
	End Select

	If Err.number = 0 Then
		FncInsertRow = True                                           '☜: Processing is OK
	End If

    Set gActiveElement = document.ActiveElement
End Function

'==================================================================================================
' 현재 선택한 행을 삭제한다.
'--------------------------------------------------------------------------------------------------
Function FncDeleteRow()
	Dim szTemp
	Dim i, j, k
    Dim iColSep, iRowSep

	FncDeleteRow = false
    iColSep = parent.gColSep
    iRowSep = parent.gRowSep

	Select Case (nCurrentSpread)
		Case 1
			If frm1.vspdData1.MaxRows < 1 Or nSpreadIndex1 < 1 Then Exit Function
			
			'--------------------------------------------------------------------------------------
			' 현재선택된 행의 하위 정보들을 모두 지운후 현재행을 제외하고 첫번째 스프레드를 다시 
			' 그린후 2 번째 및 3번째 스프레드도 다시 그린다.
			If Confirm("쿼리문과 파라메터 정보도 함께 지우시겠습니까?") Then
				'----------------------------------------------------------------------------------
				' 자 우선 자식들 정보부터 지우자 
				For i = nSpreadIndex1 - 1 To frm1.vspdData1.MaxRows - 1
					arrQuery(i, 0) = arrQuery(i+1, 0)
					arrQuery(i, 1) = arrParam(i+1, 1)
				Next

				arrQuery(10, 0) = ""
				arrQuery(10, 1) = ""
				
				'----------------------------------------------------------------------------------
				' 세개의 스프레드를 지우고 다시 그린다.
				szTemp = ""
				For i = 1 To frm1.vspdData1.MaxRows
					If i <> nSpreadIndex1 Then
					szTemp = szTemp & _
							Trim(GetSpreadText(frm1.vspdData1, C_SP1_DESC, i, "X", "X")) & iColSep & _
							Trim(GetSpreadText(frm1.vspdData1, C_SP1_RSLT, i, "X", "X")) & iRowSep
					End If
				Next

				ggoSpread.Source = frm1.vspdData1
				Call ggoSpread.ClearSpreadData()
				ggoSpread.SSShowData szTemp
				Call SetSpreadColor(1, 1, frm1.vspdData1.MaxRows)
				frm1.vspdData1.ReDraw = True

				ggoSpread.Source = frm1.vspdData2
				Call ggoSpread.ClearSpreadData()
				ggoSpread.SSShowData arrQuery(nSpreadIndex1, 1)
				Call SetSpreadColor(2, 1, frm1.vspdData2.MaxRows)
				frm1.vspdData2.ReDraw = True

				If frm1.vspdData.MaxRows > 0 Then 
					nSpreadIndex1 = 1
					If frm1.vspdData2.MaxRows > 0 Then
						nSpreadIndex2 = 1
					Else
						nSpreadIndex2 = 0
					End If
				Else
					nSpreadIndex1 = 0
					nSpreadIndex2 = 0
				End If
			Else
				Exit Function
			End If
		Case 2
			If frm1.vspdData2.MaxRows < 1 Or nSpreadIndex2 < 1 Then Exit Function
			
			'--------------------------------------------------------------------------------------
			' 현재선택된 행의 하위 정보들을 모두 지운후 현재행을 제외하고 첫번째 스프레드를 다시 
			' 그린후 2 번째 및 3번째 스프레드도 다시 그린다.
			If Confirm("파라메터 구성 정보를 지우시겠습니까?") Then

				'----------------------------------------------------------------------------------
				' 두개의 스프레드를 지우고 다시 그린다.
				arrQuery(nSpreadIndex1-1, 1) = ""
				j = 0
				For i = 1 To frm1.vspdData2.MaxRows
					If i <> nSpreadIndex2 Then
					arrQuery(nSpreadIndex1-1, 1) = arrQuery(nSpreadIndex1-1, 1) & j & iColSep& _
							 GetSpreadText(frm1.vspdData2,C_SP2_NUM,i,"X","X") & iColSep & _
							 GetSpreadText(frm1.vspdData2,C_SP2_NAM,i,"X","X") & iColSep & _
							 GetSpreadText(frm1.vspdData2,C_SP2_TYP,i,"X","X") & iColSep & _
							 GetSpreadText(frm1.vspdData2,C_SP2_LNG,i,"X","X") & iColSep & _
							 GetSpreadText(frm1.vspdData2,C_SP2_FID,i,"X","X") & iColSep & _
							 GetSpreadText(frm1.vspdData2,C_SP2_FNM,i,"X","X") & iRowSep
					j = j + 1
					End If
				Next

				ggoSpread.Source = frm1.vspdData2
				Call ggoSpread.ClearSpreadData()
				ggoSpread.SSShowData arrQuery(nSpreadIndex1-1, 1)
				Call SetSpreadColor(2, 1, frm1.vspdData2.MaxRows)
				frm1.vspdData2.ReDraw = True
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
    ggoSpread.Source = frm1.vspdData1
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

    ggoSpread.Source = frm1.vspdData1
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
          Exit Function
        End If
    End If
	
	If Not chkField(Document, "1") Then
		Exit Function
    End If

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
    
    strVal = BIZ_PGM_QUERY_ID & _
            "?txtMode=" & Parent.UID_M0001 & _
            "&txtProcID=" & Trim(frm1.txtProcID.value)
    
    Call RunMyBizASP(MyBizASP, strVal)


    DbQuery = True
End Function

'==========================================================================================================
' 조회 작업이 완료 되었을 때 자식 프레임에 의해 호출된다.
' 전달인수:
'----------------------------------------------------------------------------------------------------------
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE

    Call ggoOper.LockField(Document, "Q")
	Call SetSpreadColor(1, 1, frm1.vspdData1.MaxRows)
	Call SetSpreadColor(2, 1, frm1.vspdData2.MaxRows)
	Call SetSpreadColor(3, 1, frm1.vspdData3.MaxRows)

	nSpreadIndex1 = 1
	nSpreadIndex2 = 1
    
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
    Dim iColSep, iRowSep

    FncSave = False

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep

    ggoSpread.Source = frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then
        Exit Function
    End If

	If nSpreadIndex1 > 0 Then
		arrQuery(nSpreadIndex1 - 1, 0) = frm1.txtSQLStatements.value
		
		' 현재 vspdData2 의 자료를 저장한다.
		arrQuery(nSpreadIndex1 - 1, 1) = ""
		For i = 1 To frm1.vspdData2.MaxRows
			arrQuery(nSpreadIndex1 - 1, 1) = arrQuery(nSpreadIndex1 - 1, 1) & i & iColSep & _
						GetSpreadText(frm1.vspdData2, C_SP2_NUM, i, "X", "X") & iColSep & _
						GetSpreadText(frm1.vspdData2, C_SP2_NAM, i, "X", "X") & iColSep & _
						GetSpreadText(frm1.vspdData2, C_SP2_TYP, i, "X", "X") & iColSep & _
						GetSpreadText(frm1.vspdData2, C_SP2_LNG, i, "X", "X") & iColSep & _
						GetSpreadText(frm1.vspdData2, C_SP2_FID, i, "X", "X") & iColSep & _
						GetSpreadText(frm1.vspdData2, C_SP2_FNM, i, "X", "X") & iColSep & iRowSep
		Next
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
    Dim i
	Dim strVal1, strVal2
    Dim iColSep, iRowSep
    Dim arrTempRow

    DbSave = False

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep

'    On Error Resume Next

    Call LayerShowHide(1)                                        '☜: Protect system from crashing

	lgRetFlag = False
	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode

	    strMode = .txtMode.value

	    '-----------------------
		'Data manipulate area
        strVal1 = ""
        strVal2 = ""

        For i = 1 To .vspdData1.MaxRows
            strVal1 = strVal1 & _
					 i & iColSep & arrQuery(i - 1, 0) & iColSep &_
					 Trim(GetSpreadText(.vspdData1, C_SP1_DESC, i, "X", "X")) & iColSep & _
					 Trim(GetSpreadText(.vspdData1, C_SP1_RSLT, i, "X", "X")) & iRowSep

            strVal2 = strVal2 & arrQuery(i - 1, 1)
        Next

        .txtSpread1.value = strVal1
        .txtSpread2.value = strVal2
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
		Case 1
			With frm1.vspdData1
				.ReDraw = False
				ggoSpread.Source = frm1.vspdData1
				ggoSpread.SSSetRequired	 C_SP1_DESC, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP1_RSLT, pvStartRow, pvEndRow
				.ReDraw = True
			End With
		Case 2
			With frm1.vspdData2
				.ReDraw = False
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.SSSetRequired	 C_SP2_NAM, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP2_TYP, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP2_LNG, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP2_FID, pvStartRow, pvEndRow
				ggoSpread.SSSetRequired	 C_SP2_FNM, pvStartRow, pvEndRow
				ggoSpread.SSSetProtected C_SP2_FID, pvStartRow, pvEndRow
				ggoSpread.SSSetProtected C_SP2_FNM, pvStartRow, pvEndRow
				.ReDraw = True
			End With
		Case 3
			With frm1.vspdData3
				.ReDraw = False
				ggoSpread.Source = frm1.vspdData3
				ggoSpread.SSSetProtected C_SP3_NUM, pvStartRow, pvEndRow
				ggoSpread.SSSetProtected C_SP3_SEQ, pvStartRow, pvEndRow
				ggoSpread.SSSetProtected C_SP3_NAM, pvStartRow, pvEndRow
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
Sub vspdData1_MouseDown(Button, Shift, x, y)
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
'
'  참고: 저장된 행과 현재 클릭된 행이 일치 하지 않을 경우 
'		 vspdData1 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'		 vspdData2 의 값을 저장하고 현재 클릭된 행에 일치하는 데이터로 치환해 준다.
'--------------------------------------------------------------------------------------------------
Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
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
    If Button = 2 And gMouseClickStatus = "SP3C" Then
       gMouseClickStatus = "SP3CR"
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
Sub vspdData1_Click(ByVal Col, ByVal Row)
	'Exit Sub
	Dim i, j
	
	If nSpreadIndex1 <> Row And nSpreadIndex1 > 0 Then
		' 현재 vspdData1 의 자료를 저장한다.		

		If nSpreadIndex2 > 0 Then
			' 현재 vspdData2 의 자료를 저장한다.
			arrQuery(nSpreadIndex1-1, 1) = ""
			For i = 1 To frm1.vspdData2.MaxRows
				arrQuery(nSpreadIndex1-1, 1) = arrQuery(nSpreadIndex1-1, 1) & i & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP2_NUM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP2_NAM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP2_TYP, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP2_LNG, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP2_FID, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData2, C_SP2_FNM, i, "X", "X") & _
								Chr(11) & Chr(12)
			Next
		End If

		If Row > 0 Then
			' 현재 vspdData2 의 자료를 변경한다.
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ClearSpreadData()

			If Not ISEmpty(arrQuery) Then
				frm1.vspdData2.ReDraw = False
				ggoSpread.SSShowData arrQuery(Row - 1, 1)
				Call SetSpreadColor(2, 1, frm1.vspdData2.MaxRows)
				frm1.vspdData2.ReDraw = True
				nSpreadIndex2 = 1
			End If	
			
		End If
		
		'------ Developer Coding part (Start)
		frm1.txtSQLStatements.value = arrQuery(Row - 1, 0)
 		'------ Developer Coding part (End)
	End If

	nSpreadIndex1 = Row
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
	If nSpreadIndex2 = 0 Or frm1.vspdData2.MaxRows = 0 Then
		MsgBox "결합할 행을 선택하십시요"
		Exit Sub
	End If

	Call SetSpreadValue(frm1.vspdData2, C_SP2_FID, frm1.vspdData2.ActiveRow, _
						Row, "", "")
	Call SetSpreadValue(frm1.vspdData2, C_SP2_FNM, frm1.vspdData2.ActiveRow, _
						GetSpreadText(frm1.vspdData3, C_SP3_NAM, Row, "X", "X"), "", "")
End Sub

Function txtSQLStatements_Onchange()
    If frm1.vspdData1.ActiveRow < 1 Then
		Exit Function
    End if

    arrQuery(frm1.vspdData1.ActiveRow-1, 0) = frm1.txtSQLStatements.value
	lgBlnFlgChgValue = True
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>검증로직등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>업무코드</TD>
									<TD CLASS="TD656" NOWRAP>
									    <INPUT NAME="txtProcID" MAXLENGTH="20" SIZE=20 ALT ="업무코드" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcID" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(txtProcID.value, 0)">
										<INPUT NAME="txtProcNm" MAXLENGTH="80" SIZE=40 ALT ="업 무 명" tag="14X"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top HEIGHT=30%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD WIDTH=100% HEIGHT=30%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData1 NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>

								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top HEIGHT=30%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>검증SQL</TD>
									<TD CLASS="TD656" NOWRAP><TEXTAREA cols=100 name=txtSQLStatements tag="2" rows=12 MAXLENGTH=3600></TEXTAREA></TD>
								</TR>
							</TABLE>
						</FIELDSET>	
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top HEIGHT=30%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD WIDTH=60% HEIGHT=100%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData2 NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>

								</TD>
								<TD WIDTH=40% HEIGHT=30%>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hProcId" TAG="24" TABINDEX="-1">
<TEXTAREA CLASS=hidden NAME=txtSpread1 tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA CLASS=hidden NAME=txtSpread2 tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" TAG="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" TAG="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
