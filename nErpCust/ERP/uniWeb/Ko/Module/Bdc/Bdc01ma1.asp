<%@ LANGUAGE="VBSCRIPT" %>
<!--'==============================================================================================
'*  1. Module Name          : BDC
'*  2. Function Name        : 
'*  3. Program ID           : BDC04MA1
'*  4. Program Name         : BDC ������� 
'*  5. Program Desc         : BDC ������ ����Ÿ�� ������Ʈ ���� �Է� 
'*  6. Component List       : BDC001
'*  7. Modified date(First) : 2005/01/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kweon, SoonTae
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
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
'��: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->
'==================================================================================================
' ��� �� ���� ���� 
'--------------------------------------------------------------------------------------------------
Const BIZ_PGM_QUERY_ID = "BDC01MB1.ASP"								'��: ��ȸ �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID  = "BDC01MB2.ASP"								'��: ���� �����Ͻ� ���� ASP�� 

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
' �������� ��Ʈ�� ������ ���� ���� 
Dim nCurrentSpread
Dim nSpreadIndex1
Dim nSpreadIndex2
Dim nSpreadIndex3
Dim nSpreadIndex4

Dim strMode
Dim IsOpenPop
Dim lgRetFlag

' �������� �ٸ��� �������� �� ���������� ���� ������ �迭�� �����Ѵ�.
Dim arrParams(10)
Dim arrJoins(10, 20)

'==================================================================================================
' ������ �ε尡 �Ϸ�Ǹ� �ڵ����� ȣ��Ǵ� �Լ�.
' �ʱ�ȭ ��ƾ�� �̰��� ���߽��� �־�� ��.
' ../../inc/incCliMAMain.vbs ���Ͽ� �� �Լ��� ȣ�� �ϵ��� �ϴ� ����� �ֽ� 
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
' �ý��ۿ� ������ ȭ�����, ����ڵ�, ������ �������� �ʱ�ȭ �ϴ� �Լ�.
' ../../inc/incCliVariables.vbs �� ../../ComAsp/LoadInfTB19029.asp  ���Ͽ� �������̴�.
'--------------------------------------------------------------------------------------------------
Sub LoadInfTB19029()
<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
End Sub

'==================================================================================================
' �������� �ʱ�ȭ �Լ� 
' ���α׷��� ���� ����ڵ��� ������ �־�� �ϴ� �κ� 
'--------------------------------------------------------------------------------------------------
Sub InitSpreadSheet()
	With frm1.vspdData
        .ReDraw = False
		.MaxCols = C_SP1_TRET_DSC
        .MaxRows = 0

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit			'"V20021121",,parent.gAllowDragDropSpread
		ggoSpread.SSSetEdit   C_SP1_COMP_NAM,  "������Ʈ", 25, , , 80
		ggoSpread.SSSetEdit   C_SP1_METH_NAM,  "�� �� ��", 25, , , 80
		ggoSpread.SSSetEdit   C_SP1_TRET_DSC,  "��    ��", 44, , , 128
		.ReDraw = True
	End With
	
	With frm1.vspdData1
        .ReDraw = False
		.MaxCols = C_SP2_REQU_FLG
        .MaxRows = 0

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit			'"V20021121",,parent.gAllowDragDropSpread
		ggoSpread.SSSetEdit   C_SP2_PARA_NAM,  "���ڸ�", 20, , , 80
		ggoSpread.SSSetCombo  C_SP2_PARA_TYP,  "����", 8, 2, False
		ggoSpread.SSSetCombo  C_SP2_REQU_FLG,  "�ʼ�", 8, 2, False
		.ReDraw = True
	End With

	With frm1.vspdData2
        .ReDraw = False
		.MaxCols = C_SP3_ATTH_CHR
        .MaxRows = 0

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit			'"V20021121",,parent.gAllowDragDropSpread
		ggoSpread.SSSetEdit   C_SP3_FILD_NUM,  "", 2, , , 2
		ggoSpread.SSSetEdit   C_SP3_SHET_NUM,  "��Ʈ", 8, , , 1
		ggoSpread.SSSetEdit   C_SP3_FILD_SEQ,  "�ʵ�", 8, , , 2
		ggoSpread.SSSetEdit   C_SP3_FILD_NAM,  "�� �� ��", 15, , , 40
		ggoSpread.SSSetCombo  C_SP3_ATTH_CHR,  "÷��", 8, 2, False
		Call ggoSpread.SSSetColHidden(C_SP3_FILD_NUM, C_SP3_FILD_NUM, True)

		.ReDraw = True
	End With

	ggoSpread.Source = frm1.vspdData3
	ggoSpread.Spreadinit				'"V20021121",,parent.gAllowDragDropSpread
	With frm1.vspdData3
        .ReDraw = False
		.MaxCols = C_SP4_FILD_SND
        .MaxRows = 0

		ggoSpread.SSSetEdit   C_SP4_SHET_NUM,  "��Ʈ", 8, , , 1
		ggoSpread.SSSetEdit   C_SP4_FILD_SEQ,  "�ʵ�", 8, , , 2
		ggoSpread.SSSetEdit   C_SP4_FILD_NAM,  "�� �� ��", 25, , , 40
		ggoSpread.SSSetCombo  C_SP4_FILD_TYP,  "Ÿ��", 8, 2, False
		ggoSpread.SSSetCombo  C_SP4_FILD_FLG,  "�ʼ�", 8, 2, False
		ggoSpread.SSSetEdit   C_SP4_PANT_FLD,  "����", 8, , , 2
		ggoSpread.SSSetButton C_SP4_FILD_SND   
		.ReDraw = True
	End With
	

	
End Sub

'==================================================================================================
' ���� �������� �ʱ�ȭ ��Ų��.
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
' ���������Ʈ �̿��� �޺��ڽ����� �ʱ�ȭ �Ѵ�.
'--------------------------------------------------------------------------------------------------
Sub InitComboBox()
End Sub

'==================================================================================================
' �������� ��Ʈ�� �޺��ڽ��� ���� �ʱ�ȭ �Ѵ�.
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
' �����ڵ� ���� �˾� â�� ������Ų��.
'--------------------------------------------------------------------------------------------------
Function OpenPopup(Byval StrCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_BDC_MASTER"			' TABLE ��Ī 
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = " USE_FLAG= " & Filtervar("Y", "''", "S")				' Code Condition
	arrParam(5) = "����"

	arrField(0) = "PROCESS_ID"				' Field��(0)
	arrField(1) = "PROCESS_NAME"			' Field��(1)

	arrHeader(0) = "�����ڵ�"				' Header��(0)
	arrHeader(1) = "������"				' Header��(1)

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
			' �ι�° ��������� ����° ���������� ���� �մ����� �˻� �� ����.
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
				' �ι�° ��������� ����° ���������� ���� �迭�� �о� ���δ�.
				arrParams(nSpreadIndex1-1) = ""
				For i = 1 To frm1.vspdData1.MaxRows
					arrParams(nSpreadIndex1-1) = arrParams(nSpreadIndex1-1) & _
								Chr(11) & GetSpreadText(frm1.vspdData1,C_SP2_PARA_NAM,i,"X","X") & _
								Chr(11) & GetSpreadText(frm1.vspdData1,C_SP2_PARA_TYP,i,"X","X") & _
								Chr(11) & GetSpreadText(frm1.vspdData1,C_SP2_REQU_FLG,i,"X","X") & _
								Chr(11) & Chr(12)
				Next

				If nSpreadIndex2 > 0 Then
					' ���� vspdData2 �� �ڷḦ �����Ѵ�.
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
			' ���� ����ġ�� ���ο� ���� ������Ű�� ������ ���ĸ� �ϳ��� �ڷ� �о��.
			' ���ο� ���� ������Ʈ ������ �߰� �ǹǷ� 
			' ���� ������Ʈ�� ����Ű�� �ִ� �� ���Ŀ� ���� ������Ʈ ID�� �ϳ��� �з��� �Ѵ�.
			' ���� �ش� ������Ʈ�� ����Ű�� �ִ� �Ķ���� �迭�� ��ĭ�� �з��� �Ѵ�.
			' �Ķ���Ͱ� ����Ű�� �ִ� �������� �迭�� ��ĭ�� �з��� �Ѵ�.
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
			' �ι�° ��������� ����° �������带 �ʱ�ȭ �����ش�.
			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.ClearSpreadData()
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ClearSpreadData()
			
			'--------------------------------------------------------------------------------------
			' ù��° �������忡 ���ο� ���� �߰� ��Ų��.
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
			' ����° ���������� ���� �մ����� �˻� �� ����.
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
				MsgBox "�޼ҵ尡 ���õ��� �ʾҽ��ϴ�."
				Exit Function
			End If

			If nSpreadIndex1 > 0 And nSpreadIndex2 > 0 And frm1.vspdData.MaxRows > 0 Then
				'--------------------------------------------------------------------------------------
				' ����° ���������� ���� �迭�� �о� ���δ�.
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
			' ���� ����ġ�� ���ο� ���� ������Ű�� ������ ���ĸ� �ϳ��� �ڷ� �о��.
			' ���ο� ���� ������Ʈ ������ �߰� �ǹǷ� 
			' ���� ������Ʈ�� ����Ű�� �ִ� �� ���Ŀ� ���� ������Ʈ ID�� �ϳ��� �з��� �Ѵ�.
			' ���� �ش� ������Ʈ�� ����Ű�� �ִ� �Ķ���� �迭�� ��ĭ�� �з��� �Ѵ�.
			' �Ķ���Ͱ� ����Ű�� �ִ� �������� �迭�� ��ĭ�� �з��� �Ѵ�.
			' Dim arrParams(10)
			' Dim arrJoins(10, 20)
			For i = 20 To nSpreadIndex2 Step -1
				If arrJoins(nSpreadIndex1-1, i) <> "" Then
					arrJoins(nSpreadIndex1-1, i+1) = arrJoins(nSpreadIndex1-1, i)
					arrJoins(nSpreadIndex1-1, i) = ""
				End If
			Next

			'--------------------------------------------------------------------------------------
			' ����° �������带 �ʱ�ȭ �����ش�.
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ClearSpreadData()

			'--------------------------------------------------------------------------------------
			' ù��° �������忡 ���ο� ���� �߰� ��Ų��.
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
				MsgBox "�Ķ���Ͱ� ���õ��� �ʾҽ��ϴ�."
			End If
		Case 3
			'--------------------------------------------------------------------------------------
			' ��Ʈ�� �߰��Ǿ����Ƿ� ���� �߰��Ǵ� �� ������ �ʵ� ���̵� 1 �����ϹǷ� 
			' ������ ���̵� ���� ��� �迭���� ������ ������ �־�� �Ѵ�.
			' ������ �� �밡���̴�.

			If nSpreadIndex1 > 0 And nSpreadIndex2 > 0 Then
				'--------------------------------------------------------------------------------------
				' ����° ���������� ���� �迭�� �о� ���δ�.
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
					' ���� �ش� �ʵ� ���̵� ���� nSpreadIndex4 ���� ũ��...
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
			' ���� arrJoins(10,20) �迭�� ��� ������ ���� �ʵ� ���̵� ������ ���� �����ϴ��� ����.
			' ���� �����ϸ� ������ ��� ������ �ش�.
			Dim szTempRow
			For i = 0 To 10
				If arrParams(i) <> "" Then
					For j = 0 To 20
						If arrJoins(i, j) <> "" Then
							'----------------------------------------------------------------------
							' �켱 ���ڿ��� �ະ�� �����Ͽ� �迭�� �ִ´�.
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
		FncInsertRow = True                                                  '��: Processing is OK
	End If

    Set gActiveElement = document.ActiveElement
End Function

'==================================================================================================
' ���� ������ ���� �����Ѵ�.
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
			' ���缱�õ� ���� ���� �������� ��� ������ �������� �����ϰ� ù��° �������带 �ٽ� 
			' �׸��� 2 ��° �� 3��° �������嵵 �ٽ� �׸���.
			If Confirm("���� �Ķ���� �� �Ķ���� ���� ������ ����ðڽ��ϱ�?") Then
				'----------------------------------------------------------------------------------
				' �� �켱 �ڽĵ� �������� ������ 
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
				' ������ �������带 ����� �ٽ� �׸���.
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

				' ���� vspdData2 �� �ڷḦ �����Ѵ�.
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
			' ���缱�õ� ���� ���� �������� ��� ������ �������� �����ϰ� ù��° �������带 �ٽ� 
			' �׸��� 2 ��° �� 3��° �������嵵 �ٽ� �׸���.
			If Confirm("���� �Ķ���� ���� ������ ����ðڽ��ϱ�?") Then
				'----------------------------------------------------------------------------------
				' �� �켱 �ڽĵ� �������� ������ 
				For j = nSpreadIndex2 - 1 To frm1.vspdData1.MaxRows - 1
					arrJoins(nSpreadIndex1-1, j) = arrJoins(nSpreadIndex1-1, j+1)
				Next

				arrJoins(nSpreadIndex1-1, j) = ""

				'----------------------------------------------------------------------------------
				' �ΰ��� �������带 ����� �ٽ� �׸���.
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

				' ���� vspdData2 �� �ڷḦ �����Ѵ�.
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
			' ���缱�õ� ���� ���� �������� ��� ������ �������� �����ϰ� ù��° �������带 �ٽ� 
			' �׸��� 2 ��° �� 3��° �������嵵 �ٽ� �׸���.
			If Confirm("�������� ����ðڽ��ϱ�?") Then
				'----------------------------------------------------------------------------------
				' �ڽ��� �������带 ����� �ٽ� �׸���.
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

				' ���� vspdData2 �� �ڷḦ �����Ѵ�.
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

			If Confirm("�������� ����ðڽ��ϱ�?") Then
				'--------------------------------------------------------------------------------------
				' ���� ���� ��������� �˻��Ѵ�.
				For i = 1 To frm1.vspdData2.MaxRows
					If GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i,"X","X") = nSpreadIndex4 Then
						MsgBox "�����Ϸ��� ���ڵ尡 ������Դϴ�."
						Exit Function
					End If
				Next

				For i = 0 To 10
					If arrParams(i) <> "" Then
						For j = 0 To 20
							If arrJoins(i, j) <> "" Then
								'----------------------------------------------------------------------
								' �켱 ���ڿ��� �ະ�� �����Ͽ� �迭�� �ִ´�.
								arrTempRow = Split(arrJoins(i, j), iRowSep)
								For k = 0 To UBound(arrTempRow) -1
									arrTempCol = Split(arrTempRow(k), iColSep)
									If CInt(arrTempCol(1)) = nSpreadIndex4 Then
										MsgBox "�����Ϸ��� ���ڵ尡 ������Դϴ�."
										Exit Function
									End If
								Next
							End If
						Next
					End If
				Next
				'--------------------------------------------------------------------------------------
				' ���� �࿡ ���� ��ũ ������ �ִ��� �˻��Ѵ�.
				For i = 1 To frm1.vspdData3.MaxRows
					If GetSpreadText(frm1.vspdData3, C_SP4_PANT_FLD, i,"X","X") <> "" Then
						If CInt(GetSpreadText(frm1.vspdData3, C_SP4_PANT_FLD, i,"X","X")) = nSpreadIndex4 Then
							MsgBox "�����Ϸ��� �࿡ ���� ��Ʈ���������� �����մϴ�."
							Exit Function
						End If
					End If
				Next

				'----------------------------------------------------------------------------------
				' �ڽ��� �������带 ����� �ٽ� �׸���.
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
				
				' ���� vspdData3 �� �ڷḦ �����Ѵ�.
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
' ���� �۾��� ����Ѵ�.
'--------------------------------------------------------------------------------------------------
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
    Call InitData
End Function

'==========================================================================================================
' �޴����� ��ȸ ��ư�� �������� ȣ��Ǵ� �޼��� �ڵ鷯�̴�.
' �����μ�:
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
' �޴����� ��ȸ ��ư�� �������� ȣ��Ǵ� �޼��� �ڵ鷯�̴�.
' �����μ�:
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
' ��ȸ �۾��� �Ϸ� �Ǿ��� �� �ڽ� �����ӿ� ���� ȣ��ȴ�.
' �����μ�:
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
' �޴����� ���� ��ư�� �������� ȣ��Ǵ� �޼��� �ڵ鷯�̴�.
' �����μ�:
' ����: ����ڰ� �Է��� ���� �����Ѵ�.
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
		' ���� vspdData1 �� �ڷḦ �����Ѵ�.
		arrParams(nSpreadIndex1-1) = ""
		For i = 1 To frm1.vspdData1.MaxRows
			arrParams(nSpreadIndex1-1) = arrParams(nSpreadIndex1-1) & _
								Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_PARA_NAM, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_PARA_TYP, i, "X", "X") & _
								Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_REQU_FLG, i, "X", "X") & _
								Chr(11) & Chr(12)
		Next

		If nSpreadIndex2 > 0 Then
			' ���� vspdData2 �� �ڷḦ �����Ѵ�.
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
' FncSave �Լ��� ���ؼ� ȣ��Ǵ� �Լ��� ����ڰ� �ۼ��� ����Ÿ�� �����Ͽ� �����Ͻ� ������ �ִ� 
' ���α׷��� ������ �ش�.
' �����μ�:
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

    Call LayerShowHide(1)                                        '��: Protect system from crashing

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
		' ���� arrJoins(10,20) �迭�� ��� ������ ���� �ʵ� ���̵� ������ ���� �����ϴ��� ����.
		' ���� �����ϸ� ������ ��� ������ �ش�.
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
						' �켱 ���ڿ��� �ະ�� �����Ͽ� �迭�� �ִ´�.
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

    DbSave = True                                                           '��: Processing is NG
End Function

'==================================================================================================
' ���� �۾��� �Ϸ��� ���ϵ� ���������� ȣ���ϴ� �޼ҵ� �̴�.
' �����μ�:
' ��    ��: ����� ���� ������ ���� ��� ���� ������������ �����.
'--------------------------------------------------------------------------------------------------
Function DbSaveOk()
    Call InitVariables
    frm1.vspdData.MaxRows = 0
    Call MainQuery()
End Function

'==================================================================================================
' ���� �۾� �������� ������ ȣ��Ǵ� �޼��� �ڵ鷯�̴�.
' �����μ�:
' ��    ��: ����� ���� ������ ���� ��� ���� ������������ �����.
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
' ���������� �ʼ��Է� �ʵ带 �����Ѵ�.
' �����μ�:
'  Flag			: �������� ������ 
'  pvStartRow	: ������ 
'  pvEndRow		: ������ 
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
' ù��° ���������� ���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'
'  ����: ����� ��� ���� Ŭ���� ���� ��ġ ���� ���� ��� 
'		 vspdData1 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'		 vspdData2 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
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
' ù��° ���������� ���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'
'  ����: ����� ��� ���� Ŭ���� ���� ��ġ ���� ���� ��� 
'		 vspdData1 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'		 vspdData2 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
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
' ù��° ���������� ���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'  Col		  : Ŭ���� �� 
'  Row		  : Ŭ���� �� 
'
'  ����: ����� ��� ���� Ŭ���� ���� ��ġ ���� ���� ��� 
'		 vspdData1 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'		 vspdData2 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
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
' ù��° ���������� ���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'  Col		  : Ŭ���� �� 
'  Row		  : Ŭ���� �� 
'
'  ����: ����� ��� ���� Ŭ���� ���� ��ġ ���� ���� ��� 
'		 vspdData1 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'		 vspdData2 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
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
' ù��° ���������� ���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'  Col		  : Ŭ���� �� 
'  Row		  : Ŭ���� �� 
'
'  ����: ����� ��� ���� Ŭ���� ���� ��ġ ���� ���� ��� 
'		 vspdData1 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'		 vspdData2 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'--------------------------------------------------------------------------------------------------
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim i, j
	If nSpreadIndex1 <> Row And nSpreadIndex1 > 0 Then
		' ���� vspdData1 �� �ڷḦ �����Ѵ�.
		arrParams(nSpreadIndex1-1) = ""
		For i = 1 To frm1.vspdData1.MaxRows
            arrParams(nSpreadIndex1-1) = arrParams(nSpreadIndex1-1) & _
							Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_PARA_NAM, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_PARA_TYP, i, "X", "X") & _
							Chr(11) & GetSpreadText(frm1.vspdData1, C_SP2_REQU_FLG, i, "X", "X") & _
							Chr(11) & Chr(12)
		Next

		If nSpreadIndex2 > 0 Then
			' ���� vspdData2 �� �ڷḦ �����Ѵ�.
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
			' ���� vspdData1 �� �ڷḦ �����Ѵ�.
			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.ClearSpreadData()
			ggoSpread.SSShowData arrParams(Row - 1)
			Call SetSpreadColor(1, 1, frm1.vspdData1.MaxRows)
			frm1.vspdData1.ReDraw = True

			' ���� vspdData2 �� �ڷḦ �����Ѵ�.
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
' ù��° ���������� ���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'  Col		  : Ŭ���� �� 
'  Row		  : Ŭ���� �� 
'
'  ����: ����� ��� ���� Ŭ���� ���� ��ġ ���� ���� ��� 
'		 vspdData1 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'		 vspdData2 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'--------------------------------------------------------------------------------------------------
Sub vspdData1_Click(ByVal Col, ByVal Row)
	Dim i, j
	
	If nSpreadIndex2 <> Row And nSpreadIndex1 > 0 And nSpreadIndex2 > 0 Then
		' ���� vspdData1 �� �ڷḦ �����Ѵ�.
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
			' ���� vspdData1 �� �ڷḦ �����Ѵ�.
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
' ù��° ���������� ���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'  Col		  : Ŭ���� �� 
'  Row		  : Ŭ���� �� 
'
'  ����: ����� ��� ���� Ŭ���� ���� ��ġ ���� ���� ��� 
'		 vspdData1 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'		 vspdData2 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'--------------------------------------------------------------------------------------------------
Sub vspdData2_Click(ByVal Col, ByVal Row)
	Dim i
	If nSpreadIndex3 <> Row And nSpreadIndex1 > 0 And nSpreadIndex2 > 0 Then
		' ���� vspdData1 �� �ڷḦ �����Ѵ�.
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
' ù��° ���������� ���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'  Col		  : Ŭ���� �� 
'  Row		  : Ŭ���� �� 
'
'  ����: ����� ��� ���� Ŭ���� ���� ��ġ ���� ���� ��� 
'		 vspdData1 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'		 vspdData2 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'--------------------------------------------------------------------------------------------------
Sub vspdData3_Click(ByVal Col, ByVal Row)
	nSpreadIndex4 = Row
End Sub

'==========================================================================================================
' �׹�° ���������� ��ư���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'  Col		  : Ŭ���� �� 
'  Row		  : Ŭ���� �� 
'  ButtonDown : �׻� 0 ��(�����ص� ��)
'
'  ����: vspdData2�� ���� �࿡ �ʵ������� ������.
'----------------------------------------------------------------------------------------------------------
Sub vspdData3_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	'If nSpreadIndex3 = 0 Or frm1.vspdData2.MaxRows = 0 Then
	'	MsgBox "@@������ ���� �����Ͻʽÿ�"
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
' �׹�° ���������� ���� ����Ǿ����� �޼��� �ڵ鷯�̴�.
' �����μ�:
'  Col		  : Ŭ���� �� 
'  Row		  : Ŭ���� �� 
'
'  ����: ���� ������ �ڵ����� Field_ID �� �������Ѿ� �Ѵ�.
'        ���� ���ڵ忡 �����ϴ� Field_ID�� �ִ밪�� 1 �� ���Ѱ��� ���ο� Field_Id�� �Ѵ�.
'----------------------------------------------------------------------------------------------------------
Sub vspdData3_Change(ByVal Col, ByVal Row)
	Dim i, j, k
	Exit Sub
	'----------------------------------------------------------------------------------------------
	' ��Ʈ�� �߰��Ǿ����Ƿ� ���� �߰��Ǵ� �� ������ �ʵ� ���̵� 1 �����ϹǷ� 
	' ������ ���̵� ���� ��� �迭���� ������ ������ �־�� �Ѵ�.
	' ������ �� �밡���̴�.

	If nSpreadIndex1 > 0 And nSpreadIndex2 > 0 Then
		'------------------------------------------------------------------------------------------
		' ����° ���������� ���� �迭�� �о� ���δ�.
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
			' ���� �ش� �ʵ� ���̵� ���� nSpreadIndex4 ���� ũ��...
			If CInt(GetSpreadText(frm1.vspdData2, C_SP3_FILD_NUM, i, "X", "X")) > nSpreadIndex4 Then
				Call SetSpreadValue(frm1.vspdData2, 1, i, _
									CInt(GetSpreadText(frm1.vspdData2, _
													   C_SP3_FILD_NUM, i, "X", "X")) + 1, "", "")
			End If
		Next
	End If

	'--------------------------------------------------------------------------------------
	' ���� arrJoins(10,20) �迭�� ��� ������ ���� �ʵ� ���̵� ������ ���� �����ϴ��� ����.
	' ���� �����ϸ� ������ ��� ������ �ش�.
	Dim szTempRow
	For i = 0 To 10
		If arrParams(i) <> "" Then
			For j = 0 To 20
				If arrJoins(i, j) <> "" Then
					'----------------------------------------------------------------------
					' �켱 ���ڿ��� �ະ�� �����Ͽ� �迭�� �ִ´�.
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
' ���� ��ư 1 �̺�Ʈ �ڵ鷯 
'--------------------------------------------------------------------------------------------------
Function Radio1_OnClick()
	frm1.hUseFlag.value = frm1.Radio1.value
End Function

'==================================================================================================
' ���� ��ư 2 �̺�Ʈ �ڵ鷯 
'--------------------------------------------------------------------------------------------------
Function Radio2_OnClick()
	frm1.hUseFlag.value = frm1.Radio2.value
End Function

'==================================================================================================
' ���� ��ư 3 �̺�Ʈ �ڵ鷯 
'--------------------------------------------------------------------------------------------------
Function Radio3_OnClick()
	frm1.hJoinMethod.value = frm1.Radio3.value
End Function

'==================================================================================================
' ���� ��ư 4 �̺�Ʈ �ڵ鷯 
'--------------------------------------------------------------------------------------------------
Function Radio4_OnClick()
	frm1.hJoinMethod.value = frm1.Radio4.value
End Function

'==================================================================================================
' ���� ��ư 6 �̺�Ʈ �ڵ鷯 
'--------------------------------------------------------------------------------------------------
Function Radio6_OnClick()
	frm1.hTranFlag.value = frm1.Radio6.value
End Function

'==================================================================================================
' ���� ��ư 7 �̺�Ʈ �ڵ鷯 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�������</font></td>
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
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD656" NOWRAP>
									    <INPUT NAME="txtProcID" MAXLENGTH="20" SIZE=20 ALT ="�����ڵ�" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcID" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(txtProcID.value, 0)">
										<INPUT NAME="txtProcNm" MAXLENGTH="80" SIZE=40 ALT ="�� �� ��" tag="14X"  STYLE="TEXT-ALIGN:left"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=60 WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtNProcNm" ALT="������" MAXLENGTH="80" SIZE=40 STYLE="TEXT-ALIGN: left" tag ="22N">
								</TD>
                                <TD CLASS="TD5">��뿩��</TD>
                                <TD CLASS="TD6">
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoUseFlag" ID="Radio1" VALUE="Y" TAG = "11X" CHECKED>
										<LABEL FOR="Radio1">���</LABEL>&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoUseFlag" ID="Radio2" VALUE="N" TAG = "11X">
										<LABEL FOR="Radio2">�̻��</LABEL>&nbsp;&nbsp;
                                </TD>
							</TR>
							<TR>
                                <TD CLASS="TD5">������</TD>
                                <TD CLASS="TD6">
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoJoinFlag" ID="Radio3" VALUE="N" TAG = "11X" CHECKED>
										<LABEL FOR="Radio3">�ܵ�</LABEL>&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoJoinFlag" ID="Radio4" VALUE="S" TAG = "11X">
										<LABEL FOR="Radio4">��Ʈ</LABEL>&nbsp;&nbsp;
                                </TD>
                                <TD CLASS="TD5">Ʈ������</TD>
                                <TD CLASS="TD6">
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoTranFlag" ID="Radio6" VALUE="Y" TAG = "11X" CHECKED>
										<LABEL FOR="Radio6">���</LABEL>&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoTranFlag" ID="Radio7" VALUE="N" TAG = "11X">
										<LABEL FOR="Radio7">�̻��</LABEL>&nbsp;&nbsp;
                                </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�� �� ��</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtStartRow" ALT="�� �� ��" MAXLENGTH="2" SIZE=10 STYLE="TEXT-ALIGN: left" tag ="22N">
								</TD>
								<TD CLASS=TD5 NOWRAP>����ð�</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtRunTime" ALT="����ð�" MAXLENGTH="5" SIZE=10 STYLE="TEXT-ALIGN: left" tag ="22N">
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

