<%@ LANGUAGE="VBSCRIPT" %>
<!--'==============================================================================================
'*  1. Module Name          : BDC
'*  2. Function Name        : 
'*  3. Program ID           : BDC02MA1
'*  4. Program Name         : BDC �������� ��� 
'*  5. Program Desc         : BDC SQL�� Mapping ���� �Է� 
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
Const BIZ_PGM_QUERY_ID = "BDC02MB1.ASP"								'��: ��ȸ �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID  = "BDC02MB2.ASP"								'��: ���� �����Ͻ� ���� ASP�� 

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
' �������� �ٸ��� �������� �� ���������� ���� ������ �迭�� �����Ѵ�.
Dim arrQuery(10, 1)

' �������� ��Ʈ�� ������ ���� ���� 
Dim nCurrentSpread
Dim nSpreadIndex1
Dim nSpreadIndex2

Dim strMode
Dim IsOpenPop
Dim lgRetFlag


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
	With frm1.vspdData1
        .ReDraw = False
		.MaxCols = C_SP1_RSLT
        .MaxRows = 0

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit			'"V20050321",,parent.gAllowDragDropSpread
		ggoSpread.SSSetEdit  C_SP1_DESC,  "���ǹ���", 50, , , 40
		ggoSpread.SSSetEdit  C_SP1_RSLT,  "�����", 30, , , 40
		.ReDraw = True
	End With
	
	With frm1.vspdData2
        .ReDraw = False
		.MaxCols = C_SP2_FNM
        .MaxRows = 0

		ggoSpread.Source =	frm1.vspdData2
		ggoSpread.Spreadinit			'"V20021121",,parent.gAllowDragDropSpread
		ggoSpread.SSSetEdit  C_SP2_NUM,  "��ȣ", 3, , , 20
		ggoSpread.SSSetEdit  C_SP2_NAM,  "�̸�", 12, , , 20
		ggoSpread.SSSetCombo C_SP2_TYP,  "����", 8, 2, False
		ggoSpread.SSSetEdit  C_SP2_LNG,  "����", 8, , , 2
		ggoSpread.SSSetEdit  C_SP2_FID,  "�ʵ�ID", 8, , , 2
		ggoSpread.SSSetEdit  C_SP2_FNM,  "�ʵ��", 12, , ,40
		Call ggoSpread.SSSetColHidden(C_SP2_NUM, C_SP2_NUM, True)
		.ReDraw = True
	End With

	ggoSpread.Source = frm1.vspdData3
	ggoSpread.Spreadinit			'"V20021121",,parent.gAllowDragDropSpread
	With frm1.vspdData3
        .ReDraw = False
		.MaxCols = C_SP3_SND
        .MaxRows = 0

		ggoSpread.SSSetEdit   C_SP3_NUM,  "��Ʈ", 6, , , 1
		ggoSpread.SSSetEdit   C_SP3_SEQ,  "�ʵ�", 6, , , 2
		ggoSpread.SSSetEdit   C_SP3_NAM,  "�ʵ��", 15, , , 40
		ggoSpread.SSSetButton C_SP3_SND
		.ReDraw = True
	End With
End Sub

'==================================================================================================
' ���� �������� �ʱ�ȭ ��Ų��.
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
' ���������Ʈ �̿��� �޺��ڽ����� �ʱ�ȭ �Ѵ�.
'--------------------------------------------------------------------------------------------------
Sub InitComboBox()
End Sub

'==================================================================================================
' �������� ��Ʈ�� �޺��ڽ��� ���� �ʱ�ȭ �Ѵ�.
'--------------------------------------------------------------------------------------------------
Sub InitGridComboBox()
	ggoSpread.Source = frm1.vspdData2
    ggoSpread.SetCombo "VarChar" & vbTab & "NVarChar" & vbTab & "Integer", C_SP2_TYP
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

	arrParam(0) = "�����˾�"			' �˾� ��Ī 
	arrParam(1) = "B_BDC_MASTER"			' TABLE ��Ī 
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = " USE_FLAG= " & Filtervar("Y", "''", "S")						' Code Condition
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
    Dim i
    Dim iColSep, iRowSep

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep

    FncInsertRow = False
	Select Case (nCurrentSpread)
		Case 1
			'--------------------------------------------------------------------------------------
			' �ι�° ��������� ����° ���������� ���� �մ����� �˻� �� ����.
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
			    
				' ���� vspdData2 �� �ڷḦ �����Ѵ�.
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
			' ���� ����ġ�� ���ο� ���� ������Ű�� ������ ���ĸ� �ϳ��� �ڷ� �о��.
			' ���ο� ���� ������Ʈ ������ �߰� �ǹǷ� 
			' ���� ������Ʈ�� ����Ű�� �ִ� �� ���Ŀ� ���� ������Ʈ ID�� �ϳ��� �з��� �Ѵ�.
			' ���� �ش� ������Ʈ�� ����Ű�� �ִ� �Ķ���� �迭�� ��ĭ�� �з��� �Ѵ�.
			' �Ķ���Ͱ� ����Ű�� �ִ� �������� �迭�� ��ĭ�� �з��� �Ѵ�.
			
			frm1.txtSQLStatements.value = ""
			'--------------------------------------------------------------------------------------
			' �ι�° �������带 �ʱ�ȭ �����ش�.
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ClearSpreadData()
			
			'--------------------------------------------------------------------------------------
			' ù��° �������忡 ���ο� ���� �߰� ��Ų��.
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
		FncInsertRow = True                                           '��: Processing is OK
	End If

    Set gActiveElement = document.ActiveElement
End Function

'==================================================================================================
' ���� ������ ���� �����Ѵ�.
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
			' ���缱�õ� ���� ���� �������� ��� ������ �������� �����ϰ� ù��° �������带 �ٽ� 
			' �׸��� 2 ��° �� 3��° �������嵵 �ٽ� �׸���.
			If Confirm("�������� �Ķ���� ������ �Բ� ����ðڽ��ϱ�?") Then
				'----------------------------------------------------------------------------------
				' �� �켱 �ڽĵ� �������� ������ 
				For i = nSpreadIndex1 - 1 To frm1.vspdData1.MaxRows - 1
					arrQuery(i, 0) = arrQuery(i+1, 0)
					arrQuery(i, 1) = arrParam(i+1, 1)
				Next

				arrQuery(10, 0) = ""
				arrQuery(10, 1) = ""
				
				'----------------------------------------------------------------------------------
				' ������ �������带 ����� �ٽ� �׸���.
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
			' ���缱�õ� ���� ���� �������� ��� ������ �������� �����ϰ� ù��° �������带 �ٽ� 
			' �׸��� 2 ��° �� 3��° �������嵵 �ٽ� �׸���.
			If Confirm("�Ķ���� ���� ������ ����ðڽ��ϱ�?") Then

				'----------------------------------------------------------------------------------
				' �ΰ��� �������带 ����� �ٽ� �׸���.
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
' ���� �۾��� ����Ѵ�.
'--------------------------------------------------------------------------------------------------
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData1
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
' �޴����� ��ȸ ��ư�� �������� ȣ��Ǵ� �޼��� �ڵ鷯�̴�.
' �����μ�:
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
' ��ȸ �۾��� �Ϸ� �Ǿ��� �� �ڽ� �����ӿ� ���� ȣ��ȴ�.
' �����μ�:
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
' �޴����� ���� ��ư�� �������� ȣ��Ǵ� �޼��� �ڵ鷯�̴�.
' �����μ�:
' ����: ����ڰ� �Է��� ���� �����Ѵ�.
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
		
		' ���� vspdData2 �� �ڷḦ �����Ѵ�.
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
' FncSave �Լ��� ���ؼ� ȣ��Ǵ� �Լ��� ����ڰ� �ۼ��� ����Ÿ�� �����Ͽ� �����Ͻ� ������ �ִ� 
' ���α׷��� ������ �ش�.
' �����μ�:
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

    Call LayerShowHide(1)                                        '��: Protect system from crashing

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

    DbSave = True                                                           '��: Processing is NG
End Function

'==================================================================================================
' ���� �۾��� �Ϸ��� ���ϵ� ���������� ȣ���ϴ� �޼ҵ� �̴�.
' �����μ�:
' ��    ��: ����� ���� ������ ���� ��� ���� ������������ �����.
'--------------------------------------------------------------------------------------------------
Function DbSaveOk()
    Call InitVariables
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
' ���������� �ʼ��Է� �ʵ带 �����Ѵ�.
' �����μ�:
'  Flag			: �������� ������ 
'  pvStartRow	: ������ 
'  pvEndRow		: ������ 
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
' ù��° ���������� ���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'
'  ����: ����� ��� ���� Ŭ���� ���� ��ġ ���� ���� ��� 
'		 vspdData1 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'		 vspdData2 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
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
' ù��° ���������� ���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'
'  ����: ����� ��� ���� Ŭ���� ���� ��ġ ���� ���� ��� 
'		 vspdData1 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
'		 vspdData2 �� ���� �����ϰ� ���� Ŭ���� �࿡ ��ġ�ϴ� �����ͷ� ġȯ�� �ش�.
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
    If Button = 2 And gMouseClickStatus = "SP3C" Then
       gMouseClickStatus = "SP3CR"
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
Sub vspdData1_Click(ByVal Col, ByVal Row)
	'Exit Sub
	Dim i, j
	
	If nSpreadIndex1 <> Row And nSpreadIndex1 > 0 Then
		' ���� vspdData1 �� �ڷḦ �����Ѵ�.		

		If nSpreadIndex2 > 0 Then
			' ���� vspdData2 �� �ڷḦ �����Ѵ�.
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
			' ���� vspdData2 �� �ڷḦ �����Ѵ�.
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
' �׹�° ���������� ��ư���� Ŭ���Ǿ������� �޼��� �ڵ鷯�̴�.
' �����μ�:
'  Col		  : Ŭ���� �� 
'  Row		  : Ŭ���� �� 
'  ButtonDown : �׻� 0 ��(�����ص� ��)
'
'  ����: vspdData2�� ���� �࿡ �ʵ������� ������.
'----------------------------------------------------------------------------------------------------------
Sub vspdData3_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	If nSpreadIndex2 = 0 Or frm1.vspdData2.MaxRows = 0 Then
		MsgBox "������ ���� �����Ͻʽÿ�"
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�����������</font></td>
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
									<TD CLASS="TD5" NOWRAP>�����ڵ�</TD>
									<TD CLASS="TD656" NOWRAP>
									    <INPUT NAME="txtProcID" MAXLENGTH="20" SIZE=20 ALT ="�����ڵ�" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcID" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(txtProcID.value, 0)">
										<INPUT NAME="txtProcNm" MAXLENGTH="80" SIZE=40 ALT ="�� �� ��" tag="14X"></TD>
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
									<TD CLASS="TD5" NOWRAP>����SQL</TD>
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
