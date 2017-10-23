<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ���������� 
'*  3. Program ID           : W3135MA1
'*  4. Program Name         : W3135MA1.asp
'*  5. Program Desc         : ��40ȣ(��) ��ȭ�ڻ�� �������� ��������(��)
'*  6. Modified date(First) : 2005/01/20
'*  7. Modified date(Last)  : 2005/01/20
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  �α����� ������ �����ڵ带 ����ϱ� ����  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "W3135MA1"
Const BIZ_PGM_ID		= "W3135mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID	    = "W3135OA1"

Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.

' -- �׸��� �÷� ���� 
Dim C_SEQ_NO
Dim C_W1_CD
Dim C_W1
Dim C_W2
Dim C_W2_P
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W8
Dim C_W9
Dim C_W10

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(2)

Dim lgW2, lgMonth	' ������, ����������� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	
	C_SEQ_NO	= 1
	C_W1_CD		= 2
	C_W1		= 3	' ���� 
	C_W2		= 4	' ��ȭ���� 
	C_W2_P		= 5 ' 
	C_W3		= 6	' ��ȭ�ݾ� 
	C_W4		= 7	' ����ȯ�� 
	C_W5		= 7	' ��ȭ�ݾ� 
	C_W6		= 8	' 
	C_W7		= 9
	C_W8		= 9	' 
	C_W9		= 10 ' 
	C_W10		= 11 ' �򰡼��� 
	
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False

    lgCurrGrid = TYPE_1
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  �Ű��� �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	' ��ȸ����(����)
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))

End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1) = frm1.vspdData0
	'Set lgvspdData(TYPE_2) = frm1.vspdData1
	
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("6","4","2")
	Call AppendNumberPlace("7","15","2")
	
	' 1�� �׸��� 
	For iRow = TYPE_1 To TYPE_1
	
		With lgvspdData(iRow)
				
			ggoSpread.Source = lgvspdData(iRow)	
			'patch version
			ggoSpread.Spreadinit "V20041222_" & iRow,,parent.gAllowDragDropSpread    
    
			.ReDraw = false

			.MaxCols = C_W10 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.Col = .MaxCols									'��: ����� �� Hidden Column
			.ColHidden = True    
 
			.MaxRows = 0
			ggoSpread.ClearSpreadData

			ggoSpread.SSSetEdit		C_SEQ_NO,	"����", 10,,,10,1	' �����÷� 
			ggoSpread.SSSetEdit		C_W1_CD,	"�ڵ�", 7,,,10,1
			ggoSpread.SSSetEdit		C_W1,		"(1)����", 7,,,50,1
			ggoSpread.SSSetEdit		C_W2,		"(2)��ȭ" & vbCrlf & "����", 7,2,,3,1
			ggoSpread.SSSetButton   C_W2_P
			ggoSpread.SSSetFloat	C_W3,		"(3)��ȭ�ݾ�"	, 15, "7",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W5,		"(5)����ȯ��"	, 10, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W6,		"(6)��ȭ�ݾ�"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
			ggoSpread.SSSetFloat	C_W8,		"(8)����ȯ��"	, 10, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W9,		"(9)��ȭ�ݾ�" 	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W10,		"(10)�򰡼���"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	

			'If iRow = TYPE_1 Then
				'����� 3�ٷ�    
				.ColHeaderRows = 2    
								      
				' �׸��� ��� ��ħ 
				ret = .AddCellSpan(C_W1_CD	, -1000, 1, 2)	' ���� 2�� ��ħ 
				ret = .AddCellSpan(C_W1		, -1000, 1, 2)	
				ret = .AddCellSpan(C_W2		, -1000, 1, 2)
				ret = .AddCellSpan(C_W2_P	, -1000, 1, 2)
				ret = .AddCellSpan(C_W3		, -1000, 1, 2)
				ret = .AddCellSpan(C_W4		, -1000, 2, 1)
				ret = .AddCellSpan(C_W7		, -1000, 2, 1)
    
				' ù��° ��� ��� ���� 
				.Row = -1000
				.Col = C_W4	: .Text = "(4) �� �� �� ��"
				.Col = C_W7	: .Text = "(7) �� �� �� ��"
				.Col = C_W10: .Text = "(10)�� ����"
			
				.Row = -999
				.Col = C_W5	: .Text = "(5)����ȯ��"
				.Col = C_W6	: .Text = "(6)��ȭ�ݾ�"
				.Col = C_W8	: .Text = "(8)����ȯ��"
				.Col = C_W9	: .Text = "(9)��ȭ�ݾ�"
				.Col = C_W10: .Text = "�ڻ�[(9)-(6)]" & vbCrLf & "��ä[(6)-(9)]"
						
				.rowheight(-999) = 20	' ���� ������	(2���� ���, 1���� 15)
			
				Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
				Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_W1_CD,True)
				
			'Else
			'	.ColHeadersShow = False
			'End If

					
			'Call SetSpreadLock(iRow)
					
			.ReDraw = true	
				
		End With 
	Next
 
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

End Sub

Sub InsertFirstRow()
	Dim iMaxRows, iRow, iType, ret

	iMaxRows = 5 ' �ϵ��ڵ��Ǵ� ��� 

	With lgvspdData(TYPE_1)
		ggoSpread.Source = lgvspdData(TYPE_1)
		.Redraw = False

		ggoSpread.InsertRow , iMaxRows
		Call SetSpreadLock
		
		iRow = 1
		
		.Row = iRow		
		.Col = C_SEQ_NO : .Value = iRow		: iRow = iRow + 1
		.Col = C_W1_CD	: .Value = "0"
		.Col = C_W1		: .value = "��" & vbCrLf & "��"
		.TypeEditMultiLine = True
		.TypeHAlign = 2 : .TypeVAlign = 2
		
		.Row = iRow		
		.Col = C_SEQ_NO : .Value = SUM_SEQ_NO	: iRow = iRow + 1
		.Col = C_W1_CD	: .Value = "09"
		.Col = C_W2		: .CellType = 1	: .Text = "�հ�"	: .TypeHAlign = 2	
		.Col = C_W2_P	: .CellType = 1
		'.Col = C_W6		: .Formula = "SUM(H1:H2)"
		ggoSpread.SpreadLock C_W1, .Row, C_W10, .Row
		ret = .AddCellSpan(C_W1	, .Row - 1, 1, 2)
		
		.Row = iRow		
		.Col = C_SEQ_NO : .Value = "2"	: iRow = iRow + 1
		.Col = C_W1_CD	: .Value = "1"
		.Col = C_W1		: .value = "��" & vbCrLf & "ä"
		.TypeEditMultiLine = True
		.TypeHAlign = 2 : .TypeVAlign = 2
		
		.Row = iRow		
		.Col = C_SEQ_NO : .Value = SUM_SEQ_NO	: iRow = iRow + 1
		.Col = C_W1_CD	: .Value = "19"
		.Col = C_W2		: .CellType = 1	: .Text = "�հ�"	: .TypeHAlign = 2	
		.Col = C_W2_P	: .CellType = 1
		'.Col = C_W6		: .Formula = "SUM(H3:H4)"
		ggoSpread.SpreadLock C_W1, .Row, C_W10, .Row
		ret = .AddCellSpan(C_W1	, .Row - 1, 1, 2)

		.Row = iRow		
		.Col = C_SEQ_NO : .Value = SUM_SEQ_NO	: iRow = iRow + 1
		.Col = C_W1_CD	: .Value = "99"
		.Col = C_W1		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2	
		.Col = C_W2_P	: .CellType = 1
		'.Col = C_W6		: .Formula = "SUM(H2,H3)"
		ret = .AddCellSpan(C_W1	, .Row, 3, 1)
		ggoSpread.SpreadLock C_W1, .Row, C_W10, .Row
					
		.Redraw = True
	
	End With
	'Call SetSpreadLock(iType)
End Sub

Sub WriteLeftHead(pType)
	With lgvspdData(pType)
		.Col = C_W1 : .Row = 1
		.TypeEditMultiLine = True
		.TypeHAlign = 2 : .TypeVAlign = 2

		If pType = TYPE_1 Then
			.Text = "ȭ" & vbCrLf & "��" & vbCrLf & "��" & vbCrLf & "��" & vbCrLf & "ȭ" & vbCrLf & "��" & vbCrLf & "��"
		Else
			.Text = "ȭ" & vbCrLf & "��" & vbCrLf & "��" & vbCrLf & "��" & vbCrLf & "ȭ" & vbCrLf & "��" & vbCrLf & "ä"
		End If
	End With
End Sub

Sub SetSpreadLock()

	ggoSpread.Source = lgvspdData(TYPE_1)	

	ggoSpread.SpreadLock C_SEQ_NO, -1, C_W1
	ggoSpread.SpreadLock C_W10, -1, C_W10
	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

	ggoSpread.Source = lgvspdData(TYPE_1)	

	ggoSpread.Protected C_SEQ_NO, pvStartRow, pvEndRow
	ggoSpread.Protected C_W1_CD, pvStartRow, pvEndRow
	ggoSpread.Protected C_W1, pvStartRow, pvEndRow
	ggoSpread.Protected C_W10, pvStartRow, pvEndRow

End Sub

Sub SetSpreadTotalLine()
	Dim iRow
	For iRow = TYPE_1 To TYPE_2
		ggoSpread.Source = lgvspdData(iRow)
		With lgvspdData(iRow)
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W9 : .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
				ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
			End If
		End With
	Next
End Sub 

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W_TYPE	= iCurColumnPos(2)
            C_W13		= iCurColumnPos(3)
            C_W1		= iCurColumnPos(4)
            C_W2		= iCurColumnPos(5)
            C_W13		= iCurColumnPos(6)
            C_W3		= iCurColumnPos(7)
            C_W4		= iCurColumnPos(8)
            C_W5		= iCurColumnPos(9)
            C_W13		= iCurColumnPos(10)
            C_W15		= iCurColumnPos(11)
            C_W16		= iCurColumnPos(12)
            C_W9		= iCurColumnPos(13)
            C_W_TYPE	= iCurColumnPos(14)
            C_W1		= iCurColumnPos(15)
            C_W2		= iCurColumnPos(16)
    End Select    
End Sub

'============================== ���۷��� �Լ�  ========================================

Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 

End Function

Function OpenCurrency()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��ȭ �˾�"						<%' �˾� ��Ī %>
	arrParam(1) = "b_currency"						<%' TABLE ��Ī %>

	With lgvspdData(TYPE_1)
		.Col = .ActiveCol -1 : .Row = .ActiveRow
		arrParam(2) = Trim(.Text)		<%' Code Condition%>
	End With
	
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "��ȭ"							<%' �����ʵ��� �� ��Ī %>
	
    arrField(0) = "currency"						<%' Field��(0)%>
    arrField(1) = "currency_desc"					<%' Field��(1)%>
    
    arrHeader(0) = "��ȭ"							<%' Header��(0)%>
    arrHeader(1) = "���"							<%' Header��(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCurrency(arrRet)
	End If	
	
End Function

Function SetCurrency(Byval arrRet)
	With lgvspdData(TYPE_1)
		.Col =C_W2 : .Row = .ActiveRow
		.Text = arrRet(0)
		lgBlnFlgChgValue = True
	End With
End Function


Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    'iCalledAspName = AskPRAspName("W5105RA1")
    
'    If Trim(iCalledAspName) = "" Then
 '       IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W5105RA1", "x")
  '      IsOpenPop = False
   '     Exit Function
'    End If

    arrRet = window.showModalDialog("../W5/W5105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function
'====================================== �� �Լ� =========================================

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100000111")										<%'��ư ���� ���� %>
	  
	' �����Ѱ� 
	Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)
	
	Call InitData 
	Call FncQuery
     
    
End Sub


'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub

Sub cboREP_TYPE_onChange()	' �Ű������ �ٲٸ�..
	'Call GetFISC_DATE
End Sub

'============================================  �׸��� �̺�Ʈ   ====================================
' -- 0�� �׸��� 
Sub vspdData0_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_1
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_1
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData0_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_1
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData0_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_1
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_GotFocus()
	lgCurrGrid = TYPE_1
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData0_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_1
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData0_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_1
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData0_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_1
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData0_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_1
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub

' -- 1�� �׸��� 
Sub vspdData1_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_2
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_2
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_2
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_GotFocus()
	lgCurrGrid = TYPE_2
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData1_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_2
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_2
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_2
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, dblW3, dblW5, dblW6, dblW8, dblW9, dblW1_CD
	
	lgBlnFlgChgValue= True ' ���濩�� 
    lgvspdData(lgCurrGrid).Row = Row
    lgvspdData(lgCurrGrid).Col = Col

    If lgvspdData(Index).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(lgvspdData(Index).text) < CDbl(lgvspdData(Index).TypeFloatMin) Then
         lgvspdData(Index).text = lgvspdData(Index).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(Index)
    ggoSpread.UpdateRow Row

	' --- �߰��� �κ� 
	With lgvspdData(Index)

	If Index = TYPE_1 Then	'1�� �׸� 
	
		Select Case Col
			Case C_W3, C_W5, C_W8
				.Col = C_W3	: dblW3 = UNICDbl(.Value)
				.Col = C_W5	: dblW5 = UNICDbl(.Value)
				.Col = C_W8	: dblW8 = UNICDbl(.Value)
				.Col = C_W6	: .Value = dblW3 * dblW5
				.Col = C_W9	: .Value = dblW3 * dblW8
				.Col = C_W1_CD	: dblW1_CD = .Value

				Call SetSum2Col(C_W6, dblW1_CD)
				Call SetSum2Col(C_W9, dblW1_CD)
				Call SetW10(Row)
			Case C_W6, C_W9
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.Value)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.Value = 0
				End If
				.Col = C_W1_CD
				Call SetSum2Col(Col, .Value)
				Call SetW10(Row)
		End Select
	
	End If
	
	End With
	
End Sub

' �򰡼��� ��� 
Function SetW10(pRow)
	Dim dblW6, dblW9, dblW10, dblW1_CD
	With lgvspdData(TYPE_1)
		.Row = pRow
		.Col = C_W6		: dblW6 = UNICDbl(.Value)
		.Col = C_W9		: dblW9 = UNICDbl(.Value)
		.Col = C_W1_CD	: dblW1_CD = UNICDbl(.Value)
		
		If dblW1_CD = "0" Then
			dblW10 = dblW9 - dblW6
		Else
			dblW10 = dblW6 - dblW9
		End If
		.Col = C_W10
		.Value = dblW10
		
		.Col = C_W1_CD
		Call SetSum2Col(C_W10, .Value)
	End With
End Function

' ���� �÷��� �������� �հ� ����� �� �� ����Ѵ�.
Function SetSum2Col(Byval pCol, Byval pW1_CD)
	Dim dblSum09, dblSum19, dblSum99, dblSumCol, iRow, sW1_CD, iMaxRows, iDx
	sW1_CD = "" : iDx = 0
	
	With lgvspdData(TYPE_1)	' ��Ŀ���� �׸��� 
		iMaxRows = .MaxRows
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		For iRow = 1 To iMaxRows
			.Row = iRow	: .Col = C_W1_CD
			
			If .Text = pW1_CD Then	' ���� ���� 
				.Col = pCol
				dblSumCol = dblSumCol + UNICDbl(.Value)
			ElseIf .Text = pW1_CD & "9" Then
				' �հ� ��� 
				.Col = pCol
				.Value = dblSumCol
				ggoSpread.UpdateRow iRow
				
				If pW1_CD = "0" Then
					' ����(19)�� �� ���� �о�´�.
					dblSum09	= dblSumCol
					dblSum19	= GetSum2Col(pCol, "19")
					dblSum99	= dblSum09 + dblSum19
				Else
					' ����(09)�� �� ���� �о�´�.
					dblSum19	= dblSumCol
					dblSum09	= GetSum2Col(pCol, "09")
					dblSum99	= dblSum09 + dblSum19
				End If
				
				.Row = .MaxRows	: .Col = pCol	: .Value = dblSum99
				
				ggoSpread.UpdateRow .MaxRows
				Exit Function
			End If
			
		Next
		
	End With
End Function

Function GetSum2Col(Byval pCol, Byval pW1_CD)
	Dim iRow, iMaxRows
	
	With lgvspdData(TYPE_1)	' ��Ŀ���� �׸��� 
		iMaxRows = .MaxRows
		
		For iRow = 1 To iMaxRows	
			.Row = iRow	: .Col = C_W1_CD
			
			If .Text = pW1_CD Then	' ���� ����	
				.Col = pCol
				GetSum2Col = UNICDbl(.Value)
				Exit Function
			End If		
		Next
		
	End With
End Function

Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = lgvspdData(Index)
   
    If lgvspdData(Index).MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = lgvspdData(Index)
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	lgvspdData(Index).Row = Row
End Sub

Sub vspdData_ColWidthChange(Index, ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = lgvspdData(Index)
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(Index, ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If lgvspdData(Index).MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData_GotFocus(Index)
    ggoSpread.Source = lgvspdData(Index)
    lgCurrGrid = Index
End Sub

Sub vspdData_MouseDown(Index, Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	lgCurrGrid = Index
	ggoSpread.Source = lgvspdData(Index)
End Sub    

Sub vspdData_ScriptDragDropBlock(Index, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(Index, ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if lgvspdData(Index).MaxRows < NewTop + VisibleRowCnt(lgvspdData(Index),NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub vspdData_ButtonClicked(Index, ByVal Col, ByVal Row, Byval ButtonDown)
	With lgvspdData(Index)
		If Row > 0 And Col = C_W2_P Then
		    .Row = Row
		    .Col = C_W2_P

		    Call OpenCurrency()
		End If
    End With
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'============================================  �������� �Լ�  ====================================

Function FncQuery() 
    Dim IntRetCD , i, blnChange
    
    FncQuery = False                                                        
    blnChange = False
    
    Err.Clear                                                               <%'Protect system from crashing%>

	
<%  '-----------------------
    'Check previous data area
    '----------------------- %>
	'For i = TYPE_1 To TYPE_6
	'	ggoSpread.Source = lgvspdData(i)
	'	If ggoSpread.SSCheckChange = True Then
	'		blnChange = True
	'		Exit For
	'	End If
    'Next
    
    If lgBlnFlgChgValue Or blnChange Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables													<%'Initializes local global variables%>
    Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, i, sMsg
    
    blnChange = False
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	    
    For i = TYPE_1 To TYPE_1
    
		ggoSpread.Source = lgvspdData(i)
		If ggoSpread.SSCheckChange = False Then
			Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
			Exit Function
		End If
	Next
	
    'If Verification = False Then Exit Function
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function

' ----------------------  ���� -------------------------
Function  Verification()
	Dim dblW11, dblW12, dblW16, dblW14, dblW15, dblW13
	
	Verification = False
	
	With lgvspdData(TYPE_1)
		.Row = .MaxRows
		'1. W11 < W12
		.Col = C_W11 : dblW11 = UNICDbl(.Value)
		.Col = C_W12 : dblW12 = UNICDbl(.Value)
		
		If dblW11 < dblW12 Then
			Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, "(12)���� �غ�� ȯ�Ծ�", "(11)��λ� �غ�� �����ܾ�")                          <%'No data changed!!%>
			Exit Function
		End If
		
		'2. W11 < W14+W15
		.Col = C_W14 : dblW14 = UNICDbl(.Value)
		.Col = C_W15 : dblW15 = UNICDbl(.Value)
		If dblW11 < dblW14 + dblW15 Then
			Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "������[(W14)+(W15)]", "(11)��λ� �غ�� �����ܾ�")                          <%'No data changed!!%>
			Exit Function
		End If

		'3. W11 < W16
		.Col = C_W16 : dblW16 = UNICDbl(.Value)
		If dblW11 < dblW16 Then
			Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "������[(W16)]", "(11)��λ� �غ�� �����ܾ�")                          <%'No data changed!!%>
			Exit Function
		End If
		
		'4. W11 < W13
		.Col = C_W13 : dblW13 = UNICDbl(.Value)
		If dblW11 < dblW13 Then
			Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "(13)�غ�� ���� �����", "(11)��λ� �غ�� �����ܾ�")                          <%'No data changed!!%>
			Exit Function
		End If
	End With
	
	Verification = True	
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData

    Call SetToolbar("1100110100000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

    If lgvspdData(lgCurrGrid).MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If lgvspdData(lgCurrGrid).ActiveRow > 0 Then
			lgvspdData(lgCurrGrid).focus
			lgvspdData(lgCurrGrid).ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor lgvspdData(lgCurrGrid).ActiveRow, lgvspdData(lgCurrGrid).ActiveRow

			lgvspdData(lgCurrGrid).Col = C_W13
			lgvspdData(lgCurrGrid).Text = ""
    
			lgvspdData(lgCurrGrid).Col = C_W3
			lgvspdData(lgCurrGrid).Text = ""
			
			lgvspdData(lgCurrGrid).Col = C_W4
			lgvspdData(lgCurrGrid).Text = ""
			
			lgvspdData(lgCurrGrid).Col = C_W5
			lgvspdData(lgCurrGrid).Text = ""
			
			lgvspdData(lgCurrGrid).ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 

    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    If lgvspdData(lgCurrGrid).MaxRows = 2 Then
		ggoSpread.EditUndo                                                 
	End If
	ggoSpread.EditUndo 
    Call CheckReCalc()				' �Ѷ����� ��ҵǸ� ���� 
End Function


Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo, sW1_CD

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG
    
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If
 
	With lgvspdData(lgCurrGrid)	' ��Ŀ���� �׸��� 
			
		ggoSpread.Source = lgvspdData(lgCurrGrid)
			
		iRow = .ActiveRow
		lgvspdData(lgCurrGrid).ReDraw = False
		
		If iRow = .MaxRows Then Exit Function
		
		If .MaxRows = 0 Then	' ù InsertRow�� 1��+�հ��� 

			Call InsertFirstRow()
		
		Else
			.Row = iRow
			If iRow <1 Then iRow = 1
			.Col = C_W1_CD
			If Right(.Value, 1) = "9" Then	' �հ� �� 
				sW1_CD = Left(.Value, 1)
				.Row = iRow - 1
				ggoSpread.InsertRow iRow-1 , imRow 
				SetSpreadColor iRow, iRow + imRow - 1	
				
				Call SetDefaultVal(iRow, imRow, sW1_CD)			
			Else
				sW1_CD = Left(.Value, 1)
				.Row = iRow		
				ggoSpread.InsertRow ,imRow
				SetSpreadColor iRow+1, iRow + imRow
				
				Call SetDefaultVal(iRow+1, imRow, sW1_CD)
			End If
			
		End If
		
		lgvspdData(lgCurrGrid).ReDraw = True
	End With
	

	'Call CheckW7Status(lgCurrGrid)	' ������ ���� üũ 

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function



' GetREF ���� ���� �����µ� ȣ��� 
Function InsertTotalLine(Index)
	With lgvspdData(Index)
	
	ggoSpread.Source = lgvspdData(Index)
	
	If .MaxRows = 0 Then	' ���� �߰� 
		ggoSpread.InsertRow ,1
		
		.Row = 1
		.Col = C_SEQ_NO : .Text = SUM_SEQ_NO
		.Col = C_W9		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
		
		ggoSpread.SpreadLock C_W1, 1, C_W6, 1
	End If
	End With
End Function

' �׸��忡 SEQ_NO, TYPE �ִ� ���� 
Function SetDefaultVal(iRow, iAddRows, pW1_CD)
	
	Dim i, iSeqNo
	
	With lgvspdData(TYPE_1)	' ��Ŀ���� �׸��� 

		ggoSpread.Source = lgvspdData(TYPE_1)
	
		If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
			.Row = iRow
			.Col = C_W1_CD : .Value = pW1_CD
			MaxSpreadVal lgvspdData(TYPE_1), C_SEQ_NO, iRow
		Else
			iSeqNo = MaxSpreadVal(lgvspdData(TYPE_1), C_SEQ_NO, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
			
			For i = iRow to iRow + iAddRows -1
				.Row = i
				.Col = C_W1_CD	: .Value = pW1_CD
				.Col = C_SEQ_NO : .Value = iSeqNo : iSeqNo = iSeqNo + 1
			Next
		End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows

	With lgvspdData(lgCurrGrid)
		.focus
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		lDelRows = ggoSpread.DeleteRow
	End With
	
	Call CheckReCalc()				' �Ѷ����� ��ҵǸ� ���� 
	
End Function

' ���� 
Function CheckReCalc()
	Dim dblSum, sW1_CD
	
	With lgvspdData(lgCurrGrid)
		If .MaxRows = 0 Then Exit Function
		ggoSpread.Source = lgvspdData(lgCurrGrid)	
	
		.Row = .ActiveRow : .Col = C_W1_CD : sW1_CD = .Value
		
		Call SetSum2Col(C_W6, sW1_CD)
		Call SetSum2Col(C_W9, sW1_CD)
		Call SetSum2Col(C_W10, sW1_CD)
				
	End With
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '��: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'��: ȭ�� ���� %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'��:ȭ�� ����, Tab ���� %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function

'============================================  DB �＼�� �Լ�  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key   
        'strVal = strVal     & "&txtMaxRows="         & lgvspdData(lgCurrGrid).MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

Function ReDrawGRidColHead()
	' -- �׸��� �÷������ �� �����Ѵ�.
	Dim iRow, ret
	
	With lgvspdData(TYPE_1)
		.Redraw = False
		
		Call SetSpreadLock
		
		iRow = ReDrawW1("0", 1)
		iRow = ReDrawW1("1", iRow)

		.Row = iRow		
		.Col = C_W1		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2	
		.Col = C_W2_P	: .CellType = 1
		ret = .AddCellSpan(C_W1	, .Row, 3, 1)
		ggoSpread.SpreadLock C_W1, .Row, C_W10, .Row
				
		.Redraw = True
	End With
End Function

Function MakeCrLf(Byval iCnt)
	Dim i, sTmp
	If iCnt < 1 Then Exit Function
	For i = 1 to iCnt
		sTmp = sTmp & vbCrLf 
	Next
	MakeCrLf = sTmp
End Function

Function ReDrawW1(Byval pW1_CD, Byval pRow)
	Dim iRow, iMaxRows, iRowLoc , iRowSpanCnt, ret
	
	pRow = pRow 
	iRowLoc = pRow : iRowSpanCnt = 0

	With lgvspdData(TYPE_1)
		iMaxRows = .MaxRows
		.Row = pRow		: .Col = C_W1	

		Do Until False
			.Row = pRow	: .Col = C_W1_CD
			If Left(.Value, 1) = pW1_CD Then
				iRowSpanCnt = iRowSpanCnt + 1
			Else
				' -- �հ� 
				.Row = pRow - 1
				.Col = C_W2		: .CellType = 1	: .Text = "�հ�"	: .TypeHAlign = 2	
				.Col = C_W2_P	: .CellType = 1
				ggoSpread.SpreadLock C_W1, .Row, C_W10, .Row
				ret = .AddCellSpan(C_W1	, iRowLOc, 1, iRowSpanCnt)
				Exit Do
			End If
			pRow = pRow + 1
		Loop
		
		ReDrawW1 = pRow 

		.Row = iRowLoc
		If pW1_CD = "0" Then
			.value = "��" & MakeCrLf(iRowSpanCnt/2) & "��"
		Else
			.value = "��" & MakeCrLf(iRowSpanCnt/2) & "ä"
		End If
		If iRowSpanCnt > 1 Then
			.TypeEditMultiLine = True
		End If		
		.TypeHAlign = 2 : .TypeVAlign = 2
		
	End With
End Function
		
Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	If lgvspdData(TYPE_1).MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		
		Call ReDrawGRidColHead()
		
		' �������� ���� : ���ߵǸ� ���ȴ�.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 ����üũ : �׸��� �� 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1

			'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
			Call SetToolbar("1101110100000111")										<%'��ư ���� ���� %>

		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100110000000111")										<%'��ư ���� ���� %>
		End If
	Else
		Call SetToolbar("1100111100000111")										<%'��ư ���� ���� %>
	End If
	
	'Call SetSpreadTotalLine ' - �հ���� �籸�� 
	
	'lgvspdData(lgCurrGrid).focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow, lCol, lMaxRows, lMaxCols , i    
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    For i = TYPE_1 To TYPE_1	' ��ü �׸��� ���� 
    
		With lgvspdData(i)
	
			ggoSpread.Source = lgvspdData(i)
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
			
			' ----- 1��° �׸��� 
			For lRow = 1 To .MaxRows
    
		       .Row = lRow
		       .Col = 0
		    
		       Select Case .Text
		           Case  ggoSpread.InsertFlag                                      '��: Insert
		                                              strVal = strVal & "C"  &  Parent.gColSep
		           Case  ggoSpread.UpdateFlag                                      '��: Update
		                                              strVal = strVal & "U"  &  Parent.gColSep
		           Case  ggoSpread.DeleteFlag                                      '��: Delete
		                                              strDel = strDel & "D"  &  Parent.gColSep
		       End Select
		       
			  ' ��� �׸��� ����Ÿ ����     
			  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
					For lCol = C_SEQ_NO To lMaxCols
						.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
					Next
					strVal = strVal & Trim(.Text) &  Parent.gRowSep
			  End If  
			Next
		
		End With

		document.all("txtSpread" & CStr(i)).value =  strDel & strVal
		strDel = "" : strVal = ""
	Next

	'Frm1.txtSpread.value      = strDel & strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	frm1.txtHeadMode.value	  =  lgIntFlgMode
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' ���� ������ ���� ���� %>
	Call InitVariables
	
    Call MainQuery()
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
    strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key            
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 width=300>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white>��40ȣ(��) ��ȭ�ڻ� ��������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenRefMenu">�ҵ�ݾ��հ�ǥ��ȸ</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">�������</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/w3135ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
									<TD CLASS="TD5">���θ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">�Ű���</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="�Ű���" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=100% VALIGN=TOP>
							     <script language =javascript src='./js/w3135ma1_vspdData0_vspdData0.js'></script>
							    </TD>
							</TR>
<!--							 <TR>
							     <TD width="100%" HEIGHT=50%>
							     <script language =javascript src='./js/w3135ma1_vspdData1_vspdData1.js'></script>
							    </TD>
							</TR> -->
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
		        <TR>
				        <TD WIDTH=10>&nbsp;</TD>
				         <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><����>ȭ�󼺿�ȭ�ڻ�</LABEL>&nbsp;
				             <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><����>ȭ�󼺿�ȭ��ä</LABEL>&nbsp;
				        </TD>
				            
		
                </TR>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('PRINT')"   Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread0" tag="24" style="display:'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHeadMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

