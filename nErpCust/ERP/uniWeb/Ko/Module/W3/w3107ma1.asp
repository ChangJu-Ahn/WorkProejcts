
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �� ���� ���� 
'*  3. Program ID           : W3107MA1
'*  4. Program Name         : W3107MA1.asp
'*  5. Program Desc         : ������� �Է� 
'*  6. Modified date(First) : 2005/01/05
'*  7. Modified date(Last)  : 2006/01/23
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : HJO
'* 10. Comment              : ERP�Ϲݰ������� ó�� 
'								������ ��ȸ�� ���� ��Ͽ� �ݾװ� ���並 �ݾװ� �ŷ�ó�� �����Է� 8������ ���ε� �����ε���Ÿ�� 
'								�������� ��ձݺҷ����� (������ ��������Ÿ�� ������ ���ݵ�� ����Ÿ �������� 
' ����� ó�����...
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID = "W3107MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID2 = "W3107MB2.asp"											 '��: �����Ͻ� ���� ASP�� 

Const TAB1 = 1																	'��: Tab�� ��ġ 
Const TAB2 = 2


' -- �׸��� �÷� ���� 
Dim C_SEQ_NO1
Dim C_W1
Dim C_W1_BT
Dim C_W1_NM
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W7_NM
Dim C_W8
Dim C_W9
Dim C_W9_NM
Dim C_W10
Dim C_W11

Dim C_SEQ_NO2
Dim C_W12
Dim C_W12_BT
Dim C_W12_NM
Dim C_W13
Dim C_W14
Dim C_W15
Dim C_W16
Dim C_W17
Dim C_W18
Dim C_W19

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgFISC_START_DT, lgFISC_END_DT
Dim gCurrGrid

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
    gCurrGrid = 1
	
	C_SEQ_NO1 = 1
	C_W1 = 2
	C_W1_BT = 3
	C_W1_NM = 4
	C_W2 = 5
	C_W3 = 6
	C_W4 = 7
	C_W5 = 8
	C_W6 = 9
	C_W7 = 10
	C_W7_NM = 11
	C_W8 = 12
	C_W9 = 13
	C_W9_NM = 14
	C_W10 = 15
	C_W11 = 16

	C_SEQ_NO2 = 1
	C_W12 = 2
	C_W12_BT = 3
	C_W12_NM = 4
	C_W13 = 5
	C_W14 = 6
	C_W15 = 7
	C_W16 = 8
	C_W17 = 9
	C_W18 = 10
	C_W19 = 11

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
    
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  �Ű��� �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))

'	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1050' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'    Call SetCombo2(frm1.cboAPP_TYPE ,lgF0  ,lgF1  ,Chr(11))

'    Call SetCombo2(frm1.cboCONF_TYPE ,"Y" & Chr(11) & "N" & Chr(11)   ,"����" & Chr(11) & "������" & Chr(11)   ,Chr(11))
   
End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	
    Call initSpreadPosVariables()  

	' 1�� �׸��� 
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
	   'patch version
	    ggoSpread.Spreadinit "V20041222",,parent.gAllowDragDropSpread    
	    
		.ReDraw = false
	
	    .MaxCols = C_W11 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
		       
	    .MaxRows = 0
	    ggoSpread.ClearSpreadData
	    
		'����� 2�ٷ�    
	    .ColHeaderRows = 2
	    'Call AppendNumberPlace("6","3","2")
	
	    ggoSpread.SSSetEdit		C_SEQ_NO1,	"����"		, 10,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W1,		"(1)�����ڵ�"	, 10,,,50,1	
        ggoSpread.SSSetButton 	C_W1_BT       		'4
		ggoSpread.SSSetEdit		C_W1_NM,	"(1)������"	, 15,,,50,1	
	    ggoSpread.SSSetDate		C_W2,		"(2)��������"      , 10, 2, parent.gDateFormat
	    ggoSpread.SSSetEdit		C_W3,		"(3)�ŷ�ó" & vbCrLf & "(ä�ǳ���)", 10,,,50,1
	    ggoSpread.SSSetDate		C_W4,		"(4)��������"      , 10, 2, parent.gDateFormat
	    ggoSpread.SSSetFloat	C_W5,		"(5)�ݾ�", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W6,		"(6)�ݾ�", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetCombo	C_W7,		"(7)�ú���", 10
	    ggoSpread.SSSetCombo	C_W7_NM,	"(7)�ú���", 10
	    ggoSpread.SSSetFloat	C_W8,		"(8)�ݾ�", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetCombo	C_W9,		"(9)�ú���", 10
	    ggoSpread.SSSetCombo	C_W9_NM,	"(9)�ú���", 10
	    ggoSpread.SSSetFloat	C_W10,		"(10)�������ܾ�", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetEdit		C_W11,		"(11)��ջ���"	, 10,,,50,1	

	    ret = .AddCellSpan(0, -1000, 1, 2)
	    ret = .AddCellSpan(1, -1000, 1, 2)
	    ret = .AddCellSpan(2, -1000, 3, 2)
	    ret = .AddCellSpan(5, -1000, 1, 2)
	    ret = .AddCellSpan(6, -1000, 1, 2)
	    ret = .AddCellSpan(7, -1000, 1, 2)
	    ret = .AddCellSpan(8, -1000, 1, 2)
	    ret = .AddCellSpan(9, -1000, 3, 1)
	    ret = .AddCellSpan(12, -1000, 3, 1)
	    ret = .AddCellSpan(15, -1000, 1, 2) 
	    ret = .AddCellSpan(16, -1000, 1, 2) 

	    
	    ' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W6
		.Text = "������ݻ��"
		.Col = C_W8
		.Text = "���ձݰ��"
	
		' �ι�° ��� ��� ���� 
		.Row = -999
		.Col = C_W6
		.Text = "(6)�ݾ�"
		.Col = C_W7
		.Text = "(7)�ú���"
		.Col = C_W7_NM
		.Text = "(7)�ú���"
		.Col = C_W8
		.Text = "(8)�ݾ�"
		.Col = C_W9
		.Text = "(9)�ú���"
		.Col = C_W9_NM
		.Text = "(9)�ú���"
		.rowheight(-999) = 20	' ���� ������ 

		Call ggoSpread.MakePairsColumn(C_W1,C_W1_BT)
	
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W7,C_W7,True)
		Call ggoSpread.SSSetColHidden(C_W9,C_W9,True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO1,C_SEQ_NO1,True)

		Call InitSpreadComboBox()
					
		.ReDraw = true
		
	    'Call SetSpreadLock 
    
    End With

 	' -----  2�� �׸��� 
	With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2	
	   'patch version
	    ggoSpread.Spreadinit "V20041222_2",,parent.gAllowDragDropSpread    
	    
		.ReDraw = false
	    
	    .MaxCols = C_W19 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
		       
	    .MaxRows = 0
	    ggoSpread.ClearSpreadData
	
		'����� 2�ٷ�    
	    'Call AppendNumberPlace("6","3","2")
	
	    ggoSpread.SSSetEdit		C_SEQ_NO2,	"����", 10,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W12,		"(12)�����ڵ�", 10,,,50,1
        ggoSpread.SSSetButton 	C_W12_BT       		'4
		ggoSpread.SSSetEdit		C_W12_NM,	"(12)������", 10,,,50,1
	    ggoSpread.SSSetDate		C_W13,		"(13)��������"      , 10, 2, parent.gDateFormat '6    
		ggoSpread.SSSetEdit		C_W14,		"(14)ä�ǳ���" & vbCrLf & "(�ŷ�ó)", 10,,,50,1
		ggoSpread.SSSetDate		C_W15,		"(15)��������"      , 10, 2, parent.gDateFormat '6    
	    ggoSpread.SSSetFloat	C_W16,		"(16)�ݾ�",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
	    ggoSpread.SSSetFloat	C_W17,		"(17)������αݾ�" ,		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0","" 
		ggoSpread.SSSetFloat	C_W18,		"(18)�ܾ�",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetEdit		C_W19,		"(19)������" & vbCrLf & "��ջ���", 10,,,50,1

	    ret = .AddCellSpan(2, -1000, 3, 1)

		.rowheight(-1000) = 25	' ���� ������ 
		Call ggoSpread.MakePairsColumn(C_W12,C_W12_BT)
	
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO2,C_SEQ_NO2,True)
		
		.ReDraw = true
		
	    Call SetSpreadLock 
    
    End With
       
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
    
	Call GetFISC_DATE

End Sub

Sub InitSpreadComboBox()

    Dim IntRetCD1

	' �ú��� ���� 
	IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1050' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W7
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W7_NM
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W9
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W9_NM

	End If
		  
End Sub

Function OpenAccount(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strWhere

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"					<%' �˾� ��Ī %>
	arrParam(1) = "TB_ACCT_MATCH"					<%' TABLE ��Ī %>
	

	If iWhere = 1 then
		frm1.vspdData.Col = C_W1
	    arrParam(2) = frm1.vspdData.Text		<%' Code Condition%>
	ElseIf iWhere = 2 then
		frm1.vspdData.Col = C_W12
	    arrParam(2) = frm1.vspdData2.Text		<%' Code Condition%>
	ElseIf iWhere = 3 then
	    arrParam(2) = frm1.txtACCT_CD.value		<%' Code Condition%>
	End If
	arrParam(3) = ""							<%' Name Cindition%>

	strWhere = " MATCH_CD = '07'"
	strWhere = strWhere & " AND CO_CD = '" & frm1.txtCO_CD.value & "' "
	strWhere = strWhere & " AND FISC_YEAR = '" & frm1.txtFISC_YEAR.text & "' "
	strWhere = strWhere & " AND REP_TYPE = '" & frm1.cboREP_TYPE.value & "' "

	arrParam(4) = strWhere							<%' Where Condition%>
	arrParam(5) = "����"						<%' �����ʵ��� �� ��Ī %>
	
    arrField(0) = "ED5" & Chr(11) & "ACCT_GP_CD" & Chr(11)					<%' Field��(0)%>
    arrField(1) = "ED10" & Chr(11) & "dbo.ufn_GetCodeName('W1085', ACCT_GP_CD)" & Chr(11)					<%' Field��(1)%>
    arrField(2) = "ED7" & Chr(11) & "ACCT_CD" & Chr(11)					<%' Field��(2)%>
    arrField(3) = "ED15" & Chr(11) & "ACCT_NM" & Chr(11)					<%' Field��(3)%>
    
    arrHeader(0) = "��ǥ�����ڵ�"					<%' Header��(0)%>
    arrHeader(1) = "��ǥ������"						<%' Header��(1)%>
    arrHeader(2) = "�����ڵ�"					<%' Header��(2)%>
    arrHeader(3) = "������"						<%' Header��(3)%>
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAccount(arrRet,iWhere)
	End If	
	
End Function

Function SetAccount(byval arrRet,Byval iWhere)
    With frm1
		If iWhere = 1 Then 'Spread1(Condition)
			.vspdData.Col = C_W1
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_W1_NM
			.vspdData.Text = arrRet(1)
		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.UpdateRow frm1.vspdData.ActiveRow
			lgBlnFlgChgValue = True
		ElseIf iWhere = 2 Then 'Spread2(Condition)
			.vspdData2.Col = C_W12
			.vspdData2.Text = arrRet(0)
			.vspdData2.Col = C_W12_NM
			.vspdData2.Text = arrRet(1)
		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.UpdateRow frm1.vspdData.ActiveRow
			lgBlnFlgChgValue = True
		ElseIf iWhere = 3 Then 'Header
			.txtACCT_CD.Value = arrRet(0)
			.txtACCT_NM.Value = arrRet(1)
		End If
	End With
End Function

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    .vspdData2.ReDraw = False
        
    ggoSpread.Source = frm1.vspdData
        
	ggoSpread.SSSetRequired C_W1, -1, -1
	ggoSpread.SSSetProtected C_W1_NM, -1, -1
	ggoSpread.SSSetRequired C_W2, -1, -1
	ggoSpread.SSSetProtected C_W5, -1, -1
	ggoSpread.SSSetProtected C_W10, -1, -1
    .vspdData.ReDraw = True


    ggoSpread.Source = frm1.vspdData2	

	ggoSpread.SSSetRequired C_W12, -1, -1
	ggoSpread.SSSetProtected C_W12_NM, -1, -1
	ggoSpread.SSSetRequired C_W13, -1, -1
	ggoSpread.SSSetRequired C_W16, -1, -1
	ggoSpread.SSSetProtected C_W18, -1, -1
    .vspdData2.ReDraw = True
	
    End With
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim iRow
    With frm1

	If gCurrGrid = 1 Then
		.vspdData.ReDraw = False
 
		ggoSpread.Source = frm1.vspdData

		For iRow = pvStartRow To pvEndRow
		    .vspdData.Row = iRow
			if iRow <>  .vspdData.MaxRows then
				ggoSpread.SSSetRequired C_W1, iRow, iRow
				ggoSpread.SSSetProtected C_W1_NM, iRow, iRow
				ggoSpread.SSSetRequired C_W2, iRow, iRow
				ggoSpread.SSSetProtected C_W5, iRow, iRow
				ggoSpread.SSSetProtected C_W10, iRow, iRow

			    .vspdData.Col = C_W6
			    If UNICdbl(.vspdData.Text) > 0 Then
					ggoSpread.SSSetRequired C_W7, iRow, iRow
					ggoSpread.SSSetRequired C_W7_NM, iRow, iRow
			    End If
			    .vspdData.Col = C_W8
			    If UNICdbl(.vspdData.Text) > 0 Then
					ggoSpread.SSSetRequired C_W9, iRow, iRow
					ggoSpread.SSSetRequired C_W9_NM, iRow, iRow
			    End If
		    End If
	
		   	.vspdData.col = C_SEQ_NO1	 
		
		    if .vspdData.text = "999999" and .vspdData.MaxRows > 0 then
	'		    ggoSpread.SpreadLock C_W1_BT, iRow, C_W1_BT, iRow
			    ggoSpread.SSSetProtected -1 , iRow, iRow
		
		    End If
	    Next
	
		.vspdData.ReDraw = True
    Else
		.vspdData2.ReDraw = False
 
		ggoSpread.Source = frm1.vspdData2

		For iRow = pvStartRow To pvEndRow
		    .vspdData2.Row = iRow
			if iRow <>  .vspdData2.MaxRows then
				ggoSpread.SSSetRequired C_W12, iRow, iRow
				ggoSpread.SSSetProtected C_W12_NM, iRow, iRow
				ggoSpread.SSSetRequired C_W13, iRow, iRow
				ggoSpread.SSSetRequired C_W16, iRow, iRow
				ggoSpread.SSSetProtected C_W18, iRow, iRow

		    End If
	
		   	.vspdData2.col = C_SEQ_NO2	 
		
		    if .vspdData2.text = "999999" and .vspdData2.MaxRows > 0 then
	'		    ggoSpread.SpreadLock C_W12_BT, iRow, C_W12_BT, iRow
			    ggoSpread.SSSetProtected -1 , iRow, iRow
		
		    End If
	    Next
		    
		.vspdData2.ReDraw = True    
    End If
    
    End With
End Sub

Sub SetSpreadTotalLine()
	Dim iTmpCurrGrid
	Dim iRow

	iTmpCurrGrid = gCurrGrid
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		If .MaxRows > 0 Then
			.Row = .MaxRows
			Call .AddCellSpan(C_W1, .MaxRows, 3, 1) 
			.Col = C_W1		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
			gCurrGrid = 1
			Call SetSpreadColor(1, .MaxRows)
		End If
	End With

	ggoSpread.Source = frm1.vspdData2
	With frm1.vspdData2
		If .MaxRows > 0 Then
			.Row = .MaxRows
			Call .AddCellSpan(C_W12, .MaxRows, 3, 1) 
			.Col = C_W12		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
			gCurrGrid = 2
			Call SetSpreadColor(1, .MaxRows)
		End If
	End With
	
	gCurrGrid = iTmpCurrGrid
End Sub 

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO1	= iCurColumnPos(1)
            C_W1	= iCurColumnPos(2)
            C_W1_BT	= iCurColumnPos(3)
            C_W1_NM	= iCurColumnPos(4)
            C_W2	= iCurColumnPos(5)
            C_W3	= iCurColumnPos(6)
            C_W4	= iCurColumnPos(7)
            C_W5	= iCurColumnPos(8)
            C_W6	= iCurColumnPos(9)
            C_W7	= iCurColumnPos(10)
            C_W7_NM	= iCurColumnPos(11)
            C_W8	= iCurColumnPos(12)
            C_W9	= iCurColumnPos(13)
            C_W9_NM	= iCurColumnPos(14)
            C_W10	= iCurColumnPos(15)
            C_W11	= iCurColumnPos(16)
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO2	= iCurColumnPos(1)
            C_W12		= iCurColumnPos(2)
            C_W12_BT	= iCurColumnPos(3)
            C_W12_NM	= iCurColumnPos(4)
            C_W13		= iCurColumnPos(5)
            C_W14		= iCurColumnPos(6)
            C_W15		= iCurColumnPos(7)
            C_W16		= iCurColumnPos(8)
            C_W17		= iCurColumnPos(9)
            C_W18		= iCurColumnPos(10)
            C_W19		= iCurColumnPos(11)
	End Select    

End Sub

'============================== ���۷��� �Լ�  ========================================

Sub GetFISC_DATE()	' ��û������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		lgFISC_START_DT = CDate(lgF0)
	Else
		lgFISC_START_DT = ""
	End if

    If lgF1 <> "" Then 
		lgFISC_END_DT = CDate(lgF1)
	Else
		lgFISC_END_DT = ""
	End if
	
End Sub

Function GetRef()	' ����������ձݰ������� ��ũ Ŭ���� 
    Dim IntRetCD , i
    Dim sFiscYear, sRepType, sCoCd
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value 
    
    
	If gSelframeFlg = TAB1 Then Exit Function
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
'	ggoSpread.Source = frm1.vspdData2
'	If ggoSpread.SSCheckChange = True Then
'		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
'		If IntRetCD = vbNo Then
'			Exit Function
'		End If
'	End If
   'add logic check about data exist or not
   Call CommonQueryRs(" count(seq_no) "," TB_BED_DEBT_CON "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   If lgF0>0 Then
		IntRetCD = DisplayMsgBox("W30010", parent.VB_YES_NO, "X", "X")			    <%'exist data%>
		If IntRetCD = vbNo Then
			Exit Function
		Else
			frm1.txtMode.value="MD"
			Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		End If
	End If
	
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
'    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables													<%'Initializes local global variables%>
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery2()
    	
End Function

Function OpenRef()	'������ ��ȸ 

    Dim arrRet
    Dim arrParam(4)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
	Dim arrRowVal
    Dim arrColVal, lLngMaxRow
    Dim iDx
    
    If IsOpenPop = True Then Exit Function

	If gSelframeFlg = TAB2 Then Exit Function

    IsOpenPop = True

   ' iCalledAspName = AskPRAspName("W3107RA1")
    
    

	arrParam(0) = frm1.txtCO_CD.Value
	arrParam(1) = frm1.txtCO_NM.Value		
	arrParam(2) = frm1.txtFISC_YEAR.Text		
	arrParam(3) = frm1.cboREP_TYPE.Value		

    arrRet = window.showModalDialog("W3107RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0,0) <> "" Then
		arrRowVal = Split(arrRet(0,0), Parent.gRowSep)                                 '��: Split Row    data
		lLngMaxRow = UBound(arrRowVal)
		
		For iDx = 1 To lLngMaxRow
		    arrColVal = Split(arrRowVal(iDx-1), Parent.gColSep)    

			Call FncInsertRow(1)
			Frm1.vspdData.Col	= C_W3
			Frm1.vspdData.Text	= arrColVal(C_W3)
			Frm1.vspdData.Col	= C_W5
			Frm1.vspdData.Text	= arrColVal(C_W5)
			Call CheckReCalc(C_W5, Frm1.vspdData.ActiveRow)
		Next
		
	End IF
    
    IsOpenPop = False
    
    
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
    'End If

    arrRet = window.showModalDialog("../W5/W5105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

'====================================== �� �Լ� =========================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
    gCurrGrid = 1
	'Call ShowTabLInk(TAB1)


End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
    gCurrGrid = 2
	Call ShowTabLInk(TAB2)

End Function


' -- �Ǻ� ��ũ �����ֱ� 
Function ShowTabLInk(pType)
	Dim pObj1, pObj2, i
	Set pObj1 = document.all("myTabRef")
	
	'For i = 0 To 1
	'	pObj1(i).style.display = "none"
	'Next	
	'pObj1(pType-1).style.display = ""
	
	pObj1.style.display = ""
End Function



'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100111100000111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
	Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData 
	gSelframeFlg = TAB1
	' �������� üũȣ�� 
	
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
	Call GetFISC_DATE
End Sub


'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' �� Template ȭ�鿡���� ���� ������, �޺�(Name)�� ����Ǹ� �޺�(Code, Hidden)�� ��������ִ� ���� 
	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case  C_W7
				.Col = Col
				intIndex = .Value
				.Col = C_W7_NM
				.Value = intIndex	
			Case  C_W7_NM
				.Col = Col
				intIndex = .Value
				.Col = C_W7
				.Value = intIndex		
			Case  C_W9
				.Col = Col
				intIndex = .Value
				.Col = C_W9_NM
				.Value = intIndex	
			Case  C_W9_NM
				.Col = Col
				intIndex = .Value
				.Col = C_W9
				.Value = intIndex		
		End Select
	End With

End Sub

Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim arrVal
	Dim iDblW5, iDblW6, iDblW8
	
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If

    If Col = C_W1 Then
		frm1.vspdData.Col = C_W1

		If Len(frm1.vspdData.Text) > 0 Then
			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_W1

'				If CommonQueryRs("ACCT_NM", " TB_WORK_6 (NOLOCK)" , "ACCT_CD = '" & Frm1.vspdData.Text &"' AND ACCT_CD IN (SELECT ACCT_CD FROM TB_ACCT_MATCH (NOLOCK) WHERE MATCH_CD = '7')", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			If CommonQueryRs("MINOR_NM", " B_MINOR " , "MAJOR_CD = 'W1085' AND MINOR_CD = '" & Frm1.vspdData.Text &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		    	frm1.vspdData.Col	= C_W1_NM
		    	arrVal				= Split(lgF0, Chr(11))
				frm1.vspdData.Text	= arrVal(0)
			Else
				frm1.vspdData.Text	= ""
				frm1.vspdData.Col	= C_W1_NM
				frm1.vspdData.Text	= ""
			End If
		Else
			frm1.vspdData.Col = C_W1_NM
			frm1.vspdData.Text = ""
		End If
	Else
		'(5)�ݾ��� (6)+(8)�̾�� �Ѵ�. 
		With Frm1.vspdData
			If Col = C_W6 Or Col = C_W8 Then
			    .Col = C_W6 :	iDblW6 = unicdbl(.text)
				
			    .Col = C_W8 :	iDblW8 = unicdbl(.text)
			    
			    .Col = C_W5 :	iDblW5 = unicdbl(iDblW6+iDblW8)
			    
			    '(5) < (6) �̸� ���� (�޼��� WC0010)
'			    If iDblW5 < iDblW6 Then
'			        Call DisplayMsgBox("WC0010", "X", "(6)������ݻ��ݾ�", "(5)�ݾ�")
'				    .Col = Col :	.text = 0
'			    '(5) < (8) �̸� ���� (�޼��� WC0010)
'			    ElseIf iDblW5 < iDblW8 Then
'			        Call DisplayMsgBox("WC0010", "X", "(8)���ձݰ��ݾ�", "(5)�ݾ�")
'				    .Col = Col :	.text = 0
			    '(5) < (6) + (8) �̸� ���� (�޼��� WC0012)
			    
			'    If iDblW5 <> (iDblW6 + iDblW8) Then
			 '       Call DisplayMsgBox("WC0012", "X", "(6) + (8)", "(5)�ݾ�")
			'	    .Col = Col :	.text = 0
				
					If iDblW6 > 0 Then
					    ggoSpread.Source = frm1.vspdData
						ggoSpread.SSSetRequired C_W7, Row, Row
						ggoSpread.SSSetRequired C_W7_NM, Row, Row
					End If
					If iDblW8 > 0 Then
					    ggoSpread.Source = frm1.vspdData
						ggoSpread.SSSetRequired C_W9, Row, Row
						ggoSpread.SSSetRequired C_W9_NM, Row, Row
					End If

			End IF
		End With
		
		Call CheckReCalc(Col, Row)
	End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    ggoSpread.UpdateRow frm1.vspdData.MaxRows

End Sub

Sub CheckReCalc(ByVal Col , ByVal Row)
	Dim iDblAmt, iDblW6, iDblW8
	Dim dblSum
	Dim iStrCd1, iStrCd2

	With Frm1.vspdData
		'10 (10)�ܾ��� (7),(9)�ú����� ����ڰ� "����"�� ������ ��� (6)+(8)�ݾ��� ����Ѵ�.
		If  Col = C_W6 Or Col = C_W7 Or Col = C_W7_NM Or Col = C_W8 Or Col = C_W9 Or Col = C_W9_NM Then
		    .Col = C_W6 :	iDblW6 = unicdbl(.text)			
		    .Col = C_W8 :	iDblW8 = unicdbl(.text)
		    
		    .Col = C_W7 :	iStrCd1 = .text
	
		    .Col = C_W9 :	iStrCd2 = .text
			.Col =C_W5	:	.text=iDblW6 +iDblW8
			iDblAmt = 0
			If iStrCd1 = "2" Then
			    iDblAmt = iDblAmt + iDblW6
			End If
			If iStrCd2 = "2" Then
			    iDblAmt = iDblAmt + iDblW8
			End If			
		    .Col = C_W10 :	.text = iDblAmt
			
		End If
	End With

	With Frm1.vspdData
		If .MaxRows > 0 Then
			'dblSum = FncSumSheet(Frm1.vspdData, C_W5, 1, .MaxRows - 1, true, .MaxRows, C_W5, "V")	' �հ� 
			dblSum = FncSumSheet(Frm1.vspdData, C_W6, 1, .MaxRows - 1, true, .MaxRows, C_W6, "V")	' �հ� 
			dblSum = FncSumSheet(Frm1.vspdData, C_W8, 1, .MaxRows - 1, true, .MaxRows, C_W8, "V")	' �հ� 
			dblSum = FncSumSheet(Frm1.vspdData, C_W10, 1, .MaxRows - 1, true, .MaxRows, C_W10, "V")	' �հ� 
		End If
	End With


End Sub

Sub vspdData2_Change(ByVal Col , ByVal Row )
	Dim arrVal
	Dim iDblW16, iDblW17
	
    Frm1.vspdData2.Row = Row
    Frm1.vspdData2.Col = Col

    If Frm1.vspdData2.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData2.text) < UNICDbl(Frm1.vspdData2.TypeFloatMin) Then
         Frm1.vspdData2.text = Frm1.vspdData2.TypeFloatMin
      End If
    End If
	
    If Col = C_W12 Then
		frm1.vspdData2.Col = C_W12

		If Len(frm1.vspdData2.Text) > 0 Then
			frm1.vspdData2.Row = Row
			frm1.vspdData2.Col = C_W12

'				If CommonQueryRs("ACCT_NM", " TB_WORK_6 (NOLOCK)" , "ACCT_CD = '" & Frm1.vspdData2.Text &"' AND ACCT_CD IN (SELECT ACCT_CD FROM TB_ACCT_MATCH (NOLOCK) WHERE MATCH_CD = '7')", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			If CommonQueryRs("MINOR_NM", " B_MINOR " , "MAJOR_CD = 'W1085' AND MINOR_CD = '" & Frm1.vspdData2.Text &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		    	frm1.vspdData2.Col	= C_W12_NM
		    	arrVal				= Split(lgF0, Chr(11))
				frm1.vspdData2.Text	= arrVal(0)
			Else
				frm1.vspdData2.Text	= ""
				frm1.vspdData2.Col	= C_W12_NM
				frm1.vspdData2.Text	= ""
			End If
		Else
			frm1.vspdData2.Col = C_W12_NM
			frm1.vspdData2.Text = ""
		End If
	ElseIf Col = C_W16 Or Col = C_W17 Then
	    Frm1.vspdData2.Col = C_W16
	    iDblW16 = unicdbl(Frm1.vspdData2.text)
		
	    Frm1.vspdData2.Col = C_W17
	    iDblW17 = unicdbl(Frm1.vspdData2.text)
		
	    '(16) < (17) �̸� ���� (�޼��� WC0010)
	    If iDblW16 < iDblW17 Then
	        Call DisplayMsgBox("WC0010", "X", "(17)������αݾ�", "(16)�ݾ�")
		    Frm1.vspdData2.Col = Col
		    Frm1.vspdData2.text = 0
	    End If
		Call CheckReCalc2(Col, Row)
	End If
	
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row
    ggoSpread.UpdateRow frm1.vspdData2.MaxRows

End Sub


Sub CheckReCalc2(ByVal Col , ByVal Row)
	Dim iDblAmt
	Dim dblSum

	With Frm1.vspdData2
		'(18) �ܾ��� (16) - (17) �� ����Ͽ� ����Ѵ�.
		If Col = C_W16 Or Col = C_W17 Or Col = C_W18 Then
		    .Col = C_W16 :	iDblAmt = unicdbl(.text)
			
		    .Col = C_W17 :	iDblAmt = iDblAmt - unicdbl(.text)
		    
		    .Col = C_W18 :	.text = iDblAmt
		
		End If
	End With

	With Frm1.vspdData2
		If .MaxRows > 0 Then
			dblSum = FncSumSheet(Frm1.vspdData2, C_W16, 1, .MaxRows - 1, true, .MaxRows, C_W16, "V")	' �հ� 
			dblSum = FncSumSheet(Frm1.vspdData2, C_W17, 1, .MaxRows - 1, true, .MaxRows, C_W17, "V")	' �հ� 
			dblSum = FncSumSheet(Frm1.vspdData2, C_W18, 1, .MaxRows - 1, true, .MaxRows, C_W18, "V")	' �հ� 
		End If
	End With
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
    	Exit Sub
       ggoSpread.Source = frm1.vspdData
      
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
      
       Exit Sub
    End If

	frm1.vspdData.Row = Row
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData2
   
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
    	Exit Sub
       ggoSpread.Source = frm1.vspdData2
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
      
       Exit Sub
    End If

	frm1.vspdData2.Row = Row
End Sub


Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	gCurrGrid = 1
	ggoSpread.Source = Frm1.vspdData
End Sub    

Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	gCurrGrid = 2
	ggoSpread.Source = Frm1.vspdData2
End Sub  

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim strTemp
Dim intPos1
   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_W1_BT Then
        .Col = Col - 1
        .Row = Row
        
        Call OpenAccount(1)
        
    End If
    
    End With
      
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim strTemp
Dim intPos1
   
	With frm1.vspdData2 
	
    ggoSpread.Source = frm1.vspdData2
   
    If Row > 0 And Col = C_W12_BT Then
        .Col = Col - 1
        .Row = Row
        
        Call OpenAccount(2)
        
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
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	ggoSpread.Source = frm1.vspdData2
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	If blnChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call InitVariables													<%'Initializes local global variables%>
'    Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
	FncQuery = True
    
End Function

Function FncSave() 
    Dim blnChange, i
    Dim bRtn1, bRtn2
    blnChange = False
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    On Error Resume Next                                                   <%'��: Protect system from crashing%>    

<%  '-----------------------
    'Precheck area
    '----------------------- %>
	ggoSpread.Source = frm1.vspdData
	bRtn1 = ggoSpread.SSCheckChange


	ggoSpread.Source = frm1.vspdData2
	bRtn2 = ggoSpread.SSCheckChange

	If bRtn1 <> True And bRtn2 <> True Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

	IF bRtn1 = True Then
		ggoSpread.Source = frm1.vspdData
		If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
		      Exit Function
		End If    
	End If
	IF bRtn2 = True Then
		ggoSpread.Source = frm1.vspdData2
		If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
		      Exit Function
		End If    
	End If    
	 
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
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

    Call SetToolbar("1100111100000111")										<%'��ư ���� ���� %>

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
	Dim iActiveRow
	
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

	If gCurrGrid = 1 Then
	    ggoSpread.Source = Frm1.vspdData
		With Frm1.vspdData
		    If .MaxRows < 1 Or .ActiveRow = .MaxRows Then
		       Exit Function
		    End If
		
			If .ActiveRow > 0 Then
				iActiveRow = .ActiveRow
				.focus
				.ReDraw = False
			
				ggoSpread.CopyRow
				SetSpreadColor iActiveRow, iActiveRow + 1
		
'				.Col = C_W10
'				.Text = ""
						
				.ReDraw = True
			End If
		End With
		Call CheckReCalc(C_W5, iActiveRow + 1)
	Else
	    ggoSpread.Source = Frm1.vspdData2
		With Frm1.vspdData2
		    If .MaxRows < 1 Or .ActiveRow = .MaxRows Then
		       Exit Function
		    End If
		
			If .ActiveRow > 0 Then
				iActiveRow = .ActiveRow
				.focus
				.ReDraw = False
			
				ggoSpread.CopyRow
				SetSpreadColor iActiveRow, iActiveRow + 1
		
'				.Col = C_W18
'				.Text = ""
				
				.ReDraw = True
			End If
		End With
		Call CheckReCalc2(C_W16, iActiveRow + 1)
	End If
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    Dim lDelRows, dblSum 

	If gCurrGrid = 1 Then
		With Frm1.vspdData
			.focus
			lDelRows = .ActiveRow
			
			ggoSpread.Source = Frm1.vspdData
			If .MaxRows <= 0 Then
				Exit Function
			ElseIf CheckTotRow(Frm1.vspdData, .ActiveRow) = True Then
				MsgBox "�հ�� ������ �� �����ϴ�.", vbCritical
				Exit Function
			Else
				.focus
			ggoSpread.Source = Frm1.vspdData
				lDelRows = ggoSpread.EditUndo(lDelRows)
				lgBlnFlgChgValue = True
				lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
				If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
			End If
			
		End With

		Call CheckReCalc(C_W5, frm1.vspdData.ActiveRow)
		ggoSpread.UpdateRow frm1.vspdData.MaxRows

	Else
		With Frm1.vspdData2
			.focus
			lDelRows = .ActiveRow
			
			ggoSpread.Source = Frm1.vspdData2
			If .MaxRows <= 0 Then
				Exit Function
			ElseIf CheckTotRow(Frm1.vspdData2, .ActiveRow) = True Then
				MsgBox "�հ�� ������ �� �����ϴ�.", vbCritical
				Exit Function
			Else
				lDelRows = ggoSpread.EditUndo(lDelRows)
				lgBlnFlgChgValue = True
				lDelRows = CheckLastRow(Frm1.vspdData2, lDelRows)
				If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
			End If
			
		End With

		Call CheckReCalc2(C_W16, frm1.vspdData2.ActiveRow)
		ggoSpread.UpdateRow frm1.vspdData2.MaxRows

	End If
	
End Function

' -- �հ� ������ üũ(Header Grid)
Function CheckTotRow(Byref pObj, Byval pRow) 
	CheckTotRow = False
	pObj.Col = C_SEQ_NO1 : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If pObj.Text = 999999 Then	 ' �հ� �� 
		CheckTotRow = True
	End If
End Function

' -- Detail Data�� �����ϴ��� üũ 
Function CheckLastRow(Byref pObj, Byval pRow) 
	Dim iCnt, iRow, iMaxRow
	CheckLastRow = 0
	iCnt = 0
	
	With pObj

		For iRow = 1 To .MaxRows
			.Row = iRow : .Col = 0
			If .Text <> ggoSpread.DeleteFlag Then
				iCnt = iCnt + 1
				iMaxRow = iRow
			End If
		Next
		.Col = C_SEQ_NO1	:	.Row = iMaxRow
		If .Text = 999999 and iCnt = 1 Then
			CheckLastRow = iMaxRow
		End If
	End With
	
End Function


Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo
    Dim ret

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
   
	With frm1	
	
		If gCurrGrid = 1 Then
		
			.vspdData.focus
			ggoSpread.Source = .vspdData
		
			'.vspdData.ReDraw = False
			iSeqNo = .vspdData.MaxRows+1
		
			if 	.vspdData.MaxRows = 0 then
			
			     ggoSpread.InsertRow  imRow 
			     .vspdData.Col	= C_SEQ_NO1
				.vspdData.Text	= 1
			     ggoSpread.InsertRow  imRow 

				Call .vspdData.AddCellSpan(C_W1, .vspdData.MaxRows, 3, 1) 
			    .vspdData.row = .vspdData.MaxRows
			    .vspdData.Col	= C_SEQ_NO1
				.vspdData.Text	= "999999"
				.vspdData.Col = C_W1	:	.vspdData.CellType = 1	:	.vspdData.Text = "��"	:	.vspdData.TypeHAlign = 2
				.vspdData.Col = C_W7_NM	:	.vspdData.CellType = 1
				.vspdData.Col = C_W9_NM	:	.vspdData.CellType = 1
				 SetSpreadColor 1, .vspdData.MaxRows
				.vspdData.Row  = 1
				.vspdData.ActiveRow = 1
			else
				'.vspdData.ReDraw = False	' �� ���� ActiveRow ���� ������� ��, Ư���� �� ������ �ƴ϶� ReDraw�� �����. - �ֿ��� 
				iRow = .vspdData.ActiveRow

				If iRow = .vspdData.MaxRows Then	' -- ������ �հ��ٿ��� InsertRow�� �ϸ� ������ �߰��Ѵ�.
					ggoSpread.InsertRow iRow-1 , imRow 
					SetSpreadColor iRow, iRow + imRow - 1
	
					Call SetDefaultVal(iRow, imRow)
				Else
					ggoSpread.InsertRow ,imRow
					SetSpreadColor iRow+1, iRow + imRow
	
					Call SetDefaultVal(iRow+1, imRow)
					iRow = iRow + 1
				End If   
				.vspdData.Row  = iRow
				.vspdData.ActiveRow = iRow
	        End if 	
		    .vspdData.Col	= C_W2
			.vspdData.Text	= lgFISC_END_DT
   
		Else
			.vspdData.focus
			ggoSpread.Source = .vspdData2
		
			'.vspdData2.ReDraw = False
			iSeqNo = .vspdData2.MaxRows+1
		
			if 	.vspdData2.MaxRows = 0 then
			
			     ggoSpread.InsertRow  imRow 
			     .vspdData2.Col	= C_SEQ_NO2
				.vspdData2.Text	= 1
			     ggoSpread.InsertRow  imRow 

				Call .vspdData2.AddCellSpan(C_W12, .vspdData2.MaxRows, 3, 1) 
			     .row = .vspdData2.MaxRows
			    .vspdData2.Col	= C_SEQ_NO2
				.vspdData2.Text	= "999999"
				.vspdData2.Col = C_W12	:	.vspdData2.CellType = 1	:	.vspdData2.Text = "��"	:	.vspdData2.TypeHAlign = 2

				 SetSpreadColor 1, .vspdData2.MaxRows
				.vspdData2.Row  = 1
				 
			else
				'.vspdData2.ReDraw = False	' �� ���� ActiveRow ���� ������� ��, Ư���� �� ������ �ƴ϶� ReDraw�� �����. - �ֿ��� 
		     
				iRow = .vspdData2.ActiveRow
		
				If iRow = .vspdData2.MaxRows Then	' -- ������ �հ��ٿ��� InsertRow�� �ϸ� ������ �߰��Ѵ�.
					ggoSpread.InsertRow iRow-1 , imRow 
					SetSpreadColor iRow, iRow + imRow - 1
	
					Call SetDefaultVal(iRow, imRow)
				Else
					ggoSpread.InsertRow ,imRow
					SetSpreadColor iRow+1, iRow + imRow
	
					Call SetDefaultVal(iRow+1, imRow)
					iRow = iRow + 1
				End If   
				.vspdData2.Row  = iRow
	        End if 	
		    .vspdData2.Col	= C_W13
			.vspdData2.Text	= lgFISC_END_DT
			'.vspdData2.ReDraw = True
		End If
		
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

' �׸��忡 SEQ_NO, TYPE �ִ� ���� 
Function SetDefaultVal(iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With frm1	
	
		If gCurrGrid = 1 Then
		
			ggoSpread.Source = .vspdData
		
			If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
				.vspdData.Row = iRow
				.vspdData.Value = MaxSpreadVal(.vspdData, C_SEQ_NO1, iRow)
			Else
				iSeqNo = MaxSpreadVal(.vspdData, C_SEQ_NO1, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
				
				For i = iRow to iRow + iAddRows -1
					.vspdData.Row = i
					.vspdData.Col = C_SEQ_NO1 : .vspdData.Value = iSeqNo : iSeqNo = iSeqNo + 1
				Next
			End If
		Else
			ggoSpread.Source = .vspdData2
		
			If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
				.vspdData2.Row = iRow
				.vspdData2.Value = MaxSpreadVal(.vspdData2, C_SEQ_NO2, iRow)
			Else
				iSeqNo = MaxSpreadVal(.vspdData2, C_SEQ_NO2, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
				
				For i = iRow to iRow + iAddRows -1
					.vspdData2.Row = i
					.vspdData2.Col = C_SEQ_NO2 : .vspdData2.Value = iSeqNo : iSeqNo = iSeqNo + 1
				Next
			End If
		End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows, iActiveRow, dblSum 

	If gCurrGrid = 1	Then
		With frm1.vspdData 
			.focus
	
			ggoSpread.Source = Frm1.vspdData
			If .MaxRows <= 0 Then
				Exit Function
			ElseIf CheckTotRow(Frm1.vspdData, .ActiveRow) = True Then
				MsgBox "�հ�� ������ �� �����ϴ�.", vbCritical
				Exit Function
			Else
				lDelRows = ggoSpread.DeleteRow
				lgBlnFlgChgValue = True
				lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
				If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
			End If
			Call CheckReCalc(C_W5, .ActiveRow)
			ggoSpread.UpdateRow .MaxRows
		End With
    Else
		With frm1.vspdData2 
			.focus
	
			ggoSpread.Source = Frm1.vspdData2
			If .MaxRows <= 0 Then
				Exit Function
			ElseIf CheckTotRow(Frm1.vspdData2, .ActiveRow) = True Then
				MsgBox "�հ�� ������ �� �����ϴ�.", vbCritical
				Exit Function
			Else
				lDelRows = ggoSpread.DeleteRow
				lgBlnFlgChgValue = True
				lDelRows = CheckLastRow(Frm1.vspdData2, lDelRows)
				If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
			End If
			Call CheckReCalc2(C_W16, .ActiveRow)
			ggoSpread.UpdateRow .MaxRows
		End With    
    End If

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
	Dim bRtn1, bRtn2
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData
    bRtn1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData2
    bRtn2 = ggoSpread.SSCheckChange
    If bRtn1 = True Or bRtn2 = True Then
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
        strVal = strVal     & "&txtACCT_CD="        & Frm1.txtACCT_CD.Value
        strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQuery2() 

    DbQuery2 = False
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal


    With Frm1
    
		strVal = BIZ_PGM_ID2 & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtCO_CD="			& Frm1.txtCO_CD.value      '��: Query Key        
        strVal = strVal     & "&txtFISC_YEAR="		& Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="		& Frm1.cboREP_TYPE.Value      '��: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  

    DbQuery2 = True  
  
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	If frm1.vspdData.MaxRows > 0 Or _
		frm1.vspdData2.MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		    
	    Call SetToolbar("1101111100000111")										<%'��ư ���� ���� %>
'		Call SetToolbar("1111111100111111")										<%'��ư ���� ���� %>
	Else
		Call SetToolbar("1100111100000111")										<%'��ư ���� ���� %>
	End If
	
	Call SetSpreadTotalLine ' - �հ���� �籸�� 
End Function

Function DbQueryOk2()													<%'��ȸ ������ ������� %>
	
    Dim lRow    
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	If frm1.vspdData2.MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		    
	    Call SetToolbar("1100111100000111")										<%'��ư ���� ���� %>
'		Call SetToolbar("1111111100111111")										<%'��ư ���� ���� %>
	Else
		Call SetToolbar("1100111100000111")										<%'��ư ���� ���� %>
	End If
	
	Call SetSpreadTotalLine ' - �հ���� �籸�� 
	Call CheckReCalc2(C_W18, frm1.vspdData2.MaxRows)
	With Frm1.vspdData2
		' ----- 1��° �׸��� 
		For lRow = 1 To .MaxRows
    
           .Row = lRow
           .Col = 0
        
           .Text = ggoSpread.InsertFlag                                      '��: Insert
		Next
	End With
'	frm1.vspdData2.focus
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
	Dim pP21011
    Dim lRow        
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
    
	With Frm1
		' ----- 1��° �׸��� 
		For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
        
               Case  ggoSpread.InsertFlag                                      '��: Insert
                                                  strVal = strVal & "C"  &  Parent.gColSep
                                                'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO1   : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W1		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W1_NM		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2        : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W3		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W4		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W5		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W6		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W7		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W8		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W9		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W10		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W11		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep

 
                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '��: Update
                                                  strVal = strVal & "U"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO1   : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W1		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W1_NM		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W2        : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W3		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W4		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W5		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W6		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W7		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W8		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W9		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W10		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
                    .vspdData.Col = C_W11		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
                    
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '��: Delete
'                                                  strDel = strDel & "D"  &  Parent.gColSep
                                                  'strVal = strVal & lRow &  Parent.gColSep
                                                  strVal = strVal & "D"  &  Parent.gColSep
                    .vspdData.Col = C_SEQ_NO1   : strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep
   
                    lGrpCnt = lGrpCnt + 1
           End Select
		Next
		
       .txtSpread.value      = strVal
       strVal = ""
       
		' ----- 2��° �׸��� 
 		For lRow = 1 To .vspdData2.MaxRows
    
           .vspdData2.Row = lRow
           .vspdData2.Col = 0
        
           Select Case .vspdData2.Text
        
               Case  ggoSpread.InsertFlag                                      '��: Insert
													strVal = strVal & "C"  &  Parent.gColSep
													'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData2.Col = C_SEQ_NO2		: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W12			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W12_NM		: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W13			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W14			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W15			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W16			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W17			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W18			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W19			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gRowSep
 
                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '��: Update
													strVal = strVal & "U"  &  Parent.gColSep
													'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData2.Col = C_SEQ_NO2		: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W12			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W12_NM		: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W13			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W14			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W15			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W16			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W17			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W18			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gColSep
                    .vspdData2.Col = C_W19			: strVal = strVal & Trim(.vspdData2.Text) &  Parent.gRowSep
 
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '��: Delete
													strVal = strVal & "D"  &  Parent.gColSep
													'strVal = strVal & lRow &  Parent.gColSep
                    .vspdData2.Col = C_SEQ_NO2      : strVal = strVal & Trim(.vspdData2.Text) &  Parent.gRowSep
   
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
      
		.txtMode.value        =  Parent.UID_M0002
		'.txtUpdtUserId.value  =  Parent.gUsrID
		'.txtInsrtUserId.value =  Parent.gUsrID
		.txtMaxRows.value     = lGrpCnt-1 
		.txtSpread2.value      = strVal
		.txtFlgMode.value     = lgIntFlgMode
		
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' ���� ������ ���� ���� %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
	
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()" width=200>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ձݵ��</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" width=200>
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����������ձ�����</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>
						<!--<DIV id="myTabRef"><a href="vbscript:OpenRef()">��������ȸ</A>&nbsp;</DIV>-->
						<DIV id="myTabRef" STYLE="display:'none'"><a href="vbscript:GetRef()">����������ձݰ�������</A>|<A href="vbscript:OpenRefMenu">�ҵ�ݾ��հ�ǥ��ȸ</A>&nbsp;</DIV>
						</TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w3107ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script></TD>
									<TD CLASS="TD5">���θ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">�Ű���</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="�Ű���" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5">����</TD>
									<TD CLASS="TD6"><INPUT NAME="txtACCT_CD" MAXLENGTH="10" SIZE=10 ALT ="�����ڵ�" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenAccount(3)"> <INPUT NAME="txtACCT_NM" MAXLENGTH="30" SIZE=20 ALT ="����" tag="14X"></TD>
								</TR>
<!--								<TR>
									<TD CLASS="TD5">�ú���</TD>
									<TD CLASS="TD6"><SELECT NAME="cboAPP_TYPE" ALT="�ú���" STYLE="WIDTH: 50%" tag="1X"></SELECT>
									</TD>
									<TD CLASS="TD5">������</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/w3107ma1_txtDEBT_START_DT_txtDEBT_START_DT.js'></script>&nbsp;~&nbsp;
									<script language =javascript src='./js/w3107ma1_txtDEBT_END_DT_txtDEBT_END_DT.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">���ο���</TD>
									<TD CLASS="TD6"><SELECT NAME="cboCONF_TYPE" ALT="���ο���" STYLE="WIDTH: 50%" tag="1X"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD-->
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
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=15%>
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="100%">
											<script language =javascript src='./js/w3107ma1_vspdData_vspdData.js'></script>
										</TD>
									</TR>
								</TABLE>
								</DIV>
						
								<DIV ID="TabDiv" SCROLL=no>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="30%">
											<script language =javascript src='./js/w3107ma1_vspdData2_vspdData2.js'></script>
										</TD>
									</TR>
								</TABLE>
								</DIV>

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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows2" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
