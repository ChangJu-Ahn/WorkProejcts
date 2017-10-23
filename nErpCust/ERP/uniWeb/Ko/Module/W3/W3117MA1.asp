
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �����ޱ� �� ������ ������� 
'*  3. Program ID           : W3117MA1
'*  4. Program Name         : W3117MA1.asp
'*  5. Program Desc         : �����ޱ� �� ������ ������� 
'*  6. Modified date(First) : 2005/01/20
'*  7. Modified date(Last)  : 2006/01/24
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : HJO
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
' �λ� DB���� ���� �̺�.
' ���깮�� : ������ �ο� ���� ó�� ���� �����..
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
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "W3117MA1"
Const BIZ_PGM_ID		= "W3117MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "W3117MB2.asp"
Const EBR_RPT_ID		= "W3117OA1"

' -- 1�� �հ� �׸��� 
Dim C_SEQ_NO
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W7_NM
Dim C_W8
Dim C_W8_NM

' -- 2,3�� �����ޱ�/������ �׸��� 
Dim C_CHILD_SEQ_NO
Dim C_W9
Dim C_W10
Dim C_W11
Dim C_W12
Dim C_W13
Dim C_W14
Dim C_W15
Dim C_W16

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgOldCol, lgOldRow , lgChgFlg
Dim lgFISC_START_DT, lgFISC_END_DT, lgRateOver, lgDefaultRate

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	lgCurrGrid	= 1
	lgOldRow	= 0
	lgOldCol	= 2
	lgChgFlg	= False

	'--1���׸��� 
	C_SEQ_NO = 1
	C_W1 = 2
	C_W2 = 3
	C_W3 = 4
	C_W4 = 5
	C_W5 = 6
	C_W6 = 7
	C_W7 = 8
	C_W7_NM = 9
	C_W8 = 10
	C_W8_NM = 11

	'--2�� 3���׸��� 
	C_CHILD_SEQ_NO	= 2
	C_W9		= 3
	C_W10		= 4
	C_W11		= 5
	C_W12		= 6
	C_W13		= 7
	C_W14		= 8
	C_W15		= 9
	C_W16		= 10

	
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
     
End Sub

Sub InitSpreadSheet()
	Dim ret
	
    Call initSpreadPosVariables()  

	' 1�� �׸��� 
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
		frm1.vspdData.ScriptEnhanced = True
	   'patch version
	    ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
	    
		.ReDraw = false
	
	    .MaxCols = C_W8_NM + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
		       
	    ggoSpread.ClearSpreadData
	    .MaxRows = 0
	    
	    'Call AppendNumberPlace("6","3","2")
	
	    ggoSpread.SSSetEdit		C_SEQ_NO,	"����",				5,,,6,1	' �����÷� 
	    ggoSpread.SSSetCombo	C_W1,		"(1)�뿩����",		10
		ggoSpread.SSSetEdit		C_W2,		"(2)����(���θ�)",	15,,,50,1	
		ggoSpread.SSSetFloat	C_W3,		"(3)�����ޱ� ����",	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W4,		"(4)������ ����",	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W5,		"(5)������",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W6,		"(6)���ڼ���",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetCombo	C_W7,		"(7)�ҵ�ó��", 		10
	    ggoSpread.SSSetCombo	C_W7_NM,	"(7)�ҵ�ó��", 		10
	    ggoSpread.SSSetCombo	C_W8,		"(8)��������������", 10
	    ggoSpread.SSSetCombo	C_W8_NM,	"(8)��������������", 15
	
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W7,C_W7,True)
		Call ggoSpread.SSSetColHidden(C_W8,C_W8,True)
						
		.ReDraw = true
	
    End With

 	' -----  2�� �׸��� 
	With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2	
	   'patch version
	    ggoSpread.Spreadinit "V20041222_2",,parent.gForbidDragDropSpread    
	    
		.ReDraw = false
	    
	    .MaxCols = C_W16 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
		       
	    ggoSpread.ClearSpreadData
	   .MaxRows = 0
	 
	    'Call AppendNumberPlace("6","3","2")
	
		ggoSpread.SSSetEdit		C_SEQ_NO,	"�θ����", 5,,,6,1	' �����÷� 
	    ggoSpread.SSSetEdit		C_CHILD_SEQ_NO,	"�ڽļ���", 5,,,6,1	' �����÷� 
	    ggoSpread.SSSetEdit		C_W9,		"(9)����(���θ�)",	15,,,50,1	
	    ggoSpread.SSSetDate		C_W10,		"(10)����",			10, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_W11,		"(11)����",			15,,,50,1
	    ggoSpread.SSSetFloat	C_W12,		"(12)�����ݾ�" ,	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W13,		"(13)�뺯�ݾ�" ,	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W14,		"(14)�ܾ�",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
	    ggoSpread.SSSetEdit		C_W15,		"(15)�ϼ�",			10,1,,50,1
	    ggoSpread.SSSetFloat	C_W16,		"(16)����",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
	
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_CHILD_SEQ_NO,C_CHILD_SEQ_NO,True)
					
		.ReDraw = true
	
    End With

 	' -----  3�� �׸��� 
	With frm1.vspdData3
	
		ggoSpread.Source = frm1.vspdData3	
	   'patch version
	    ggoSpread.Spreadinit "V20041222_2",,parent.gForbidDragDropSpread    
	    
		.ReDraw = false
	    
	    .MaxCols = C_W16 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
		       
	    ggoSpread.ClearSpreadData
	   .MaxRows = 0
	 
	    'Call AppendNumberPlace("6","3","2")
	
		ggoSpread.SSSetEdit		C_SEQ_NO,	"�θ����", 5,,,6,1	' �����÷� 
	    ggoSpread.SSSetEdit		C_CHILD_SEQ_NO,	"�ڽļ���", 5,,,6,1	' �����÷� 
	    ggoSpread.SSSetEdit		C_W9,		"(9)����(���θ�)",	15,,,50,1	
	    ggoSpread.SSSetDate		C_W10,		"(10)����",			10, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_W11,		"(11)����",			15,,,50,1
	    ggoSpread.SSSetFloat	C_W12,		"(12)�����ݾ�" ,	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W13,		"(13)�뺯�ݾ�" ,	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetFloat	C_W14,		"(14)�ܾ�",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
	    ggoSpread.SSSetEdit		C_W15,		"(15)�ϼ�",			10,1,,50,1
	    ggoSpread.SSSetFloat	C_W16,		"(16)����",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_CHILD_SEQ_NO,C_CHILD_SEQ_NO,True)
					
		.ReDraw = true
    
    End With
    
	Call InitSpreadComboBox()
    Call SetSpreadLock 
           
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitSpreadComboBox()
    Dim IntRetCD1

	' �뿩���� 
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo "�����ڱ�" & vbTab & "��Ÿ", C_W1

	' �������� �ҵ�ó�� 
	IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1060' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W7
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W7_NM
	End If

	' ���������� ���� 
	IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1059' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W8
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W8_NM
	End If

End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    .vspdData2.ReDraw = False
    .vspdData3.ReDraw = False

	' 1�� �׸��� 
    ggoSpread.Source = frm1.vspdData
        
	ggoSpread.SSSetRequired C_W1, -1, -1
	ggoSpread.SSSetRequired C_W2, -1, -1
'	ggoSpread.SSSetRequired C_W6, -1, -1
	ggoSpread.SSSetRequired C_W7, -1, -1
	ggoSpread.SSSetRequired C_W7_NM, -1, -1
	ggoSpread.SSSetRequired C_W8, -1, -1
	ggoSpread.SSSetRequired C_W8_NM, -1, -1
    ggoSpread.SpreadLock C_W3, -1, C_W3
    ggoSpread.SpreadLock C_W4, -1, C_W4
    ggoSpread.SpreadLock C_W5, -1, C_W5    
    
    ' 2�� �׸��� 
    ggoSpread.Source = frm1.vspdData2	

    ggoSpread.SpreadLock C_W9, -1, C_W9
	ggoSpread.SSSetRequired C_W10, -1, -1
    ggoSpread.SpreadLock C_W14, -1, C_W14
    ggoSpread.SpreadLock C_W15, -1, C_W15
    ggoSpread.SpreadLock C_W16, -1, C_W16

	' 3�� �׸��� 
    ggoSpread.Source = frm1.vspdData3	

    ggoSpread.SpreadLock C_W9, -1, C_W9
	ggoSpread.SSSetRequired C_W10, -1, -1
    ggoSpread.SpreadLock C_W14, -1, C_W14
    ggoSpread.SpreadLock C_W15, -1, C_W15
    ggoSpread.SpreadLock C_W16, -1, C_W16

    .vspdData.ReDraw = True
    .vspdData2.ReDraw = True
    .vspdData3.ReDraw = True

    End With
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

	'If lgCurrGrid = 1 Then
		'.vspdData.ReDraw = False
 
		ggoSpread.Source = .vspdData
	
		ggoSpread.SSSetRequired C_W1, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W2, pvStartRow, pvEndRow
'		ggoSpread.SSSetRequired C_W6, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W7, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W7_NM, pvStartRow, pvEndRow
		If lgRateOver Then
			ggoSpread.SSSetRequired C_W8, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W8_NM, pvStartRow, pvEndRow
		Else
			ggoSpread.SpreadLock C_W8, pvEndRow, C_W8
			ggoSpread.SpreadLock C_W8_NM, pvEndRow, C_W8_NM
		End If
	    ggoSpread.SpreadLock C_W3, pvEndRow, C_W3
	    ggoSpread.SpreadLock C_W4, pvEndRow, C_W4
	    ggoSpread.SpreadLock C_W5, pvEndRow, C_W5    
		    
		'.vspdData.ReDraw = True

    'End If
    
    End With
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColorDetail2(ByVal pvEndRow)
    With frm1
    
		' 2�� �׸��� 
		ggoSpread.Source = frm1.vspdData2	

		ggoSpread.SpreadLock C_W9, pvEndRow, C_W9
		ggoSpread.SSSetRequired C_W10, pvEndRow, pvEndRow
		ggoSpread.SpreadLock C_W14, pvEndRow, C_W14
		ggoSpread.SpreadLock C_W15, pvEndRow, C_W15
		ggoSpread.SpreadLock C_W16, pvEndRow, C_W16
    
    End With
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColorDetail3(ByVal pvEndRow)
    With frm1
    
		' 2�� �׸��� 
		ggoSpread.Source = frm1.vspdData3	

		ggoSpread.SpreadLock C_W9, pvEndRow, C_W9
		ggoSpread.SSSetRequired C_W10, pvEndRow, pvEndRow
		ggoSpread.SpreadLock C_W14, pvEndRow, C_W14
		ggoSpread.SpreadLock C_W15, pvEndRow, C_W15
		ggoSpread.SpreadLock C_W16, pvEndRow, C_W16
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W7		= iCurColumnPos(2)
            C_W9		= iCurColumnPos(3)
            C_W8		= iCurColumnPos(4)
            C_W8_NM		= iCurColumnPos(5)
            C_W9		= iCurColumnPos(6)
            C_W10		= iCurColumnPos(7)
            C_W11		= iCurColumnPos(8)
            C_W12       = iCurColumnPos(9)
            C_W13		= iCurColumnPos(10)
            C_W15		= iCurColumnPos(11)
            C_W16		= iCurColumnPos(12)
            C_W17		= iCurColumnPos(13)
            C_W18		= iCurColumnPos(14)
            C_W19		= iCurColumnPos(15)
            C_W20		= iCurColumnPos(16)
    End Select    
End Sub


'============================== ����� ���� �Լ�  ========================================
Sub InsertRow2Head()
	' fncNew, onLoad�ÿ� ȣ���ؼ� �⺻������ 3ĭ�� �Է��� 
	Dim ret, iRow, iLoop
	
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
			
		.ReDraw = False

		iRow = 1
		ggoSpread.InsertRow ,1
		Call SetSpreadColor(iRow, iRow) 
		.Col = C_SEQ_NO : .Row = iRow: .Text = iRow	
		Call SetDefaultW8(iRow)		' ���������� ���� 
		
		iRow = 2
		ggoSpread.InsertRow ,1
		Call SetSpreadColor(iRow, iRow) 
		.Col = C_SEQ_NO : .Row = iRow: .Text = "999999"	
		Call .AddCellSpan(C_W1, .MaxRows, 2, 1)
		
		.col = C_W1 : .CellType = 1 : .text = "��" : .TypeHAlign = 2
				
		ggoSpread.SpreadLock C_W1, iRow, C_W8_NM, iRow
		
		.ReDraw = True		
		.focus
		'.SetActiveCell 2, 1
					
	End With

'	Call InsertRow2Detail2(1)
'	Call InsertRow2Detail3(1)
	
	Call vspdData_Click(C_W1, 1)
End Sub

Sub InsertRow2Detail2(Byval pSeqNo)

	' �۾������ �׸��� �߰� 
	Dim ret, iRow, iLoop, iLastRow
	
	With frm1.vspdData2
		
		.focus
		ggoSpread.Source = frm1.vspdData2

		iLastRow = .MaxRows
		.SetActiveCell C_W9, iLastRow	
		
		.ReDraw = False
		'ggoSpread.ClearSpreadData

		iRow = 1
		ggoSpread.InsertRow ,1
		.Row = iLastRow+iRow
		.Col = C_CHILD_SEQ_NO	: .Text = iRow
		.Col = C_SEQ_NO			: .Text = pSeqNo
		Call SetSpreadColorDetail2(iLastRow+iRow) 
		.RowHidden = True

		iRow = 2
		ggoSpread.InsertRow ,1
		.Row = iLastRow+iRow
		.Col = C_CHILD_SEQ_NO	: .Text = "999999"
		.Col = C_SEQ_NO			: .Text = pSeqNo
		.Col = C_W9				: .Text = "��"	: .TypeHAlign = 0	
		Call .AddCellSpan(C_W9, .MaxRows, 7, 1)
		Call SetSpreadColorDetail2(iLastRow+iRow) 

		ggoSpread.SpreadLock C_W9, iLastRow+iRow, C_W16, iLastRow+iRow
		.RowHidden = True	
		
		'.vspdData2.SetActiveCell 2, 1	
		.ReDraw = True		

	End With
	
End Sub

Sub InsertRow2Detail3(Byval pSeqNo)

	' �۾������ �׸��� �߰� 
	Dim ret, iRow, iLoop, iLastRow
	' ��Ÿ���Աݾ� �׸���	
	With frm1.vspdData3
		
		.focus
		ggoSpread.Source = frm1.vspdData3

		iLastRow = .MaxRows
		.SetActiveCell C_W9, iLastRow	
		
		.ReDraw = False
		'ggoSpread.ClearSpreadData

		iRow = 1
		ggoSpread.InsertRow ,1
		.Row = iLastRow+iRow
		.Col = C_CHILD_SEQ_NO	: .Text = iRow
		.Col = C_SEQ_NO			: .Text = pSeqNo
		Call SetSpreadColorDetail3(iLastRow+iRow) 
		.RowHidden = True

		iRow = 2
		ggoSpread.InsertRow ,1
		.Row = iLastRow+iRow
		.Col = C_CHILD_SEQ_NO	: .Text = "999999"
		.Col = C_SEQ_NO			: .Text = pSeqNo
		.Col = C_W9				: .Text = "��"	: .TypeHAlign = 0	
		Call .AddCellSpan(C_W9, .MaxRows, 7, 1)
		Call SetSpreadColorDetail3(iLastRow+iRow) 

		ggoSpread.SpreadLock C_W9, iLastRow+iRow, C_W16, iLastRow+iRow
		.RowHidden = True	
		
		'.vspdData2.SetActiveCell 2, 1	
		.ReDraw = True		
	
	End With
End Sub

' -- ����� �׸��� ������ 
Sub RedrawSumRow()
	Dim iRow, iMaxRows, lSeqNo, ret
	
	With frm1.vspdData
		iMaxRows = .MaxRows
		ggoSpread.Source = frm1.vspdData
		
		For iRow = 1 to iMaxRows
			.Col = C_SEQ_NO : .Row = iRow : lSeqNo = .Value
			
			If lSeqNo = 999999 Then ' �հ��� 
				ret = .AddCellSpan(C_W1, iRow, 2, 1)	' �հ� ���� �������� ���� ��ħ	
				
				.col = C_W1 : .CellType = 1 : .text = "��" : .TypeHAlign = 2

				ggoSpread.SpreadLock C_W1, iRow, C_W8_NM, iRow

			Else
				Call SetSpreadColor(iRow, iRow)
			End If
		Next
	End With
End Sub

' --  2��° �׸��� �հ� ������ 
Sub RedrawSumRow2()
	Dim iRow, iMaxRows, lSeqNo, ret
	
	With frm1.vspdData2
		iMaxRows = .MaxRows
		ggoSpread.Source = frm1.vspdData2
		
		For iRow = 1 to iMaxRows
			.Col = C_CHILD_SEQ_NO : .Row = iRow : lSeqNo = .Value
			
			If lSeqNo = 999999 Then ' �հ��� 
			
				.Col = C_W9			: .Text = "��"	: .TypeHAlign = 0	
				Call SetSpreadColorDetail2(iRow) 

				ggoSpread.SpreadLock C_W9, iRow, C_W16, iRow
				ret = .AddCellSpan(C_W9, iRow, 7, 1)	' �հ� ���� �������� ���� ��ħ	
				ggoSpread.SpreadLock 1, iRow, C_W16, iRow	
			End If
		Next
	End With
End Sub

' -- 3��° �׸��� �հ� ������ 
Sub RedrawSumRow3()
	Dim iRow, iMaxRows, lSeqNo, ret
	
	With frm1.vspdData3
		iMaxRows = .MaxRows
		ggoSpread.Source = frm1.vspdData3
		
		For iRow = 1 to iMaxRows
			.Col = C_CHILD_SEQ_NO : .Row = iRow : lSeqNo = .Value
			
			If lSeqNo = 999999 Then ' �հ��� 
			
				.Col = C_W9			: .Text = "��"	: .TypeHAlign = 0	
				Call SetSpreadColorDetail3(iRow) 

				ret = .AddCellSpan(C_W9, iRow, 7, 1)	' �հ� ���� �������� ���� ��ħ	
				ggoSpread.SpreadLock 1, iRow, C_W16, iRow	
			End If
		Next
	End With
End Sub

' -- ���� ���� ó�� 
Function ShowRowHidden(Byref pObj, Byval pSeqNo)
	Dim iRow, iSeqNo, iMaxRows, iFirstRow
	
	With pObj
	
	iMaxRows = .MaxRows : iFirstRow = 0
	
	For iRow = 1 To iMaxRows
		.Col = C_SEQ_NO : .Row = iRow : iSeqNo = .Value
		If iSeqNo = pSeqNo Then	' ���� ������..
			.RowHidden = False
			If iFirstRow = 0 Then iFirstRow = iRow
		Else
			.RowHidden = True
		End If
	Next
	
	ShowRowHidden = iFirstRow
	End With
End Function

' -- �հ� ������ üũ 
Function CheckTotalRow(Byref pObj, Byval pRow) 
	CheckTotalRow = False
	pObj.Col = C_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If UNICDbl(pObj.Text) = 999999 Then	 ' �հ� �� 
		CheckTotalRow = True
	End If
End Function

' -- �հ� ������ üũ 
Function CheckTotalRow2(Byref pObj, Byval pRow) 
	CheckTotalRow2 = False
	pObj.Col = C_CHILD_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If UNICDbl(pObj.Text) = 999999 Then	 ' �հ� �� 
		CheckTotalRow2 = True
	End If
End Function

' -- Detail Data�� �����ϴ��� üũ 
Function CheckDetailData(Byref pObj, Byref pObjDe, Byval pRow) 
	Dim iSeqNo, iRow
	CheckDetailData = 0
	pObj.Col = C_SEQ_NO : pObj.Row = pRow	:	iSeqNo = Trim(pObj.Text)
	
	With pObjDe
		For iRow = 1 To .MaxRows
			.Row = iRow	:	.Col = C_SEQ_NO
			If Trim(.Text) = iSeqNo Then
				.Col = 0
				If .Text <> ggoSpread.DeleteFlag Then
					CheckDetailData = CheckDetailData + 1
				End If
			End If
		Next
	End With
End Function

' -- �հ��̿��� ����Ÿ�� �ִ��� �����ϴ��� üũ 
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
		.Col = C_SEQ_NO	:	.Row = iMaxRow
		If .Text = 999999 and iCnt = 1 Then
			CheckLastRow = iMaxRow
		End If
	End With
	
End Function

' -- �հ��̿��� ����Ÿ�� �ִ��� �����ϴ��� üũ 
Function CheckLastRow2(Byref pObj, Byval pRow) 
	Dim iCnt, iRow, iMaxRow, iSeqNo, iTmpRow
	CheckLastRow2 = 0
	iCnt = 0
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_SEQ_NO
	iSeqNo = frm1.vspdData.Text
	With pObj

		For iRow = 1 To .MaxRows
			.Row = iRow
			.Col = C_SEQ_NO
			If .Text = iSeqNo Then
				.Col = 0
				If .Text <> ggoSpread.DeleteFlag Then
					iCnt = iCnt + 1
					iMaxRow = iRow
				End If
				.Col = C_CHILD_SEQ_NO
				If .Text = 999999 Then
					iTmpRow = iRow
				End If
			End If
		Next
		.Col = C_CHILD_SEQ_NO	:	.Row = iMaxRow
		If .Text = 999999 and iCnt = 1 Then
			CheckLastRow2 = iMaxRow
		ElseIf iCnt = 1 Then
			CheckLastRow2 = iTmpRow
		End If
	End With
	
End Function


' ----------- Grid 0 Process
Function Fn_GridCalc(ByVal pCol, ByVal pRow)
	Dim dblSum

	With Frm1.vspdData
		Select Case pCol
			Case C_W2		' ����(���θ�)
				.Col = C_W2	:	.Row = pRow
				Call SetW2ToChildGrid(.Text)	' ���� �̸��� ���� �׸��忡 �ִ´�.
		End Select

		' C_W3 : �����ޱ� ���� 
		dblSum = FncSumSheet(frm1.vspdData, C_W3, 1, .MaxRows - 1, false, -1, -1, "V")
		.Col = C_W3 : .Row = .MaxRows : .Value = dblSum	

		' C_W4 : ������ ���� 
		dblSum = FncSumSheet(frm1.vspdData, C_W4, 1, .MaxRows - 1, false, -1, -1, "V")
		.Col = C_W4 : .Row = .MaxRows : .Value = dblSum	

		' C_W5 : ������ 
		dblSum = FncSumSheet(frm1.vspdData, C_W5, 1, .MaxRows - 1, false, -1, -1, "V")
		.Col = C_W5 : .Row = .MaxRows : .Value = dblSum	

		' C_W6 : ���ڼ��� 
		dblSum = FncSumSheet(frm1.vspdData, C_W6, 1, .MaxRows - 1, false, -1, -1, "V")
		.Col = C_W6 : .Row = .MaxRows : .Value = dblSum	
	End With
End Function

' ----------- Grid 2 Process
Function Fn_GridCalc2(ByVal pCol, ByVal pRow)
	Dim dblSum

	With Frm1.vspdData2
	    Select Case pCol
			Case 0, C_W10, C_W12, C_W13
				Call SetG2SumValue(C_W12, pRow)	' ���� �÷� ���հ� 
				Call SetG2SumValue(C_W13, pRow)	' �뺯 �÷� ���հ� 
				Call SetG2W14(pRow)			' �ܾ� ��� 
				If pRow <> 0 Then
					.Col = C_W14 : .Row = pRow
					If UNICDbl(.Text) < 0 Then
						Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "�����ޱ�(14)�ܾ�", "X")           '��: "���ڸ� ���� ���߱�.."	
						Exit Function
					End If
				End If
				Call SetG2W15(pRow)			' �ϼ� ��� 
				Call SetG2W16(pRow)			' ���� ��� 
				'Call SetG2SumValue(C_W14, pRow)	' �ܾ� ���հ� 
				Call SetG2SumValue(C_W15, pRow)	' �ϼ� ���հ� 
				Call SetG2SumValue(C_W16, pRow)	' ���� ���հ� 
				
				Call SetG2W11(pRow)				' ���� �ݿ� 
				Call SetW3_W4()				' ��� �׸��� �ݿ� 
	    End Select
	End With
End Function

' ----------- Grid 3 Process
Function Fn_GridCalc3(ByVal pCol, ByVal pRow)
	Dim dblSum

	With Frm1.vspdData3
	    Select Case pCol
			Case 0, C_W10, C_W12, C_W13
				Call SetG3SumValue(C_W12, pRow)	' ���� �÷� ���հ� 
				Call SetG3SumValue(C_W13, pRow)	' �뺯 �÷� ���հ� 
				Call SetG3W14(pRow)			' �ܾ� ���  
				If pRow <> 0 Then
					.Col = C_W14 : .Row = pRow
					If UNICDbl(.Text) < 0 Then
						Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "������(14)�ܾ�", "X")           '��: "���ڸ� ���� ���߱�.."	
						Exit Function
					End If
				End If
				Call SetG3W15(pRow)			' �ϼ� ��� 
				Call SetG3W16(pRow)			' ���� ��� 
				'Call SetG3SumValue(C_W14, pRow)	' �ܾ� ���հ� 
				Call SetG3SumValue(C_W15, pRow)	' �ϼ� ���հ� 
				Call SetG3SumValue(C_W16, pRow)	' ���� ���հ� 
				
				Call SetG3W11(pRow)				' ���� �ݿ� 
				Call SetW3_W4()				' ��� �׸��� �ݿ� 
	    End Select

	End With
End Function

' -- ���� ������ �Ʒ� �׸��忡 ǥ�� 
Sub	SetW2ToChildGrid(Byval pW2)
	Dim i, iMaxRows, iLastRow, iSeqNo
	
	frm1.vspdData.Col = C_SEQ_NO: frm1.vspdData.Row = frm1.vspdData.ActiveRow
	iSeqNo = frm1.vspdData.Text
	
	With frm1.vspdData2
		iMaxRows = .MaxRows 
		For i = 1 to iMaxRows
			.Col = C_SEQ_NO : .Row = i 
			If iSeqNo = .VAlue Then
				If CheckTotalRow2(frm1.vspdData2, i) = False Then 
					.Col = C_W9 : .Row = i : .text = pW2
					
				End If
			End If
		Next
	End With

	With frm1.vspdData3
		iMaxRows = .MaxRows
		For i = 1 to iMaxRows
			.Col = C_SEQ_NO : .Row = i 
			If iSeqNo = .VAlue Then
				If CheckTotalRow2(frm1.vspdData3, i) = False Then 
					.Col = C_W9 : .Row = i : .text = pW2 
				End If
			End If
		Next
	End With
	
End Sub 

Function SetDefaultW8(ByVal pRow)
	Dim iIndex
	
	With Frm1.vspdData
		.Row = pRow
		If lgRateOver = False Then
			.Col = C_W8 : .Text = lgDefaultRate :	iIndex = .Value
			.Col = C_W8_NM : .Value = iIndex
		End If
	End With
End Function


' --- W5�� ����Ÿ ��� 
Function SetW5(Byval pRow)
	Dim dblW3, dblW4
	With frm1.vspdData
		.Col = C_W3	: .Row = pRow	: dblW3 = UNICDbl(.Value)
		.Col = C_W4	: .Row = pRow	: dblW4 = UNICDbl(.Value)
		.Col = C_W5	: .Row = pRow
		.Value = (dblW3 - dblW4)
	End With
End Function

' --- Grid2 W14�� ����Ÿ ��� 
Function SetG2W14(Byval pRow)
	Dim iRow, bIsFirstRow
	Dim iSeqNo, iChildSeqNo
	Dim dblW12, dblW13, dblW14
	
	If pRow = 0 Then Exit Function
	With frm1.vspdData2
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.text)
		.Col = C_CHILD_SEQ_NO	: .Row = pRow	: iChildSeqNo = UNICDbl(.text)
		bIsFirstRow = True : dblW14 = 0
		For iRow = 1 To .MaxRows-1
			.Row = iRow : .Col = C_SEQ_NO
			if UNICDbl(.text) = iSeqNo Then
				If bIsFirstRow = False Then
					.Col = C_SEQ_NO	: .Row = iRow - 1
					If iSeqNo = UNICDbl(.text) Then
						.Col = C_W14
						dblW14 = UNICDbl(.text)
					End If
				End If
				bIsFirstRow = False
				
				.Row = iRow 
				.Col = C_W12	: dblW12 = UNICDbl(.text)
				.Col = C_W13	: dblW13 = UNICDbl(.text)
				.Col = C_W14	: .Value = (dblW14 + dblW12 - dblW13)
				
			End If
		Next		
	End With
End Function

' --- Grid3 W14�� ����Ÿ ��� 
Function SetG3W14(Byval pRow)
	Dim iRow, bIsFirstRow
	Dim iSeqNo, iChildSeqNo
	Dim dblW12, dblW13, dblW14
	
	If pRow = 0 Then Exit Function
	With frm1.vspdData3
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.text)
		.Col = C_CHILD_SEQ_NO	: .Row = pRow	: iChildSeqNo = UNICDbl(.text)
		bIsFirstRow = True : dblW14 = 0
		For iRow = 1 To .MaxRows-1
			.Row = iRow : .Col = C_SEQ_NO
			if UNICDbl(.text) = iSeqNo Then
				If bIsFirstRow = False Then
					.Col = C_SEQ_NO	: .Row = iRow - 1
					If iSeqNo = UNICDbl(.text) Then
						.Col = C_W14
						dblW14 = UNICDbl(.text)
					End If
				End If
				bIsFirstRow = False
				
				.Row = iRow 
				.Col = C_W12	: dblW12 = UNICDbl(.text)
				.Col = C_W13	: dblW13 = UNICDbl(.text)
				.Col = C_W14	: .Value = (dblW14 + dblW12 - dblW13)
			End If
		Next
	End With
End Function

' --- Grid2 W15(�ϼ�)�� ����Ÿ ��� 
Function SetG2W15(Byval pRow)
	Dim datW10, datW10_DOWN, dblSum, iRow, blnPrintLast
	Dim dblW12, dblW13, iSeqNo
	
	If pRow = 0 Then Exit Function
	With frm1.vspdData2
		blnPrintLast = False
		
		.Col = C_W12	: .Row = pRow	: dblW12 = UNICDbl(.Text)
		.Col = C_W13	: .Row = pRow	: dblW13 = UNICDbl(.Text)
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.Value)
		If dblW12 = 0 And dblW13 = 0 Then
			Exit Function
		End If
		
		For iRow = .MaxRows-1 To 1 Step -1
			.Row = iRow : .Col = C_SEQ_NO
			if UNICDbl(.Value) = iSeqNo Then
			
				.Row = iRow
				.Col = C_W10	
				If .Text = "" Then		' ���ڰ� �����̸� �հ��� �����Ѵ�.
					
				Else		
	
					datW10 = CDate(.Text)

					If blnPrintLast = False Then	' �������� �ϼ� �����Ѱ�� 
						If frm1.cboREP_TYPE.value = "2" Then
							.Col = C_W15	: .Value = DateDiff("d", datW10, DateAdd("m", 6, lgFISC_START_DT)-1)+1
						Else
							.Col = C_W15	: .Value = DateDiff("d", datW10, lgFISC_END_DT)+1
						End If
						'.Col = C_W15	: .Text = DateDiff("d", datW10, lgFISC_END_DT)+1
						blnPrintLast = True
					Else
						.Col = C_W10	: .Row = iRow+1	
						
						If .Text <> "" Then	' �����Ҷ�.
							datW10_DOWN = CDate(.Text)	' ���� �������� ���ڸ� ��� 
							.Col = C_W15	: .Row = iRow	: .Text = DateDiff("d", datW10,  datW10_DOWN)	
						Else
							.Col = C_W15	: .Row = iRow	: .Text = ""
						End If
					End If
				
				End If
			End If
		Next
		
		dblSum = FncSumSheet(frm1.vspdData2, C_W15, 1, .MaxRows - 1, true, .MaxRows, C_W15, "V")	' �հ� 
	End With	

End Function

' --- Grid3 W15(�ϼ�)�� ����Ÿ ��� 
Function SetG3W15(Byval pRow)
	Dim datW10, datW10_DOWN, dblSum, iRow, blnPrintLast
	Dim dblW12, dblW13, iSeqNo
	
	If pRow = 0 Then Exit Function
	With frm1.vspdData3
		blnPrintLast = False

		.Col = C_W12	: .Row = pRow	: dblW12 = UNICDbl(.Text)
		.Col = C_W13	: .Row = pRow	: dblW13 = UNICDbl(.Text)
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.Value)
		If dblW12 = 0 And dblW13 = 0 Then
			Exit Function
		End If
		
		For iRow = .MaxRows-1 To 1 Step -1
			.Row = iRow : .Col = C_SEQ_NO
			if UNICDbl(.Value) = iSeqNo Then
				.Row = iRow
				.Col = C_W10	
			
				If .Text = "" Then		' ���ڰ� �����̸� �հ��� �����Ѵ�.
					
				Else		
	
					datW10 = CDate(.Text)
			
					If blnPrintLast = False Then	' �������� �ϼ� �����Ѱ�� 
						.Col = C_W15	: .Text = DateDiff("d", datW10, lgFISC_END_DT)+1
						blnPrintLast = True
					Else
						.Col = C_W10	: .Row = iRow+1	
						
						If .Text <> "" Then	' �����Ҷ�.
							datW10_DOWN = CDate(.Text)	' ���� �������� ���ڸ� ��� 
							.Col = C_W15	: .Row = iRow	: .Text = DateDiff("d", datW10,  datW10_DOWN)	
						Else
							.Col = C_W15	: .Row = iRow	: .Text = ""
						End If
					End If
				
				End If
			End If
		Next
		
		dblSum = FncSumSheet(frm1.vspdData3, C_W15, 1, .MaxRows - 1, true, .MaxRows, C_W15, "V")	' �հ� 
	End With	
End Function

' --- Grid2 W16�� ����Ÿ ��� 
Function SetG2W16(Byval pRow)
	Dim dblW14, dblW15, iRow
	With frm1.vspdData2
		For iRow =  1 To .MaxRows
			.Col = C_W14	: .Row = iRow	: dblW14 = UNICDbl(.Text)
			.Col = C_W15	: .Row = iRow
			If Trim(.Text) <> "" Then
				dblW15 = UNICDbl(.Text)
				.Col = C_W16	: .Row = iRow
				.Text = (dblW14 * dblW15)
			End If
		Next
	End With
End Function

' --- Grid3 W16�� ����Ÿ ��� 
Function SetG3W16(Byval pRow)
	Dim dblW14, dblW15, iRow
	With frm1.vspdData3
		For iRow =  1 To .MaxRows
			.Col = C_W14	: .Row = iRow	: dblW14 = UNICDbl(.Text)
			.Col = C_W15	: .Row = iRow
			If Trim(.Text) <> "" Then
				dblW15 = UNICDbl(.Text)
				.Col = C_W16	: .Row = iRow
				.Text = (dblW14 * dblW15)
			End If
		Next
	End With
End Function

' 2�� �׸��忡 ���� �ݿ� 
Function SetG2W11(ByVal pRow)
	Dim dblW12, dblW13
	Dim datW10
	Dim strDesc
	
	strDesc = ""
	If pRow = 0 Then Exit Function
	With frm1.vspdData2
		.Row = pRow
		.Col = C_W12 : dblW12 = UNICDbl(.Text)
		.Col = C_W13 : dblW13 = UNICDbl(.Text)
		.Col = C_W10

		If dblW12 > 0 Then
			strDesc = "�뿩"
		ElseIf dblW13 > 0 Then
			strDesc = "��ȯ"
		End If

		If Trim(.Value) <> "" Then
			datW10 = CDate(.Text)
			If Month(datW10) = 1 And Day(datW10) = 1 And dblW12 > 0 Then strDesc = "�����̿�"
		End IF
		
		.Col = C_W11 : .Row = pRow : .Text = strDesc
	End With
End Function

' 3�� �׸��忡 ���� �ݿ� 
Function SetG3W11(ByVal pRow)
	Dim dblW12, dblW13
	Dim datW10
	Dim strDesc
	
	strDesc = ""
	If pRow = 0 Then Exit Function
	With frm1.vspdData3
		.Row = pRow
		.Col = C_W12 : dblW12 = UNICDbl(.Text)
		.Col = C_W13 : dblW13 = UNICDbl(.Text)
		.Col = C_W10

		If dblW12 > 0 Then
			strDesc = "�뿩"
		ElseIf dblW13 > 0 Then
			strDesc = "��ȯ"
		End If

		If Trim(.Text) <> "" Then
			datW10 = CDate(.Text)
			If Month(datW10) = 1 And Day(datW10) = 1 And dblW12 > 0 Then strDesc = "�����̿�"
		End IF
		
		.Col = C_W11 : .Row = pRow : .Text = strDesc
	End With
End Function

Function SetG2SumValue(ByVal pCol, ByVal pRow)
	Dim iRow
	Dim iSeqNo, iChlSeqNo, dblSum
	
	dblSum = 0
	
	If pRow = 0 Then Exit Function
	
	With frm1.vspdData2
		
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.Text)
		For iRow = 1 To .MaxRows
			.Row = iRow
			.Col = C_SEQ_NO
			
			If iSeqNo = UNICDbl(.Text) Then
				.Col = pCol
			
				If .Text = "" Then		' ���ڰ� �����̸� �հ��� �����Ѵ�.					
				Else		
	
					.Col = C_CHILD_SEQ_NO
					If UNICDbl(.Text) <> 999999 Then
						.Col = pCol : dblSum = dblSum + UNICDbl(.Text)
					Else
						IF pCol= C_W16 Then
							.Col = pCol : .Text = dblSum
							ggoSpread.Source = frm1.vspdData2
							ggoSpread.UpdateRow iRow
						Else
							.Col = pCol : .Text = ""
						End If
					End If
				
				End If
			End If
		Next
		
	End With	
End Function

Function SetG3SumValue(ByVal pCol, ByVal pRow)
	Dim iRow
	Dim iSeqNo, iChlSeqNo, dblSum
	
	dblSum = 0
	If pRow = 0 Then Exit Function
	
	With frm1.vspdData3
		
		.Col = C_SEQ_NO	: .Row = pRow	: iSeqNo = UNICDbl(.Text)
		For iRow = 1 To .MaxRows
			.Row = iRow
			.Col = C_SEQ_NO
			
			If iSeqNo = UNICDbl(.Text) Then
				.Col = pCol
			
				If .Text = "" Then		' ���ڰ� �����̸� �հ��� �����Ѵ�.
					
				Else		
	
					.Col = C_CHILD_SEQ_NO
					If UNICDbl(.Text) <> 999999 Then
						.Col = pCol : dblSum = dblSum + UNICDbl(.Text)
					Else
						.Col = pCol : .Text = dblSum
						ggoSpread.Source = frm1.vspdData3
						ggoSpread.UpdateRow iRow
					End If
				
				End If
			End If
		Next
		
	End With	
End Function

' 1���׸��� W3(�����ޱ�����), W4(������ ����)�� �ֱ� 
Function SetW3_W4()
	Dim dblGrid2Sum, dblGrid3Sum, dblSum, iG1Row, iSeqNo, iRow
	
	iG1Row = frm1.vspdData.ActiveRow
	With frm1.vspdData
		.Col = C_SEQ_NO	: .Row = iG1Row	: iSeqNo = UNICDbl(.Value)
	End With
	
	With frm1.vspdData3
		If .MaxRows = 0 Then
			dblGrid3Sum = 0
		Else
			For iRow = 1 To .MaxRows
				.Row = iRow
				.Col = C_SEQ_NO
				If UNICDbl(.value) = iSeqNo Then
					.Col = C_CHILD_SEQ_NO
					If UNICDbl(.value) = 999999 Then
						.Col = C_W16 : dblGrid3Sum = UNICDbl(.Value)
						.Col = 0
						If .Text = ggoSpread.DeleteFlag Then dblGrid3Sum = 0
					End If
				End If
			Next
		End If
	End With
	
	With frm1.vspdData2
		If .MaxRows = 0 Then
			dblGrid2Sum = 0
		Else
			For iRow = 1 To .MaxRows
				.Row = iRow
				.Col = C_SEQ_NO
				If UNICDbl(.value) = iSeqNo Then
					.Col = C_CHILD_SEQ_NO
					If UNICDbl(.value) = 999999 Then
						.Col = C_W16 : dblGrid2Sum = UNICDbl(.Value)
						.Col = 0
						If .Text = ggoSpread.DeleteFlag Then dblGrid2Sum = 0
					End If
				End If
			Next
		End If
	End With

	With frm1.vspdData
		
		.Col = C_W3	: .Row = iG1Row	: .Value = dblGrid2Sum
		.Col = C_W4	: .Row = iG1Row	: .Value = dblGrid3Sum

		dblSum = FncSumSheet(frm1.vspdData, C_W3, 1, .MaxRows - 1, false, -1, -1, "V")	' ���� �÷� ���հ� 
		.Col = C_W3 : .Row = .MaxRows : .Value = dblSum
		dblSum = FncSumSheet(frm1.vspdData, C_W4, 1, .MaxRows - 1, false, -1, -1, "V")	' ���� �÷� ���հ� 
		.Col = C_W4 : .Row = .MaxRows : .Value = dblSum
	End With
	
	Call Setw5(iG1Row)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow iG1Row

	With frm1.vspdData
		dblSum = FncSumSheet(frm1.vspdData, C_W5, 1, .MaxRows - 1, false, -1, -1, "V")	' ���� �÷� ���հ� 
		.Col = C_W5 : .Row = .MaxRows : .Value = dblSum
		ggoSpread.UpdateRow .MaxRows
	End With
	
End Function

Function ChkG2W10(ByVal pRow)
	Dim datCurW10
	ChkG2W10 = False

	If Frm1.vspdData2.MaxRows <= 0 Then
		ChkG2W10 = True
		Exit Function
	End If
	
	With Frm1.vspdData2
		.Row = pRow : .Col = C_W10
		If Trim(.Value) = "" Then
			ChkG2W10 = True
		Else
			datCurW10 = CDate(.Text)
			If pRow -1  >= 1 Then
				.Row = pRow - 1
				If Trim(.Value) = "" Then
					ChkG2W10 = True
				ElseIf DateDiff("d", CDate(.Text), datCurW10) > 0 Then
					ChkG2W10 = True
				Else
					ChkG2W10 = False
					Exit Function
				End If
			End If
			If pRow + 1 < .MaxRows Then
				.Row = pRow + 1
				If Trim(.Value) = "" Then
					ChkG2W10 = True
				ElseIf DateDiff("d", datCurW10, CDate(.Text)) > 0 Then
					ChkG2W10 = True
				Else
					ChkG2W10 = False
				End If
			ElseIf pRow + 1 = .MaxRows Then
				ChkG2W10 = True
			End If
		End If
	End With
End Function

Function ChkG3W10(ByVal pRow)
	Dim datCurW10
	ChkG3W10 = False

	If Frm1.vspdData3.MaxRows <= 0 Then
		ChkG3W10 = True
		Exit Function
	End If
	
	With Frm1.vspdData3
		.Row = pRow : .Col = C_W10
		If Trim(.Value) = "" Then
			ChkG3W10 = True
		Else
			datCurW10 = CDate(.Text)
			If pRow -1  >= 1 Then
				.Row = pRow - 1
				If Trim(.Value) = "" Then
					ChkG3W10 = True
				ElseIf DateDiff("d", CDate(.Text), datCurW10) > 0 Then
					ChkG3W10 = True
				Else
					ChkG3W10 = False
					Exit Function
				End If
			End If
			If pRow + 1 < .MaxRows Then
				.Row = pRow + 1
				If Trim(.Value) = "" Then
					ChkG3W10 = True
				ElseIf DateDiff("d", datCurW10, CDate(.Text)) > 0 Then
					ChkG3W10 = True
				Else
					ChkG3W10 = False
				End If
			ElseIf pRow + 1 = .MaxRows Then
				ChkG3W10 = True
			End If
		End If
	End With
End Function

'============================== ���۷��� �Լ�  ========================================

Sub GetFISC_DATE()	' ��û������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd
	Dim dblConfRate, dblRateLoan
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

	call CommonQueryRs(" ISNULL(MAX(W1), 0)"," TB_LOAN_CALC_SUM "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		dblRateLoan = UNICDbl(lgF0)
	Else
		dblRateLoan = 0
	End if

	call CommonQueryRs(" CONVERT(NUMERIC(5,2), REFERENCE) * 100"," B_CONFIGURATION "," MAJOR_CD = 'W2006' AND MINOR_CD = '1' AND SEQ_NO = 1 ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		dblConfRate = UNICDbl(lgF0)
	Else
		dblConfRate = 0
	End if
	
	If dblRateLoan >= 0 And dblRateLoan < dblConfRate Then
		lgRateOver = False		' ������������ ���´����������� �����Ѵ�.
		lgDefaultRate = "1"
	Else
		lgRateOver = True		' ������������ ����ڰ� �����Ѵ�.
	End If
End Sub

Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' �������� ���� : ���ߵǸ� ���ȴ�.
'	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
			
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
	' 2. ��������ǥ�� �ڻ��Ѱ�, ��ä�Ѱ�-�����޹��μ�, �ں���+�����޹��μ�+�ֽĹ����ʰ���+��������-�ֽĹ�����������-�������� �������� 
End Function

Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("W1111RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W1111RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    With frm1
        If .vspdData.ActiveRow > 0 then 
            arrParam(0) = GetSpreadText(.vspdData, 3, .vspdData.ActiveRow, "X", "X")
            arrParam(1) = GetSpreadText(.vspdData, 4, .vspdData.ActiveRow, "X", "X")
        End If            
    End With    

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    'Call InsertRow2Head
    'Call InsertRow2Detail(1)
    
    Call SetToolbar("1110110100100111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

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
	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
    Call GetFISC_DATE

End Sub

Sub cboREP_TYPE_onChange()	' �Ű������ �ٲٸ�..
	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	Call GetFISC_DATE
End Sub



Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' �� Template ȭ�鿡���� ���� ������, �޺�(Name)�� ����Ǹ� �޺�(Code, Hidden)�� ��������ִ� ���� 
	With frm1.vspdData
		.Row = Row

	End With
End Sub

'==========================================================================================
Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	Call GetFISC_DATE

End Sub


'============================================  1�� �׸��� �̺�Ʈ  ====================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' �� Template ȭ�鿡���� ���� ������, �޺�(Name)�� ����Ǹ� �޺�(Code, Hidden)�� ��������ִ� ���� 
	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case C_W1		' ����(���θ�)
				.Col = Col
				If .Text = "�����ڱ�" Then
					.Col = C_W7_NM : .Text = "��" : intIndex = .Value
					.Col = C_W7 : .Value = intIndex		
				Else
					.Col = C_W7_NM : .Text = "��Ÿ�������" : intIndex = .Value
					.Col = C_W7 : .Value = intIndex		
				End If
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
			Case  C_W8
				.Col = Col
				intIndex = .Value
				.Col = C_W8_NM
				.Value = intIndex	
			Case  C_W8_NM
				.Col = Col
				intIndex = .Value
				.Col = C_W8
				.Value = intIndex		
		End Select
	End With

End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim dblSum
	
	With Frm1.vspdData
		.Row = Row
		.Col = Col
	
		If .CellType = parent.SS_CELL_TYPE_FLOAT Then
			If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
			   .text = .TypeFloatMin
			End If
		End If
			
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row
		ggoSpread.UpdateRow .maxRows
		
		
		Call Fn_GridCalc(Col, Row)
	End With
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If Row <> NewRow Then
		Call vspdData_Click(Col, NewRow)
	End If
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
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
	
	Dim iSeqNo, IntRetCD, iLastRow
	
	If lgOldRow = Row  Then Exit Sub
	
    ggoSpread.Source = frm1.vspdData
  
    If Row = frm1.vspdData.MaxRows Then
		iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
		iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)

	Else
		With frm1.vspdData
			.Col = C_SEQ_NO : .Row = Row : iSeqNo = .Value
			
			' ���� �׸��� ǥ�÷�ƾ'
			iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
			frm1.vspdData2.SetActiveCell C_W10, iLastRow
			
'			If iLastRow = 0 Then 
'				Call InsertRow2Detail2(iSeqNo)
'				iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
'			End If

			iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
			frm1.vspdData3.SetActiveCell C_W10, iLastRow
	
'			If iLastRow = 0 Then 
'				Call InsertRow2Detail3(iSeqNo)
'				iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
'			End If

			.focus
		End With	
	End If
  
    lgOldRow = Row	: lgOldCol = Col


End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
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

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	lgCurrGrid = 1
	ggoSpread.Source = Frm1.vspdData
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
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



'============================================  2�� �׸��� �̺�Ʈ  ====================================

Sub vspdData2_Change(ByVal Col , ByVal Row )
	Dim dblSum
	
	With Frm1.vspdData2
		.Row = Row
		.Col = Col
	
		If .CellType = parent.SS_CELL_TYPE_FLOAT Then
			If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
			   .text = .TypeFloatMin
			End If
		End If
		
		If ChkG2W10(Row) = False Then
			Call DisplayMsgBox("WC0016", parent.VB_INFORMATION, "X", "X")           '��: "���ڸ� ���� ���߱�.."	
			.Row = Row : .text = ""		
			Exit Sub
		End If
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.UpdateRow Row

		Call Fn_GridCalc2(Col, Row)
	End With
	lgChgFlg = True ' ����Ÿ ���� 
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y )
	lgCurrGrid = 2
	ggoSpread.Source = Frm1.vspdData2
	
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)

End Sub

'============================================  3�� �׸��� �̺�Ʈ  ====================================
Sub vspdData3_Change(ByVal Col , ByVal Row )
	Dim dblSum
	
	With Frm1.vspdData3
		.Row = Row
		.Col = Col
	
		If .CellType = parent.SS_CELL_TYPE_FLOAT Then
			If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
			   .text = .TypeFloatMin
			End If
		End If
		
		If ChkG3W10(Row) = False Then
			Call DisplayMsgBox("WC0016", parent.VB_INFORMATION, "X", "X")           '��: "���ڸ� ���� ���߱�.."	
			.Row = Row : .text = ""		
			Exit Sub
		End If
		ggoSpread.Source = frm1.vspdData3
		ggoSpread.UpdateRow Row
		
		Call Fn_GridCalc3(Col, Row)
	End With
	lgChgFlg = True ' ����Ÿ ���� 
End Sub

Sub vspdData3_MouseDown(Button , Shift , x , y )
	lgCurrGrid = 3
	ggoSpread.Source = Frm1.vspdData3
End Sub

'============================================  �������� �Լ�  ====================================

Function FncQuery() 
    Dim IntRetCD , i, blnChange
    
    FncQuery = False                                                        
    blnChange = False
    
    Err.Clear                                                                <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
	ggoSpread.Source = Frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	ggoSpread.Source = Frm1.vspdData2
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

	ggoSpread.Source = Frm1.vspdData3
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If
    
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
'    Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
	Dim blnChange
        
    FncSave = False           
    blnChange = False                                              
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    
'    If lgChgFlg = False Then
    
	    ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange <> False Then
			blnChange = True
		End If

	    ggoSpread.Source = frm1.vspdData2
		If ggoSpread.SSCheckChange <> False Then
			blnChange = True
'		    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
'		    Exit Function
		End If

	    ggoSpread.Source = frm1.vspdData3
		If ggoSpread.SSCheckChange <> False Then
			blnChange = True
		End If


		If blnChange = False Then
		    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		    Exit Function
		End If

	    ggoSpread.Source = frm1.vspdData
		If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
		      Exit Function
		End If    


	    ggoSpread.Source = frm1.vspdData2
		If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
		      Exit Function
		End If    


	    ggoSpread.Source = frm1.vspdData3
		If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
		      Exit Function
		End If    
		
'	End If
	
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

    Call InsertRow2Head
    
    Call SetToolbar("1110111100000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If .vspdData.ActiveRow > 0 Then
			.vspdData.focus
			.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

			.vspdData.Col = C_W9
			.vspdData.Text = ""
    
			.vspdData.Col = C_W10
			.vspdData.Text = ""
			
			.vspdData.Col = C_W11
			.vspdData.Text = ""
			
			.vspdData.Col = C_W12
			.vspdData.Text = ""
			
			.vspdData.ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    Dim lDelRows

	Select Case lgCurrGrid 
		CAse  1	
			With frm1.vspdData 
				.focus
				ggoSpread.Source = frm1.vspdData 
				If CheckTotalRow(frm1.vspdData, .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				ElseIf CheckDetailData(Frm1.vspdData, Frm1.vspdData2, .ActiveRow)  > 0  Or CheckDetailData(Frm1.vspdData, Frm1.vspdData3, .ActiveRow)  > 0 Then
					MsgBox "���� ����Ÿ�� �����Ͽ� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.EditUndo

					lDelRows = ggoSpread.EditUndo
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
					If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
				End If
				Call SetW3_W4()
			End With
		CAse 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				If CheckTotalRow2(frm1.vspdData2, .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.EditUndo

					lDelRows = ggoSpread.EditUndo
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow2(Frm1.vspdData2, lDelRows)
					If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
				End If
				Call SetW3_W4()
			End With    
 		CAse 3
			With frm1.vspdData3 
				.focus
				ggoSpread.Source = frm1.vspdData3
				If CheckTotalRow2(frm1.vspdData3, .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.EditUndo

					lDelRows = ggoSpread.EditUndo
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow2(Frm1.vspdData3, lDelRows)
					If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
				End If
				Call SetW3_W4()
			End With     
	End Select
  
	lgChgFlg = True                                                '��: Protect system from crashing
End Function


Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo, iLastRow
    Dim iStrNm

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
   
	Frm1.vspdData.Col = C_W2
	Frm1.vspdData.Row = Frm1.vspdData.ActiveRow
   	iStrNm = Frm1.vspdData.Text

	With frm1	

		' ù���� ��� �հ���� �ִ� ��ƾ 
		If .vspdData.MaxRows = 0 Then
			Call InsertRow2Head
			Call SetToolbar("1110111100000111")
			Exit Function
		End If
		
		Select Case lgCurrGrid
			Case 1	' 1�� �׸��� 
		
			.vspdData.focus
			ggoSpread.Source = .vspdData
			
			iRow = .vspdData.ActiveRow	' ������ 
			
			.vspdData.ReDraw = False
			
			If iRow = .vspdData.MaxRows Then
		
				' SEQ_NO �� �׸��忡 �ִ� ���� 
				iSeqNo = GetMaxSpreadVal(.vspdData , C_SEQ_NO)	' �ִ�SEQ_NO�� ���ؿ´�.
			
				ggoSpread.InsertRow iRow-1 ,imRow	' �׸��� �� �߰�(����� ��� ����)
				SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1	' �׸��� ���󺯰� 
		
				For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1	' �߰��� �׸����� SEQ_NO�� �����Ѵ�.
					.vspdData.Row = iRow
					.vspdData.Col = C_SEQ_NO
					.vspdData.Text = iSeqNo
					iSeqNo = iSeqNo + 1		' SEQ_NO�� �����Ѵ�.

					Call SetDefaultW8(iRow)		' ���������� ���� 

				Next				

				'ggoSpread.InsertRow iRow-1 , imRow 
				'SetSpreadColor iRow, iRow + imRow - 1
				'MaxSpreadVal2 .vspdData, C_SEQ_NO, iRow	
				'vspdData.Col = C_SEQ_NO : .vspdData.Row = Row : iSeqNo = .vspdData.Value
			Else

				' SEQ_NO �� �׸��忡 �ִ� ���� 
				iSeqNo = GetMaxSpreadVal(.vspdData , C_SEQ_NO)	' �ִ�SEQ_NO�� ���ؿ´�.
			
				ggoSpread.InsertRow ,imRow	' �׸��� �� �߰�(����� ��� ����)
				SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1	' �׸��� ���󺯰� 
		
				For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1	' �߰��� �׸����� SEQ_NO�� �����Ѵ�.
					.vspdData.Row = iRow
					.vspdData.Col = C_SEQ_NO
					.vspdData.Text = iSeqNo
					iSeqNo = iSeqNo + 1		' SEQ_NO�� �����Ѵ�.

					Call SetDefaultW8(iRow)		' ���������� ���� 
				Next			
				'ggoSpread.InsertRow ,imRow
				'SetSpreadColor iRow+1, iRow+1
				'MaxSpreadVal .vspdData, C_SEQ_NO, iRow+1
				'.vspdData.Col = C_SEQ_NO : .vspdData.Row = Row+1 : iSeqNo = .vspdData.Value
			End If

			.vspdData.ReDraw = True	
						
			' ���� �׸��� ǥ�÷�ƾ'
			iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
			frm1.vspdData2.SetActiveCell C_W7, iLastRow
			
			iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
			frm1.vspdData3.SetActiveCell C_W7, iLastRow
			
			Call vspdData_Click(.vspdData.Col, .vspdData.ActiveRow)
		Case 2	' 2�� �׸��� 
			.vspdData2.focus
			ggoSpread.Source = .vspdData2

			.vspdData.Col = C_SEQ_NO : .vspdData.Row = .vspdData.ActiveRow : iSeqNo = .vspdData.Value

			' ù���� ��� �հ���� �ִ� ��ƾ 
			If .vspdData.ActiveRow = .vspdData.MaxRows Then
				Exit Function
			ElseIf ShowRowHidden(frm1.vspdData2, iSeqNo) = 0 Then
				Call InsertRow2Detail2(iSeqNo)
				Call ShowRowHidden(frm1.vspdData2, iSeqNo)
			Else
				'.vspdData2.ReDraw = False	' �� ���� ActiveRow ���� ������� ��, Ư���� �� ������ �ƴ϶� ReDraw�� �����. - �ֿ��� 
				iRow = .vspdData2.ActiveRow
				
				If iRow = .vspdData2.MaxRows Then
					ggoSpread.InsertRow iRow-1 , imRow 
					SetSpreadColorDetail2 iRow
					MaxSpreadVal2 .vspdData2, C_SEQ_NO, C_CHILD_SEQ_NO, iRow , iSeqNo
				Else
					ggoSpread.InsertRow ,imRow
					SetSpreadColorDetail2 iRow+1
					MaxSpreadVal2 .vspdData2, C_SEQ_NO, C_CHILD_SEQ_NO, iRow+1, iSeqNo	
				End If	
			End If
			Call SetW2ToChildGrid(iStrNm)	' ���� �̸��� ���� �׸��忡 �ִ´�.

		Case 3	' 3�� �׸��� 
			.vspdData3.focus
			ggoSpread.Source = .vspdData3

			.vspdData.Col = C_SEQ_NO : .vspdData.Row = .vspdData.ActiveRow : iSeqNo = .vspdData.Value

			' ù���� ��� �հ���� �ִ� ��ƾ 
			If .vspdData.ActiveRow = .vspdData.MaxRows Then
				Exit Function
			ElseIf ShowRowHidden(frm1.vspdData3, iSeqNo) = 0 Then
				Call InsertRow2Detail3(iSeqNo)
				Call ShowRowHidden(frm1.vspdData3, iSeqNo)
			Else
				'.vspdData3.ReDraw = False	' �� ���� ActiveRow ���� ������� ��, Ư���� �� ������ �ƴ϶� ReDraw�� �����. - �ֿ��� 
				iRow = .vspdData3.ActiveRow
				
				If iRow = .vspdData3.MaxRows Then
					ggoSpread.InsertRow iRow-1 , imRow 
					SetSpreadColorDetail3 iRow
					MaxSpreadVal2 .vspdData3, C_SEQ_NO, C_CHILD_SEQ_NO, iRow	, iSeqNo
				Else
					ggoSpread.InsertRow ,imRow
					SetSpreadColorDetail3 iRow+1
					MaxSpreadVal2 .vspdData3, C_SEQ_NO, C_CHILD_SEQ_NO, iRow+1, iSeqNo	
				End If	
			End If
			Call SetW2ToChildGrid(iStrNm)	' ���� �̸��� ���� �׸��忡 �ִ´�.
		End Select
		
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows

	Select Case lgCurrGrid 
		Case  1	
			With frm1.vspdData 
				.focus
				ggoSpread.Source = frm1.vspdData 
				If CheckTotalRow(frm1.vspdData, .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				ElseIf CheckDetailData(Frm1.vspdData, Frm1.vspdData2, .ActiveRow)  > 0  Or CheckDetailData(Frm1.vspdData, Frm1.vspdData3, .ActiveRow)  > 0 Then
					MsgBox "���� ����Ÿ�� �����Ͽ� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.DeleteRow

					lDelRows = ggoSpread.DeleteRow
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
					If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
				End If
				Call SetW3_W4()
			End With
		Case 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				If CheckTotalRow2(frm1.vspdData2, .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.DeleteRow

					lDelRows = ggoSpread.DeleteRow
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow2(Frm1.vspdData2, lDelRows)
					If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
				End If
				Call SetW3_W4()
			End With    
 		Case 3
			With frm1.vspdData3 
				.focus
				ggoSpread.Source = frm1.vspdData3
				If CheckTotalRow2(frm1.vspdData3, .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
'					lDelRows = ggoSpread.DeleteRow

					lDelRows = ggoSpread.DeleteRow
					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow2(Frm1.vspdData3, lDelRows)
					If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
				End If
				Call SetW3_W4()
			End With     
	End Select
	
	lgChgFlg = True
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
	
    ggoSpread.Source = frm1.vspdData	
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
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr, iSeqNo, iLastRow
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
        
    Call SetToolbar("1111111100110111")										<%'��ư ���� ���� %>
	
	Call RedrawSumRow
	Call RedrawSumRow2
	Call RedrawSumRow3

	With frm1.vspdData
		.Col = C_SEQ_NO : .Row = .ActiveRow : iSeqNo = .Value
			
		' ���� �׸��� ǥ�÷�ƾ'
		iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)

		' ���� �׸��� ǥ�÷�ƾ'
		iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
	End With			
	frm1.vspdData.focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow , lCol, lGrpCnt, lMaxRows, lMaxCols
    Dim lStartRow, lEndRow , lChkAmt
    Dim strVal
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    lGrpCnt = 1
    
	With frm1.vspdData
		' ----- 1��° �׸��� 
		ggoSpread.Source = frm1.vspdData
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
				
		For lRow = 1 To lMaxRows
		    
		   .Row = lRow : .Col = 0
		   
		   ' I/U/D �÷��� ó�� 
		   Select Case .Text
		       Case  ggoSpread.InsertFlag                                      '��: Insert
		                                          strVal = strVal & "C"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1
		                    
		       Case  ggoSpread.UpdateFlag                                      '��: Update                                                  
		                                           strVal = strVal & "U"  &  Parent.gColSep                                                 
		            lGrpCnt = lGrpCnt + 1                                                 
		       Case  ggoSpread.DeleteFlag                                      '��: Delete
		                                          strVal = strVal & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
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

    frm1.txtSpread.value      = strVal
    strVal = ""

 	With frm1.vspdData2
		' ----- 2��° �׸��� 
		ggoSpread.Source = frm1.vspdData2
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
				
		For lRow = 1 To lMaxRows
		    
		   .Row = lRow : .Col = 0

		   ' I/U/D �÷��� ó�� 
		   Select Case .Text
		       Case  ggoSpread.InsertFlag                                      '��: Insert
                                          strVal = strVal & "C"  &  Parent.gColSep	
		            lGrpCnt = lGrpCnt + 1
		                    
		       Case  ggoSpread.UpdateFlag                                      '��: Update                                                  
                                          strVal = strVal & "U"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1                                                 
		       Case  ggoSpread.DeleteFlag                                      '��: Delete
                                          strVal = strVal & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		 
		 '�ܾװ��� 20060125 by HJO
		 .Col = C_W14 : .Row = lRow
			If UNICDbl(.Text) < 0 Then
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "�������ޱ�(14)�ܾ�", "X")           '��: "���ڸ� ���� ���߱�.."	
				Call  LayerShowHide(0)
				Exit Function
			Else
			.Col=0
			  ' ��� �׸��� ����Ÿ ����     
			  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
					For lCol = C_SEQ_NO To lMaxCols
	'					If lCol = C_W10 Then
	'						.Col = lCol : strVal = strVal & CDate(.Text) &  Parent.gColSep
	'					Else
							.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
	'					End If
					Next
					strVal = strVal & Trim(.Text) &  Parent.gRowSep
			  End If  
			End If
		Next
	End With

    frm1.txtSpread2.value      = strVal
    strVal = ""
    	
	With frm1.vspdData3
		' ----- 3��° �׸��� 
		ggoSpread.Source = frm1.vspdData3
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
				
		For lRow = 1 To lMaxRows
		    
		   .Row = lRow : .Col = 0
		   
		   ' I/U/D �÷��� ó�� 
		   Select Case .Text
		       Case  ggoSpread.InsertFlag                                      '��: Insert
		                                          strVal = strVal & "C"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1
		                    
		       Case  ggoSpread.UpdateFlag                                      '��: Update                                                  
		                                           strVal = strVal & "U"  &  Parent.gColSep                                                 
		            lGrpCnt = lGrpCnt + 1                                                 
		       Case  ggoSpread.DeleteFlag                                      '��: Delete
		                                          strVal = strVal & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		 '�ܾװ��� 20060125 by HJO
		 .Col = C_W14 : .Row = lRow
		If UNICDbl(.Text) < 0 Then
			Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "������(14)�ܾ�", "X")           '��: "���ڸ� ���� ���߱�.."	
			Call  LayerShowHide(0)
			Exit Function
		End If
		.Col=0
		  ' ��� �׸��� ����Ÿ ����     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = C_SEQ_NO To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
 
		Next
	End With
	
    frm1.txtSpread3.value      = strVal
    strVal = ""	  
        
	frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
	frm1.txtFlgMode.value     = lgIntFlgMode

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
    
	frm1.vspdData3.MaxRows = 0
    ggoSpread.Source = frm1.vspdData3
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
<BODY TABINDEX="-1" SCROLL="">
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<!--<a href="vbscript:GetRef">�λ絥��Ÿ �ҷ�����</A>  -->
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
									<TD CLASS="TD6"><script language =javascript src='./js/w3117ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=1>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP>
                                   <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT"> 1. </LEGEND>
                                   <script language =javascript src='./js/w3117ma1_vaSpread_vspdData.js'></script>
								  </FIELDSET>
								  <BR>
                                   <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT"> 2. </LEGEND>
                                   <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT"> ��. �����ޱ�</LEGEND>
                                   <script language =javascript src='./js/w3117ma1_vaSpread2_vspdData2.js'></script>
								  </FIELDSET>
								  <BR>
									<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN="LEFT"> ��. ������</LEGEND>
									<script language =javascript src='./js/w3117ma1_vaSpread3_vspdData3.js'></script>
								  </FIELDSET>
								  <BR>
								  </FIELDSET>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
			    <TR>
				        <TD WIDTH=10>&nbsp;</TD>
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><����>�������� �������</LABEL>&nbsp;
				            <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><����>�������� �������</LABEL>&nbsp;
				            
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" STYLE="Display:none"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" STYLE="Display:none"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread3" tag="24" STYLE="Display:none"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOLD_CO_CD" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOLD_FISC_YEAR" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOLD_REP_TYPE" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

