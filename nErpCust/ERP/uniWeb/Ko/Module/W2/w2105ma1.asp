<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ���Աݾ����� 
'*  3. Program ID           : W1111MA1
'*  4. Program Name         : W1111MA1.asp
'*  5. Program Desc         : ��16ȣ ���Աݾ� �������� 
'*  6. Modified date(First) : 2005/01/03
'*  7. Modified date(Last)  : 2006/01/23
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
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "w2105MA1"
Const BIZ_PGM_ID = "w2105mb1.asp"	
Const EBR_RPT_ID = "w2105OA1"										 '��: �����Ͻ� ���� ASP�� 

' -- 1�� ���Աݾ�������� �׸��� 
Dim C_SEQ_NO
Dim C_W1_CD
Dim C_W1_NM
Dim C_W2_CD
Dim C_W2_NM
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_DESC1

' -- 2�� ���Աݾ� ������ ��. �۾�������� ���� ���Աݾ� �׸��� 
Dim C_CHILD_SEQ_NO
Dim C_W2_NM2
Dim C_W7
Dim C_W8
Dim C_W9
Dim C_W10
Dim C_W11
Dim C_W12
Dim C_W13
Dim C_W14
Dim C_W15
Dim C_W16

' -- 3�� ���Աݾ� ������ ��. ��Ÿ ���Աݾ� �׸��� 
Dim C_W17
Dim C_W18
Dim C_W19
Dim C_W20
Dim C_DESC2

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgOldCol, lgOldRow , lgChgFlg

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	lgCurrGrid	= 1
	lgOldRow	= 0
	lgOldCol	= 2
	lgChgFlg	= False

	'--1�����Աݾ��������׸��� 
	C_SEQ_NO	= 1
	C_W1_CD		= 2
	C_W1_NM		= 3
	C_W2_CD		= 4
	C_W2_NM		= 5
	C_W3		= 6
	C_W4		= 7
	C_W5		= 8
	C_W6		= 9
	C_DESC1		= 10

	'--2�����Աݾ���������.�۾�����������Ѽ��Աݾױ׸��� 
	C_CHILD_SEQ_NO	= 2
	C_W2_NM2	= 3
	C_W7		= 4
	C_W8		= 5
	C_W9		= 6
	C_W10		= 7
	C_W11		= 8
	C_W12		= 9
	C_W13		= 10
	C_W14		= 11
	C_W15		= 12
	C_W16		= 13

	'--3�����Աݾ���������.��Ÿ���Աݾױ׸��� 
	C_CHILD_SEQ_NO	= 2
	'C_W2_NM2	= 3
	C_W17		= 4
	C_W18		= 5
	C_W19		= 6
	C_W20		= 7
	C_DESC2		= 8
	
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
    lgOldRow = 0
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
   
    ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_DESC1 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols									'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    ggoSpread.ClearSpreadData
    .MaxRows = 0
    
	'����� 2�ٷ�    
    .ColHeaderRows = 2    
    ' 
    Call AppendNumberPlace("6","3","2")

    ggoSpread.SSSetEdit		C_SEQ_NO,	"����"		, 5,,,6,1	' �����÷� 
	ggoSpread.SSSetEdit		C_W1_CD,	"(1)�׸�"	, 10,,,50,1	
	ggoSpread.SSSetEdit		C_W1_NM,	"(1)�׸�"	, 15,,,50,1	
	ggoSpread.SSSetEdit		C_W2_CD,	"(2)����"	, 10,,,50,1	
	ggoSpread.SSSetEdit		C_W2_NM,	"(2)����"	, 15,,,50,1	
	ggoSpread.SSSetFloat	C_W3,		"(3)��꼭��" & vbCrLf & "���Աݾ�"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	ggoSpread.SSSetFloat	C_W4,		"(4)����"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	ggoSpread.SSSetFloat	C_W5,		"(5)����"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
    ggoSpread.SSSetFloat	C_W6,		"(6)������ ���Աݾ�" & vbCrLf & "[(3) + (4) - (5)]", 15,	Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
    ggoSpread.SSSetEdit		C_DESC1,	"�� ��", 20,,,20,1

	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
	Call ggoSpread.SSSetColHidden(C_W1_CD,C_W1_CD,True)
	Call ggoSpread.SSSetColHidden(C_W2_CD,C_W2_CD,True)
					
	'Call InitSpreadComboBox()	�޺����� 

	' �׸��� ��� ��ħ ���� 
	ret = .AddCellSpan(C_SEQ_NO, -1000, 1, 2)	' SEQ_NO ��ħ 
    ret = .AddCellSpan(C_W1_CD, -1000, 4, 1)	' �������� �� ��ħ 
    ret = .AddCellSpan(C_W3, -1000, 1, 2)	' ��꼭����Աݾ� �� ��ħ 
    ret = .AddCellSpan(C_W4, -1000, 2, 1)	' ���� �� ��ħ 
    ret = .AddCellSpan(C_W6, -1000, 1, 2)	' �����ļ��Աݾ� �� ��ħ 
    ret = .AddCellSpan(C_DESC1, -1000, 1, 2)	' ��� �� ��ħ 

       
     ' ù��° ��� ��� ���� 
	.Row = -1000
	.Col = C_W1_CD
	.Text = "�� �� �� ��"
	.Col = C_W4
	.Text = "��       ��"
		
	' �ι�° ��� ��� ���� 
	.Row = -999	
	.Col = C_W1_NM
	.Text = "(1)�� ��"
	.Col = C_W2_NM
	.Text = "(2)�� ��"
	.Col = C_W4	
	.Text = "(4)�� ��"
	.Col = C_W5
	.Text = "(5)�� ��"

	.rowheight(-999) = 15	' ���� ������ 
   	
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
 
	'����� 2�ٷ�    
    .ColHeaderRows = 2
    'Call AppendNumberPlace("6","3","2")

	ggoSpread.SSSetEdit		C_SEQ_NO,	"�θ����", 5,,,6,1	' �����÷� 
    ggoSpread.SSSetEdit		C_CHILD_SEQ_NO,	"�ڽļ���", 5,,,6,1	' �����÷� 
    ggoSpread.SSSetEdit		C_W2_NM2,	"�� ��"	, 10,,,50,1	
	ggoSpread.SSSetEdit		C_W7,		"(7)�����", 10,,,50,1
	ggoSpread.SSSetEdit		C_W8,		"(8)������", 10,,,50,1
    ggoSpread.SSSetFloat	C_W9,		"(9)���ޱݾ�" ,		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0","" 
	ggoSpread.SSSetFloat	C_W10,		"(10)���ػ�� ������" & vbCrLf & "�Ѱ���� ������" ,		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
    ggoSpread.SSSetFloat	C_W11,		"(11)�Ѱ��� ������",		13,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
    ggoSpread.SSSetEdit		C_W12,		"(12)������",		10, 2,,10,2
    ggoSpread.SSSetFloat	C_W13,		"(13)�����ͱ�" & vbCrLf & "���Ծ�" & vbCrLf & "[(9) * (12)]" ,13,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0",""
    ggoSpread.SSSetFloat	C_W14,		"(14)���⸻" & vbCrLf & "��������" & vbCrLf & "����" ,13,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0",""
    ggoSpread.SSSetFloat	C_W15,		"(15)���ȸ��" & vbCrLf & "���԰���" ,13,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0",""
    ggoSpread.SSSetFloat	C_W16,		"(16)������" & vbCrLf & "[(13) - (14) - (15)]" ,14,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0",""

	' �ۼ�Ʈ �� ���� 
    .Col = C_W12
    .Row = -1
    .CellType = 14
    .TypeHAlign = 2
    '.TypePercentDecimal = 0
    .TypePercentMax = 999
    '.TypePercentMin = 0
    '.TypePercentDecPlaces = 0
    
	' �׸��� ��� ��ħ ���� 
	'ret = .AddCellSpan(C_SEQ_NO, -1000, 1, 2)	' SEQ_NO �� ��ħ 
    'ret = .AddCellSpan(C_CHILD_SEQ_NO, -1000, 1, 2)	' SEQ_NO �� ��ħ 
    ret = .AddCellSpan(C_W2_NM2, -1000, 1, 2)	' ����� 
    ret = .AddCellSpan(C_W7, -1000, 1, 2)	' ����� 
    ret = .AddCellSpan(C_W8, -1000, 1, 2)	' ������ 
    ret = .AddCellSpan(C_W9, -1000, 1, 2)	' ���ޱݾ� 
    ret = .AddCellSpan(C_W10, -1000, 3, 1)	' �۾��������� 
    ret = .AddCellSpan(C_W13, -1000, 1, 2)	' �Ա޻��Ծ� 
    ret = .AddCellSpan(C_W14, -1000, 1, 2)	' ���⸻ 
    ret = .AddCellSpan(C_W15, -1000, 1, 2)	' ���ȸ�� 
    ret = .AddCellSpan(C_W16, -1000, 1, 2)	' ���� 
    
    ' ù��° ��� ��� ���� 
	.Row = -1000
	.Col = C_W10
	.Text = "�۾���������"

	' �ι�° ��� ��� ���� 
	.Row = -999
	.Col = C_W10
	.Text = "(10)���ػ�� ������" & vbCrLf & "�Ѱ���� ������"
	.Col = C_W11
	.Text = "(11)�Ѱ��� ������"
	.Col = C_W12
	.Text = "(12)������" & vbCrLf & "[(10)/(11)]"
	.rowheight(-999) = 20	' ���� ������ 
	
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_CHILD_SEQ_NO,True)
				
	'Call InitSpreadComboBox()
	
	.ReDraw = true
	
  'Call SetSpreadLock 
    
    End With

 	' -----  3�� �׸��� 
	With frm1.vspdData3
	
	ggoSpread.Source = frm1.vspdData3
   'patch version
    ggoSpread.Spreadinit "V20041222_2",,parent.gForbidDragDropSpread    
    
	.ReDraw = false
    
    .MaxCols = C_DESC2 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols									'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    ggoSpread.ClearSpreadData
   .MaxRows = 0
 
    'Call AppendNumberPlace("6","3","2")

    ggoSpread.SSSetEdit		C_SEQ_NO,	"�θ����", 5,,,6,1	' �����÷� 
    ggoSpread.SSSetEdit		C_CHILD_SEQ_NO,	"�ڽļ���", 5,,,6,1	' �����÷� 
    ggoSpread.SSSetEdit		C_W2_NM2,	"�� ��"	, 10,,,50,1	
	ggoSpread.SSSetEdit		C_W17,		"(17)�� ��", 20,,,50,1
	ggoSpread.SSSetEdit		C_W18,		"(18)�ٰŹ���", 20,,,50,1
    ggoSpread.SSSetFloat	C_W19,		"(19)���Աݾ�" ,		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0","" 
	ggoSpread.SSSetFloat	C_W20,		"(20)��������" ,		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	ggoSpread.SSSetEdit		C_DESC2,	"�� ��", 20,,,50,1
		
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_CHILD_SEQ_NO,True)
				
	'Call InitSpreadComboBox()
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
           
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    .vspdData2.ReDraw = False
    .vspdData3.ReDraw = False

	' 1�� �׸��� 
    ggoSpread.Source = frm1.vspdData
        
	ggoSpread.SpreadLock C_SEQ_NO, -1, C_W1_CD
	ggoSpread.SSSetRequired C_W1_NM, -1, -1
	ggoSpread.SSSetRequired C_W2_NM, -1, -1
	ggoSpread.SSSetRequired C_W3, -1, -1
    ggoSpread.SpreadLock C_W4, -1, C_W4
    ggoSpread.SpreadLock C_W5, -1, C_W5
    ggoSpread.SpreadLock C_W6, -1, C_W6    
    
    ' 2�� �׸��� 
    ggoSpread.Source = frm1.vspdData2	

    ggoSpread.SpreadLock C_W2_NM2, -1, C_W2_NM2
    ggoSpread.SpreadLock C_W12, -1, C_W12
    ggoSpread.SpreadLock C_W13, -1, C_W13
    ggoSpread.SpreadLock C_W16, -1, C_W16

	
	' 3�� �׸��� 
    ggoSpread.Source = frm1.vspdData3	

    'ggoSpread.SpreadLock C_W17, -1, C_W17
	ggoSpread.SpreadLock C_W2_NM2, -1, C_W2_NM2
    
	'ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
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
	
  		ggoSpread.SSSetProtected C_SEQ_NO, pvEndRow, pvEndRow
		ggoSpread.SSSetProtected C_CHILD_SEQ_NO, pvEndRow, pvEndRow
		ggoSpread.SSSetRequired C_W1_NM, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W2_NM, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W3, pvStartRow, pvEndRow 		
 		
 		ggoSpread.SSSetProtected C_W4, -1, -1
 		ggoSpread.SSSetProtected C_W5, -1, -1
 		ggoSpread.SSSetProtected C_W6, -1, -1
		    
		'.vspdData.ReDraw = True

    'End If
    
    End With
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColorDetail2(ByVal pvEndRow)
    With frm1
    
		' 2�� �׸��� 
		ggoSpread.Source = frm1.vspdData2	

		ggoSpread.SpreadLock C_SEQ_NO, -1, C_CHILD_SEQ_NO
		ggoSpread.SSSetProtected C_W2_NM2, pvEndRow, pvEndRow
		ggoSpread.SSSetProtected C_W12, pvEndRow, pvEndRow
		ggoSpread.SSSetProtected C_W13, pvEndRow, pvEndRow
		ggoSpread.SSSetProtected C_W16, pvEndRow, pvEndRow
    
    End With
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColorDetail3(ByVal pvEndRow)
    With frm1
    
		' 2�� �׸��� 
		ggoSpread.Source = frm1.vspdData3	

		ggoSpread.SpreadLock C_SEQ_NO, -1, C_CHILD_SEQ_NO
		ggoSpread.SSSetProtected C_W2_NM2, pvEndRow, pvEndRow

    
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

Sub InsertRow2Head()
	' fncNew, onLoad�ÿ� ȣ���ؼ� �⺻������ 3ĭ�� �Է��� 
	Dim ret, iRow, iLoop, iSeqNo
	
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
			
		.ReDraw = False

		iRow = 1
		ggoSpread.InsertRow ,1
		Call SetSpreadColor(iRow, iRow) 
		iSeqNo = MaxSpreadVal(frm1.vspdData, C_SEQ_NO, iRow)
		
		iRow = 2
		ggoSpread.InsertRow ,1
		Call SetSpreadColor(iRow, iRow) 
		.Col = C_SEQ_NO : .Row = iRow: .value = SUM_SEQ_NO
		.col = C_W1_CD : .text = "��" : .TypeHAlign = 2
				
		ggoSpread.SpreadLock C_W1_CD, iRow, C_DESC1-1, iRow
		ret = .AddCellSpan(C_W1_CD, iRow, 4, 1)	' �հ� ���� �������� ���� ��ħ 
		
		.ReDraw = True		
		.focus
		.SetActiveCell C_W1_NM, 1
					
	End With

	Call InsertRow2Detail2(iSeqNo)
	Call InsertRow2Detail3(iSeqNo)
	
	Call vspdData_Click(C_W1_NM, 1)
End Sub

Sub InsertRow2Detail2(Byval pSeqNo)

	' �۾������ �׸��� �߰� 
	Dim ret, iRow, iLoop, iLastRow, sW2_NM
	
	sW2_NM = GetGrid(frm1.vspdData, C_W2_NM, frm1.vspdData.ActiveRow)
	
	With frm1.vspdData2
		
		.focus
		ggoSpread.Source = frm1.vspdData2

		iLastRow = .MaxRows
		.SetActiveCell C_W2_NM2, iLastRow	
		
		.ReDraw = False
		For iRow = 1 to 2
			If iRow Mod 2 = 0 Then	' �հ��� 
			
				ggoSpread.InsertRow ,1
				.Row = iLastRow+iRow
				.Col = C_CHILD_SEQ_NO	: .value = SUM_SEQ_NO
				.Col = C_SEQ_NO			: .Text = pSeqNo
				.Col = C_W2_NM2			: .Text = "��"	: .TypeHAlign = 2	
				Call SetSpreadColorDetail2(iLastRow+iRow) 

				ggoSpread.SpreadLock C_W9, iLastRow+iRow, C_W16, iLastRow+iRow
				ret = .AddCellSpan(C_W2_NM2, iLastRow+iRow, 3, 1)	' �հ� ���� �������� ���� ��ħ	
				ggoSpread.SpreadLock 1, iLastRow+iRow, C_W16, iLastRow+iRow	
				.RowHidden = True	
			Else	 ' �׿� 
			
				ggoSpread.InsertRow ,1
				.Row = iLastRow+iRow
				.Col = C_CHILD_SEQ_NO	: .Text = iRow
				.Col = C_SEQ_NO			: .Text = pSeqNo
				.Col = C_W2_NM2			: .Text = sW2_NM
				Call SetSpreadColorDetail2(iLastRow+iRow) 
				.RowHidden = True
			End If
		Next

		.ReDraw = True		
		.SetActiveCell C_W2_NM2, iLastRow+1	

	End With
	
End Sub

Sub InsertRow2Detail3(Byval pSeqNo)

	' �۾������ �׸��� �߰� 
	Dim ret, iRow, iLoop, iLastRow, sW2_NM
	
	sW2_NM = GetGrid(frm1.vspdData, C_W2_NM, frm1.vspdData.ActiveRow)
	
	' ��Ÿ���Աݾ� �׸���	
	With frm1.vspdData3
		
		.focus
		ggoSpread.Source = frm1.vspdData3

		iLastRow = .MaxRows
		.SetActiveCell C_W2_NM2, iLastRow	
	
		.ReDraw = False
		For iRow = 1 to 2
			If iRow Mod 2 = 0 Then	' �հ��� 
			
				ggoSpread.InsertRow ,1
				.Row = iLastRow+iRow
				.Col = C_CHILD_SEQ_NO	: .Text = "999999"
				.Col = C_SEQ_NO			: .Text = pSeqNo
				.Col = C_W2_NM2			: .Text = "��"	: .TypeHAlign = 2	
				Call SetSpreadColorDetail3(iLastRow+iRow) 
				
				'ggoSpread.SpreadLock C_W9, iLoop+iRow-1, C_W16, iLoop+iRow-1
				ret = .AddCellSpan(C_W2_NM2, iLastRow+iRow, 3, 1)	' �հ� ���� �������� ���� ��ħ	
				ggoSpread.SpreadLock 1, iLastRow+iRow, C_DESC2, iLastRow+iRow	
				.RowHidden = True	
			Else	 ' �׿� 
			
				ggoSpread.InsertRow ,1
				.Row = iLastRow+iRow
				.Col = C_CHILD_SEQ_NO	: .Text = iRow
				.Col = C_SEQ_NO			: .Text = pSeqNo	
				.Col = C_W2_NM2			: .Text = sW2_NM
				Call SetSpreadColorDetail3(iLastRow+iRow) 

				.RowHidden = True
			End If
		Next
	
		.ReDraw = True	
		.SetActiveCell C_W2_NM2, iLastRow+1	
	
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
				.col = C_W1_CD : .text = "��" : .TypeHAlign = 2
				
				ggoSpread.SpreadLock C_W1_CD, iRow, C_DESC1-1, iRow
				ret = .AddCellSpan(C_W1_CD, iRow, 4, 1)	' �հ� ���� �������� ���� ��ħ	
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
			
				.Col = C_W2_NM2			: .Text = "��"	: .TypeHAlign = 2	
				Call SetSpreadColorDetail2(iRow) 

				ggoSpread.SpreadLock C_W9, iRow, C_W16, iRow
				ret = .AddCellSpan(C_W2_NM2, iRow, 3, 1)	' �հ� ���� �������� ���� ��ħ	
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
			
				.Col = C_W2_NM2			: .Text = "��"	: .TypeHAlign = 2	
				Call SetSpreadColorDetail3(iRow) 

				ret = .AddCellSpan(C_W2_NM2, iRow, 3, 1)	' �հ� ���� �������� ���� ��ħ	
				ggoSpread.SpreadLock 1, iRow, C_DESC2, iRow	
			End If
		Next
	End With
End Sub

'============================== ����� ���� �Լ�  ========================================
' -- ���� ���� ó�� 
Function ShowRowHidden(Byref pObj, Byval pSeqNo)
	Dim iRow, iSeqNo, iMaxRows, iFirstRow
	
	With pObj
	
	iMaxRows = .MaxRows : iFirstRow = 0
	pObj.ReDraw = False
	For iRow = 1 To iMaxRows
		.Col = C_SEQ_NO : .Row = iRow : iSeqNo = .Value
		If iSeqNo = pSeqNo Then	' ���� ������..
			.RowHidden = False
			If iFirstRow = 0 Then iFirstRow = iRow
		Else
			.RowHidden = True
		End If
	Next
	pObj.ReDraw = True
	ShowRowHidden = iFirstRow
	End With
End Function

Function ShowRowHidden2(Byref pObj, Byval pSeqNo)
	Dim iRow, iSeqNo, iMaxRows, iFirstRow
	
	With pObj
	
	iMaxRows = .MaxRows : iFirstRow = 0
	pObj.ReDraw = False
	For iRow = 1 To iMaxRows
		.Col = C_CHILD_SEQ_NO : .Row = iRow : iSeqNo = .Value
		If iSeqNo = pSeqNo Then	' ���� ������..
			.RowHidden = False
			If iFirstRow = 0 Then iFirstRow = iRow
		Else
			.RowHidden = True
		End If
	Next
	pObj.ReDraw = True
	ShowRowHidden2 = iFirstRow
	End With
End Function

' -- �հ� ������ üũ 
Function CheckTotalRow(Byref pObj, Byval pRow) 
	CheckTotalRow = False
	pObj.Col = C_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If CDbl(pObj.Text) = 999999 Then	 ' �հ� �� 
		CheckTotalRow = True
	End If
End Function

' -- �հ� ������ üũ 
Function CheckTotalRow2(Byref pObj, Byval pRow) 
	CheckTotalRow2 = False
	pObj.Col = C_CHILD_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If CDbl(pObj.Text) = 999999 Then	 ' �հ� �� 
		CheckTotalRow2 = True
	End If
End Function

' -- ���� ������ �Ʒ� �׸��忡 ǥ�� 
Sub	SetW2ToChildGrid(Byval pW2)
	Dim i, iMaxRows, iLastRow, iSeqNo
	
	frm1.vspdData2.Col = C_SEQ_NO: frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
	iSeqNo = frm1.vspdData2.Value
	
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		iMaxRows = .MaxRows 
		For i = 1 to iMaxRows
			.Col = C_SEQ_NO : .Row = i 
			If iSeqNo = .VAlue Then
				If CheckTotalRow2(frm1.vspdData2, i) = False Then 
					.Col = C_W2_NM2 : .Row = i : .Value = pW2 
					ggoSpread.UpdateRow .Row
				End If
			End If
		Next
	End With

	With frm1.vspdData3
		ggoSpread.Source = frm1.vspdData3
		iMaxRows = .MaxRows
		For i = 1 to iMaxRows
			.Col = C_SEQ_NO : .Row = i 
			If iSeqNo = .VAlue Then
				If CheckTotalRow2(frm1.vspdData3, i) = False Then 
					.Col = C_W2_NM2 : .Row = i : .Value = pW2 
					ggoSpread.UpdateRow .Row
				End If
			End If
		Next
	End With
	
End Sub 

' --- W12�� ����Ÿ ��� 
Function SetW12(Byval pRow)
	Dim dblW10, dblW11
	With frm1.vspdData2
		.Col = C_W10	: .Row = pRow	: dblW10 = CDbl(.Value)
		.Col = C_W11	: .Row = pRow	: dblW11 = CDbl(.Value)
		.Col = C_W12	: .Row = pRow
		If dblW11 > 0 Then
			.Value = dblW10/dblW11
		Else
			.Value = 0
		End If
	End With
End Function

' --- W13�� ����Ÿ ��� 
Function SetW13(Byval pRow)
	Dim dblW9, dblW12
	With frm1.vspdData2
		.Col = C_W9		: .Row = pRow	: dblW9 = CDbl(.Value)
		.Col = C_W12	: .Row = pRow	: dblW12 = CDbl(.Value)
		.Col = C_W13	: .Row = pRow
		If dblW12 > 0 Then
			.Value = (dblW9*dblW12)
		Else
			.Value = 0
		End If
	End With
End Function

' --- W16�� ����Ÿ ��� 
Function SetW16(Byval pRow)
	Dim dblW13, dblW14, dblW15
	With frm1.vspdData2
		.Col = C_W13	: .Row = pRow	: dblW13 = CDbl(.Value)
		.Col = C_W14	: .Row = pRow	: dblW14 = CDbl(.Value)
		.Col = C_W15	: .Row = pRow	: dblW15 = CDbl(.Value)
		.Col = C_W16	: .Row = pRow	: .Value = dblW13 - dblW14 - dblW15
	End With
End Function

' 1���׸��� W4(����)�� �ֱ� 
Function SetW4_W5()
	Dim dblGrid2Sum, dblW19Sum, dblW20Sum, dblSum
	
	' 
	With frm1.vspdData3
		dblW19Sum = GetSum(frm1.vspdData3, C_W19)
		dblW20Sum = GetSum(frm1.vspdData3, C_W20)
	End With
	
	With frm1.vspdData2
		dblGrid2Sum = GetSum(frm1.vspdData2, C_W16)
	End With

	With frm1.vspdData
		
		If 	dblGrid2Sum > 0 Then
			.Col = C_W4	: .Row = .ActiveRow	: .Value = dblGrid2Sum + dblW19Sum
			.Col = C_W5	: .Row = .ActiveRow	: .Value = dblW20Sum
		Else
			.Col = C_W5	: .Row = .ActiveRow	: .Value = ABS(dblGrid2Sum) + dblW20Sum
			.Col = C_W4	: .Row = .ActiveRow	: .Value = dblW19Sum
		End If

			dblSum = FncSumSheet(frm1.vspdData, C_W4, 1, .MaxRows - 1, false, -1, -1, "V")	' ���� �÷� ���հ� 
			.Col = C_W4 : .Row = .MaxRows : .Value = dblSum
			dblSum = FncSumSheet(frm1.vspdData, C_W5, 1, .MaxRows - 1, false, -1, -1, "V")	' ���� �÷� ���հ� 
			.Col = C_W5 : .Row = .MaxRows : .Value = dblSum
			
			Call SetW6(.ActiveRow)
	End With
	
	
End Function

' -- ���� ���̴� �׸����� Ư���÷��� �հ����� ���� �о�´�.
Function GetSum(Byref pGrid, Byval pCol)
	Dim iRow, iMaxRows, iSeqNo
	
	With pGrid
		iMaxRows = .MaxRows
		.Row = .ActiveRow	: .Col = C_SEQ_NO : iSeqNo = .Value
		For iRow = 1 To iMaxRows
			.Row = iRow : .Col = C_SEQ_NO
			If .Value = iSeqNo Then
				.Col = C_CHILD_SEQ_NO 
				If UNICDbl(.Value) = SUM_SEQ_NO Then
					.Col = pCol
					GetSum = UNICDbl(.Value)
					Exit Function
				End If
			End If
		Next
	End With
End Function

' -- W6(�����ļ��Աݾ�)
Function SetW6(Byval Row)
	Dim dblW3, dblW4, dblW5, dblSum
	With frm1.vspdData
		.Col = C_W3	: .Row = Row	: dblW3 = CDbl(.value)
		.Col = C_W4	: .Row = Row	: dblW4 = CDbl(.value)
		.Col = C_W5	: .Row = Row	: dblW5 = CDbl(.value)
		.Col = C_W6	: .Row = Row	: .value = dblW3 + dblW4 - dblW5
		
		dblSum = FncSumSheet(frm1.vspdData, C_W6, 1, .MaxRows - 1, false, -1, -1, "V")	' ���� �÷� ���հ� 
		.Col = C_W6 : .Row = .MaxRows : .Value = dblSum
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow .ActiveRow
		ggoSpread.UpdateRow .MaxRows
	End With
End Function

'============================== ���۷��� �Լ�  ========================================
Function GetRef()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD, iMaxRows, iRow
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

	arrParam(0) = frm1.txtCO_CD.value 
	arrParam(1) = frm1.txtFISC_YEAR.text 
	arrParam(2) = frm1.cboREP_TYPE.value 

    arrRet = window.showModalDialog("w2105ra1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
	If arrRet(0, 0) = "" Then
	    Exit Function
	End If	

	
	With frm1.vspdData
		.Redraw = False
		
		lgBlnFlgChgValue = True
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData3
		ggoSpread.ClearSpreadData
		iMaxRows = UBound(arrRet, 1)
		lgCurrGrid = 1
		Call FncInsertRow(1)
		If iMaxRows > 1 Then Call FncInsertRow(iMaxRows-1)
		
		For iRow = 1 To iMaxRows
			.Row = iRow	
			Call vspdData_Click(C_W1_NM, iRow)
			.Row = iRow	
			'.Col = C_W1_CD	: .Value = arrRet(iRow, 3)
			.Col = C_W1_NM	: .Value = arrRet(iRow-1, 5)
			.Col = C_W2_CD	: .Value = arrRet(iRow-1, 3)
			.Col = C_W2_NM	: .Value = arrRet(iRow-1, 4)
			.Col = C_W3		: .Value = arrRet(iRow-1, 2)
			.Col = C_W6		: .Value = arrRet(iRow-1, 2)

			Call vspdData_Change(C_W2_NM, iRow)
			Call vspdData_Change(C_W3, iRow)
		Next
		.Row = 1
		.Col = 	C_W1_NM
		.Action = 0
		.Redraw = True
		
		Call vspdData_Click(C_W1_NM, 1)
	End With
	
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
    
    Call SetToolbar("1100110100000111")										<%'��ư ���� ���� %>

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
End Sub



'============================================  1�� �׸��� �̺�Ʈ  ====================================
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
	    
	Select Case Col
		Case C_W2_NM	' ����� 
			.Col = C_W2_NM
			Call SetW2ToChildGrid(.Value)	' ���� ������ ���� �׸��忡 �ִ´�.
		Case C_W3		' ��꼭�� ���Աݾ� 
			dblSum = FncSumSheet(frm1.vspdData, Col, 1, .MaxRows - 1, false, -1, -1, "V")
			.Col = Col : .Row = .MaxRows : .Value = dblSum	
			Call SetW6(Row)
	End Select
	End With
End Sub


Sub vspdData_Click(ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("0001011111") 

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

	' -- Ŀ�� �̵��� �ϴ� ���� (_Click�̺�Ʈ���� �̵�: 200603)
	Dim iSeqNo, IntRetCD, iLastRow
	
	If lgOldRow = Row  Then Exit Sub
	
	ggoSpread.Source = frm1.vspdData
  
	If Row = frm1.vspdData.MaxRows Then
		iLastRow = ShowRowHidden2(frm1.vspdData2, iSeqNo)
		iLastRow = ShowRowHidden2(frm1.vspdData3, iSeqNo)

	Else
		With frm1.vspdData
			.Col = C_SEQ_NO : .Row = Row : iSeqNo = .Value
				
			' ���� �׸��� ǥ�÷�ƾ'
			iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
			frm1.vspdData2.SetActiveCell C_W7, iLastRow
				
			If iLastRow = 0 Then 
				Call InsertRow2Detail2(iSeqNo)
				iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
			End If

			iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
			frm1.vspdData3.SetActiveCell C_W17, iLastRow
	
			If iLastRow = 0 Then 
				Call InsertRow2Detail3(iSeqNo)
				iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
			End If

			.focus
		End With	
	End If
  
	lgOldRow = Row	: lgOldCol = Col

End Sub

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	If Row <> NewRow And (Row > 0 And NewRow > 0)  Then
		window.setTimeout "vbscript:vspdData_Click " & NewCol & "," & NewRow & " ", 500
	End If
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
	Dim dblSum, dblW10, dblW11
	
	With Frm1.vspdData2
	.Row = Row
	.Col = Col

	If .CellType = parent.SS_CELL_TYPE_FLOAT Then
		If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
		   .text = .TypeFloatMin
		End If
	End If
		
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
	
    Select Case Col
		Case C_W9, C_W10, C_W11, C_W12, C_W13, C_W14, C_W15
			.Col = C_W10	: dblW10 = UNICDbl(.value)
			.Col = C_W11	: dblW11 = UNICDbl(.value)
			If dblW11 < dblW10 And (dblW11 <> 0 And dblW10 <> 0) Then
				Call DisplayMsgBox("WC0010", "X", GetGrid(frm1.vspdData2, C_W10, -999), GetGrid(frm1.vspdData2, C_W11, -999)) 
				.Row = Row
				.Col = Col	: .value = 0
				frm1.vspdData2.focus
			End If
		
			dblSum = ufn_FncSumSheet(frm1.vspdData2, Col)
			Call SetW12(Row)			' ����� ��� 
			Call SetW13(Row)			' �ͱݻ��Ծ� ��� 
			Call SetW16(Row)			' ������ ��� 
			dblSum = ufn_FncSumSheet(frm1.vspdData2, C_W16)
			
			Call SetW4_W5				' ��� �׸��� �ݿ� 

    End Select

	End With
	lgChgFlg = True ' ����Ÿ ���� 
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y )
	lgCurrGrid = 2
	ggoSpread.Source = Frm1.vspdData2
	
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)

End Sub

' -- ���� ���̴� �θ������ ��ġ�ϴ� ������� ���� ���Ѵ�.
Function ufn_FncSumSheet(Byref pGrid, Byval pCol)
	Dim dblSeqNo, iRow, iMaxRows, dblSum
	
	With pGrid
		.ReDraw = False
		iMaxRows = .MaxRows
		.Col = C_SEQ_NO	: dblSeqNO = UNICDbl(.value)
		
		For iRow = 1 To iMaxRows
			.Col = C_SEQ_NO : .Row = iRow
			If UNICDbl(.value) = dblSeqNo Then	' �θ������ ���� �� 
				.Col = C_CHILD_SEQ_NO	
				If UNICDbl(.value) < SUM_SEQ_NO Then
					.Col = pCol	: dblSum = dblSum + UNICDbl(.value)
				ElseIf UNICDbl(.value) = SUM_SEQ_NO Then
					.Col = pCol	: .Value = dblSum	' �հ��࿡ ����Ÿ ��� 
					ggoSpread.UpdateRow .Row
					ufn_FncSumSheet = dblSum
					Exit For
				End If
			End If
		Next
		.ReDraw = True
	End With
End Function

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
		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.UpdateRow Row
	
    Select Case Col
		Case C_W19, C_W20
			dblSum = ufn_FncSumSheet(frm1.vspdData3, Col)
			
			Call SetW4_W5()
			
    End Select

	End With
	lgChgFlg = True ' ����Ÿ ���� 
End Sub

Sub vspdData3_MouseDown(Button , Shift , x , y )
	lgCurrGrid = 3
	ggoSpread.Source = Frm1.vspdData3
End Sub

'============================================  �������� �Լ�  ====================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
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
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
        
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    
    If lgChgFlg = False Then
    
		If ggoSpread.SSCheckChange = False Then
		    Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		    Exit Function
		End If
		
	End If
	
	If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
	      Exit Function
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

    Call InsertRow2Head
    'Call InsertRow2Detail(1)
    
    Call SetToolbar("1100111100000111")

	Call InitData
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
				Else
					lDelRows = ggoSpread.EditUndo
				End If
				
			End With
		CAse 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				If CheckTotalRow2(frm1.vspdData2, .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.EditUndo
				End If
			End With    
 		CAse 3
			With frm1.vspdData3 
				.focus
				ggoSpread.Source = frm1.vspdData3
				If CheckTotalRow2(frm1.vspdData3, .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.EditUndo
				End If
			End With     
	End Select
  
	lgChgFlg = True                                                '��: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo, iLastRow, sW2_NM

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG
    
    ' ��Ƽ ���� �������� �ʴ´�. �����׸��尡 �����׸��忡 �����־� ������ 
	imRow = CInt(pvRowCnt)
		 
	' ù���� ��� �հ���� �ִ� ��ƾ 
	If frm1.vspdData.MaxRows = 0 Then
		Call InsertRow2Head
		Call SetToolbar("1100111100000111")
		Exit Function
	End If
		
	Select Case lgCurrGrid
		Case 1	' 1�� �׸��� 
		
		With frm1.vspdData
			
		.focus
		ggoSpread.Source = frm1.vspdData
			
		iRow = .ActiveRow	' ������ 
			
		.ReDraw = False
			
		If iRow = .MaxRows Then
			ggoSpread.InsertRow iRow-1 , imRow 
			SetSpreadColor iRow, iRow+imRow	' �׸��� ���󺯰� 
			iSeqNo = MaxSpreadVal(frm1.vspdData, C_SEQ_NO, iRow)
		Else
			ggoSpread.InsertRow ,imRow
			SetSpreadColor iRow+1, iRow+imRow	' �׸��� ���󺯰� 
			iSeqNo = MaxSpreadVal(frm1.vspdData, C_SEQ_NO, iRow+1)
		End If	

		.ReDraw = True	
						
		' ���� �׸��� ǥ�÷�ƾ'
		iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)
		frm1.vspdData2.SetActiveCell C_W7, iLastRow
			
		If iLastRow = 0 Then Call InsertRow2Detail2(iSeqNo)

		iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
		frm1.vspdData3.SetActiveCell C_W7, iLastRow
			
		If iLastRow = 0 Then Call InsertRow2Detail3(iSeqNo)
			
		Call vspdData_Click(.Col, .ActiveRow)
		
		frm1.vspdData.SetActiveCell C_W1_NM, .ActiveRow
		End With
		
	Case 2	' 2�� �׸��� 
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_W2_NM
		sW2_NM = frm1.vspdData.value
		
		With frm1.vspdData2
		.focus
		ggoSpread.Source = frm1.vspdData2
			
		iRow = .ActiveRow
		.ReDraw = False
					
		iSeqNo = GetGRid(frm1.vspdData, C_SEQ_NO, frm1.vspdData.ActiveRow)	' �θ�׸����� ��ġ���� 

		If .MaxRows = 0 Then
			Call InsertRow2Detail2(iSeqNo)			
		ElseIf iRow = .MaxRows And iRow > 0 Then
			ggoSpread.InsertRow iRow-1 , imRow 
			SetSpreadColorDetail2 iRow-1
			MaxSpreadVal2 frm1.vspdData2, C_SEQ_NO, C_CHILD_SEQ_NO, iRow , iSeqNo
			.Row = iRow
		Else
			ggoSpread.InsertRow ,imRow
			SetSpreadColorDetail2 iRow+1
			MaxSpreadVal2 frm1.vspdData2, C_SEQ_NO, C_CHILD_SEQ_NO, iRow+1, iSeqNo	
			.Row = iRow+1
		End If
		.Col = C_W2_NM2 : .value = sW2_NM
		.ReDraw = True
		End With
		
	Case 3	' 3�� �׸��� 
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_W2_NM
		sW2_NM = frm1.vspdData.value
		
		With frm1.vspdData3
		.focus
		ggoSpread.Source = frm1.vspdData3

		iRow = .ActiveRow
		.ReDraw = False
		
		iSeqNo = GetGRid(frm1.vspdData, C_SEQ_NO, frm1.vspdData.ActiveRow)	' �θ�׸����� ��ġ���� 
		
		If .MaxRows = 0 Then
			Call InsertRow2Detail3(iSeqNo)
		ElseIf iRow = .MaxRows And iRow > 0 Then
			ggoSpread.InsertRow iRow-1 , imRow 
			SetSpreadColorDetail3 iRow-1
			MaxSpreadVal2 frm1.vspdData3, C_SEQ_NO, C_CHILD_SEQ_NO, iRow	, iSeqNo
			.Row = iRow
		Else
			ggoSpread.InsertRow iRow,imRow
			SetSpreadColorDetail3 iRow+1
			MaxSpreadVal2 frm1.vspdData3, C_SEQ_NO, C_CHILD_SEQ_NO, iRow+1, iSeqNo	
			.Row = iRow+1
		End If	
		.Col = C_W2_NM2 : .value = sW2_NM
		.ReDraw = True
		End With
	End Select

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

' -- �׸��� ��ǥ�� �� �б� 
Function GetGrid(Byref pGrid, Byval pCol, Byval pRow)
	With pGrid
		.Col = pCol : .Row = pRow : GetGrid = .Value
	End With
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
				Else
					lDelRows = ggoSpread.DeleteRow
				End If
				
			End With
		Case 2
			With frm1.vspdData2 
				.focus
				ggoSpread.Source = frm1.vspdData2
				If CheckTotalRow(frm1.vspdData2, .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.DeleteRow
				End If
				lDelRows = ggoSpread.DeleteRow
			End With    
 		Case 3
			With frm1.vspdData3 
				.focus
				ggoSpread.Source = frm1.vspdData3
				If CheckTotalRow(frm1.vspdData3, .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.DeleteRow
				End If
				lDelRows = ggoSpread.DeleteRow
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
    If frm1.vspdData.MaxRows > 0 Then
    
		lgIntFlgMode = parent.OPMD_UMODE
		    
		Call SetToolbar("1101111100000111")										<%'��ư ���� ���� %>
	
		Call RedrawSumRow
		Call RedrawSumRow2
		Call RedrawSumRow3

		With frm1.vspdData
			.Col = C_SEQ_NO : .Row = 1 : iSeqNo = .Value
				
			' ���� �׸��� ǥ�÷�ƾ'
			iLastRow = ShowRowHidden(frm1.vspdData2, iSeqNo)

			' ���� �׸��� ǥ�÷�ƾ'
			iLastRow = ShowRowHidden(frm1.vspdData3, iSeqNo)
		End With	
	
		' �������� ���� : ���ߵǸ� ���ȴ�.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
		If wgConfirmFlg = "Y" Then
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLock -1, -1
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLock -1, -1
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SpreadLock -1, -1
			
			Call SetToolbar("1100000000000111")	
		Else
			Call vspdData_Click(C_W1_NM, 1)
		End If
	Else
		Call SetToolbar("1100110100000111")	
	End If
	lgOldRow =0
	
	frm1.vspdData.focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow , lCol, lGrpCnt, lMaxRows, lMaxCols
    Dim lStartRow, lEndRow , lChkAmt
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
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
		                                          strDel = strDel & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		 .Col = 0
		  ' ��� �׸��� ����Ÿ ����     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = C_SEQ_NO To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Value) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
 
		Next
	End With

    frm1.txtSpread.value      = strDel & strVal
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
					.Col = C_W16 : lChkAmt = .Value
					If lChkAmt = 0 Then
		                                          strVal = strVal & "I"  &  Parent.gColSep ' ���� �߰��� �ڵ� 
		            Else
		                                          strVal = strVal & "C"  &  Parent.gColSep	
		            End If
		            lGrpCnt = lGrpCnt + 1
		                    
		       Case  ggoSpread.UpdateFlag                                      '��: Update                                                  
					.Col = C_W16 : lChkAmt = .Value
					If lChkAmt = 0 Then
		                                          strVal = strVal & "I"  &  Parent.gColSep ' ���� �߰��� �ڵ� 
		            Else
		                                          strVal = strVal & "U"  &  Parent.gColSep
		            End If                                                   
		            lGrpCnt = lGrpCnt + 1                                                 
		       Case  ggoSpread.DeleteFlag                                      '��: Delete
		                                          strDel = strDel & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		 .Col = 0
		  ' ��� �׸��� ����Ÿ ����     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = C_SEQ_NO To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Value) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
 
		Next
	End With

    frm1.txtSpread2.value      = strDel & strVal
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
		                                          strDel = strDel & "D"  &  Parent.gColSep
		            lGrpCnt = lGrpCnt + 1  
		  End Select
		 .Col = 0
		  ' ��� �׸��� ����Ÿ ����     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = C_SEQ_NO To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Value) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
 
		Next
	End With
	
    frm1.txtSpread3.value      = strDel & strVal
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">���Աݾ� ��ȸ</A>  
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
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="�������" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
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
						<TABLE <%=LR_SPACE_TYPE_20%> BORDER=0>
                            <TR HEIGHT=10>
                                <TD WIDTH="100%">1. ���Աݾ� �������</TD>
                            </TR>
                            <TR HEIGHT=30%>
								<TD WIDTH="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="25" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
                            <TR HEIGHT=10>
                                <TD WIDTH="100%">2. ���Աݾ� ������</TD>
                            </TR>
                            <TR HEIGHT=60%>
								<TD WIDTH="100%" VALIGN=TOP HEIGHT=100%>
								<TABLE <%=LR_SPACE_TYPE_20%> BORDER=0>
									<TR HEIGHT=10>
									    <TD WIDTH="100%">&nbsp;&nbsp;&nbsp;��. �۾�������� ���� ���Աݾ�</TD>
									</TR>
									<TR HEIGHT=60%>
										<TD WIDTH="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="25" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR HEIGHT=10>
									    <TD WIDTH="100%">&nbsp;&nbsp;&nbsp;��. ��Ÿ ���Աݾ�</TD>
									</TR>
									<TR HEIGHT=40%>
										<TD WIDTH="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="25" TITLE="SPREAD" id=vaSpread3> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
								</TABLE>
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><����>���Աݾ��������</LABEL>&nbsp;
				                                 <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><����>��.�۾����࿡ ���� ���Աݾ�</LABEL>&nbsp;
				                                 <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check3" ><LABEL FOR="prt_check3"><����>��.��Ÿ���Աݾ�</LABEL>&nbsp;</TD>
                </TR>
			
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('PRINT')"   Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread3" tag="24"></TEXTAREA>
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

