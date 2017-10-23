<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��Ÿ���� 
'*  3. Program ID           : w9103mA1
'*  4. Program Name         : w9103mA1.asp
'*  5. Program Desc         : ��47ȣ �ֿ��������(��)
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 2005/02/24
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : 
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

Const BIZ_MNU_ID		= "w9103mA1"
Const BIZ_PGM_ID		= "w9103mB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "w9103mB2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID		= "w9103OA1"

Const TYPE_1	= 1		' �׸��带 �������� ���� ��� 
Const TYPE_2	= 2		
Const TYPE_3	= 3		 
Const TYPE_4	= 4		 

' -- �׸��� �÷� ���� 
Dim C_W1
Dim C_W1_NM1
Dim C_W1_NM2
Dim C_W2
Dim C_W2_NM
Dim C_W3
Dim C_W3_NM
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7

Dim C_W8
Dim C_W8_NM
Dim C_W9
Dim C_W10
Dim C_W11
Dim C_W12
Dim C_W13

Dim C_W14
Dim C_W15
Dim C_W16
Dim C_W17
Dim C_W18
Dim C_W19

Dim C_W20
Dim C_W20_NM
Dim C_W21
Dim C_W22
Dim C_W23
Dim C_W24

Dim C_101
Dim C_102
Dim C_103
Dim C_104
Dim C_105
Dim C_106
Dim C_107

Dim C_108
Dim C_109
Dim C_110

Dim C_111
Dim C_112
Dim C_113

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(5)

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_W1		= 1
	C_W1_NM1	= 2
	C_W1_NM2	= 3
	C_W2		= 4
	C_W2_NM		= 5
	C_W3		= 6
	C_W3_NM		= 7
	C_W4		= 8
	C_W5		= 9
	C_W6		= 10
	C_W7		= 11

	C_W8		= 1
	C_W8_NM		= 2
	C_W9		= 3
	C_W10		= 4
	C_W11		= 5
	C_W12		= 6
	C_W13		= 7

	C_W14		= 1
	C_W15		= 2
	C_W16		= 3
	C_W17		= 4
	C_W18		= 5
	C_W19		= 6

	C_W20		= 1
	C_W20_NM	= 2
	C_W21		= 3
	C_W22		= 4
	C_W23		= 5
	C_W24		= 6
	
	' ���� ���� ��ȣ ���� 
	C_101		= 1
	C_102		= 2
	C_103		= 3
	C_104		= 4
	C_105		= 5
	C_106		= 6
	C_107		= 7

	C_108		= 1
	C_109		= 2
	C_110		= 3

	C_111		= 1
	C_112		= 2
	C_113		= 3

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



'============================================  �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	' ��ȸ����(����)
	Dim IntRetCD1
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
	
End Sub


Sub InitSpreadComboBox()
	Dim IntRetCD1

	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " MAJOR_CD='W1057' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		Call Spread_SetCombo(Replace(lgF0, chr(11), chr(9)), C_W2, 		C_101, C_104)
		Call Spread_SetCombo(Replace(lgF1, chr(11), chr(9)), C_W2_NM,	C_101, C_104)
		Call Spread_SetCombo(Replace(lgF0, chr(11), chr(9)), C_W3, 		C_101, C_104)
		Call Spread_SetCombo(Replace(lgF1, chr(11), chr(9)), C_W3_NM,	C_101, C_104)
	End If

	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " MAJOR_CD='W1058' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  
	
	If IntRetCD1 <> False Then
		Call Spread_SetCombo(Replace(lgF0, chr(11), chr(9)), C_W2,		C_105, C_106)
		Call Spread_SetCombo(Replace(lgF1, chr(11), chr(9)), C_W2_NM,	C_105, C_106)
		Call Spread_SetCombo(Replace(lgF0, chr(11), chr(9)), C_W3,		C_105, C_106)
		Call Spread_SetCombo(Replace(lgF1, chr(11), chr(9)), C_W3_NM,	C_105, C_106)
	End If

End Sub


' Col, Row1~Row2 ���� �޺��� �����. : ǥ�ؿ� ��� ���� ������ 
Sub Spread_SetCombo(pVal, pCol1, pRow1, pRow2)

	With lgvspdData(TYPE_1)

		.BlockMode = True
		.Col = pCol1	: .Col2 = pCol1
		.Row = pRow1	: .Row2 = pRow2
		.CellType = 8	'SS_CELL_TYPE_COMBOBOX

		.TypeComboBoxList = pVal	

		.TypeComboBoxEditable = False
		.TypeComboBoxMaxDrop = 0
		' Select the first item in the list
		'.TypeComboBoxCurSel = 0
		' Set the width to display the widest item in the list
		'.TypeComboBoxWidth = 1 
		.BlockMode = False
	End With

End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1)		= frm1.vspdData1
	Set lgvspdData(TYPE_2)		= frm1.vspdData2
	Set lgvspdData(TYPE_3)		= frm1.vspdData3
	Set lgvspdData(TYPE_4)		= frm1.vspdData4
		
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("6","3","2")
	
	' 1�� �׸��� 

	With lgvspdData(TYPE_1)
				
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W7 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
 
		.MaxRows = 0
		ggoSpread.ClearSpreadData
		
		'����� 2�ٷ�    
	    .ColHeaderRows = 2

	    ggoSpread.SSSetEdit		C_W1,		"(1)�ڻ꺰",		5,,,6,1	' ���� PK �÷� 
	    ggoSpread.SSSetEdit		C_W1_NM1,	"(1)�ڻ꺰",		5,,,21,1
		ggoSpread.SSSetEdit		C_W1_NM2,	"(1)�ڻ꺰",		14,,,15,1
	    ggoSpread.SSSetCombo	C_W2,		"(2)�Ű���",		10
	    ggoSpread.SSSetCombo	C_W2_NM,	"(2)�Ű���",		10
	    ggoSpread.SSSetCombo	C_W3,		"(3)�򰡹��",		10
	    ggoSpread.SSSetCombo	C_W3_NM,	"(3)�򰡹��",		10
		ggoSpread.SSSetFloat	C_W4,		"(4)ȸ����ݾ�",	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W5,		"(5)�Ű���",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W6,		"(6)���Լ����",	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W7,		"(7)������{(5)�Ǵ� (5)�� (6)�� ū �ݾ�-(4)}",	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 

		ret = .AddCellSpan(0		, -1000, 1, 2)
		ret = .AddCellSpan(C_W1		, -1000, 1, 2)
		ret = .AddCellSpan(C_W1_NM1, -1000, 2, 2)	' (1)�ڻ꺰 ��ħ 
		ret = .AddCellSpan(C_W2, -1000, 2, 2)		' (2)�Ű��� 
		ret = .AddCellSpan(C_W3, -1000, 2, 2)	' (3)�򰡹�� 
		ret = .AddCellSpan(C_W4, -1000, 1, 2)	' (4)ȸ����ݾ� 
		ret = .AddCellSpan(C_W5, -1000, 2, 1)	' (5)�Ű��� 
		ret = .AddCellSpan(C_W7, -1000, 1, 2)	' (7)������ 

	    ' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W5
		.Text = "�������ݾ�"
		.Col = C_W7
		.Text = "(7)������          {(5)�Ǵ� (5)�� (6)�� ū �ݾ�-(4)}"

		' �ι�° ��� ��� ���� 
		.Row = -999
		.Col = C_W5
		.Text = "(5)�Ű���"
		.Col = C_W6
		.Text = "(6)���Լ����"

		.rowheight(-999) = 15	' ���� ������ 

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W1,C_W1,True)
		Call ggoSpread.SSSetColHidden(C_W2,C_W2,True)
		Call ggoSpread.SSSetColHidden(C_W3,C_W3,True)

		.ReDraw = true	
				
	End With 

 
	' 2�� �׸��� 
	With lgvspdData(TYPE_2)
				
		ggoSpread.Source = lgvspdData(TYPE_2)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W13 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
	
		.MaxRows = 0
		ggoSpread.ClearSpreadData

	    ggoSpread.SSSetEdit		C_W8,	"����Ÿ����",		5,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W8_NM,	"(8)����",		20,,,50,1
		ggoSpread.SSSetFloat	C_W9,		"(9)�ݾ�",					15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W10,		"(10)��� �����ڻ갡��",	17,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W11,		"(11)ȸ��ձݰ���",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W12,		"(12)�ѵ��ʰ���{(11)-(10)}",21,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W13,		"(13)�̻����ͱݻ��Ծ�",	19,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W8,C_W8,True)
		
		.ReDraw = true	
				
	End With 

	' 3�� �׸��� 

	With lgvspdData(TYPE_3)
				
		ggoSpread.Source = lgvspdData(TYPE_3)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_3,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W19 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
	
		.MaxRows = 0
		ggoSpread.ClearSpreadData

  		'����� 2�ٷ�    
		.ColHeaderRows = 2
		
		ggoSpread.SSSetFloat	C_W14,		"(14)�����ޱ�",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W15,		"(15)������",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W16,		"(16)����{(14)-(15)}",	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W17,		"(17)��������",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W18,		"(18)ȸ�����",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W19,		"(19)������" & vbCrLf & "{(17)-(18)}",15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	
		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(0		, -1000, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W14	, -1000, 3, 1)	' ���� 2�� ��ħ 
		'ret = .AddCellSpan(C_W16	, -1000, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W17	, -1000, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W18	, -1000, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W19	, -1000, 1, 2)	
    
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W14		: .Text = "��     ��"
		
		' �ι�° ��� ��� ���� 
		.Row = -999
		.Col = C_W14
		.Text = "(14)�����ޱ�"
		.Col = C_W15
		.Text = "(15)������"
		.Col = C_W16
		.Text = "(16)����{(14)-(15)}"

		.rowheight(-999) = 15	' ���� ������ 
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		.ReDraw = true	
				
	End With

	' 4�� �׸��� 

	With lgvspdData(TYPE_4)
				
		ggoSpread.Source = lgvspdData(TYPE_4)
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_4,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W24 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
	
		.MaxRows = 0
		ggoSpread.ClearSpreadData

	    ggoSpread.SSSetEdit		C_W20	,	"(20)����",		5,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W20_NM,	"(20)����",		15,,,50,1
		ggoSpread.SSSetFloat	C_W21,		"(21)�Ǽ��ڱ�����",					15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W22,		"(22)ȸ�����",					15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W23,		"(23)�󰢴���ڻ��",				15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
		ggoSpread.SSSetFloat	C_W24,		"(24)����������{(21)-(22)-(23)}",	15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 

		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(C_W20		, -1000, 2, 1)	' ���� 2�� ��ħ 
		.rowheight(-1000) = 20	' ���� ������ 
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W20,C_W20,True)

				
		.ReDraw = true	
				
	End With
	
	Call InitSpreadRow()
'	Call SetSpreadLock()
'	Call InitSpreadComboBox()
					
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call GetFISC_DATE
	
	'Exit Sub
		
End Sub


Sub SetSpreadLock()

	With lgvspdData(TYPE_1)
		ggoSpread.Source = lgvspdData(TYPE_1)

		ggoSpread.SpreadUnLock C_W1, -1, C_W7	' ��ü ���� 
'		ggoSpread.SSSetRequired C_W3_NM, -1, -1
		ggoSpread.SpreadLock C_W1,   -1, C_W1_NM2
		ggoSpread.SpreadLock C_W7,   -1, C_W7
		ggoSpread.SpreadLock C_W2,   C_107, C_W7,   C_107
	End With

	With lgvspdData(TYPE_2)
		ggoSpread.Source = lgvspdData(TYPE_2)

		ggoSpread.SpreadUnLock C_W8, -1, C_W13	' ��ü ���� 
		ggoSpread.SpreadLock C_W8,   -1, C_W8_NM
		ggoSpread.SpreadLock C_W12,   -1, C_W12
	End With

	With lgvspdData(TYPE_3)
		ggoSpread.Source = lgvspdData(TYPE_3)

		ggoSpread.SpreadUnLock C_W14, -1, C_W19	' ��ü ���� 
		ggoSpread.SpreadLock C_W16,   -1, C_W16
		ggoSpread.SpreadLock C_W19,   -1, C_W19
	End With

	With lgvspdData(TYPE_4)
		ggoSpread.Source = lgvspdData(TYPE_4)

		ggoSpread.SpreadUnLock C_W20, -1, C_W24	' ��ü ���� 
		ggoSpread.SpreadLock C_W20,   -1, C_W20_NM
		ggoSpread.SpreadLock C_W24,   -1, C_W24
		ggoSpread.SpreadLock C_W23,   C_112, C_W23,   C_113
		ggoSpread.SpreadLock C_W20,   C_113, C_W24,   C_113
	End With
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)
	Dim iRow

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

Sub InitSpreadRow()
	Dim ret

	With lgvspdData(TYPE_1)
				
		ggoSpread.Source = lgvspdData(TYPE_1)
		.ReDraw = False

		If .MaxRows = 0 Then	.MaxRows = C_107

	    ret = .AddCellSpan(C_W1_NM1, C_101, 2, 1)
	    ret = .AddCellSpan(C_W1_NM1, C_102, 2, 1)
	    ret = .AddCellSpan(C_W1_NM1, C_103, 2, 1)
	    ret = .AddCellSpan(C_W1_NM1, C_104, 2, 1)
	    ret = .AddCellSpan(C_W1_NM1, C_105, 1, 2)
	    ret = .AddCellSpan(C_W1_NM1, C_107, 2, 1)

	    ' ù��° �� ��� ���� 
		.Col = C_W1_NM1
		.Row = C_101	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(101)��ǰ�׻�ǰ"
		.Row = C_102	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(102)����ǰ�����ǰ"
		.Row = C_103	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(103)�����"
		.Row = C_104	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(104)����ǰ"
		.Row = C_105	:	.TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Text = " ���� ����"
		.Row = C_107	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(107)��      ��"
	
	    ' �ι�° �� ��� ���� 
		.Col = C_W1_NM2
		.Row = C_105	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(105)ä��"
		.Row = C_106	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(106)��Ÿ"

		' �⺻�ڵ尪�Է��ϱ� 
		.Col = C_W1
		.Row = C_101	:	.Text = "101"
		.Row = C_102	:	.Text = "102"
		.Row = C_103	:	.Text = "103"
		.Row = C_104	:	.Text = "104"
		.Row = C_105	:	.Text = "105"
		.Row = C_106	:	.Text = "106"
		.Row = C_107	:	.Text = "107"
		
		' �հ� ���� �޺����ֱ� 
		.Row = C_107
		.Col = C_W2_NM	:	.CellType = 1	:	.Text = ""
		.Col = C_W3_NM	:	.CellType = 1	:	.Text = ""
		
		.ReDraw = True
	End With 

	With lgvspdData(TYPE_2)
				
		ggoSpread.Source = lgvspdData(TYPE_2)
		.ReDraw = False

		If .MaxRows = 0 Then	.MaxRows = C_110

	    ' ù��° �� ��� ���� 
		.Col = C_W8_NM
		.Row = C_108	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(108)�������ݵ�"
		.Row = C_109	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(109)����δ��"
		.Row = C_110	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(110)��������"

		' �⺻�ڵ尪�Է��ϱ� 
		.Col = C_W8
		.Row = C_108	:	.Text = "108"
		.Row = C_109	:	.Text = "109"
		.Row = C_110	:	.Text = "110"

		.ReDraw = True
	End With 

	With lgvspdData(TYPE_3)
				
		ggoSpread.Source = lgvspdData(TYPE_3)
		.ReDraw = False

		If .MaxRows = 0 Then	.MaxRows = 1

		.ReDraw = True
	End With 

	With lgvspdData(TYPE_4)
				
		ggoSpread.Source = lgvspdData(TYPE_4)
		.ReDraw = False

		If .MaxRows = 0 Then	.MaxRows = C_113

	    ' ù��° �� ��� ���� 
		.Col = C_W20_NM
		.Row = C_111	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(111)�Ǽ��Ϸ��ڻ��"
		.Row = C_112	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(112)�Ǽ������ڻ��"
		.Row = C_113	:	.TypeEditMultiLine = True	:	.TypeHAlign = 0	:	.TypeVAlign = 2
		.Text = "(113)��{(111)+(112)}"
	
		' �⺻�ڵ尪�Է��ϱ� 
		.Col = C_W20
		.Row = C_111	:	.Text = "111"
		.Row = C_112	:	.Text = "112"
		.Row = C_113	:	.Text = "113"

		.ReDraw = True
	End With 

	Call SetSpreadLock()
	Call InitSpreadComboBox()

End Sub

'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' �·ε�� ���۷����޽��� �����´�.
     wgRefDoc = GetDocRef(sCoCd,sFiscYear, sRepType, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
	 
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData

	Call InitSpreadRow()
			
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
End Function

Function GetRefOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr, iSeqNo, iLastRow, iRow
    Dim iDx
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
	Call Fn_GridCalc(TYPE_1, C_W4, C_101)
	Call Fn_GridCalc(TYPE_1, C_W4, C_102)
	Call Fn_GridCalc(TYPE_1, C_W4, C_103)
	Call Fn_GridCalc(TYPE_1, C_W4, C_104)

	Call Fn_GridCalc(TYPE_3, C_W14, C_101)
	lgBlnFlgChgValue = True
	Frm1.vspdData1.focus			
End Function


Sub GetFISC_DATE()	' ������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.

		
End Sub

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
	  
	' �����Ѱ� 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	
 
	Call InitComboBox	' �����ؾ� �Ѵ�. ����� ȸ��������� �о���� ���� 
	Call InitData
	Call FncQuery
     
    
End Sub

'============================================  ����� �Լ�  ====================================
Function Fn_GridCalc(ByVal Index, ByVal pCol, ByVal pRow)
	Dim iRow, dblSum
	Dim dblW4, dblW5, dblW6, dblW7							' Grid 1 Variable
	Dim dblW10, dblW11, dblW12								' Grid 2 Variable
	Dim dblW14, dblW15, dblW16, dblW17, dblW18, dblW19		' Grid 3 Variable
	Dim dblW21, dblW22, dblW23, dblW24						' Grid 4 Variable
	
	If lgvspdData(Index).MaxRows <= 0 Then Exit Function

    ggoSpread.Source = lgvspdData(Index)
    iRow = pRow

	With lgvspdData(Index)
		Select Case Index
			Case TYPE_1
				If iRow >= C_101 And iRow <= C_106 Then
					.Row = iRow	:	.Col = C_W4	:	dblW4 = UNICdbl(.Text)
					.Row = iRow	:	.Col = C_W5	:	dblW5 = UNICdbl(.Text)
					.Row = iRow	:	.Col = C_W6	:	dblW6 = UNICdbl(.Text)

					.Row = iRow	:	.Col = C_W7
					If dblW5 > dblW6 Then
						dblW7 = dblW5 - dblW4
					Else
						dblW7 = dblW6 - dblW4
					End If
					.Text = dblW7
				End If
				dblSum = FncSumSheet(lgvspdData(Index), C_W4, 1, C_106, true, C_107, C_W4, "V")	' �հ� 
				dblSum = FncSumSheet(lgvspdData(Index), C_W5, 1, C_106, true, C_107, C_W5, "V")	' �հ� 
				dblSum = FncSumSheet(lgvspdData(Index), C_W6, 1, C_106, true, C_107, C_W6, "V")	' �հ� 
				dblSum = FncSumSheet(lgvspdData(Index), C_W7, 1, C_106, true, C_107, C_W7, "V")	' �հ� 

			Case TYPE_2
				If iRow = C_108 Or iRow = C_109 Or iRow = C_110 Then
					.Row = iRow	:	.Col = C_W10	:	dblW10 = UNICdbl(.Text)
					.Row = iRow	:	.Col = C_W11	:	dblW11 = UNICdbl(.Text)
					.Row = iRow	:	.Col = C_W12	:	dblW12 = dblW11 - dblW10
					.Text = dblW12
				End If

			Case TYPE_3
				iRow = 1
				.Row = iRow	:	.Col = C_W14	:	dblW14 = UNICdbl(.Text)
				.Row = iRow	:	.Col = C_W15	:	dblW15 = UNICdbl(.Text)
				.Row = iRow	:	.Col = C_W16	:	dblW16 = dblW14 - dblW15
				.Text = dblW16

				.Row = iRow	:	.Col = C_W17	:	dblW17 = UNICdbl(.Text)
				.Row = iRow	:	.Col = C_W18	:	dblW18 = UNICdbl(.Text)
				.Row = iRow	:	.Col = C_W19	:	dblW19 = dblW17 - dblW18
				.Text = dblW19
				
			Case TYPE_4
				If iRow = C_111 Or iRow = C_112 Then
					.Row = iRow	:	.Col = C_W21	:	dblW21 = UNICdbl(.Text)
					.Row = iRow	:	.Col = C_W22	:	dblW22 = UNICdbl(.Text)
					.Row = iRow	:	.Col = C_W23	:	dblW23 = UNICdbl(.Text)
					.Row = iRow	:	.Col = C_W24	:	dblW24 = dblW21 - dblW22 - dblW23
					.Text = dblW24
				End If
				dblSum = FncSumSheet(lgvspdData(Index), C_W21, 1, C_112, true, C_113, C_W21, "V")	' �հ� 
				dblSum = FncSumSheet(lgvspdData(Index), C_W22, 1, C_112, true, C_113, C_W22, "V")	' �հ� 
				dblSum = FncSumSheet(lgvspdData(Index), C_W23, 1, C_112, true, C_113, C_W23, "V")	' �հ� 
				dblSum = FncSumSheet(lgvspdData(Index), C_W24, 1, C_112, true, C_113, C_W24, "V")	' �հ� 
		End Select
	End With

End Function


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

'============================================  �׸��� �̺�Ʈ   ====================================
' -- 1�� �׸��� 
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_1
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_1
	Call vspdData_Change(TYPE_1, Col, Row)
End Sub

Sub vspdData1_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_1
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_1
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_1
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_GotFocus()
	lgCurrGrid = TYPE_1
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData1_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_1
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_1
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_1
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_1
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub

' -- 2�� �׸��� 
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_2
	Call vspdData_Change(TYPE_2, Col, Row)
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_2
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_2
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_GotFocus()
	lgCurrGrid = TYPE_2
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_2
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_2
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_2
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_2
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub


' -- 3�� �׸��� 
Sub vspdData3_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_3
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData3_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_3
	Call vspdData_Change(TYPE_3, Col, Row)
End Sub

Sub vspdData3_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_3
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_3
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData3_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_3
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData3_GotFocus()
	lgCurrGrid = TYPE_3
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData3_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_3
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData3_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_3
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_3
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

' -- 4�� �׸��� 
Sub vspdData4_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_4
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData4_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_4
	Call vspdData_Change(TYPE_4, Col, Row)
End Sub

Sub vspdData4_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_4
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_4
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData4_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_4
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData4_GotFocus()
	lgCurrGrid = TYPE_4
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData4_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_4
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData4_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_4
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData4_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_4
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub


'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)
	Dim iIdx, iRow, sW3, sW4, dblW2

	If Index <> TYPE_1 Then Exit Sub
	With lgvspdData(Index)
		Select Case Col
			Case C_W2, C_W3
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col +1
				.Value = iIdx
			Case C_W2_NM, C_W3_NM
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col -1
				.Value = iIdx
		End Select
		
	End With
End Sub



Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum
	
	lgBlnFlgChgValue= True ' ���濩�� 
    lgvspdData(Index).Row = Row
    lgvspdData(Index).Col = Col

    If lgvspdData(Index).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(lgvspdData(Index).text) < UNICDbl(lgvspdData(Index).TypeFloatMin) Then
         lgvspdData(Index).text = lgvspdData(Index).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(Index)
    ggoSpread.UpdateRow Row

	' --- �߰��� �κ� 
	Call Fn_GridCalc(Index, Col, Row)	' ��� 
	
End Sub

Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = lgvspdData(Index)
   
    If lgvspdData(Index).MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    Exit Sub
   	    
    If Row <= 0 Then
    	Exit Sub
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
    ggoSpread.Source = lgvspdData(Index)
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
'    Call GetSpreadColumnPos("A")
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
		If Row > 0 And Col = C_W1_BT Then
		    .Row = Row
		    .Col = C_W1_BT

		    Call OpenAdItem(Index)
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
	For i = TYPE_1 To TYPE_4
		ggoSpread.Source = lgvspdData(i)
		If ggoSpread.SSCheckChange = True Then
			blnChange = True
			Exit For
		End If
    Next
    
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
'    Call InitSpreadRow()

    															
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

    'If Verification = False Then Exit Function
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function

' ---------------------- ���ĳ� ���� -------------------------
Function  Verification()
	Dim dblW11, dblW12, dblW16, dblW14, dblW15, dblW13
	
	Verification = False

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
    Call InitSpreadRow()

    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
    lgIntFlgMode = parent.OPMD_CMODE

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
	Dim iActiveRow
	
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

    If lgvspdData(lgCurrGrid).MaxRows < 1 Then
       Exit Function
    End If

    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 

End Function

Function FncInsertRow(ByVal pvRowCnt) 

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function


Function FncDeleteRow() 

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
	
    If lgBlnFlgChgValue = True Then
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

		
Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	'-----------------------
	'Reset variables area
	'-----------------------

	lgIntFlgMode = parent.OPMD_UMODE
		
    Call SetToolbar("1101100000000111")										<%'��ư ���� ���� %>
	
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
    
    For i = TYPE_1 To TYPE_4	' ��ü �׸��� ���� 
    
		With lgvspdData(i)
	
			ggoSpread.Source = lgvspdData(i)
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
			
			For lRow = 1 To .MaxRows
    
		       .Row = lRow
		       .Col = 0
		    
		       Select Case .Text
		           Case  ggoSpread.InsertFlag                                      '��: Insert
		                                              strVal = strVal & "C"  &  Parent.gColSep
		           Case  ggoSpread.UpdateFlag                                      '��: Update
		                                              strVal = strVal & "U"  &  Parent.gColSep
			       Case  ggoSpread.DeleteFlag                                      '��: Delete
			                                          strVal = strVal & "D"  &  Parent.gColSep
			       Case Else
			                                          strVal = strVal & ""  &  Parent.gColSep
		       End Select
		       
			  ' ��� �׸��� ����Ÿ ����     
'			  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
					For lCol = 1 To lMaxCols
						.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
					Next
					strVal = strVal & Trim(.Text) &  Parent.gRowSep
'			  End If  
			Next
		
		End With

		document.all("txtSpread" & CStr(i)).value =  strVal
		strVal = ""
	Next

	'Frm1.txtSpread.value      = strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	Frm1.txtFlgMode.Value	=	lgIntFlgMode
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' ���� ������ ���� ���� %>
	Call InitVariables
	Call FncNew
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
					<TD WIDTH=* align=right><A href="vbscript:GetRef">�ݾ׺ҷ�����</A></TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w9103ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script></TD>
									<TD CLASS="TD5">���θ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
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
						<DIV ID="ViewDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : ������ ������ ������ ũ�⿡ ���� ��ũ�ѹٰ� �����ǰ� �Ѵ� %>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=210>
                                   <table <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
									   <TR>
										   <TD width="100%" HEIGHT=10 CLASS="CLSFLD"><br>&nbsp;1. ����ڻ���������� ��</TD>
									   </TR>
									   <TR>
										   <TD width="100%"><script language =javascript src='./js/w9103ma1_vspdData1_vspdData1.js'></script>
										  </TD>
									  </TR>
								  </table>
								</TD>
							</TR>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=120>
									<table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
									   <TR>
										   <TD width="100%" HEIGHT=10 CLASS="CLSFLD"><br>&nbsp;2. �������� �������δ�ݤ��������� �ձݻ��� ����</TD>
									   </TR>
									   <TR>
										   <TD width="100%"><script language =javascript src='./js/w9103ma1_vspdData2_vspdData2.js'></script>
										  </TD>
									  </TR>
								  </table>
								</TD>
							</TR>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=120>
									<table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
									   <TR>
										   <TD width="100%" HEIGHT=10 CLASS="CLSFLD"><br>&nbsp;3. �����ޱ� �� �������� ����</TD>
									   </TR>
									   <TR>
										   <TD width="100%"><script language =javascript src='./js/w9103ma1_vspdData3_vspdData3.js'></script>
										  </TD>
									  </TR>
								  </table>
								</TD>
							</TR>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=114>
									<table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
									   <TR>
										   <TD width="100%" HEIGHT=10 CLASS="CLSFLD"><br>&nbsp;4. �Ǽ��ڱ����� ����</TD>
									   </TR>
									   <TR>
										   <TD width="100%"><script language =javascript src='./js/w9103ma1_vspdData4_vspdData4.js'></script>
										  </TD>
									  </TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
   <TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
	
		
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
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24" style="display:'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" style="display:'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread3" tag="24" style="display:'none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread4" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
