<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��Ÿ���� 
'*  3. Program ID           : W9107MA1
'*  4. Program Name         : W9107MA1.asp
'*  5. Program Desc         : �� 52ȣ Ư�������ڰ� �ŷ����� 
'*  6. Modified date(First) : 2005/01/31
'*  7. Modified date(Last)  : 2005/01/31
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

Const BIZ_MNU_ID		= "W9107MA1"
Const BIZ_PGM_ID		= "W9107MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID		= "W9107OA"

Const TYPE_1	= 0		' �׸��带 �������� ���� ��� 
Const TYPE_2	= 1		

' -- �׸��� �÷� ���� 
Dim	C_SEQ_NO
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
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

Dim C_W17
Dim C_W18
Dim C_W19
Dim C_W20
Dim C_W21
Dim C_W22
Dim C_W23
Dim C_W24
Dim C_W25
Dim C_W26
Dim C_W27
Dim C_W28
Dim C_W29_1
Dim C_W29_2
Dim C_W30_1
Dim C_W30_2
'Dim C_W31_NM
Dim C_W31

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgCurrGrid, lgvspdData(1), IsRunEvents


'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_SEQ_NO	= 1
	C_W1		= 2
	C_W2		= 3
	C_W3		= 4
	C_W4		= 5
	C_W5		= 6
	C_W6		= 7
	C_W7		= 8
	C_W8		= 9
	C_W9		= 10
	C_W10		= 11
	C_W11		= 12
	C_W12		= 13
	C_W13		= 14
	C_W14		= 15
	C_W15		= 16
	C_W16		= 17

	C_W17		= 2
	C_W18		= 3
	C_W19		= 4
	C_W20		= 5
	C_W21		= 6
	C_W22		= 7
	C_W23		= 8
	C_W24		= 9
	C_W25		= 10
	C_W26		= 11
	C_W27		= 12
	C_W28		= 13
	C_W29_1		= 14
	C_W29_2		= 15
	C_W30_1		= 16
	C_W30_2		= 17
	C_W31		= 18

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

    lgCurrGrid = TYPE_1
    IsRunEvents = False
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

End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1)		= frm1.vspdData0
	Set lgvspdData(TYPE_2)		= frm1.vspdData1

		
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("6","3","2")	' -- ����(����)
	
	' 1�� �׸��� 

	With lgvspdData(TYPE_1)
				
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20051222_" & TYPE_1,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W16 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
 
  		'����� 3�ٷ�    
		.ColHeaderRows = 3  
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����", 10,,,15,1
		ggoSpread.SSSetEdit		C_W1,		"(1)���θ�" & vbCrLf & "(��ȣ �Ǵ� ����)", 10,,,25,1
		ggoSpread.SSSetEdit		C_W2,		"(2)����ڵ��" & vbCrLf & "��ȣ(�Ǵ�" & vbCrLf & "�ֹε�Ϲ�ȣ)", 12, 2,,14,1
		ggoSpread.SSSetFloat	C_W3,		"(3)��"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W4,		"(4)����ڻ�"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W5,		"(5)��Ÿ"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W6,		"(6)�����ڻ�"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W7,		"(7)�뿪"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W8,		"(8)�������"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W9,		"(9)��Ÿ"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W10,		"(10)��"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W11,		"(11)����ڻ�"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W12,		"(12)��Ÿ"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W13,		"(13)�����ڻ�"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W14,		"(14)�뿪"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W15,		"(15)�������"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W16,		"(16)��Ÿ"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	

		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(C_SEQ_NO	, -1000, 1, 3)	
		ret = .AddCellSpan(C_W1		, -1000, 2, 1)	
		ret = .AddCellSpan(C_W3		, -1000, 7, 1)	
		ret = .AddCellSpan(C_W10	, -1000, 7, 1)

		ret = .AddCellSpan(C_W1		,  -999, 1, 2)
		ret = .AddCellSpan(C_W2		,  -999, 1, 2)
		ret = .AddCellSpan(C_W3		,  -999, 1, 2)
		ret = .AddCellSpan(C_W4		,  -999, 2, 1)
		ret = .AddCellSpan(C_W6		,  -999, 1, 2)
		ret = .AddCellSpan(C_W7		,  -999, 1, 2)
		ret = .AddCellSpan(C_W8		,  -999, 1, 2)
		ret = .AddCellSpan(C_W9		,  -999, 1, 2)
		ret = .AddCellSpan(C_W10	,  -999, 1, 2)
		ret = .AddCellSpan(C_W11	,  -999, 2, 1)
		ret = .AddCellSpan(C_W13	,  -999, 1, 2)
		ret = .AddCellSpan(C_W14	,  -999, 1, 2)
		ret = .AddCellSpan(C_W15	,  -999, 1, 2)
		ret = .AddCellSpan(C_W16	,  -999, 1, 2)
    
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W1		: .Text = "�ŷ�����"
		.Col = C_W3		: .Text = "�� �� �� �� ��"
		.Col = C_W10	: .Text = "�� �� �� �� ��"
		
		.Row = -999
		.Col = C_W1		: .Text = "(6)���θ�" & vbCrLf & "(��ȣ �Ǵ� ����)"
		.Col = C_W2		: .Text = "(7)����ڵ��" & vbCrLf & "��ȣ(�Ǵ�" & vbCrLf & "�ֹε�Ϲ�ȣ)"
		.Col = C_W3		: .Text = "(8)��"
		.Col = C_W4		: .Text = "�����ڻ�"
		.Col = C_W6		: .Text = "(11)����" & vbCrLf & "�ڻ�"
		.Col = C_W7		: .Text = "(12)�뿪"
		.Col = C_W8		: .Text = "(13)����" & vbCrLf & "���"
		.Col = C_W9		: .Text = "(14)��Ÿ"
		.Col = C_W10	: .Text = "(15)��"
		.Col = C_W11	: .Text = "�����ڻ�"
		.Col = C_W13	: .Text = "(18)����" & vbCrLf & "�ڻ�"
		.Col = C_W14	: .Text = "(19)�뿪"
		.Col = C_W15	: .Text = "(20)����" & vbCrLf & "���"
		.Col = C_W16	: .Text = "(21)��Ÿ"
		
		.Row = -998
		.Col = C_W4		: .Text = "(9)����ڻ�"
		.Col = C_W5		: .Text = "(10)��Ÿ"
		.Col = C_W11	: .Text = "(16)����ڻ�"
		.Col = C_W12	: .Text = "(17)��Ÿ"
				
		.rowheight(-999) = 12	
		.rowheight(-998) = 15					
		'.rowheight(-998) = 15	' ���� ������	(2���� ���, 1���� 15)
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		
		'Call InitSpreadComboBox
		
		.ReDraw = true	
				
	End With 

 
	' 2�� �׸��� 
	With lgvspdData(TYPE_2)
				
		ggoSpread.Source = lgvspdData(TYPE_2)	
		'patch version
		ggoSpread.Spreadinit "V20051222_" & TYPE_2,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W31 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

  		'����� 3�ٷ�    
		.ColHeaderRows = 3  
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����", 10,,,15,1
		ggoSpread.SSSetEdit		C_W17,		"(22)���θ�" & vbCrLf & "(��ȣ �Ǵ� ����)", 10,,,25,1
		ggoSpread.SSSetMask		C_W18,		"(23)����ڵ��"	, 15, 2, "999-99-99999" 
		ggoSpread.SSSetCheck	C_W19,		"(24)����", 7,,,True
		ggoSpread.SSSetCheck	C_W20,		"(25)����", 7,,,True
		ggoSpread.SSSetDate		C_W21,		"(26)����",	10,		2,		Parent.gDateFormat,	-1

		ggoSpread.SSSetFloat	C_W22,		"(27)�׸��Ѿ�"	, 13, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W23,		"(28)����"	, 7, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W24,		"(29)�׸��Ѿ�"	, 13, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W25,		"(30)����"	, 7, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetCheck	C_W26,		"(31)�պ�", 7,,,True
		ggoSpread.SSSetCheck	C_W27,		"(32)����" & vbCrLf & "�պ�", 7,,,True
		ggoSpread.SSSetDate		C_W28,		"(33)����",	10,		2,		Parent.gDateFormat,	-1
		
		ggoSpread.SSSetFloat	C_W29_1,	"���ڻ�"	, 13, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W29_2,	"����"	, 7, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W30_1,	"���ڻ�"	, 13, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W30_2,	"����"	, 7, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W31,		"�պ�����"	, 7, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		
		Call MakePercentCol( lgvspdData(TYPE_2), C_W23, "", "", "")
		Call MakePercentCol( lgvspdData(TYPE_2), C_W25, "", "", "")
		Call MakePercentCol( lgvspdData(TYPE_2), C_W29_2, "", "", "")
		Call MakePercentCol( lgvspdData(TYPE_2), C_W30_2, "", "", "")
		Call MakePercentCol( lgvspdData(TYPE_2), C_W31, "", 1000, "")
		
		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(C_SEQ_NO	, -1000, 1, 3)	
		ret = .AddCellSpan(C_W17	, -1000, 2, 1)	
		ret = .AddCellSpan(C_W19	, -1000, 2, 1)	
		ret = .AddCellSpan(C_W21	, -1000, 1, 3)
		ret = .AddCellSpan(C_W22	, -1000, 2, 1)
		ret = .AddCellSpan(C_W24	, -1000, 2, 1)
		ret = .AddCellSpan(C_W26	, -1000, 2, 1)
		ret = .AddCellSpan(C_W28	, -1000, 1, 3)
		ret = .AddCellSpan(C_W29_1	, -1000, 4, 1)
		ret = .AddCellSpan(C_W31	, -1000, 1, 3)
		
		ret = .AddCellSpan(C_W17	,  -999, 1, 2)
		ret = .AddCellSpan(C_W18	,  -999, 1, 2)
		ret = .AddCellSpan(C_W19	,  -999, 1, 2)
		ret = .AddCellSpan(C_W20	,  -999, 1, 2)
		ret = .AddCellSpan(C_W22	,  -999, 1, 2)
		ret = .AddCellSpan(C_W23	,  -999, 1, 2)
		ret = .AddCellSpan(C_W24	,  -999, 1, 2)
		ret = .AddCellSpan(C_W25	,  -999, 1, 2)
		ret = .AddCellSpan(C_W26	,  -999, 1, 2)
		ret = .AddCellSpan(C_W27	,  -999, 1, 2)
		ret = .AddCellSpan(C_W29_1	,  -999, 2, 1)
		ret = .AddCellSpan(C_W30_1	,  -999, 2, 1)
    
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W17	: .Text = "�ŷ�����"
		.Col = C_W19	: .Text = "����"
		.Col = C_W21	: .Text = "(26)����"
		.Col = C_W22	: .Text = "����(�Ǵ� ����) ��"
		.Col = C_W24	: .Text = "����(�Ǵ� ����) ��"
		.Col = C_W26	: .Text = "����"
		.Col = C_W28	: .Text = "(33)����"
		.Col = C_W29_1	: .Text = "���ڻ��� �ð� �� ������"
		.Col = C_W31	: .Text = "(36)�պ�����" & vbCrLf & "1:(  )%"
		
		.Row = -999
		.Col = C_W17	: .Text = "(22)���θ�"
		.Col = C_W18	: .Text = "(23)�����" & vbCrLf & "��Ϲ�ȣ"
		.Col = C_W19	: .Text = "(24)����"
		.Col = C_W20	: .Text = "(25)����"
		.Col = C_W22	: .Text = "(27)�׸��Ѿ�"
		.Col = C_W23	: .Text = "(28)����"
		.Col = C_W24	: .Text = "(29)�׸��Ѿ�"
		.Col = C_W25	: .Text = "(30)����"
		.Col = C_W26	: .Text = "(31)�պ�"
		.Col = C_W27	: .Text = "(32)����" & vbCrLf & "�պ�"
		.Col = C_W28	: .Text = "(33)����"
		.Col = C_W29_1	: .Text = "(34)�պ����ε�"
		.Col = C_W30_1	: .Text = "(35)���պ����ε�"
				
		.Row = -998
		.Col = C_W29_1	: .Text = "���ڻ�"
		.Col = C_W29_2	: .Text = "����"
		.Col = C_W30_1	: .Text = "���ڻ�"
		.Col = C_W30_2	: .Text = "����"
				
		.rowheight(-999) = 12	
		.rowheight(-998) = 12					
		'.rowheight(-998) = 15	' ���� ������	(2���� ���, 1���� 15)
				
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
		.ReDraw = true	
				
	End With 

End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	'Call GetFISC_DATE
	
		
End Sub

Sub SetSpreadLock(pType)

	With lgvspdData(pType)
	
		ggoSpread.Source = lgvspdData(pType)	

		Select Case pType
			Case TYPE_1 
				ggoSpread.SSSetRequired C_W1, 1, .MaxRows
				ggoSpread.SSSetRequired C_W2, 1, .MaxRows
				ggoSpread.SpreadLock C_W3, -1, C_W3	' ��ü ���� 
				ggoSpread.SpreadLock C_W10, -1, C_W10	' ��ü ���� 
				ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO	' ��ü ���� 
				
			Case TYPE_2
				ggoSpread.SSSetRequired C_W17, 1, .MaxRows
				ggoSpread.SSSetRequired C_W18, 1, .MaxRows
				ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO	' ��ü ���� 
				
		End Select
		
	End With	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	With lgvspdData(pType)
		ggoSpread.Source = lgvspdData(pType)	

		If pType = TYPE_1 Then
			ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_W3, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_W10, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W1, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W2, pvStartRow, pvEndRow
		ElseIf pType = TYPE_2 Then
			ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W17, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W18, pvStartRow, pvEndRow
		End If
			
	End With	
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData0
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W1		= iCurColumnPos(2)
            C_W2		= iCurColumnPos(3)
            C_W3		= iCurColumnPos(4)
            C_W4		= iCurColumnPos(5)
            C_W5		= iCurColumnPos(6)
            C_W6		= iCurColumnPos(7)
            C_W7		= iCurColumnPos(8)
            C_W8		= iCurColumnPos(9)
            C_W9		= iCurColumnPos(10)
            C_W10		= iCurColumnPos(11)
            C_W11		= iCurColumnPos(12)
            C_W12		= iCurColumnPos(13)
            C_W13		= iCurColumnPos(14)
            C_W14		= iCurColumnPos(15)
            C_W15		= iCurColumnPos(16)
            C_W16		= iCurColumnPos(17)
 
        Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W17		= iCurColumnPos(2)
            C_W18		= iCurColumnPos(3)
            C_W19		= iCurColumnPos(4)
            C_W20		= iCurColumnPos(5)
            C_W21		= iCurColumnPos(6)
            C_W22		= iCurColumnPos(7)
            C_W23		= iCurColumnPos(8)
            C_W24		= iCurColumnPos(9)
            C_W25		= iCurColumnPos(10)
            C_W26		= iCurColumnPos(11)
            C_W27		= iCurColumnPos(12)
            C_W28		= iCurColumnPos(13)
            C_W29_1		= iCurColumnPos(14)
            C_W29_2		= iCurColumnPos(15)
            C_W31_NM	= iCurColumnPos(16)
            C_W31		= iCurColumnPos(17)           
    End Select    
End Sub


Sub SetSpreadTotalLine()
	Dim iRow, i

	For i = TYPE_1 To TYPE_2
		ggoSpread.Source = lgvspdData(i)
		With lgvspdData(i)
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W1 : .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
				'ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
			End If
		End With
	Next
End Sub 

'============================== ���۷��� �Լ�  ========================================



'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100001111")										<%'��ư ���� ���� %>
	  
	' �����Ѱ� 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
 
	Call InitComboBox	
	Call InitData
	Call FncQuery() 
	
     
    
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
Sub vspdData0_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_1
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

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
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

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

Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_2
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub


'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum,sCoCd,sFiscYear,sRepType
	
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	lgBlnFlgChgValue= True ' ���濩�� 
    lgvspdData(lgCurrGrid).Row = Row
    lgvspdData(lgCurrGrid).Col = Col

    If lgvspdData(Index).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(lgvspdData(Index).text) < UNICDbl(lgvspdData(Index).TypeFloatMin) Then
         lgvspdData(Index).text = lgvspdData(Index).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(Index)
    ggoSpread.UpdateRow Row

	' --- �߰��� �κ� 
	With lgvspdData(Index)

	If Index = TYPE_1 Then	'1�� �׸� 

		Dim dblAmt(20)
		
		Select Case Col
			Case C_W4, C_W5, C_W6, C_W7, C_W8, C_W9

				.Row = Row
				.Col = C_W4		: dblAmt(C_W4) = UNICDbl(.value)
				.Col = C_W5		: dblAmt(C_W5) = UNICDbl(.value)
				.Col = C_W6		: dblAmt(C_W6) = UNICDbl(.value)
				.Col = C_W7		: dblAmt(C_W7) = UNICDbl(.value)
				.Col = C_W8		: dblAmt(C_W8) = UNICDbl(.value)
				.Col = C_W9		: dblAmt(C_W9) = UNICDbl(.value)

				dblAmt(C_W3) = dblAmt(C_W4) + dblAmt(C_W5) + dblAmt(C_W6) + dblAmt(C_W7) + dblAmt(C_W8) + dblAmt(C_W9)
				.Col = C_W3	: .value = dblAmt(C_W3)
				
			Case C_W11, C_W12, C_W13, C_W14, C_W15, C_W16
				
				.Row = Row
				.Col = C_W11	: dblAmt(C_W11) = UNICDbl(.value)
				.Col = C_W12	: dblAmt(C_W12) = UNICDbl(.value)
				.Col = C_W13	: dblAmt(C_W13) = UNICDbl(.value)
				.Col = C_W14	: dblAmt(C_W14) = UNICDbl(.value)
				.Col = C_W15	: dblAmt(C_W15) = UNICDbl(.value)
				.Col = C_W16	: dblAmt(C_W16) = UNICDbl(.value)

				dblAmt(C_W10) = dblAmt(C_W11) + dblAmt(C_W12) + dblAmt(C_W13) + dblAmt(C_W14) + dblAmt(C_W15) + dblAmt(C_W16)
				.Col = C_W10	: .value = dblAmt(C_W10)

				
		End Select
		
	ElseIf Index = TYPE_2 Then
		
		
		Select Case Col
		
			Case C_W18	
				.Row = Row
				.COL = C_W18			
				Call CommonQueryRs("OWN_RGST_NO","TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

				IF .TEXT = Replace(lgF0,Chr(11),"") THEN
					Call DisplayMsgBox("127421", parent.VB_INFORMATION, "�ŷ������� ����ڹ�ȣ", "�������� ����ڹ�ȣ")
					.value = ""
				END IF
				
		End Select
		
	End If
	
	End With
	
End Sub

Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index
    'Call SetPopupMenuItemInf("0001011111") 

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
    ggoSpread.Source = lgvspdData(Index)
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
	If IsRunEvents = True Then Exit Sub	' �ؿ� Ÿ üũ�ڽ��� ���� ������ ���� �̺�Ʈ�� �߻��� 
	
	IsRunEvents = True
	
	' -- 200603 ����: ����/����/���þ���(�պ�/�����պ�/���þ���) : (��ɼ���, ���������Ǿ������)
	With lgvspdData(Index)
		Select Case Col
			Case C_W19
				.Col = C_W19
				If .Value = 1 Then	' -- ����ڰ� Ŭ���ؼ� üũ�Ǿ��ٸ� ���ڸ� ����(���߿� 1��)
					.Col = C_W20
					.Value = 0
				End If
			Case C_W20
				.Col = C_W20
				If .Value = 1 Then	' -- ����ڰ� Ŭ���ؼ� üũ�Ǿ��ٸ� ���ڸ� ����(���߿� 1��)
					.Col = C_W19
					.Value = 0
				End If
			Case C_W26
				.Col = C_W26
				If .Value = 1 Then	' -- ����ڰ� Ŭ���ؼ� üũ�Ǿ��ٸ� �����պ��� ����(���߿� 1��)
					.Col = C_W27
					.Value = 0
				End If
			Case C_W27
				.Col = C_W27
				If .Value = 1 Then	' -- ����ڰ� Ŭ���ؼ� üũ�Ǿ��ٸ� �����պ��� ����(���߿� 1��)
					.Col = C_W26
					.Value = 0
				End If
		End Select
    End With

    ggoSpread.Source = lgvspdData(Index)
    ggoSpread.UpdateRow Row
    
    IsRunEvents = False
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
	For i = TYPE_1 To TYPE_2
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
	For i = TYPE_1 To TYPE_2
		ggoSpread.Source = lgvspdData(i)
		If ggoSpread.SSCheckChange = True Then
			blnChange = True
			Exit For
		End If
    Next
    
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

    If Verification = False Then Exit Function
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function

' -- 200603 ���� �߰�
' ---------------------- ���ĳ� ���� -------------------------
Function  Verification()
	Dim i, blnW24, blnW25, blnW31, blnW32, iMaxRows
	
	Verification = False

	With lgvspdData(TYPE_2)
		iMaxRows = .MaxRows
		
		For i = 1 To iMaxRows
			.Row = i
			.Col = 0

			If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Then	' -- �Է�/���� �� ��� ����(����/����, �պ�,�����պ� �߿� ��� 1������ üũ�Ǿ�� ��)
				.Col = C_W19  : blnW24 = .Value		
				.Col = C_W20  : blnW25 = .Value		
				.Col = C_W26  : blnW31 = .Value		
				.Col = C_W27  : blnW32 = .Value		
			
				If blnW24 = 0 And blnW25 = 0 And blnW31 = 0 And blnW32 = 0 Then
					Call DisplayMsgBox("X", parent.VB_INFORMATION, "����/����, �պ�/�����պ��� ��� 1������ üũ�Ǿ�� �մϴ�.", "X")           '��: "Will you destory previous data"
					.Focus
					.Col = C_W17
					.Action = 0
					Exit Function
				End If
			End If
		Next

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
			SetSpreadColor lgCurrGrid, lgvspdData(lgCurrGrid).ActiveRow, lgvspdData(lgCurrGrid).ActiveRow

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
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
 End Function


Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo

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
		If iRow < 0 Then iRow = 1
		
		lgvspdData(lgCurrGrid).ReDraw = False

		If iRow = .MaxRows Then	' -- ������ �հ��ٿ��� InsertRow�� �ϸ� ������ �߰��Ѵ�.
			ggoSpread.InsertRow iRow-1 , imRow 
			SetSpreadColor lgCurrGrid, iRow, iRow + imRow - 1

			Call SetDefaultVal(lgCurrGrid, iRow, imRow)
		Else
			ggoSpread.InsertRow ,imRow
			SetSpreadColor lgCurrGrid, iRow+1, iRow + imRow

			Call SetDefaultVal(lgCurrGrid, iRow+1, imRow)
		End If   

    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

' �׸��忡 SEQ_NO, TYPE �ִ� ���� 
Function SetDefaultVal(pType, iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With lgvspdData(pType)	' ��Ŀ���� �׸��� 

	ggoSpread.Source = lgvspdData(pType)
	
	If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
		.Row = iRow
		MaxSpreadVal lgvspdData(pType), C_SEQ_NO, iRow
		
		If pType = TYPE_2 Then
			.Col = C_W19
			.Value = 1
			.Col = C_W26
			.Value = 1
		End If
	Else
		iSeqNo = MaxSpreadVal(lgvspdData(pType), C_SEQ_NO, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i
			.Col = C_SEQ_NO : .Value = iSeqNo : iSeqNo = iSeqNo + 1
			If pType = TYPE_2 Then
				.Col = C_W19
				.Value = 1
				.Col = C_W26
				.Value = 1
			End If
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
    
    lgBlnFlgChgValue = True
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
		
Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	If lgvspdData(TYPE_1).MaxRows > 0 Or lgvspdData(TYPE_2).MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		
		' �������� ���� : ���ߵǸ� ���ȴ�.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 ����üũ : �׸��� �� 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1

			'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
			Call SetToolbar("1101111100000111")										<%'��ư ���� ���� %>

		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
		End If
	Else
		Call SetToolbar("1100110100000111")										<%'��ư ���� ���� %>
	End If
	
	Call SetSpreadLock(TYPE_1)
	Call SetSpreadLock(TYPE_2)
	
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
    Dim strVal, strDel, sTmp
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    For i = TYPE_1 To TYPE_2	' ��ü �׸��� ���� 
    
		With lgvspdData(i)
	
			ggoSpread.Source = lgvspdData(i)
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
			
			' ----- 1��° �׸��� 
			For lRow = 1 To .MaxRows

    			.Col = 0
				.Row = lRow	: sTmp = ""
		    
				  ' ��� �׸��� ����Ÿ ����     
				  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
						For lCol = 1 To lMaxCols
							Select Case lCol
								'Case C_W31
								'	.Col = lCol : strVal = strVal & Trim(.Value) &  Parent.gColSep
								Case Else
									.Col = lCol : sTmp = sTmp & Trim(.Text) &  Parent.gColSep
							End Select
						Next
						sTmp = sTmp & Trim(.Text) &  Parent.gRowSep
				  End If  


				.Col = 0
				   		
				Select Case .Text
					Case  ggoSpread.InsertFlag                                      '��: Insert
				                                       strVal = strVal & "C"  &  Parent.gColSep & sTmp
				    Case  ggoSpread.UpdateFlag                                      '��: Update
				                                       strVal = strVal & "U"  &  Parent.gColSep & sTmp
				    Case  ggoSpread.DeleteFlag                                      '��: Update
				                                       strDel = strDel & "D"  &  Parent.gColSep & sTmp
				End Select

			Next
							   
		End With

		document.all("txtSpread" & CStr(i)).value =  strDel & strVal
		strDel = "" : strVal = ""
	Next

	'Frm1.txtSpread.value      = strDel & strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Call InitVariables
	lgvspdData(TYPE_1).MaxRows = 0
    ggoSpread.Source = lgvspdData(TYPE_1)
    ggoSpread.ClearSpreadData

	lgvspdData(TYPE_2).MaxRows = 0
    ggoSpread.Source = lgvspdData(TYPE_2)
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

Function ProgramJump
    Call PgmJump(JUMP_PGM_ID)
End Function

'========================================================================================
Function FncBtnPrint(byval strPrintType)
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE
	Dim StrUrl  , i, j

	Dim intCnt,IntRetCD, sWhere, sWhereSQL

	Call SetPrintCond(varCo_Cd, varFISC_YEAR, varREP_TYPE) 

	StrUrl = "varCo_Cd|"			& varCo_Cd
	StrUrl = StrUrl & "|varFISC_YEAR|"		& varFISC_YEAR
	StrUrl = StrUrl & "|varREP_TYPE|"       & varREP_TYPE
	
	For j = 1 To 2
	
		ObjName = AskEBDocumentName(EBR_RPT_ID & Cstr(j), "ebr")
			 
		if  strPrintType = "VIEW" then
		   Call FncEBRPreview(ObjName, StrUrl)
		else

			If document.all("EBAction") is Nothing Then
				Dim pObj , pHTML
				
				pHTML = "<FORM NAME=EBAction TARGET=MyBizASP METHOD=POST>" & _
				"	<INPUT TYPE=HIDDEN NAME=uname>" & _
				"	<INPUT TYPE=HIDDEN NAME=dbname>" & _
				"	<INPUT TYPE=HIDDEN NAME=filename>" & _
				"	<INPUT TYPE=HIDDEN NAME=condvar>" & _
				"	<INPUT TYPE=HIDDEN NAME=date>	" & _
				"</FORM>" 

				Set pObj = document.all("MousePT")
				Call pObj.insertAdjacentHTML("afterBegin", pHTML)
			End If
		
		   Call FncEBRPrint(EBAction,ObjName,StrUrl)
		end if	
			 
		Dim objChkBox, iCnt
			 
		Set objChkBox = document.all("prt_check")
			 
		If Not objChkBox is Nothing Then
			 
			 if document.all("prt_check" & Cstr(j)).checked = true then 

			  	ObjName = AskEBDocumentName(EBR_RPT_ID & Cstr(j) & (i+1), "ebr")
					      
				 	if  strPrintType = "VIEW" then
				 		Call FncEBRPreview(ObjName, StrUrl)
				 	else
				 		Call FncEBRPrint(MyBizASP,ObjName,StrUrl)
				 	end if	

			  	ObjName = AskEBDocumentName(EBR_RPT_ID & Cstr(j) & (i+2), "ebr")
					      
				 	if  strPrintType = "VIEW" then
				 		Call FncEBRPreview(ObjName, StrUrl)
				 	else
				 		Call FncEBRPrint(MyBizASP,ObjName,StrUrl)
				 	end if	
			  end if	

		End If

	Next
	
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
			<TABLE <%=LR_SPACE_TYPE_20%>>
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
					<TD WIDTH=* align=right></TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w9107ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=100%>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="10">&nbsp;1. ���� �� ���԰ŷ� ��						
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="45%">
											<script language =javascript src='./js/w9107ma1_vspdData0_vspdData0.js'></script>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="10">&nbsp;2. �� �� �� ��							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="45%">
											<script language =javascript src='./js/w9107ma1_vspdData1_vspdData1.js'></script>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
				<TR>
				        <TD WIDTH=10>&nbsp;</TD>
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><����>���� �� ���԰ŷ���</LABEL>&nbsp;
				        
				            <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><����>�ں��ŷ�</LABEL>&nbsp;
				           
				        </TD>
		
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
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

