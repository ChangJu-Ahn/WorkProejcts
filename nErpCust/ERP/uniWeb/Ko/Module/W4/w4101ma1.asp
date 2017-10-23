
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �غ���� 
'*  3. Program ID           : W4101MA1
'*  4. Program Name         : W4101MA1.asp
'*  5. Program Desc         : ��31ȣ(1) �߼ұ�������غ�� �������� 
'*  6. Modified date(First) : 2005/01/14
'*  7. Modified date(Last)  : 2005/01/14
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

Const BIZ_MNU_ID		= "W4101MA1"
Const BIZ_PGM_ID		= "W4101mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "W4101mb2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID	    = "W4101OA1"
Const TAB1 = 1																	'��: Tab�� ��ġ 
Const TAB2 = 2

Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.
Const TYPE_3	= 2		

' -- �׸��� �÷� ���� 
Dim C_SEQ_NO	
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
Dim C_W29
Dim C_W30
Dim C_W31
Dim C_W32
Dim C_W33

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(2), IsRunEvents

Dim lgW3, lgW2	' ������, ����������� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	
	C_SEQ_NO	= 1	' -- 1�� �׸��� 
    C_W10		= 2	' �������� 
    C_W11		= 3 ' �ݾ� 
    C_W12		= 4	' �����󰢴���� 
    C_W13		= 5	' �󰢺��δ���� 
    C_W14		= 6	' ������ 
    C_W15		= 7	' ���޼��񰡾� 
    C_W16		= 8	' �������񰡾� 

	' C_SEQ_NO ���� 
	C_W17		= 2	' ������ 
	C_W18		= 3	' ������ 
	C_W19		= 4	' ��λ��غ�� 
	C_W20		= 5 ' �����غ�� 
	C_W21		= 6	' �غ�� 
	C_W22		= 7	' ��ü�ҿ��ڱݻ��� 
	C_W23		= 8	' �̻��� 
	C_W24		= 9	' ��ü�ҿ��ڱݻ��� 
	C_W25		= 10 ' ��Ÿ 
	C_W26		= 11 ' �� 
	
	' C_SEQ_NO, C_W17 ���� 
	C_W27		= 3	' 1������ 
	C_W28		= 4	' 2������ 
	C_W29		= 5	' 3���⵵ 
	C_W30		= 6 ' �� 
	C_W31		= 7	' ȯ���ұݾ��հ� 
	C_W32		= 8	' ȸ��ȯ�Ծ� 
	C_W33		= 9	' ����ȯ�� 

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
    ' ������ 
	call CommonQueryRs("REFERENCE"," B_Configuration "," MAJOR_CD = 'W2007' AND MINOR_CD='1'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgW3 = Split(lgF0, Chr(11))
    
End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1
	Set lgvspdData(TYPE_3) = frm1.vspdData2
	
	lgvspdData(TYPE_1).ScriptEnhanced  = True
	lgvspdData(TYPE_2).ScriptEnhanced  = True
	lgvspdData(TYPE_3).ScriptEnhanced  = True
	
    Call initSpreadPosVariables()  

	' 1�� �׸��� 
	With lgvspdData(TYPE_1)
			
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W16 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
				       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'Call AppendNumberPlace("6","3","2")

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����"		, 10,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W10,		"(10)��������"	, 10,,, 20,1	
		ggoSpread.SSSetFloat	C_W11,		"(11)�� ��"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_W12,		"(12)������" & vbCrLf & "�����"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_W13,		"(13)�󰢺���" & vbCrLf & "�����"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W14,		"(14)�� �� ��" & vbCrLf & "[(11)-(12)+(13)]", 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W15,		"(15)���޼���" & vbCrLf & "����"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat	C_W16,		"(16)��������" & vbCrLf & "[(14)-(15)]"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		.rowheight(-1000) = 20	' ���� ������	(2���� ���, 1���� 15)
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
		Call SetSpreadLock(TYPE_1)
				
		.ReDraw = true	
			
	End With 
	
	' 2�� �׸��� 
	With lgvspdData(TYPE_2)
			
		ggoSpread.Source = lgvspdData(TYPE_2)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W26 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

		'����� 3�ٷ�    
		.ColHeaderRows = 3    
						       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		Call AppendNumberPlace("6","4","0")

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����"		, 10,,,6,1	' �����÷� 
		ggoSpread.SSSetMask     C_W17,	    "(17)" & vbCrLf & "�ձ�" & vbCrLf & "����" & vbCrLf & "����", 5, 2, "9999" 
		ggoSpread.SSSetFloat	C_W18,		"(18)������"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W19,		"(19)��λ�" & vbCrLf & "�غ��" & vbCrLf & "�����ܾ�"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W20,		"(20)����" & vbCrLf & "�غ��" & vbCrLf & "ȯ�Ծ�"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W21,		"(21)�غ��" & vbCrLf & "����" & vbCrLf & "�����"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W22,		"(22)��ü�ҿ�" & vbCrLf & "�ڱݻ���"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W23,		"(23)�̻���"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W24,		"(24)��ü�ҿ�" & vbCrLf & "�ڱݻ���"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W25,		"(25)��Ÿ"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W26,		"(26)��" & vbCrLf & "[(19)-(20)-(21)]"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(C_SEQ_NO , -1000, 1, 3)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W17	, -1000, 1, 3)	
		ret = .AddCellSpan(C_W18	, -1000, 1, 3)
		ret = .AddCellSpan(C_W19	, -1000, 1, 3)
		ret = .AddCellSpan(C_W20	, -1000, 1, 3)
		ret = .AddCellSpan(C_W21	, -1000, 1, 3)
		ret = .AddCellSpan(C_W22	, -1000, 4, 1)
		ret = .AddCellSpan(C_W22	, -999 , 2, 1)
		ret = .AddCellSpan(C_W24	, -999 , 2, 1)
		ret = .AddCellSpan(C_W26	, -1000, 1, 3) 
    
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W22	: .Text = "��  ��  ��"
		
		.Row = -999
		.Col = C_W22	: .Text = "3�� �̰����"
		.Col = C_W24	: .Text = "3�� �����"
		
		.Row = -998
		.Col = C_W22	: .Text = "(22)��ü�ҿ�" & vbCrLf & "�ڱݻ���"
		.Col = C_W23	: .Text = "(23)�̻���"
		.Col = C_W24	: .Text = "(24)��ü�ҿ�" & vbCrLf & "�ڱݻ���"
		.Col = C_W25	: .Text = "(25)��Ÿ"
			
		.rowheight(-998) = 20	' ���� ������	(2���� ���, 1���� 15)
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
		Call SetSpreadLock(TYPE_2)
				
		.ReDraw = true	
			
	End With 
 
	' 3�� �׸��� 
	With lgvspdData(TYPE_3)
			
		ggoSpread.Source = lgvspdData(TYPE_3)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_3,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W33 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

		'����� 2�ٷ�    
		.ColHeaderRows = 2    
						       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'Call AppendNumberPlace("6","3","2")

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����"		, 10,,,6,1	' �����÷� 
		ggoSpread.SSSetMask     C_W17,	    "(17)" & vbCrLf & "�ձ�" & vbCrLf & "����" & vbCrLf & "����", 5, 2, "9999" 
		ggoSpread.SSSetFloat	C_W27,		"(27)1������"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W28,		"(28)2������" & vbCrLf, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W29,		"(29)3������" & vbCrLf, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W30,		"(30)��" & vbCrLf & "[(27)+(28)+(29)]"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W31,		"(31)ȯ����" & vbCrLf & "�ݾ��հ�" & vbCrLf & "[(25)+(30)]"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W32,		"(32)ȸ��ȯ�Ծ�"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W33,		"(33)����ȯ��" & vbCrLf & "����ȯ��" & vbCrLf & "[(31)-(32)]"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(C_SEQ_NO , -1000, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W17	, -1000, 1, 2)	
		ret = .AddCellSpan(C_W27	, -1000, 4, 1)
		ret = .AddCellSpan(C_W31	, -1000, 1, 2)
		ret = .AddCellSpan(C_W32	, -1000, 1, 2)
		ret = .AddCellSpan(C_W33	, -1000, 1, 2)
		
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W27	: .Text = "��ü�ҿ��ڱ� ����(24)�� ȯ���� �ݾ�"
		
		.Row = -999
		.Col = C_W27	: .Text = "(27)1������"
		.Col = C_W28	: .Text = "(28)2������"
		.Col = C_W29	: .Text = "(29)3������"
		.Col = C_W30	: .Text = "(30)��" & vbCrLf & "[(27)+(28)+(29)]"

		.rowheight(-999) = 20	' ���� ������	(2���� ���, 1���� 15)
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
		Call SetSpreadLock(TYPE_3)
				
		.ReDraw = true	
			
	End With     
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
    
    frm1.txtW3.value = lgW3(1) ' ȭ��ǥ�ð� 
    frm1.txtW3_VAL.value = lgW3(0) ' ��갪 
    
	Call GetFISC_DATE

End Sub

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock(Byval pType)

	ggoSpread.Source = lgvspdData(pType)	
	
	Select Case pType
		Case TYPE_1
			ggoSpread.SSSetRequired C_SEQ_NO, -1, C_SEQ_NO
			ggoSpread.SSSetRequired C_W10, -1, C_W10
			ggoSpread.SSSetRequired C_W11, -1, C_W11
			ggoSpread.SSSetRequired C_W12, -1, C_W12
			ggoSpread.SpreadLock C_W14, -1, C_W14
			ggoSpread.SpreadLock C_W16, -1, C_W16
		Case TYPE_2
			ggoSpread.SSSetRequired C_SEQ_NO, -1, C_SEQ_NO
			ggoSpread.SSSetRequired C_W17, -1, C_W17
			ggoSpread.SSSetRequired C_W18, -1, C_W18
			ggoSpread.SSSetRequired C_W19, -1, C_W19
			ggoSpread.SpreadLock C_W25, -1, C_W25
			ggoSpread.SpreadLock C_W26, -1, C_W26
		Case TYPE_3
			ggoSpread.SSSetRequired C_SEQ_NO, -1, C_SEQ_NO
			ggoSpread.SSSetRequired C_W17, -1, C_W17
			ggoSpread.SpreadLock C_W30, -1, C_W30
			ggoSpread.SpreadLock C_W31, -1, C_W31
			ggoSpread.SpreadLock C_W33, -1, C_W33
	End Select
	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	ggoSpread.Source = lgvspdData(pType)

	Select Case pType
		Case TYPE_1
			ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow 	
			ggoSpread.SSSetRequired C_W10, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W11, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W12, pvStartRow, pvEndRow
			ggoSpread.SpreadLock C_W14, pvStartRow, C_W14,pvEndRow
			ggoSpread.SpreadLock C_W16, pvStartRow, C_W16,pvEndRow
			'ggoSpread.SSSetProtected C_W14, pvStartRow, pvEndRow 	
			'ggoSpread.SSSetProtected C_W16, pvStartRow, pvEndRow 	
		Case TYPE_2
			ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow 	
			ggoSpread.SSSetRequired C_W17, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W18, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W19, pvStartRow, pvEndRow
			ggoSpread.SpreadLock C_W25, pvStartRow, C_W14,pvEndRow
			ggoSpread.SpreadLock C_W26, pvStartRow, C_W16,pvEndRow
			'ggoSpread.SSSetProtected C_W25, pvStartRow, pvEndRow 	
			'ggoSpread.SSSetProtected C_W26, pvStartRow, pvEndRow 	
		Case TYPE_3
			ggoSpread.SpreadLock C_SEQ_NO, pvStartRow, C_SEQ_NO,pvEndRow
			ggoSpread.SpreadLock C_W30, pvStartRow, C_W30,pvEndRow
			ggoSpread.SpreadLock C_W31, pvStartRow, C_W31,pvEndRow
			ggoSpread.SpreadLock C_W33, pvStartRow, C_W33,pvEndRow
			
			'ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow 	
			ggoSpread.SSSetRequired C_W17, pvStartRow, pvEndRow
			'ggoSpread.SSSetProtected C_W30, pvStartRow, pvEndRow 	
			'ggoSpread.SSSetProtected C_W31, pvStartRow, pvEndRow 	
			'ggoSpread.SSSetProtected C_W33, pvStartRow, pvEndRow 	
	End Select

End Sub

Sub SetSpreadTotalLine()
	Dim iRow
	For iRow = TYPE_1 To TYPE_3
		ggoSpread.Source = lgvspdData(iRow)
		With lgvspdData(iRow)
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W10 : .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
				ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
			End If
		End With
	Next
End Sub 

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
End Sub

'============================== ���۷��� �Լ�  ========================================

Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
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

    iCalledAspName = "../W5/W5105RA1.ASP"
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W5105RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
   

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

Sub GetFISC_DATE()	' ������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd, iGap
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	'W2�� ��� 
	iGap = DateDiff("m", CDate(lgF0), CDate(lgF1))+1
	
	ReDim lgW2(1)
	If sRepType = "2" Then
		lgW2(1) = "6/" & iGap	' ȭ��ǥ�ð� 
		lgW2(0) = 6/iGap		' ��갪 
	Else
		lgW2(1) = "12/" & iGap 	' ȭ��ǥ�ð� 
		lgW2(0) = 12/iGap		' ��갪 
	End If
	
	frm1.txtW2.value = lgW2(1)
	frm1.txtW2_VAL.value = lgW2(0)
	
End Sub

'====================================== �� �Լ� =========================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
	lgCurrGrid = TYPE_1	' �⺻ �׸��� 
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	lgCurrGrid = TYPE_2
End Function


'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100101111")										<%'��ư ���� ���� %>
	  
	' �����Ѱ� 
	Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)

	
	Call InitData 
	
    Call MainQuery() 
    
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

Sub txtW5_Change()
	Call SetHeadReCalc
	lgBlnFlgChgValue = True
End Sub

Sub txtW7_Change()
	Call SetHeadReCalc
	lgBlnFlgChgValue = True
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
	'Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData0_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_1
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
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
	'Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_2
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

' -- 2�� �׸��� 
Sub vspdData2_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_3
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_3
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_3
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_3
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_GotFocus()
	lgCurrGrid = TYPE_3
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_3
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_3
	'Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_3
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, dblW11, dblW12, dblW13 , dblFiscYear, dblW26, dblW25, dblW24, dblW23, dblW22, dblW17, dblW15, dblW14
	Dim dblW27, dblW28, dblW29, dblW30, dblW31, dblW32, dblW33, dblW2_VAL
	
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
		Select Case Col
			Case C_W11, C_W12, C_W13, C_W15
				.Col = C_W11 : dblW11 = UNICDbl(.Value)
				.Col = C_W12 : dblW12 = UNICDbl(.Value)
				.Col = C_W13 : dblW13 = UNICDbl(.Value)
				
				
				
				
				.Col = C_W14 : dblW14 = dblW11 - dblW12 + dblW13 : .value = dblW14
				.Col = C_W15 : dblW15 = UNICDbl(.Value)
				.Col = C_W16 : .value = dblW14 - dblW15
				
			
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' �հ� 
				
				Call SetW14()	
				Call SetW16()
				Call SetHeadReCalc()
		End Select 
		ggoSpread.Source = lgvspdData(Index)
		ggoSpread.UpdateRow .MaxRows
	ElseIf Index = TYPE_2 Then	'1�� �׸� 
	
		Select Case Col
			Case C_W17	' ������ ����� 
				dblFiscYear = UNICDbl(frm1.txtFISC_YEAR.text)
				.Col = C_W17	: .Row = Row	: dblW17 = UNiCDbl(.Value)
				If dblFiscYear - 5 > dblW17 Or dblFiscYear < dblW17 Then
					Call DisplayMsgBox("W40002", parent.VB_INFORMATION, "X", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.Value = ""
					Exit Sub
				End If
			Case C_W18, C_W19, C_W20, C_W21, C_W22, C_W23, C_W24
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.Value)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.Value = 0
				End If
				
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' �հ� 
				
				Call SetW26_W25(Row)
		End Select
		ggoSpread.Source = lgvspdData(Index)
		ggoSpread.UpdateRow .MaxRows
	ElseIf Index = TYPE_3 Then
		Select Case Col
			Case C_W17
				.Col = Col	: .Row = Row	: dblW17 = UNICDbl(.value)
				Call GetW24(dblW17, dblW24, dblW25)
				
				If dblW24 = -1 Then
					Call DisplayMsgBox("W40001", parent.VB_INFORMATION, "X", "X")           '��: "(17)�Աݻ��Կ����� �߰��Ҽ������ϴ�."
					.Value = ""
					Exit Sub
				End If

				dblFiscYear = UNICDbl(frm1.txtFISC_YEAR.text)
						
				.Row = Row
				.Col = C_W27 : .Value = 0	: dblW27 = 0
				.Col = C_W28 : .Value = 0	: dblW28 = 0
				.Col = C_W29 : .Value = 0	: dblW29 = 0
				
				dblW2_VAL = UNICDbl(frm1.txtW2_VAL.value)
				If (dblFiscYear - dblW17) = 3 Then
					.Col = C_W27
					dblW27 = UNICDbl(UNIFormatNumber( (dblW24/3) * dblW2_VAL , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)) 
					.Value = dblW27
				ElseIf (dblFiscYear - dblW17) = 4 Then
					.Col = C_W28
					dblW28 = UNICDbl(UNIFormatNumber( (dblW24/2) * dblW2_VAL , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)) 
					.Value = dblW28
				ElseIf (dblFiscYear - dblW17) = 5 Then
					.Col = C_W29
					dblW29 = dblW24 * dblW2_VAL
					.Value = dblW29
				End If
		
				Call SetGridTYPE_3(Row)
			Case C_W27, C_W28, C_W29, C_W32
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.Value)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.Value = 0
				End If
				Call SetGridTYPE_3(Row)		
		End Select
		ggoSpread.Source = lgvspdData(Index)
		ggoSpread.UpdateRow .MaxRows
	End If

	End With
	
End Sub

Sub SetGridTYPE_3(Byval Row)
	Dim dblSum, dblW11, dblW12, dblW13 , dblFiscYear, dblW26, dblW25, dblW24, dblW23, dblW22, dblW17, dblW15, dblW14
	Dim dblW27, dblW28, dblW29, dblW30, dblW31, dblW32, dblW33

	With lgvspdData(TYPE_3)
		.Row = Row
		.Col = C_W17 : dblW17 = UNICDbl(.value)
		.Col = C_W27 : dblW27 = UNICDbl(.Value)
		.Col = C_W28 : dblW28 = UNICDbl(.Value)
		.Col = C_W29 : dblW29 = UNICDbl(.Value)
									
		' �հ躯�� 
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W27, 1, .MaxRows - 1, true, .MaxRows, C_W27, "V")	' �հ� 
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W28, 1, .MaxRows - 1, true, .MaxRows, C_W28, "V")	' �հ� 
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W29, 1, .MaxRows - 1, true, .MaxRows, C_W29, "V")	' �հ� 
					
		' W30 ���� 
		dblW30 = dblW27 + dblW28 + dblW29
		.Col = C_W30	: .Row = Row : .Value = dblW30
					
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W30, 1, .MaxRows - 1, true, .MaxRows, C_W30, "V")	' �հ� 
					
		' W31 ���� 
		Dim iGrid1Row
		iGrid1Row = GetRowByW17(TYPE_2, dblW17)	' �׸���1���� �ش� ������ ã�´�.
		If iGrid1Row > 0 Then	
			dblW25 = UNICDbl(GetGrid(TYPE_2, C_W25, iGrid1Row))
			dblW31 = dblW25 + dblW30
			.Col = C_W31	: .Row = Row : .Value = dblW31

			Call FncSumSheet(lgvspdData(lgCurrGrid), C_W31, 1, .MaxRows - 1, true, .MaxRows, C_W31, "V")	' �հ� 
		End If
		
		.Row = Row			
		.Col = C_W32	: dblW32 = UNICDbl(.Value)
		' W33 ���� 
		dblW33 = dblW31 - dblW32
		.Col = C_W33	: .Value = dblW33 
					
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W33, 1, .MaxRows - 1, true, .MaxRows, C_W33, "V")	' �հ�	
	End With
End Sub

' W17 (���Կ���)�� ���� ã�´� 
Function GetRowByW17(Byval pType, Byval pW17)
	Dim iMaxRows, iRow
	With lgvspdData(pType)
		iMaxRows = .MaxRows
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = C_W17 
			If CStr(.Value) = CStr(pW17) Then
				GetRowByW17 = iRow
				Exit Function
			End If
		Next
	End With
	GetRowByW17 = -1
End Function

Function GetGrid(Byval pType, Byval pCol, Byval pRow)
	With lgvspdData(pType)
		.Col = pCol : .Row = pRow : GetGrid = .value
	End With
End Function

' 3�� �׸��忡�� 2�� �׸����� ����Ÿ�� ã�Ƽ� W24�ݾ��� �����Ѵ� 
Sub GetW24(Byval pYear , Byref pdblW24, Byref pdblW25)
	Dim iRow, iMaxRows
	With lgvspdData(TYPE_2)
		iMaxRows = .MaxRows - 1
		.Col = C_W17
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			If UNICDbl(.Value) = pYear Then
				.Col = C_W24 : pdblW24 = UNICDbl(.Value)
				.Col = C_W25 : pdblW25 = UNICDbl(.Value)
				Exit Sub
			End If
		Next
		pdblW24 = -1 : pdblW25 = -1
	End With
End Sub

' ������ 
Sub SetW14()	
	Dim dblSum
	
	With lgvspdData(TYPE_1)
		dblSum = FncSumSheet(lgvspdData(TYPE_1), C_W14, 1, .MaxRows - 1, true, .MaxRows, C_W14, "V")	' �հ� 
	End With	

End Sub

' �������񰡾� 
Sub SetW16()	
	Dim dblSum
	
	With lgvspdData(TYPE_1)
		dblSum = FncSumSheet(lgvspdData(TYPE_1), C_W16, 1, .MaxRows - 1, true, .MaxRows, C_W16, "V")	' �հ� 
	End With	
End Sub

' ��� ���� 
Sub SetHeadReCalc()	
	Dim dblSum, dblW16, dblW3, dblW2, dblW4, dblW5, dblW6, dblW7
	
	If IsRunEvents Then Exit Sub
	
	IsRunEvents = True
	
	With lgvspdData(TYPE_1)
		If .MaxRows = 0 Then Exit Sub
		.Col = C_W16 : .Row = .MaxRows : dblW16 = UNICDbl(.Value)
	End With	
	
	With frm1
		.txtW1.value = dblW16
		dblW2 = UNICDbl(.txtW2_VAL.value)
		dblW3 = UNICDbl(.txtW3_VAL.value)
		dblW4 = UNICDbl(UNIFormatNumber(dblW16 * dblW2 * dblW3 , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit))  ' �Ҽ����� �����ϴ� ������ �ش��Լ��� ����� 
		.txtW4.value = dblW4
		dblW5 = UNICDbl(.txtW5.value)
		If (dblW5 - dblW4) > 0 Then
			dblW6 = dblW5 - dblW4
		Else	
			dblW6 = 0
		End If
		dblW7 = UNICDbl(.txtW7.value)
		
		.txtW6.value = dblW6
		.txtW8.value = dblW6 + dblW7
		.txtW9.value = dblW5 - dblW6 - dblW7
	End With
	
	IsRunEvents = False
	lgBlnFlgChgValue = True
End Sub

' �ܾ� �÷��� ����ɶ� ȣ��� 
Sub SetW26_W25(Row)
	Dim dblSum, dblW19, dblW20, dblW21, dblW26, dblW24, dblW23, dblW22, dblW25, dblW17, iRow
	
	With lgvspdData(TYPE_2)
		
		.Row = Row
		.Col = C_W19	: dblW19 = UNICDbl(.Value)	' ���� 
		.Col = C_W20	: dblW20 = UNICDbl(.Value)	' �뺯 
		.Col = C_W21	: dblW21 = UNICDbl(.Value)	' ���� 
		.Col = C_W22	: dblW22 = UNICDbl(.Value)	' �뺯 
		.Col = C_W23	: dblW23 = UNICDbl(.Value)	' ���� 
		.Col = C_W24	: dblW24 = UNICDbl(.Value)	' �뺯 
		.Col = C_W26	: dblW26 = dblW19 - dblW20 - dblW21	: .Value = dblW26
		.Col = C_W25	: dblW25 = dblW26 - dblW22 - dblW23 - dblW24	: .Value = dblW25

		Call FncSumSheet(lgvspdData(TYPE_2), C_W25, 1, .MaxRows - 1, true, .MaxRows, C_W25, "V")	' �հ� 
		Call FncSumSheet(lgvspdData(TYPE_2), C_W26, 1, .MaxRows - 1, true, .MaxRows, C_W26, "V")	' �հ� 
		
		If lgvspdData(TYPE_3).MaxRows > 0 Then
			dblW17 = GetGrid(TYPE_2, C_W17, Row)	' �������� ���Կ����� ���Ѵ�.
			iRow = GetRowByW17(TYPE_3, dblW17)
			If iRow > 0 Then Call vspdData_Change(TYPE_3, C_W17, iRow)	' �ձݻ��Կ����� �������� �߰ߵǸ� ..
		End If
	End With
	
End Sub

Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = lgvspdData(Index)
   
    If lgvspdData(Index).MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    'If Row <= 0 Then
     '  ggoSpread.Source = lgvspdData(Index)
     '  
     '  If lgSortKey = 1 Then
     '      ggoSpread.SSSort Col               'Sort in ascending
     '      lgSortKey = 2
     '  Else
     '      ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
     '      lgSortKey = 1
     '  End If
     '  
     '  Exit Sub
    'End If

	lgvspdData(Index).Row = Row
End Sub

Sub vspdData_ColWidthChange(Index, ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = lgvspdData(Index)
    'Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
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
    ggoSpread.Source = frm1.vspdData0
    'Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    'Call GetSpreadColumnPos("A")
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
   
	
	If lgBlnFlgChgValue = True Then
		blnChange = True
	End If
	
	
	For i = TYPE_1 To TYPE_3
    
		ggoSpread.Source = lgvspdData(i)
		If ggoSpread.SSCheckChange = True Then
			blnChange = True
		End If
	Next
	
	
	
	If blnChange = False Then
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

' ----------------------  ���� -------------------------
Function  Verification()
	Dim dblW11, dblW12, dblW15,dblW16, dblW19, dblW22, dblW23, dblW24, dblW21, dblW20
	
	Verification = False
	
	With lgvspdData(TYPE_1)
		If .MaxRows > 0 Then
		
			.Row = .MaxRows
			'1. W11 < W12
			.Col = C_W11 : dblW11 = UNICDbl(.Value)
			If dblW11 < 0 Then
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "(11) �ݾ� ", "")                          <%'No data changed!!%>
				Exit Function
			End If
			.Col = C_W12 : dblW12 = UNICDbl(.Value)
			If dblW12 < 0 Then
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "(12)�����󰢴���� ", "")                          <%'No data changed!!%>
				Exit Function
			End If
			
			If dblW11 < dblW12 Then
				Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, "(12)�����󰢴����", "(11)�ݾ�")                          <%'No data changed!!%>
				Exit Function
			End If
			
			.Col = C_W15 : dblW15 = UNICDbl(.Value)
			If dblW15 < 0 Then
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "(15)���޼��񰡾�", "")                          <%'No data changed!!%>
				Exit Function
			End If
			
			.Col = C_W16 : dblW16 = UNICDbl(.Value)
			If dblW16 < 0 Then
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "(16)�������񰡾�", "")                          <%'No data changed!!%>
				Exit Function
			End If
			
			'2. W16 < 0
			.Col = C_W16 : dblW16 = UNICDbl(.Value)
			If dblW16 < 0 Then
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "(16)�������񰡾�", "")                          <%'No data changed!!%>
				Exit Function
			End If
		End If
	End With
	
	With lgvspdData(TYPE_2)	
		If .MaxRows > 0 Then
			'3. W19 < W22 + W23
			.Col = C_W19 : dblW19 = UNICDbl(.Value)
			.Col = C_W22 : dblW22 = UNICDbl(.Value)
			.Col = C_W23 : dblW23 = UNICDbl(.Value)
			If dblW19 < dblW22 + dblW23 Then
				Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, "������[(W22)+(W23)]", "(W19)��λ� �غ�� �����ܾ�")                          <%'No data changed!!%>
				Exit Function
			End If		
		
			'4. W19 < W24
			.Col = C_W24 : dblW24 = UNICDbl(.Value)
			If dblW19 < dblW24 Then
				Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, "������[(W24)]", "(W19)��λ� �غ�� �����ܾ�")                          <%'No data changed!!%>
				Exit Function
			End If		
		
			'5. W19 < W20
			.Col = C_W20 : dblW20 = UNICDbl(.Value)
			If dblW19 < dblW20 Then
				Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, "(W20)���� �غ�� ȯ�Ծ�", "(W19)��λ� �غ�� �����ܾ�")                          <%'No data changed!!%>
				Exit Function
			End If		
		
			'6. W19 < W21
			.Col = C_W21 : dblW21 = UNICDbl(.Value)
			If dblW19 < dblW21 Then
				Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, "(W21)�غ�� ���δ����", "(W19)��λ� �غ�� �����ܾ�")                          <%'No data changed!!%>
				Exit Function
			End If		
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

    Call SetToolbar("1110110100001111")

	Call ClickTab1()
	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
 
End Function

Function FncCancel() 

    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
    
    If lgvspdData(lgCurrGrid).MaxRows = 1 Then
		ggoSpread.EditUndo 
	Else
		Call ReCalcGridSum(lgCurrGrid)
    End If

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
		lgvspdData(lgCurrGrid).ReDraw = False
					
		If .MaxRows = 0 Then	' ù InsertRow�� 1��+�հ��� 

			iRow = 1
			ggoSpread.InsertRow , 2
			Call SetSpreadColor(lgCurrGrid, iRow, iRow+1) 
			.Row = iRow		
			
			If lgCurrGrid = TYPE_1 Then
				.Col = C_SEQ_NO : .Text = iRow	
			
				iRow = 2		: .Row = iRow
				.Col = C_SEQ_NO : .Text = SUM_SEQ_NO	
				.Col = C_W10	: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
				ggoSpread.SpreadLock C_W10, iRow, C_W16, iRow
			Else
				.Col = C_SEQ_NO : .Text = iRow	
			
				iRow = 2		: .Row = iRow
				.Col = C_SEQ_NO : .Text = SUM_SEQ_NO	
				.Col = C_W17	: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
				ggoSpread.SpreadLock C_W17, iRow, .MaxCols-1, iRow
			End If		
		
		Else
				
			If iRow = .MaxRows Then	' -- ������ �հ��ٿ��� InsertRow�� �ϸ� ������ �߰��Ѵ�.
				ggoSpread.InsertRow iRow-1 , imRow 
				SetSpreadColor lgCurrGrid, iRow, iRow + imRow - 1

				'If lgCurrGrid = TYPE_1 Then
					Call SetDefaultVal(lgCurrGrid, iRow, imRow)
				'End If
			Else
				ggoSpread.InsertRow ,imRow
				SetSpreadColor lgCurrGrid, iRow+1, iRow + imRow

				'If lgCurrGrid = TYPE_1 Then
					Call SetDefaultVal(lgCurrGrid, iRow+1, imRow)
				'End If
			End If   
		End If
	End With
	
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
		.Col = C_W1		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
		
		ggoSpread.SpreadLock C_W1, 1, C_W6, 1
	End If
	End With
End Function

' �׸��忡 SEQ_NO, TYPE �ִ� ���� 
Function SetDefaultVal(Index, iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With lgvspdData(lgCurrGrid)	' ��Ŀ���� �׸��� 

	ggoSpread.Source = lgvspdData(lgCurrGrid)
	
	If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
		.Row = iRow
		MaxSpreadVal lgvspdData(lgCurrGrid), C_SEQ_NO, iRow
	Else
		iSeqNo = MaxSpreadVal(lgvspdData(lgCurrGrid), C_SEQ_NO, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i
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
		
		Call ReCalcGridSum(lgCurrGrid)
	End With

End Function

Function ReCalcGridSum(Byval pType)
	Dim iCol, iMaxRows, iMaxCols
	With lgvspdData(pType)
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		iMaxRows = .MaxRows	: iMaxCols = .MaxCols-1
		For iCol = 4 To iMaxCols
			Call FncSumSheet(lgvspdData(pType), iCol, 1, .MaxRows - 1, true, .MaxRows, iCol, "V")	' �հ� 
		Next
		If pType <> TYPE_3 Then
			Call FncSumSheet(lgvspdData(pType), 3, 1, .MaxRows - 1, true, .MaxRows, 3, "V")	' �հ� 
		End If
		ggoSpread.UpdateRow .MaxRows
		lgBlnFlgChgValue = True
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

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	
	lgIntFlgMode = parent.OPMD_UMODE
	If lgvspdData(TYPE_1).MaxRows > 0 Or _
		lgvspdData(TYPE_2).MaxRows > 0 Or _
		lgvspdData(TYPE_3).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		
		    
		' �������� ���� : ���ߵǸ� ���ȴ�.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 ����üũ : �׸��� �� 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1
			Call SetSpreadLock(TYPE_1)
			Call SetSpreadLock(TYPE_2)
			Call SetSpreadLock(TYPE_3)
			'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
			Call SetToolbar("1111111100011111")										<%'��ư ���� ���� %>
			
			'3. ������/������� ���� üũ(�������� ���� üũ)
			With frm1
				If .txtW2.value <> lgW2(1) Then
				ElseIf .txtW3.value <> lgW3(1) Then
				End If
			End With
		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1110000000011111")										<%'��ư ���� ���� %>
		End If
	Else
		Call SetToolbar("1100110100111111")										<%'��ư ���� ���� %>
	End If
	
	Call SetSpreadTotalLine ' - �հ���� �籸�� 
	
	Call ClickTab1()
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
    
    For i = TYPE_1 To TYPE_3	' ��ü �׸��� ���� 
    
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
	Frm1.txtHeadMode.value    =  lgIntFlgMode
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
	Call InitVariables
	Call FncNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
					<TD CLASS="CLSMTAB">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()" width=200>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ձݻ��Ծ�/������ڻ�</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTAB">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" width=200>
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ͱݻ��Ծ�</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
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
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=15%>
                                   <table <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
									   <TR>
										   <TD width="100%" COLSPAN=9 CLASS="CLSFLD"><br>&nbsp;1. �ձݻ��Ծ�����</TD>
									   </TR>
									   <TR>
										   <TD CLASS="TD51" width="10%" ALIGN=CENTER>(1)������ڻ갡��</TD>
										   <TD CLASS="TD51" width="10%" ALIGN=CENTER>(2)������� ����</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(3)������</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER COLSPAN=2>(4)�ѵ���  [(1) x (2) x (3)] </TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER COLSPAN=2>(5)ȸ�����</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER>(6)�ѵ��ʰ��� [(5) - (4) ]</TD>
									   </TR>
									  <TR>
											<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1" name=txtW1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD><INPUT TYPE=TEXT id="txtW2" name=txtW2 tag="24X2Z" Style="width : 100%; text-align: center"><INPUT TYPE=HIDDEN ID="txtW2_VAL" NAME="txtW2_VAL"></TD>
											
											<!--<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPMSK%> name=txtW2 CLASS=FPDS100 title=FPDOUBLESINGLE ALT="" tag="21X2Z" id=OBJECT1></OBJECT>');</SCRIPT>-->
											
											<TD><INPUT TYPE=TEXT id="txtW3" name=txtW3 tag="24X2Z" Style="width : 100%; text-align: center"><INPUT TYPE=HIDDEN ID="txtW3_VAL" NAME="txtW3_VAL"></TD>
											<TD COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4" name=txtW4 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW5" name=txtW5 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW6" name=txtW6 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									   <TR>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER COLSPAN=2>(7)�����Ѽ� ���뿡 ���� �ձݺ��ξ�</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER COLSPAN=2>(8)�ձݺһ��� �� [(6) + (7)]</TD>
								           <TD CLASS="TD51" width="10%" ALIGN=CENTER COLSPAN=2>(9)�ձݻ��Ծ� [(5) - (6) - (7)]</TD>
									       <TD CLASS="TD51" width="10%" ALIGN=CENTER COLSPAN=2>�� ��</TD>
									  </TR>
									  <TR>
											<TD COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW7" name=txtW7 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8" name=txtW8 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW9" name=txtW9 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
											<TD COLSPAN=2><INPUT TYPE=TEXT id="txtDESC1" name=txtDESC1 ALT="�� ��" tag="25X2Z" Style="width : 100%"></TD>
									  </TR>
								  </table>
								</TD>
							</TR>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
									<table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
									   <TR>
										   <TD width="100%" HEIGHT=10 CLASS="CLSFLD"><br>&nbsp;2. ������ڻ�� ������ ���</TD>
									   </TR>
									   <TR>
										   <TD width="100%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData0 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData0> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										  </TD>
									  </TR>
									</TABLE>
								</TD>
							</TR>
							</TABLE>
							</DIV>
							<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=65%>
									<table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
									   <TR>
										   <TD width="100%" HEIGHT=10 CLASS="CLSFLD"><br>&nbsp;3. �ͱݻ��Ծ��� ����</TD>
									   </TR>
									   <TR>
										   <TD width="100%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
										  </TD>
									  </TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
									<table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
									   <TR>
										   <TD width="100%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread0" tag="24" STYLE="DISPLAY: 'NONE'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24" STYLE="DISPLAY: 'NONE'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" STYLE="DISPLAY: 'NONE'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtHeadMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" STYLE="DISPLAY: 'NONE'"></iframe>
</DIV>
</BODY>
</HTML>

