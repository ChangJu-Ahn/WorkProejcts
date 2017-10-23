<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �ҵ�ݾ����� 
'*  3. Program ID           : w5103mA1
'*  4. Program Name         : w5103mA1.asp
'*  5. Program Desc         : ��15ȣ �ҵ�ݾ� �������� 
'*  6. Modified date(First) : 2005/02/02
'*  7. Modified date(Last)  : 2005/02/02
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
<SCRIPT LANGUAGE="VBScript"  SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "w5103mA1"
Const BIZ_PGM_ID		= "w5103mB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const JUMP_PGM_ID		= "W5101MA1"

Const TAB1 = 1																	'��: Tab�� ��ġ 
Const TAB2 = 2
Const TAB3 = 3

Const TYPE_1	= 0		' �׸��带 �������� ���� ��� 
Const TYPE_2	= 1		
Const TYPE_3	= 2		 
Const TYPE_4	= 3		 

' -- �׸��� �÷� ���� 
Dim C_W_TYPE
Dim C_SEQ_NO
Dim C_W1
Dim C_W1_BT
Dim C_W1_NM
Dim C_W2
Dim C_W3_NM
Dim C_W3
Dim C_W4


Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(4)

Dim lgW_NM(8), lgRB(3)

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_W_TYPE	= 1
	C_SEQ_NO	= 2
	C_W1		= 3
	C_W1_BT		= 4
	C_W1_NM		= 5
	C_W2		= 6
	C_W3_NM		= 7
	C_W3		= 8
	C_W4		= 9
	
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

	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " MAJOR_CD='W1001' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = lgvspdData(TYPE_1)
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W3
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W3_NM
	End If

	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " MAJOR_CD='W1002' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = lgvspdData(TYPE_2)
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W3
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W3_NM
	End If
End Sub

Function OpenAdItem(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�������� �˾�"					<%' �˾� ��Ī %>
	arrParam(1) = "TB_ADJUST_ITEM"					<%' TABLE ��Ī %>
	

	If iWhere = TYPE_1 then
		frm1.vspdData0.Col = C_W1
	    arrParam(2) = frm1.vspdData0.Text		<%' Code Condition%>
	ElseIf iWhere = TYPE_2 then
		frm1.vspdData1.Col = C_W1
	    arrParam(2) = frm1.vspdData1.Text		<%' Code Condition%>
	End If
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = " USE_YN = '1' "							<%' Where Condition%>
	arrParam(5) = "��������"						<%' �����ʵ��� �� ��Ī %>
	
    arrField(0) = "ITEM_CD"					<%' Field��(0)%>
    arrField(1) = "ITEM_NM"					<%' Field��(1)%>
    
    arrHeader(0) = "��������"					<%' Header��(0)%>
    arrHeader(1) = "�����"						<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAdItem(arrRet,iWhere)
		Call SetGrid34ByGrid12(iWhere, iWhere + 2, lgvspdData(iWhere).Row)	' ��3.�׸���3, 4�� �����Ѵ�.
	End If	
	
End Function

Function SetAdItem(byval arrRet,Byval iWhere)
    With frm1
		If iWhere = TYPE_1 Then 'Spread1(Condition)
			.vspdData0.Col = C_W1
			.vspdData0.Text = arrRet(0)
			.vspdData0.Col = C_W1_NM
			.vspdData0.Text = arrRet(1)
		    ggoSpread.Source = frm1.vspdData0
		    ggoSpread.UpdateRow frm1.vspdData0.ActiveRow
			lgBlnFlgChgValue = True
		ElseIf iWhere = TYPE_2 Then 'Spread2(Condition)
			.vspdData1.Col = C_W1
			.vspdData1.Text = arrRet(0)
			.vspdData1.Col = C_W1_NM
			.vspdData1.Text = arrRet(1)
		    ggoSpread.Source = frm1.vspdData1
		    ggoSpread.UpdateRow frm1.vspdData1.ActiveRow
			lgBlnFlgChgValue = True
		End If
	End With
End Function

Function GetAdItem(Byval iWhere, ByVal pCol, ByVal pRow)
	Dim arrRet(2), sWhere, bRet

	If pCol = C_W1 Then
		sWhere = " ITEM_CD LIKE '%"
	ElseIf pCol = C_W1_NM Then
		sWhere = " ITEM_NM LIKE '%"
	Else
		Exit Function
	End If

	lgvspdData(iWhere).Col = pCol
	If lgvspdData(iWhere).Text <> "" Then
		sWhere = sWhere & lgvspdData(iWhere).Text & "%' "		<%' Code Condition%>
	
		bRet = CommonQueryRs("top 1 ITEM_CD,ITEM_NM"," TB_ADJUST_ITEM ",sWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		arrRet(0) = Replace(lgF0, chr(11), "")
		arrRet(1) = Replace(lgF1, chr(11), "")
	Else
		arrRet(0) = ""
		arrRet(1) = ""
	End If
	
	Call SetAdItem(arrRet,iWhere)
	Call SetGrid34ByGrid12(iWhere, iWhere + 2, pRow)	' ��3.�׸���3, 4�� �����Ѵ�.
	
End Function


Sub InitSpreadSheet()
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1)		= frm1.vspdData0
	Set lgvspdData(TYPE_2)		= frm1.vspdData1
	Set lgvspdData(TYPE_3)		= frm1.vspdData2
	Set lgvspdData(TYPE_4)		= frm1.vspdData3
		
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("6","3","2")
	
	' 1�� �׸��� 

	With lgvspdData(TYPE_1)
				
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W4 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
 
		.MaxRows = 0
		ggoSpread.ClearSpreadData

	    ggoSpread.SSSetEdit		C_W_TYPE,	"����Ÿ����",		5,,,6,1	' �����÷� 
	    ggoSpread.SSSetEdit		C_SEQ_NO,	"����",				5,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W1,		"(1)����",			7,,,10,1
	    ggoSpread.SSSetButton 	C_W1_BT
		ggoSpread.SSSetEdit		C_W1_NM,	"(1)�����",		20,,,50,1
		ggoSpread.SSSetFloat	C_W2,		"(2)�ݾ�",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetCombo	C_W3_NM,	"(3)ó��",		10
	    ggoSpread.SSSetCombo	C_W3,		"(3)ó��",		10
		ggoSpread.SSSetEdit		C_W4,		"(4)��������",	50,,,100,1

		ret = .AddCellSpan(C_W1, -1000, 3, 1)	' ���� 2�� ��ħ 

		Call ggoSpread.MakePairsColumn(C_W1,C_W1_BT)

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W3,C_W3,True)
		
		Call SetSpreadLock(TYPE_1)

		.ReDraw = true	
				
	End With 

 
	' 2�� �׸��� 
	With lgvspdData(TYPE_2)
				
		ggoSpread.Source = lgvspdData(TYPE_2)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W4 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
	
		.MaxRows = 0
		ggoSpread.ClearSpreadData

	    ggoSpread.SSSetEdit		C_W_TYPE,	"����Ÿ����",		5,,,6,1	' �����÷� 
	    ggoSpread.SSSetEdit		C_SEQ_NO,	"����",				5,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W1,		"(1)����",			7,,,10,1
	    ggoSpread.SSSetButton 	C_W1_BT
		ggoSpread.SSSetEdit		C_W1_NM,	"(1)�����",		20,,,50,1
		ggoSpread.SSSetFloat	C_W2,		"(2)�ݾ�",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetCombo	C_W3_NM,	"(3)ó��",		10
	    ggoSpread.SSSetCombo	C_W3,		"(3)ó��",		10
		ggoSpread.SSSetEdit		C_W4,		"(4)��������",	50,,,100,1

		ret = .AddCellSpan(C_W1, -1000, 3, 1)	' ���� 2�� ��ħ 

		Call ggoSpread.MakePairsColumn(C_W1,C_W1_BT)

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W3,C_W3,True)
		
		Call SetSpreadLock(TYPE_2)

		.ReDraw = true	
				
	End With 

	' 3�� �׸��� 

	With lgvspdData(TYPE_3)
				
		ggoSpread.Source = lgvspdData(TYPE_3)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_3,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W4 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
	
  		'����� 3�ٷ�    
		.ColHeaderRows = 3
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData

	    ggoSpread.SSSetEdit		C_W_TYPE,	"����Ÿ����",		5,,,6,1	' �����÷� 
	    ggoSpread.SSSetEdit		C_SEQ_NO,	"����",				5,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W1,		"(1)����",			7,,,10,1
	    ggoSpread.SSSetButton 	C_W1_BT
		ggoSpread.SSSetEdit		C_W1_NM,	"(1)�����",		15,,,50,1
		ggoSpread.SSSetFloat	C_W2,		"(2)�ݾ�",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetEdit		C_W3_NM,	"ó��",		10,,,50,1
	    ggoSpread.SSSetEdit		C_W3,		"�ڵ�",		10,,,50,1
		ggoSpread.SSSetEdit		C_W4,		"(4)��������",	20,,,100,1

		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(0		, -1000, 1, 3)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W1		, -1000, 7, 1)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W1		, -999, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W1_BT	, -999, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W1_NM	, -999, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W2		, -999, 1, 2)	
		ret = .AddCellSpan(C_W3_NM	, -999, 2, 1)
		ret = .AddCellSpan(C_W4 	, -999, 1, 2)
    
    
    
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W1		: .Text = "�ͱݻ��� �� �ձݺһ���"
		
		' ù��° ��� ��� ���� 
		.Row = -999
		.Col = C_W1_NM	: .Text = "(1)����"
		.Col = C_W2		: .Text = "(2)�ݾ�"
		.Col = C_W3_NM	: .Text = "(3)�ҵ�ó��"
		
		.Row = -998
		.Col = C_W3_NM	: .Text = "ó��"
		.Col = C_W3		: .Text = "�ڵ�"
								
		.rowheight(-999) = 15	' ���� ������ 
		.rowheight(-998) = 15	' ���� ������ 
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W1,C_W1,True)
		Call ggoSpread.SSSetColHidden(C_W1_BT,C_W1_BT,True)
		Call ggoSpread.SSSetColHidden(C_W4,C_W4,True)
				
		Call SetSpreadLock(TYPE_3)

		.ReDraw = true	
				
	End With

	' 4�� �׸��� 

	With lgvspdData(TYPE_4)
				
		ggoSpread.Source = lgvspdData(TYPE_4)
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_4,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W4 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
	
  		'����� 3�ٷ�    
		.ColHeaderRows = 3
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData

	    ggoSpread.SSSetEdit		C_W_TYPE,	"����Ÿ����",		5,,,6,1	' �����÷� 
	    ggoSpread.SSSetEdit		C_SEQ_NO,	"����",				5,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W1,		"(1)����",			7,,,10,1
	    ggoSpread.SSSetButton 	C_W1_BT
		ggoSpread.SSSetEdit		C_W1_NM,	"(1)�����",		15,,,50,1
		ggoSpread.SSSetFloat	C_W2,		"(2)�ݾ�",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
	    ggoSpread.SSSetEdit		C_W3_NM,	"�ڵ�",		10,,,50,1
	    ggoSpread.SSSetEdit		C_W3,		"ó��",		10,,,50,1
		ggoSpread.SSSetEdit		C_W4,		"(4)��������",	20,,,100,1

		' �׸��� ��� ��ħ 
		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(0		, -1000, 1, 3)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W1		, -1000, 7, 1)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W1		, -999, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W1_BT	, -999, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W1_NM	, -999, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W2		, -999, 1, 2)	
		ret = .AddCellSpan(C_W3_NM	, -999, 2, 1)
		ret = .AddCellSpan(C_W4 	, -999, 1, 2)
    
    
    
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W1		: .Text = "�ձݻ��� �� �ͱݺһ���"
		
		' ù��° ��� ��� ���� 
		.Row = -999
		.Col = C_W1_NM	: .Text = "(1)����"
		.Col = C_W2		: .Text = "(2)�ݾ�"
		.Col = C_W3_NM	: .Text = "(3)�ҵ�ó��"
		
		.Row = -998
		.Col = C_W3_NM	: .Text = "ó��"
		.Col = C_W3		: .Text = "�ڵ�"
								
		.rowheight(-999) = 15	' ���� ������ 
		.rowheight(-998) = 15	' ���� ������ 
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W1,C_W1,True)
		Call ggoSpread.SSSetColHidden(C_W1_BT,C_W1_BT,True)
		Call ggoSpread.SSSetColHidden(C_W4,C_W4,True)

		Call SetSpreadLock(TYPE_4)
				
		.ReDraw = true	
				
	End With
	
	Call InitSpreadComboBox
	
					
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


Sub SetSpreadLock(pType)

	With lgvspdData(pType)
	
		ggoSpread.Source = lgvspdData(pType)	

		Select Case pType
			Case TYPE_1 
				ggoSpread.SpreadUnLock C_W_TYPE, -1, C_W4	' ��ü ���� 
				ggoSpread.SSSetRequired C_W1, -1, -1
				ggoSpread.SSSetRequired C_W1_NM, -1, -1
				ggoSpread.SSSetRequired C_W2, -1, -1
				ggoSpread.SSSetRequired C_W3, -1, -1
				ggoSpread.SSSetRequired C_W3_NM, -1, -1

			Case TYPE_2
				ggoSpread.SpreadUnLock C_W_TYPE, -1, C_W4	' ��ü ���� 
				ggoSpread.SSSetRequired C_W1, -1, -1
				ggoSpread.SSSetRequired C_W1_NM, -1, -1
				ggoSpread.SSSetRequired C_W2, -1, -1
				ggoSpread.SSSetRequired C_W3, -1, -1
				ggoSpread.SSSetRequired C_W3_NM, -1, -1

			Case TYPE_3
				ggoSpread.SpreadLock C_W_TYPE,   -1, C_W4

			Case TYPE_4
				ggoSpread.SpreadLock C_W_TYPE,   -1, C_W4

		End Select
		
	End With	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)
	Dim iRow

	With lgvspdData(pType)
		If pType < TYPE_3 Then
			ggoSpread.Source = lgvspdData(pType)
			For iRow = pvStartRow To pvEndRow
				.Col = C_SEQ_NO
				.Row = iRow
				If UNICDbl(.Text) = 999999 Then
					ggoSpread.SpreadLock C_W_TYPE,   iRow, C_W4, iRow
				Else
					ggoSpread.SpreadUnLock C_W_TYPE, iRow, C_W4, iRow	' ��ü ���� 
					ggoSpread.SSSetRequired C_W1, iRow, iRow
					ggoSpread.SSSetRequired C_W1_NM, iRow, iRow
					ggoSpread.SSSetRequired C_W2, iRow, iRow
					ggoSpread.SSSetRequired C_W3, iRow, iRow
					ggoSpread.SSSetRequired C_W3_NM, iRow, iRow
				End If
			Next
		Else
			ggoSpread.SpreadLock C_W_TYPE, pvStartRow, C_W4, pvEndRow
		End If
			
	End With	
End Sub

Sub SetSpreadTotalLine()
	Dim iRow, ret
	For iRow = TYPE_1 To TYPE_2
		ggoSpread.Source = lgvspdData(iRow)
		With lgvspdData(iRow)
			If .MaxRows > 0 Then
				.Row = .MaxRows
'				.Col	= C_W_TYPE	:	.Text	= 1
'				.Col	= C_SEQ_NO	:	.Text	= 999999
				ret = .AddCellSpan(C_W1	, .MaxRows, 3, 1)	' ���� 2�� ��ħ 
				.Col	= C_W1	:	.CellType = 1	:	.Text	= "��"	:	.TypeHAlign = 2
				.Col	= C_W3_NM	:	.CellType = 1	:	.Text	= ""
				.Col	= C_W3	:	.CellType = 1	:	.Text	= ""
				SetSpreadColor iRow, .MaxRows, .MaxRows

			End If
		End With
	Next
	For iRow = TYPE_3 To TYPE_4
		ggoSpread.Source = lgvspdData(iRow)
		With lgvspdData(iRow)
			If .MaxRows > 0 Then
				.Row = .MaxRows
'				.Col	= C_W_TYPE	:	.Text	= 1
'				.Col	= C_SEQ_NO	:	.Text	= 999999
'				ret = .AddCellSpan(C_W1	, .MaxRows, 3, 1)	' ���� 2�� ��ħ 
				.Col	= C_W1_NM	:	.CellType = 1	:	.Text	= "��"	:	.TypeHAlign = 2
				SetSpreadColor iRow, .MaxRows, .MaxRows

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
Function GetRef()	'
	' 2. �˾� 
	Dim arrRet, sParam
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
Exit Function
	IsOpenPop = True
    
	arrParam(0) = frm1.txtCO_CD.value
	arrParam(1) = frm1.txtFISC_YEAR.text
	arrParam(2) = frm1.cboREP_TYPE.value

	arrRet = window.showModalDialog(BIZ_REF_PGM_ID, Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCO_CD.focus
	    Exit Function
	End If	
	
	lgvspdData(TYPE_1).focus
End Function



Sub GetFISC_DATE()	' ������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.

		
End Sub


' -- �Ǻ� ��ũ �����ֱ� 
Function ShowTabLInk(pType)
	Dim pObj1, pObj2, i
	Set pObj1 = document.all("myTabRef")
	Set pObj2 = document.all("myTabRef2")
	
	For i = 0 To 2
		pObj1(i).style.display = "none"
		pObj2(i).style.display = "none"
	Next
	
	pObj1(pType-1).style.display = ""
	pObj2(pType-1).style.display = ""

End Function


'====================================== �� �Լ� =========================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
	lgCurrGrid = TYPE_1	' �⺻ �׸��� 
    Call SetToolbar("1101111100000111")										<%'��ư ���� ���� %>
	Call ShowTabLInk(TAB1)
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	lgCurrGrid = TYPE_2
    Call SetToolbar("1101111100000111")										<%'��ư ���� ���� %>
	Call ShowTabLInk(TAB2)
End Function

Function ClickTab3()

	If gSelframeFlg = TAB3 Then Exit Function
	
	Call changeTabs(TAB3)
	gSelframeFlg = TAB3
	lgCurrGrid = TYPE_3
    Call SetToolbar("1101100000000111")										<%'��ư ���� ���� %>
	Call ShowTabLInk(TAB3)
End Function


'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100111100100111")										<%'��ư ���� ���� %>
	  
	' �����Ѱ� 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	
 
	Call InitComboBox	' �����ؾ� �Ѵ�. ����� ȸ��������� �о���� ���� 
	Call InitData

	Call DBQuery()
	
     
    
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

Function BtnPrint(byval strPrintType)
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE,EBR_RPT_ID,EBR_RPT_ID2
	Dim StrUrl  , i

	Dim intCnt,IntRetCD


    If Not chkField(Document, "1") Then					'��: This function check indispensable field
       Exit Function
    End If
	
	varCo_Cd	 =  Trim(frm1.txtCo_Cd.value)
	varFISC_YEAR = Trim(frm1.txtFISC_YEAR.text)
	varREP_TYPE	 =  Trim(frm1.cboREP_TYPE.value)
	
    StrUrl = StrUrl & "varCo_Cd|"			& varCo_Cd
	StrUrl = StrUrl & "|varFISC_YEAR|"		& varFISC_YEAR
	StrUrl = StrUrl & "|varREP_TYPE|"       & varREP_TYPE


	

    if frm1.prt_check1.checked = true  then
      	 EBR_RPT_ID	    = "W5103OA1"
           ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")
			if  strPrintType = "VIEW" then
			Call FncEBRPreview(ObjName, StrUrl)
			else
			Call FncEBRPrint(EBAction,ObjName,StrUrl)
			end if	
			  
    
        
        
    end if
    
    
	if frm1.prt_check2.checked = true  then
	    	 EBR_RPT_ID	    = "W5103OA2"
	         ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")
				if  strPrintType = "VIEW" then
				Call FncEBRPreview(ObjName, StrUrl)
				else
				Call FncEBRPrint(EBAction,ObjName,StrUrl)
				end if	
		  
	  end if

   
   if frm1.prt_check3.checked = true  then
      	    EBR_RPT_ID	    = "W5103OA3"
           ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")
			
			
			if  strPrintType = "VIEW" then
			Call FncEBRPreview(ObjName, StrUrl)
			else
			Call FncEBRPrint(EBAction,ObjName,StrUrl)
			end if	
	  
    end if
   
     
End Function  
     
   


'============================================  �׸��� �̺�Ʈ   ====================================
' -- 1�� �׸��� 
Sub vspdData0_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_1
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_1
	Call vspdData_Change(TYPE_1, Col, Row)
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

' -- 2�� �׸��� 
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_2
	Call vspdData_Change(TYPE_2, Col, Row)
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


' -- 3�� �׸��� 
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_3
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData2_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_3
	Call vspdData_Change(TYPE_3, Col, Row)
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
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_3
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

' -- 4�� �׸��� 
Sub vspdData3_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_4
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData3_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_4
	Call vspdData_Change(TYPE_4, Col, Row)
End Sub

Sub vspdData3_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_4
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_4
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData3_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_4
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData3_GotFocus()
	lgCurrGrid = TYPE_4
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData3_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_4
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData3_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_4
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_4
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub


'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)
	Dim iIdx, iRow, sW3, sW4, dblW2

	With lgvspdData(Index)
		Select Case Col
			Case C_W3_NM
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col +1
				.Value = iIdx
			Case C_W3
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col -1
				.Value = iIdx
		End Select

		ggoSpread.Source = lgvspdData(Index)
		ggoSpread.UpdateRow Row
		
		Call SetGrid34ByGrid12(Index, Index + 2, Row)	' ��3.�׸���3, 4�� �����Ѵ�.
		

	End With
End Sub

Function SetGrid34ByGrid12(pType, pType2, pRow)
	Dim sW1, sW1_NM, dblW2, sW3, sW3_NM, sW4
	
	With lgvspdData(pType)
			.Col = C_W1		: .Row = pRow	:	sW1 = .Text
			.Col = C_W1_NM	: .Row = pRow	:	sW1_NM = .Text
			.Col = C_W2		: .Row = pRow	:	dblW2 = .Text
			.Col = C_W3		: .Row = pRow	:	sW3 = .Text
			.Col = C_W3_NM	: .Row = pRow	:	sW3_NM = .Text
			.Col = C_W4		: .Row = pRow	:	sW4 = .Text
	End With

	With lgvspdData(pType2)
			.Col = C_W1		: .Row = pRow	:	.Text = sW1
			.Col = C_W1_NM	: .Row = pRow	:	.Text = sW1_NM
			.Col = C_W2		: .Row = pRow	:	.Text = dblW2
			.Col = C_W3		: .Row = pRow	:	.Text = sW3
			.Col = C_W3_NM	: .Row = pRow	:	.Text = sW3_NM
			.Col = C_W4		: .Row = pRow	:	.Text = sW4
	End With
End Function


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
	
	If Col = C_W3_NM Then 
		Call vspdData_ComboSelChange(Index, Col, Row)
		Exit Sub
	End If
	
    ggoSpread.Source = lgvspdData(Index)
    ggoSpread.UpdateRow Row

	' --- �߰��� �κ� 
	Call GetAdItem(Index, Col, Row)			' �������� ��������		ǥ�ؿ��� ��÷� ��Ī�� �Ȱ����´�.
	Call SetGrid34ByGrid12(Index, Index + 2, Row)	' ��3.�׸���3, 4�� �����Ѵ�.

	dblSum = FncSumSheet(lgvspdData(Index), C_W2, 1, lgvspdData(Index).MaxRows - 1, true, lgvspdData(Index).MaxRows, C_W2, "V")	' �հ� 
	dblSum = FncSumSheet(lgvspdData(Index + 2), C_W2, 1, lgvspdData(Index + 2).MaxRows - 1, true, lgvspdData(Index + 2).MaxRows, C_W2, "V")	' �հ� 
	ggoSpread.UpdateRow lgvspdData(Index).MaxRows
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
    '	Exit Sub
    '   ggoSpread.Source = lgvspdData(Index)
       
    '   If lgSortKey = 1 Then
    '       ggoSpread.SSSort Col               'Sort in ascending
    '       lgSortKey = 2
    '   Else
    '       ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
    '       lgSortKey = 1
    '   End If
       
    '   Exit Sub
    'End If

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
'    Call InitVariables													<%'Initializes local global variables%>
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
    Dim blnChange, i, sMsg
    
    blnChange = False
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData0
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If

    ggoSpread.Source = frm1.vspdData1
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If
	
    If lgBlnFlgChgValue = False And blnChange = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

    ggoSpread.Source = frm1.vspdData0
	If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
	      Exit Function
	End If    

    ggoSpread.Source = frm1.vspdData1
	If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
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

    Call SetToolbar("1100111100000111")
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

    If lgCurrGrid > TYPE_2 Then
       Exit Function
    End If
 
	With lgvspdData(lgCurrGrid)
	    ggoSpread.Source = lgvspdData(lgCurrGrid)
	    iActiveRow = .ActiveRow

		If .ActiveRow > 0 Then
			.focus
			.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor lgCurrGrid, .ActiveRow + 1, .ActiveRow + 1

			.Col = C_W2
			.Text = ""
    
			.ReDraw = True
		End If
	End With

	With lgvspdData(lgCurrGrid + 2)
	    ggoSpread.Source = lgvspdData(lgCurrGrid + 2)
	    .ActiveRow = iActiveRow

		If iActiveRow > 0 Then
			.ReDraw = False
		
			ggoSpread.CopyRow iActiveRow
			SetSpreadColor lgCurrGrid + 2, iActiveRow + 1, iActiveRow + 1

			.ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    Dim lDelRows, iActiveRow, dblSum

	Select Case lgCurrGrid 
		CAse  TYPE_1, TYPE_2
			With lgvspdData(lgCurrGrid)
				.focus
				iActiveRow = .ActiveRow
				ggoSpread.Source = lgvspdData(lgCurrGrid)
				If CheckTotalRow(lgvspdData(lgCurrGrid), .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.EditUndo

					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow(lgvspdData(lgCurrGrid), lDelRows)
					If lDelRows > 0 Then ggoSpread.EditUndo lDelRows

					ggoSpread.Source = lgvspdData(lgCurrGrid + 2)
					lDelRows = ggoSpread.EditUndo(iActiveRow)
					lDelRows = CheckLastRow(lgvspdData(lgCurrGrid + 2), lDelRows)
					If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
				End If
				
			End With
	End Select
	Call SetGrid34ByGrid12(lgCurrGrid, lgCurrGrid + 2, lDelRows)	' ��3.�׸���3, 4�� �����Ѵ�.

	dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W2, 1, lgvspdData(lgCurrGrid).MaxRows - 1, true, lgvspdData(lgCurrGrid).MaxRows, C_W2, "V")	' �հ� 
	dblSum = FncSumSheet(lgvspdData(lgCurrGrid + 2), C_W2, 1, lgvspdData(lgCurrGrid + 2).MaxRows - 1, true, lgvspdData(lgCurrGrid + 2).MaxRows, C_W2, "V")	' �հ� 

End Function

' -- �հ� ������ üũ(Header Grid)
Function CheckTotalRow(Byref pObj, Byval pRow) 
	CheckTotalRow = False
	pObj.Col = C_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If pObj.Text = "999999" And pObj.MaxRows > 1 Then	 ' �հ� �� 
		CheckTotalRow = True
	End If
End Function

' -- Detail Data�� �����ϴ��� üũ 
Function CheckLastRow(Byref pObj, Byval pRow) 
	Dim iCnt, iRow, iMaxRow, iTmpRow
	CheckLastRow = 0
	iCnt = 0
	
	With pObj

		For iRow = 1 To .MaxRows
			.Row = iRow : .Col = 0
			If .Text <> ggoSpread.DeleteFlag Then
				iCnt = iCnt + 1
				iMaxRow = iRow
			End If
			.Col = C_SEQ_NO
			If .Text = 999999 Then
				iTmpRow = iRow
			End If
		Next
		.Col = C_SEQ_NO	:	.Row = iMaxRow
		If .Text = 999999 and iCnt = 1 Then
			CheckLastRow = iMaxRow
		ElseIf iCnt = 1 Then
			CheckLastRow = iTmpRow
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

    If lgCurrGrid > TYPE_2 Then
       Exit Function
    End If
 
	With lgvspdData(lgCurrGrid)
	
		.focus
		ggoSpread.Source = lgvspdData(lgCurrGrid)
	
		iSeqNo = .MaxRows+1
	
		if .MaxRows = 0 then
		
			ggoSpread.InsertRow  imRow 
			.Col	= C_W_TYPE	:	.Text	= lgCurrGrid + 1
			.Col	= C_SEQ_NO	:	.Text	= 1
			SetSpreadColor lgCurrGrid, 1, 1
			
			ggoSpread.InsertRow  imRow 
			.Row = .MaxRows
			.Col	= C_W_TYPE	:	.Text	= lgCurrGrid + 1
			.Col	= C_SEQ_NO	:	.Text	= 999999
			ret = .AddCellSpan(C_W1	, .MaxRows, 3, 1)	' ���� 2�� ��ħ 
			.Col	= C_W1	:	.CellType = 1	:	.Text	= "��"	:	.TypeHAlign = 2
			.Col	= C_W3_NM	:	.CellType = 1	:	.Text	= ""
			.Col	= C_W3	:	.CellType = 1	:	.Text	= ""
			SetSpreadColor lgCurrGrid, .MaxRows, .MaxRows
			.Row  = 1
			.ActiveRow = 1

			' �ҵ�ݾ����� �հ�ǥ�� ���� ���� �߰��Ѵ�.
			ggoSpread.Source = lgvspdData(lgCurrGrid + 2)
			ggoSpread.InsertRow  imRow 
			lgvspdData(lgCurrGrid + 2).Col	= C_W_TYPE	:	lgvspdData(lgCurrGrid + 2).Text	= lgCurrGrid + 1
			lgvspdData(lgCurrGrid + 2).Col	= C_SEQ_NO	:	lgvspdData(lgCurrGrid + 2).Text	= 1
			SetSpreadColor lgCurrGrid + 2, 1, 1
			
			ggoSpread.InsertRow  imRow 
			lgvspdData(lgCurrGrid + 2).Row = lgvspdData(lgCurrGrid + 2).MaxRows
			lgvspdData(lgCurrGrid + 2).Col	= C_W_TYPE	:	lgvspdData(lgCurrGrid + 2).Text	= lgCurrGrid + 1
			lgvspdData(lgCurrGrid + 2).Col	= C_SEQ_NO	:	lgvspdData(lgCurrGrid + 2).Text	= 999999
			lgvspdData(lgCurrGrid + 2).Col	= C_W1_NM	:	lgvspdData(lgCurrGrid + 2).CellType = 1	:	lgvspdData(lgCurrGrid + 2).Text	= "��"	:	lgvspdData(lgCurrGrid + 2).TypeHAlign = 2
			SetSpreadColor lgCurrGrid + 2, lgvspdData(lgCurrGrid + 2).MaxRows, lgvspdData(lgCurrGrid + 2).MaxRows
		else
			iRow = .ActiveRow

			If iRow = .MaxRows Then	' -- ������ �հ��ٿ��� InsertRow�� �ϸ� ������ �߰��Ѵ�.
				iRow = iRow - 1
				ggoSpread.InsertRow iRow, imRow 
				SetSpreadColor lgCurrGrid, iRow, iRow + imRow

				Call SetDefaultVal(lgCurrGrid, iRow + 1, imRow)
			Else
				ggoSpread.InsertRow ,imRow
				SetSpreadColor lgCurrGrid, iRow+1, iRow + imRow

				Call SetDefaultVal(lgCurrGrid, iRow + 1, imRow)
			End If   
			.vspdData.Row  = iRow + 1
			.vspdData.ActiveRow = iRow +1
			
			' �ҵ�ݾ����� �հ�ǥ�� ���� ���� �߰��Ѵ�.
			ggoSpread.Source = lgvspdData(lgCurrGrid + 2)
			ggoSpread.InsertRow  iRow, imRow 
			SetSpreadColor lgCurrGrid + 2, iRow, iRow + imRow
			Call SetDefaultVal(lgCurrGrid + 2, iRow + 1, imRow)
        End if 	
		
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function


' �׸��忡 SEQ_NO, TYPE �ִ� ���� 
Function SetDefaultVal(pType, iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With lgvspdData(pType)	
	
		If pType < TYPE_3 Then
		
			If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
				.Row = iRow
				.Value = MaxSpreadVal(lgvspdData(pType), C_SEQ_NO, iRow)
				.Col = C_W_TYPE	:	.Text = pType + 1
			Else
				iSeqNo = MaxSpreadVal(lgvspdData(pType), C_SEQ_NO, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
				
				For i = iRow to iRow + iAddRows -1
					.Row = i
					.Col = C_SEQ_NO : .Value = iSeqNo : iSeqNo = iSeqNo + 1
					.Col = C_W_TYPE	:	.Text = pType + 1
				Next
			End If
		Else
			If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
				Call SetGrid34ByGrid12Key(pType - 2, pType, iRow)
			Else
				For i = iRow to iRow + iAddRows -1
					Call SetGrid34ByGrid12Key(pType - 2, pType, iRow)
				Next
			End If
		End If
	End With
End Function

Function SetGrid34ByGrid12Key(pType, pType2, pRow)
	Dim iType, iSeqNo
	
	With lgvspdData(pType)
			.Col = C_W_TYPE		: .Row = pRow	:	iType = .Text
			.Col = C_SEQ_NO		: .Row = pRow	:	iSeqNo = .Text
	End With

	With lgvspdData(pType2)
			.Col = C_W_TYPE		: .Row = pRow	:	.Text = iType
			.Col = C_SEQ_NO		: .Row = pRow	:	.Text = iSeqNo
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows, iActiveRow, dblSum

	Select Case lgCurrGrid 
		CAse  TYPE_1, TYPE_2
			With lgvspdData(lgCurrGrid)
				.focus
				iActiveRow = .ActiveRow
				ggoSpread.Source = lgvspdData(lgCurrGrid)
				If CheckTotalRow(lgvspdData(lgCurrGrid), .ActiveRow) = True Then
					MsgBox "�հ� ���� ������ �� �����ϴ�.", vbCritical
					Exit Function
				Else
					lDelRows = ggoSpread.DeleteRow

					lgBlnFlgChgValue = True
					lDelRows = CheckLastRow(lgvspdData(lgCurrGrid), lDelRows)
					If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows

					ggoSpread.Source = lgvspdData(lgCurrGrid + 2)
					lDelRows = ggoSpread.DeleteRow(iActiveRow)
					lDelRows = CheckLastRow(lgvspdData(lgCurrGrid + 2), lDelRows)
					If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
					
				End If
				
			End With
	End Select

	Call SetGrid34ByGrid12(lgCurrGrid, lgCurrGrid + 2, lDelRows)	' ��3.�׸���3, 4�� �����Ѵ�.

	dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W2, 1, lgvspdData(lgCurrGrid).MaxRows - 1, true, lgvspdData(lgCurrGrid).MaxRows, C_W2, "V")	' �հ� 
	dblSum = FncSumSheet(lgvspdData(lgCurrGrid + 2), C_W2, 1, lgvspdData(lgCurrGrid + 2).MaxRows - 1, true, lgvspdData(lgCurrGrid + 2).MaxRows, C_W2, "V")	' �հ� 

	ggoSpread.Source = lgvspdData(lgCurrGrid)
    ggoSpread.UpdateRow lgvspdData(lgCurrGrid).MaxRows
    ggoSpread.Source = lgvspdData(lgCurrGrid + 2)
    ggoSpread.UpdateRow lgvspdData(lgCurrGrid + 2).MaxRows
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
	
	If lgvspdData(TYPE_1).MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		

		Call SetToolbar("11011111000111")										<%'��ư ���� ���� %>
	End If
	
	Call SetSpreadTotalLine ' - �հ���� �籸�� 
	
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
    
    For i = TYPE_1 To TYPE_2	' ��ü �׸��� ���� 
    
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
			                                          strVal = strVal & "D"  &  Parent.gColSep
		       End Select
		       
			  ' ��� �׸��� ����Ÿ ����     
			  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
					For lCol = 1 To lMaxCols
						.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
					Next
					strVal = strVal & Trim(.Text) &  Parent.gRowSep
			  End If  
			Next
		
		End With

		document.all("txtSpread" & CStr(i)).value =  strVal
		strVal = ""
	Next

	'Frm1.txtSpread.value      = strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' ���� ������ ���� ���� %>
	Call InitVariables
	frm1.vspdData1.MaxRows = 0
    ggoSpread.Source = frm1.vspdData1
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

Function ProgramJump
    Call PgmJump(JUMP_PGM_ID)
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
					<TD CLASS="CLSMTABP" width=170>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>1.�ͱݻ��ԡ��ձݺһ���</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP" width=170>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>2.�ձݻ��ԡ��ͱݺһ���</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP" width=170>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>3.�ҵ�ݾ����� �հ�ǥ</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><DIV id="myTabRef">&nbsp;</DIV>
						<DIV id="myTabRef" STYLE="display:'none'"><A href="vbscript:GetRef"></A>&nbsp;</DIV>
						<DIV id="myTabRef" STYLE="display:'none'">&nbsp;</DIV>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w5103ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script></TD>
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
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
							     <script language =javascript src='./js/w5103ma1_vspdData0_vspdData0.js'></script>
							    </TD>
							</TR>
						</TABLE>
						</DIV>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
							     <script language =javascript src='./js/w5103ma1_vspdData1_vspdData1.js'></script>
							    </TD>
							</TR>
						</TABLE>
						</DIV>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="50%" VALIGN=TOP HEIGHT=*>
							     <script language =javascript src='./js/w5103ma1_vspdData2_vspdData2.js'></script>
							    </TD>
							     <TD WIDTH="50%" VALIGN=TOP HEIGHT=*>
							     <script language =javascript src='./js/w5103ma1_vspdData3_vspdData3.js'></script>
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
            <TABLE <%=LR_SPACE_TYPE_30%>>   
                <TR>   
                <TD WIDTH=50%>&nbsp;</TD>   
                <TD WIDTH=50%>
                    <TABLE WIDTH=100%>                           
                        <TD WIDTH=* Align=right>
							<DIV ID=myTabRef2><A href="Vbscript:ProgramJump()">��15ȣ ����������</A></DIV>
							<DIV ID=myTabRef2 STYLE="display:'none'"><A href="Vbscript:ProgramJump()">��15ȣ ����������</A></DIV>
							<DIV ID=myTabRef2 STYLE="display:'none'">&nbsp;</A></DIV>
						</TD>
                        <TD WIDTH=10>&nbsp;</TD>                           
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" checked><LABEL FOR="prt_check1"><����>15-1ȣ ���񺰼ҵ�ݾ� ��������(1)</LABEL>&nbsp;
							<INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><����>15-2ȣ ���񺰼ҵ�ݾ���������(2)</LABEL>&nbsp;
				            <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check3" ><LABEL FOR="prt_check3"><����>�ҵ�ݾ��հ�ǥ</LABEL>&nbsp;
				           
				        </TD>
		
				<TR>
			    
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('PRINT')"   Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
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
