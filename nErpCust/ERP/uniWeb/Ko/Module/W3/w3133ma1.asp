<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ���������� 
'*  3. Program ID           : W3133MA1
'*  4. Program Name         : W3133MA1.asp
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

Const BIZ_MNU_ID		= "W3133MA1"
Const BIZ_PGM_ID		= "W3133mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID	    = "W3133OA1"

Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.

' -- �׸��� �÷� ���� 
Dim C_SEQ_NO
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W8
Dim C_W_TYPE
Dim C_W_TYPE_NM

Dim C_W9
Dim C_W9_1
Dim C_W10
Dim C_W11
Dim C_W12
Dim C_W13
Dim C_W14

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(2)
Dim StrSum1,StrSum2

Dim lgW6, lgFISC_START_DT, lgFISC_END_DT	' ��������ϼ� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	
	C_SEQ_NO	= 1
	
	C_W1		= 2	' ���� 
	C_W2		= 3	' ȸ������ 
	C_W3		= 4	' �����̿��� 
	C_W4		= 5	' ������Ծ� 
	C_W5		= 6	' �� 
	C_W6		= 7	' 
	C_W7		= 8
	C_W8		= 9	' 
	C_W_TYPE	= 10
	C_W_TYPE_NM	= 11
	
	C_W9		= 2 ' 
	C_W9_1		= 3
	C_W10		= 4 ' �����ͱ� 
	C_W11		= 5 ' ȸ����ͱ� 
	C_W12		= 6 ' �������� 
	C_W13		= 7 ' �������� 
	C_W14		= 8 ' �����ݾ� 
	
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
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))

End Sub


Sub InitSpreadComboBox()
    Dim IntRetCD1

	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " MAJOR_CD='W1014' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = lgvspdData(TYPE_1)
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W_TYPE
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W_TYPE_NM
	End If

End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1
	
    Call initSpreadPosVariables()  

	'Call AppendNumberPlace("6","4","2")
	
	' 1�� �׸��� 

	With lgvspdData(TYPE_1)
				
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W_TYPE_NM + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
 
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����", 10,,,10,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W1,		"(1)����", 10,,,50,1
		ggoSpread.SSSetDate		C_W2,		"(2)����" & vbCrlf & "��ȯ(ȸ��)" & vbCrlf & "����", 10, 2, Parent.gDateFormat,	-1
		ggoSpread.SSSetFloat	C_W3,		"(3)�����̿���"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W4,		"(4)������Ծ�"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W5,		"(5)��[(3)+(4)]"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetMask		C_W6,		"(6)�����" & vbCrlf & "�ϼ�/����" & vbCrlf & "�ϼ�"	, 10, 2, "999//9999" 
		ggoSpread.SSSetFloat	C_W7,		"(7)���ͱ�" & vbCrlf & "�ش��" & vbCrlf & "[(5)*(6)]" 	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W8,		"(8)�����̿���" & vbCrlf & "[(5)-(7)]" 	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetCombo	C_W_TYPE,	"����/����"	, 8
		ggoSpread.SSSetCombo	C_W_TYPE_NM,"����/����"	, 8
						
		.rowheight(-1000) = 30	' ���� ������	(2���� ���, 1���� 15)
			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
		
		Call InitSpreadComboBox
		
		.ReDraw = true	
				
	End With 

 
	' 2�� �׸��� 

	With lgvspdData(TYPE_2)
				
		ggoSpread.Source = lgvspdData(TYPE_2)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W14 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
 
 		'����� 3�ٷ�    
		.ColHeaderRows = 2   
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����", 10,,,10,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W9,		"(9)����", 15,,,100,1
		ggoSpread.SSSetEdit		C_W9_1,		"", 5,2,,50,1
		ggoSpread.SSSetFloat	C_W10,		"(10)�����ͱ�" & vbCrlf & "�ش��", 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W11,		"(11)ȸ����ͱ�" & vbCrlf & "����", 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W12,		"(12)��������" & vbCrlf & "[(10)-(11)]"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W13,		"(13)��������" & vbCrlf & "[(11)-(10)]"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W14,		"�����ݾ�" & vbCrlf & "[(11)-(10)]" & vbCrlf , 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	 

		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(C_SEQ_NO , -1000, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W9		, -1000, 2, 2)	
		ret = .AddCellSpan(C_W10	, -1000, 1, 2)
		ret = .AddCellSpan(C_W11	, -1000, 1, 2)
		ret = .AddCellSpan(C_W12	, -1000, 2, 1)
		ret = .AddCellSpan(C_W14	, -1000, 1, 2)
    
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W12	: .Text = "��  ��"
		
		.Row = -999
		.Col = C_W12	: .Text = "(12)��������" & vbCrlf & "[(10)-(11)]"
		.Col = C_W13	: .Text = "(13)��������" & vbCrlf & "[(11)-(10)]"
						
		.rowheight(-999) = 25	' ���� ������	(2���� ���, 1���� 15)

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

	Call GetFISC_DATE
	
	Call MakeGrid2
End Sub

Sub MakeGrid2()
	' 2�� �׸��� �׸� 
	Dim ret, iRow
	
	With lgvspdData(TYPE_2)
		ggoSpread.Source = lgvspdData(TYPE_2)
		ggoSpread.InsertRow , 5
		
		For iRow = 1 To 5
			.Row = iRow : .Col = C_SEQ_NO	: .Value = iRow
		Next
		
		Call SetSpreadLock(TYPE_2)
		
		Call ReDrawGrid2()
	End With
End Sub

' �׸���2 �ؽ�Ʈ �籸��: Query��, New/Form_load�� 
Sub ReDrawGrid2()
	Dim ret
	
	With lgvspdData(TYPE_2)
		.Redraw = false

		.Row = 1
		.Col = C_W9		: .value ="��.ȯ�� ���� ������ ���� [����(7)]"
		ret = .AddCellSpan(C_W9 , 1, 1, 2)	
		ret = .AddCellSpan(C_W14, 1, 1, 2)
		.TypeEditMultiLine = True
		'.TypeHAlign = 2 : .TypeVAlign = 2
		.Col = C_W9_1	: .value = "����"
		'.rowheight(1) = 15
		
		.Row = 2
		.Col = C_W9_1	: .value = "����"
		'.rowheight(2) = 15
		
		.Row = 3
		.Col = C_W9		: .value = "��.��ȭ�ڻ� ��ä�򰡼���" & vbCrLf & "[������40ȣ����(��)��(10)]"
		.TypeEditMultiLine = True
		ret = .AddCellSpan(C_W9 , 3, 2, 1)	
		.Col = C_W10	: .TypeVAlign = 2
		.Col = C_W11	: .TypeVAlign = 2
		.Col = C_W14	: .TypeVAlign = 2
		
		.rowheight(3) = 20

		.Row = 4
		.Col = C_W9		: .value = "��"	
		ret = .AddCellSpan(C_W9 , 4, 2, 2)
		.TypeHAlign = 2	: .TypeVAlign = 2
		ret = .AddCellSpan(C_W10 , 4, 1, 2)
		ret = .AddCellSpan(C_W11 , 4, 1, 2)
		ret = .AddCellSpan(C_W14 , 4, 1, 2)
		'.rowheight(4) = 15
		.Col = C_W10	: .TypeVAlign = 2
		.Col = C_W11	: .TypeVAlign = 2
						
		.Redraw = True

		.SetActiveCell C_W10, 1
	End With
End Sub

Sub SetSpreadLock(pType)

	With lgvspdData(pType)
	
		ggoSpread.Source = lgvspdData(pType)	

		If pType = TYPE_1 Then
			ggoSpread.SSSetRequired	 C_W2, 1, .MaxRows-1
			ggoSpread.SSSetRequired	 C_W3, 1, .MaxRows-1
			ggoSpread.SSSetRequired	 C_W_TYPE_NM, 1, .MaxRows-1
			ggoSpread.SpreadLock C_W5, -1, C_W5
			ggoSpread.SpreadLock C_W8, -1, C_W8
			ggoSpread.SpreadLock C_W1, .MaxRows-1, C_W_TYPE_NM, .MaxRows-1
			ggoSpread.SpreadLock C_W1, .MaxRows  , C_W_TYPE_NM, .MaxRows
		Else
			ggoSpread.SpreadLock C_W9 , -1, C_W9_1
			ggoSpread.SpreadLock C_W12, -1, C_W14
			ggoSpread.SpreadLock C_W9, .MaxRows -1  , C_W14, .MaxRows
		End If
		
	End With	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	With lgvspdData(pType)
		ggoSpread.Source = lgvspdData(pType)	

		If pType = TYPE_1 Then
			ggoSpread.SSSetRequired	 C_W2, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired	 C_W3, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired	 C_W_TYPE_NM, pvStartRow, pvEndRow

			ggoSpread.SSSetProtected C_W5, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_W8, pvStartRow, pvEndRow
		End If
			
	End With	
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

Sub InsertFirstRow()
	Dim iMaxRows, iRow, iType, ret

	iMaxRows = 3 ' �ϵ��ڵ��Ǵ� ��� 

	With lgvspdData(TYPE_1)
		ggoSpread.Source = lgvspdData(TYPE_1)
		.Redraw = False

		ggoSpread.InsertRow , iMaxRows
		Call SetSpreadLock(TYPE_1)
		
		iRow = 1
		
		.Row = iRow		
		.Col = C_SEQ_NO : .Value = iRow		: iRow = iRow + 1
		
		.Row = iRow		
		.Col = C_SEQ_NO : .Value = SUM_SEQ_NO	: iRow = iRow + 1

		.Row = iRow		
		.Col = C_SEQ_NO : .Value = SUM_SEQ_NO+1	: iRow = iRow + 1

		Call ReDrawGrid1

		.Redraw = True

	End With
	'Call SetSpreadLock(iType)
End Sub

Sub ReDrawGrid1()
	Dim iMaxRows, iRow, iType, ret

	With lgvspdData(TYPE_1)
		If .MaxRows > 0 Then
			.Row = .MaxRows -1		
			.Col = C_W1		: .CellType = 1			: .Text = "��"		: .TypeHAlign = 2	: .TypeVAlign = 2	
			.Col = C_W2		: .CellType = 1			: .Text = "����"	: .TypeHAlign = 2
			.Col = C_W_TYPE_NM	: .CellType = 1
			ret = .AddCellSpan(C_W1	, .Row, 1, 2)
			ret = .AddCellSpan(C_W6	, .Row, 1, 2)
		
			.Row = .MaxRows		
			.Col = C_W2		: .CellType = 1			: .Text = "����"	: .TypeHAlign = 2
			.Col = C_W_TYPE_NM	: .CellType = 1
		End If
	End With
	'Call SetSpreadLock(iType)
End Sub

'============================== ���۷��� �Լ�  ========================================

Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrRet = window.showModalDialog("../W5/W5105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

Sub GetFISC_DATE()	' ������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd, iGap
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	lgFISC_START_DT = CDate(lgF0)
	lgFISC_END_DT = CDate(lgF1)
		
End Sub

'====================================== �� �Լ� =========================================

'============================================  ��ȸ���� �Լ�  ====================================
Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrW1, arrW1_1, arrW2, arrW3, arrW4, iMaxRows, iRow
	Dim StrSum1,StrSum2
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	StrSum1 = 0
	StrSum2 = 0
	Dim sMesg

	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
			
	IntRetCD = CommonQueryRs("W1, W1_1, W2, W3, W_TYPE"," dbo.ufn_TB_40A_GetRef('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD Then
		ggoSpread.Source = lgvspdData(TYPE_1)
		ggoSpread.ClearSpreadData
		
		ggoSpread.Source = lgvspdData(TYPE_2)
		ggoSpread.ClearSpreadData
		
		Dim strVal

		If lgIntFlgMode = parent.OPMD_UMODE then 
		
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0004							'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
		strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key            
	
		Call RunMyBizASP(MyBizASP, strVal)
		
		End if
		
		Call MakeGrid2
		
		' 1�� �׸��� 
		With lgvspdData(TYPE_1)
			arrW1 = Split(lgF0, chr(11))
			arrW1_1 = Split(lgF1, chr(11))
			arrW2 = Split(lgF2, chr(11))
			arrW3 = Split(lgF3, chr(11))
			arrW4 = Split(lgF4, chr(11))
			iMaxRows = UBound(arrW1)
			
			.Redraw = False

			For iRow = 0 To iMaxRows-1

				If arrW1(iRow) = "1" Then 
					Call FncInsertRow(1)
					.Row = iRow+1
					.Col = C_W1		: .text = arrW1_1(iRow)
					.Col = C_W2		: .text = arrW2(iRow)
					.Col = C_W3		: .value = arrW3(iRow)
					.Col = C_W_TYPE	: .text = arrW4(iRow)
					
					Call vspdData_Change(TYPE_1, C_W2, iRow+1)
					Call vspdData_Change(TYPE_1, C_W3, iRow+1)
					Call vspdData_ComboSelChange(TYPE_1, C_W_TYPE, iRow+1)
				End If
			Next
		End With
		
	
		Call CommonQueryRs("W5"," TB_3_3 a WITH (NOLOCK) ", " CO_CD = '" & sCoCd & "' And FISC_YEAR = '" & sFiscYear & "' and REP_TYPE = '" & sRepType & "' and w4 = '44' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		StrSum1 = Replace(lgF0,Chr(10),"")
		Call CommonQueryRs("W5"," TB_3_3 a WITH (NOLOCK) " ," CO_CD = '" & sCoCd & "' And FISC_YEAR = '" & sFiscYear & "' and REP_TYPE = '" & sRepType & "' and w4 = '60' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		StrSum2 = Replace(lgF0,Chr(10),"")
		' �׸���2�� ��.��ȭ�ڻ��ä�򰡼����� "40ȣ(��)������ �� �򰡼��� "��"�� �ݾ��� �Է���.
		With lgvspdData(TYPE_2)
			.Row = 1
			.Col = C_W10
			.value = StrSum1
			.Row = 2
			.Col = C_W10
			.value = StrSum2
			.Row = 3
			.Col = C_W11
			.value = UNICDBL(StrSum1) - UNICDBL(StrSum2)
			.Col = C_W10
			For iRow = 0 To iMaxRows-1
				If arrW1(iRow) = "2" Then 
					.value = arrW3(iRow)
					Exit For
				End If
			Next
			
			Call SetGridSum2
		End With
	End If
	lgBlnFlgChgValue = True
End Function

'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1110110100101111")										<%'��ư ���� ���� %>
	  
	' �����Ѱ� 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)
	
 
	Call InitComboBox	' �����ؾ� �Ѵ�. ����� ȸ��������� �о���� ���� 
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
	Call GetFISC_DATE
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

'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)
	Dim iIdx

	With lgvspdData(TYPE_1)
		Select Case Col
			Case C_W_TYPE_NM
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col -1
				.Value = iIdx
				
				Call SetGridSum
			Case C_W_TYPE
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col +1
				.Value = iIdx
				
				Call SetGridSum
		End Select
	End With
End Sub

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, iGap1, iGap2, datW2, dblW3, dblW4, dblW5, dblW6, dblW7, dblW8, sW_TYPE, dblW10, dblW11, dblW12, dblW13, dblW14
	
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
			Case C_W2	' �ϼ��Է� 
				.Col = C_W2	: .Row = Row
				datW2 = CDate(.Text)
				If DateDiff("d", lgFISC_START_DT, datW2) <= 0 Then
					Call DisplayMsgBox("W20003", parent.VB_INFORMATION, lgFISC_START_DT, "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.Text = lgFISC_START_DT+1
				End If
				
				iGap2 = Right("    " & DateDiff("d", lgFISC_START_DT, datW2)+1, 4)
				
				If frm1.cboREP_TYPE.value = "2" Then	' -- �߰������ϰ�� 
					If DateDiff("d" , datW2, DateAdd("m", 6, lgFISC_START_DT)-1 ) > 0 Then ' �������� ��������Ϻ��� �����̸� 
						iGap1 = Right("   " & DateDiff("d", lgFISC_START_DT, datW2)+1, 3)
					Else	' �������� ����������� �ѱ�(������)�̸� 
						iGap1 = Right("   " & DateDiff("d", lgFISC_START_DT, DateAdd("m", 6, lgFISC_START_DT)-1)+1, 3)
					End If
				Else
					If DateDiff("d" , datW2, lgFISC_END_DT ) > 0 Then ' �������� ��������Ϻ��� �����̸� 
						iGap1 = Right("   " & DateDiff("d", lgFISC_START_DT, datW2)+1, 3)
					Else	' �������� ����������� �ѱ�(������)�̸� 
						iGap1 = Right("   " & DateDiff("d", lgFISC_START_DT, lgFISC_END_DT)+1, 3)
					End If
				End If
				
				' W6 ����ũ ���� 
				.Col = C_W6	: .Value = iGap1 & iGap2 
				
				' W5 * W6 = W7
				Call SetW7_8(Row, Col)	' W8 �����̿���(5-7)
				
			Case C_W3, C_W4
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.Value)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.Value = 0
				End If
				.Col = C_W3	: dblW3 = UNICDbl(.Value)
				.Col = C_W4	: dblW4 = UNICDbl(.Value)

				.Col = C_W5	: .Value = dblW3 + dblW4
				.Col = C_W_TYPE	: sW_TYPE = .Text
				
				Call SetW7_8(Row, Col)	' W8 �����̿���(5-7)
				
				If sW_TYPE = "" Then Exit Sub
				
				Call SetSum2Col(C_W3, sW_TYPE)
				Call SetSum2Col(C_W4, sW_TYPE)
				Call SetSum2Col(C_W5, sW_TYPE)
			
			Case C_W7, C_W6
				Call SetW7_8(Row, Col)	' W8 �����̿���(5-7)
		End Select
		
	ElseIf Index = TYPE_2 Then
		Select Case Col
			Case C_W10, C_W11
				.Row = Row
				.Col = C_W10	: dblW10 = UNICDbl(.value)
				.Col = C_W11	: dblW11 = UNICDbl(.value)
				
				dblW12 = dblW10 - dblW11
				dblW13 = dblW11 - dblW10
				
				If Row = 1 Then ' ��.����/���� 
					.Col = C_W12	: .value = dblW12
				ElseIf Row = 1 Then
					.Col = C_W13	: .value = dblW13
				Else	'��.��ȭ�ڻ� 
					.Col = C_W14	: .value = dblW13
				End If				
			
				Call SetGridSum2	 ' 2�� �׸��� ���� 
		End Select
	End If
	
	End With
	
End Sub

' 2�� �׸��� ���� 
Function SetGridSum2()
	Dim dblW10, dblW11, dblW12, dblW13, dblW14, iRow
	Dim dblW10Sum, dblW11Sum, dblW12Sum, dblW13Sum
	
	With lgvspdData(TYPE_2)
	    ggoSpread.Source = lgvspdData(TYPE_2)
		
		.Row = 1
		.Col = C_W10	: dblW10 = UNICDbl(.value)	: dblW10Sum = dblW10
		.Col = C_W11	: dblW11 = UNICDbl(.value)	: dblW11Sum = dblW11
		
		dblW12 = dblW10 - dblW11
		.Col = C_W12	: .Value = dblW12

		.Row = 2
		.Col = C_W10	: dblW10 = UNICDbl(.value)	: dblW10Sum = dblW10Sum + dblW10
		.Col = C_W11	: dblW11 = UNICDbl(.value)	: dblW11Sum = dblW11Sum + dblW11
		
		dblW13 = dblW11 - dblW10
		.Col = C_W13	: .Value = dblW13

		.Row = 3
		.Col = C_W10	: dblW10 = UNICDbl(.value)	: dblW10Sum = dblW10Sum + dblW10
		.Col = C_W11	: dblW11 = UNICDbl(.value)	: dblW11Sum = dblW11Sum + dblW11

		dblW14 = dblW11 - dblW10
		
		.Col = C_W14	: .Value = dblW14
		
		.Row = 4
		.Col = C_W10	: .value = dblW10Sum
		.Col = C_W11	: .value = dblW11Sum
		.Col = C_W12	: .value = dblW12
		ggoSpread.UpdateRow .Row
		
		.Row = 5
		.Col = C_W13	: .value = dblW13
		ggoSpread.UpdateRow .Row
		
	End With
End Function

' �����̿��� ��� 
Function SetW7_8(Byval pRow, Byval pCol)
	Dim dblW5, dblW6, dblW7, dblW8
	
	With lgvspdData(TYPE_1)
		.Row = pRow
		.Col = C_W5	: dblW5 = UNICDbl(.value)
		
		.Col = C_W6	
		If .Text = "" Then
			dblW6 = 0
		Else 
			dblW6 = UNICDbl(Eval(.Text))
		End If
		
		If pCol <> C_W7 Then
			dblW7 = Fix(dblW5 * dblW6)
		Else
			.Col = C_W7	: dblW7 = UNICDbl(.value)
		End If
		dblW8 = dblW5 - dblW7
		
		.Col = C_W7	: .value = dblW7
		.Col = C_W8	: .value = dblW8
		
		Call SetGridSum
	End With
End Function

' ����/���� ����� ��ü �׸��� �� ���� 
Function SetGridSum()
	Dim dblW3(2), dblW4(2), dblW5(2), dblW7(2), dblW8(2), sW_TYPE, dblSum, iRow, iMaxRows

	With lgvspdData(TYPE_1)
		iMaxRows = .MaxRows
		ggoSpread.Source = lgvspdData(TYPE_1)
		' ���� ���� ���� �� ��� 
		For iRow = 1 To iMaxRows-2
			
			.Row = iRow
			.Col = C_W_TYPE : sW_TYPE = UNICDbl(.Text)
			If sW_TYPE <> 0 Then
				.Col = C_W3	: dblW3(sW_TYPE) = dblW3(sW_TYPE) + UNICDbl(.value)
				.Col = C_W4	: dblW4(sW_TYPE) = dblW4(sW_TYPE) + UNICDbl(.value)
				.Col = C_W5	: dblW5(sW_TYPE) = dblW5(sW_TYPE) + UNICDbl(.value)
				.Col = C_W7	: dblW7(sW_TYPE) = dblW7(sW_TYPE) + UNICDbl(.value)
				.Col = C_W8	: dblW8(sW_TYPE) = dblW8(sW_TYPE) + UNICDbl(.value)
			End If

		Next
		
		.Row = .MaxRows -1 ' W_TYPE = 1(����)
		.Col = C_W3		: .Value = dblW3(1)
		.Col = C_W4		: .Value = dblW4(1)
		.Col = C_W5		: .Value = dblW5(1)
		.Col = C_W7		: .Value = dblW7(1)
		.Col = C_W8		: .Value = dblW8(1)
		 StrSum1 = dblW8(1)
		ggoSpread.UpdateRow .Row
		
		.Row = .MaxRows  ' W_TYPE = 1(����)
		.Col = C_W3		: .Value = dblW3(2)
		.Col = C_W4		: .Value = dblW4(2)
		.Col = C_W5		: .Value = dblW5(2)
		.Col = C_W7		: .Value = dblW7(2)
		.Col = C_W8		: .Value = dblW8(2)
		StrSum2 = dblW8(2)
		ggoSpread.UpdateRow .Row
	End With
	
	' 2�� �׸��忡 �ݿ� 
	Call SetW10
End Function

' ���� �÷��� �������� ���� �ڵ带 ã�� �հ踦 ����Ѵ�.
Function SetSum2Col(Byval pCol, Byval pW_TYPE)
	Dim dblSum1, dblSum2, dblSumCol, iRow, iMaxRows, iDx
	iDx = 0
	
	With lgvspdData(TYPE_1)	' ��Ŀ���� �׸��� 
		iMaxRows = .MaxRows
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		For iRow = 1 To iMaxRows -2
			.Row = iRow	: .Col = C_W_TYPE
			
			If .Text = pW_TYPE Then	' ���� �����̸� ������ �÷��� ���� ���Ѵ�.
				.Col = pCol
				dblSumCol = dblSumCol + UNICDbl(.Value)
			End If
		Next
		
		If pW_TYPE = "1" Then	' Minor_CD='1' ���� 
			.Row = .MaxRows -1
		Else
			.Row = .MaxRows 
		End If
		
		' �հ� ����� �÷��׺��� 
		.Col = pCol
		.Value = dblSumCol
		ggoSpread.UpdateRow .Row
		
		' 2�� �׸��忡 �ݿ� 
		Call SetW10
	End With
End Function

Function SetW10()
	Dim dblW8(2)
	
	With lgvspdData(TYPE_1)
		.Row = .MaxRows -1	: .Col = C_W7	: dblW8(1) = UNICDbl(.value)
		.Row = .MaxRows		: .Col = C_W7	: dblW8(2) = UNICDbl(.value)
	End With
	
	' 2�� �׸��� W10�� �� �ݿ� 
	With lgvspdData(TYPE_2)
		ggoSpread.Source = lgvspdData(TYPE_2)
		.Row = 1		: .Col = C_W10		: .Value = dblW8(1)	' ���Ͱ� 
		ggoSpread.UpdateRow .Row
		.Row = 2		: .Col = C_W10		: .Value = dblW8(2)	' ���հ� 
		ggoSpread.UpdateRow .Row
	End With
	
	Call SetGridSum2

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
    'Call InitData             
    Call MakeGrid2                 
    															
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
    
    For i = TYPE_1 To TYPE_1
		With lgvspdData(i)
			If .MaxRows > 0 Then
				ggoSpread.Source = lgvspdData(i)
				If ggoSpread.SSCheckChange = True Then
					blnChange = True
				End If
			End If
		End With
	Next

    If lgBlnFlgChgValue = False And  blnChange = False Then
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

    Call SetToolbar("1110110100001111")

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

	If lgCurrGrid = TYPE_2 Then Exit Function
    ggoSpread.Source = lgvspdData(lgCurrGrid)	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
    
    Call SetGridSum()				' �Ѷ����� ��ҵǸ� ���� 
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
 
	With lgvspdData(TYPE_1)	' ��Ŀ���� �׸��� 
			
		ggoSpread.Source = lgvspdData(TYPE_1)
			
		iRow = .ActiveRow
		lgvspdData(TYPE_1).ReDraw = False
		
		If .MaxRows = 0 Then	' ù InsertRow�� 1��+�հ��� 

			Call InsertFirstRow()
		
		Else
			
			If iRow = .MaxRows Then	' �հ� �� 
				.Row = iRow - 2
				ggoSpread.InsertRow iRow-2 , imRow 
				SetSpreadColor lgCurrGrid,iRow-1, iRow + imRow - 1	
				
				Call SetDefaultVal(iRow-1, imRow)			
			ElseIf iRow = .MaxRows -1 Then	' �հ� �� 
				.Row = iRow - 1
				ggoSpread.InsertRow iRow-1 , imRow 
				SetSpreadColor lgCurrGrid,iRow, iRow + imRow - 1	
				
				Call SetDefaultVal(iRow, imRow)				
			Else
				.Row = iRow		
				ggoSpread.InsertRow ,imRow
				SetSpreadColor lgCurrGrid,iRow+1, iRow + imRow
				
				Call SetDefaultVal(iRow+1, imRow)
			End If
			
		End If
		
		lgvspdData(TYPE_1).ReDraw = True
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
Function SetDefaultVal(iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With lgvspdData(TYPE_1)	' ��Ŀ���� �׸��� 

		ggoSpread.Source = lgvspdData(TYPE_1)
	
		If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
			.Row = iRow
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
	
	'Call CheckReCalc()				' �Ѷ����� ��ҵǸ� ���� 
	
	'Call CheckW7Status(lgCurrGrid)	' ������ ���� üũ 
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
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE
		
	Call SetSpreadLock(TYPE_1)
	Call ReDrawGrid1()
	Call ReDrawGrid2()
		
	With lgvspdData(TYPE_2)
		iMaxRows = .MaxRows
			
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = 0 : .Value = iRow
		Next
	End With
	' �������� ���� : ���ߵǸ� ���ȴ�.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 ����üũ : �׸��� �� 
	If wgConfirmFlg = "N" Then
		'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1

		'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
		Call SetToolbar("1111111100011111")										<%'��ư ���� ���� %>
			
	Else
		
		'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
		Call SetToolbar("1110000000011111")										<%'��ư ���� ���� %>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK" width=90%><font color=white>��40ȣ(��) ��ȭ�ڻ� ��������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><a href="vbscript:GetRef">�ݾ� �ҷ�����</A> | <A href="vbscript:OpenRefMenu">�ҵ�ݾ��հ�ǥ��ȸ</A></TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w3133ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=* VALIGN=TOP>
							     <script language =javascript src='./js/w3133ma1_vspdData0_vspdData0.js'></script>
							    </TD>
							</TR>
							 <TR>
							     <TD width="100%" HEIGHT=165>
							     <script language =javascript src='./js/w3133ma1_vspdData1_vspdData1.js'></script>
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

