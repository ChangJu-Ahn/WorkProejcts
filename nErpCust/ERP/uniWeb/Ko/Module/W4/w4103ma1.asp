
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �غ���� 
'*  3. Program ID           : W4103MA1
'*  4. Program Name         : W4103MA1.asp
'*  5. Program Desc         : ��31ȣ(2) ���� �� �η°����غ�� �������� 
'*  6. Modified date(First) : 2005/01/17
'*  7. Modified date(Last)  : 2005/01/19
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

Const BIZ_MNU_ID		= "W4103MA1"
Const BIZ_PGM_ID		= "W4103mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "W4103mb2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID	    = "W4103OA1"
Const TAB1 = 1																	'��: Tab�� ��ġ 
Const TAB2 = 2

Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.

' -- �׸��� �÷� ���� 
Dim C_SEQ_NO	

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

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(2)

Dim lgW2, lgMonth	' ������, ����������� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	
	C_SEQ_NO	= 1	' -- 1�� �׸��� 
	C_W9		= 2	' �ձݻ��Կ��� 
	C_W10		= 3	' ������ 
	C_W11		= 4	' ��λ��غ�� 
	C_W12		= 5 ' �����غ�� 
	C_W13		= 6	' �غ�� 
	C_W14		= 7	' ��ü�ҿ��ڱݻ��� 
	C_W15		= 8	' �̻��� 
	C_W16		= 9	' ��ü�ҿ��ڱݻ��� 
	C_W17		= 10 ' ��Ÿ 
	C_W18		= 11 ' �� 
	
	' C_SEQ_NO, C_W9 ���� 
	C_W19		= 3	' 1������ 
	C_W20		= 4	' 2������ 
	C_W21		= 5	' 3���⵵ 
	C_W22		= 6 ' �� 
	C_W23		= 7	' ȯ���ұݾ��հ� 
	C_W24		= 8	' ȸ��ȯ�Ծ� 
	C_W25		= 9	' ����ȯ�� 

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
	call CommonQueryRs("MINOR_CD, REFERENCE_2, REFERENCE_1"," dbo.ufn_B_Configuration('W2008') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboW2 ,lgF0  ,lgF1  ,Chr(11))
	
    lgW2 = Split(lgF2, Chr(11))	' ��������� 
	With frm1
		.txtW2_VAL.value = lgW2(UNICDbl(.cboW2.value) - 1)
	End With
End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1
	
    Call initSpreadPosVariables()  

	' 1�� �׸��� 
	With lgvspdData(TYPE_1)
			
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W18 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

		'����� 3�ٷ�    
		.ColHeaderRows = 3    
						       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'Call AppendNumberPlace("6","4","0")

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����"		, 10,,,6,1	' �����÷� 
		ggoSpread.SSSetMask     C_W9,	    "(9)" & vbCrLf & "�ձ�" & vbCrLf & "����" & vbCrLf & "����", 5, 2, "9999" 
		ggoSpread.SSSetFloat	C_W10,		"(10)������"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W11,		"(11)��λ�" & vbCrLf & "�غ��" & vbCrLf & "�����ܾ�"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W12,		"(12)����" & vbCrLf & "�غ��" & vbCrLf & "ȯ�Ծ�"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W13,		"(13)�غ��" & vbCrLf & "����" & vbCrLf & "�����"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W14,		"(14)������" & vbCrLf & "�η°��߼ҿ�" & vbCrLf & "�ڱݻ���"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W15,		"(15)�̻���"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W16,		"(16)������" & vbCrLf & "�η°��߼ҿ�" & vbCrLf & "�ڱݻ���"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W17,		"(17)��Ÿ"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W18,		"(18)��" & vbCrLf & "[(11)-(12)-(13)]"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(C_SEQ_NO , -1000, 1, 3)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W9		, -1000, 1, 3)	
		ret = .AddCellSpan(C_W10	, -1000, 1, 3)
		ret = .AddCellSpan(C_W11	, -1000, 1, 3)
		ret = .AddCellSpan(C_W12	, -1000, 1, 3)
		ret = .AddCellSpan(C_W13	, -1000, 1, 3)
		ret = .AddCellSpan(C_W14	, -1000, 4, 1)
		ret = .AddCellSpan(C_W14	, -999 , 2, 1)
		ret = .AddCellSpan(C_W16	, -999 , 2, 1)
		ret = .AddCellSpan(C_W18	, -1000, 1, 3) 
    
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W14	: .Text = "��  ��  ��"
		
		.Row = -999
		.Col = C_W14	: .Text = "3�� �̰����"
		.Col = C_W16	: .Text = "3�� �����"
		
		.Row = -998
		.Col = C_W14	: .Text = "(14)������" & vbCrLf & "�η°��߼ҿ�" & vbCrLf & "�ڱݻ���"
		.Col = C_W15	: .Text = "(15)�̻���"
		.Col = C_W16	: .Text = "(16)������" & vbCrLf & "�η°��߼ҿ�" & vbCrLf & "�ڱݻ���"
		.Col = C_W17	: .Text = "(17)��Ÿ"
			
		.rowheight(-998) = 28	' ���� ������	(2���� ���, 1���� 15)
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
		Call SetSpreadLock(TYPE_1)
				
		.ReDraw = true	
			
	End With 
 
	' 2�� �׸��� 
	With lgvspdData(TYPE_2)
			
		ggoSpread.Source = lgvspdData(TYPE_2)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W25 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

		'����� 2�ٷ�    
		.ColHeaderRows = 2    
						       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'Call AppendNumberPlace("6","3","2")

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����"		, 10,,,6,1	' �����÷� 
		ggoSpread.SSSetMask     C_W9,	    "(9)" & vbCrLf & "�ձ�" & vbCrLf & "����" & vbCrLf & "����", 5, 2, "9999" 
		ggoSpread.SSSetFloat	C_W19,		"(19)1������"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W20,		"(20)2������" & vbCrLf, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W21,		"(21)3������" & vbCrLf, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W22,		"(22)��" & vbCrLf & "[(19)+(20)+(21)]"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W23,		"(23)ȯ����" & vbCrLf & "�ݾ��հ�" & vbCrLf & "[(17)+(22)]"	, 13, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W24,		"(24)ȸ��ȯ�Ծ�"	, 13, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W25,		"(25)����ȯ��" & vbCrLf & "����ȯ��" & vbCrLf & "[(23)-(24)]"	, 20, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(C_SEQ_NO , -1000, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W9		, -1000, 1, 2)	
		ret = .AddCellSpan(C_W19	, -1000, 4, 1)
		ret = .AddCellSpan(C_W23	, -1000, 1, 2)
		ret = .AddCellSpan(C_W24	, -1000, 1, 2)
		ret = .AddCellSpan(C_W25	, -1000, 1, 2)
		
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W19	: .Text = "�������η°��߼ҿ��ڱ� ����(24)�� ȯ���� �ݾ�"
		
		.Row = -999
		.Col = C_W19	: .Text = "(19)1������"
		.Col = C_W20	: .Text = "(20)2������"
		.Col = C_W21	: .Text = "(21)3������"
		.Col = C_W22	: .Text = "(22)��" & vbCrLf & "[(19)+(20)+(21)]"

		.rowheight(-999) = 20	' ���� ������	(2���� ���, 1���� 15)
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
		Call SetSpreadLock(TYPE_2)
				
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

End Sub

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock(Byval pType)

	ggoSpread.Source = lgvspdData(pType)	
	
	Select Case pType
		Case TYPE_1
			ggoSpread.SSSetRequired C_W9, -1, C_W9
			ggoSpread.SSSetRequired C_W10, -1, C_W10
			ggoSpread.SSSetRequired C_W11, -1, C_W11
			ggoSpread.SpreadLock C_W17, -1, C_W17
			ggoSpread.SpreadLock C_W18, -1, C_W18
		Case TYPE_2
			ggoSpread.SSSetRequired C_W9, -1, C_W9
			ggoSpread.SpreadLock C_W22, -1, C_W22
			ggoSpread.SpreadLock C_W23, -1, C_W23
			ggoSpread.SpreadLock C_W25, -1, C_W25
	End Select
	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	ggoSpread.Source = lgvspdData(pType)

	Select Case pType
		Case TYPE_1
			ggoSpread.SSSetRequired C_W9, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W10, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired C_W11, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_W17, pvStartRow, pvEndRow 	
			ggoSpread.SSSetProtected C_W18, pvStartRow, pvEndRow 	
		Case TYPE_2
			ggoSpread.SSSetRequired C_W9, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_W22, pvStartRow, pvEndRow 	
			ggoSpread.SSSetProtected C_W23, pvStartRow, pvEndRow 	
			ggoSpread.SSSetProtected C_W25, pvStartRow, pvEndRow 	
	End Select

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
			
	call CommonQueryRs("W1, W6"," dbo.ufn_TB_31_2_GetRef('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)


	' 1�� �׸��� 
	With frm1
			
		.txtW1.Value = UNIFormatNumber( UNICDbl(Replace(lgF0, chr(11), "")) , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
		.txtW6.Value = UNIFormatNumber( UNICDbl(Replace(lgF1, chr(11), "")) , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	
	End With	
	
	Call SetHeadReCalc
End Function

Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("W5105RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W5105RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
   

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

' ��� ���� 
Sub SetHeadReCalc()	
	Dim dblSum, dblW1, dblW2, dblW2_VAL, dblW3, dblW4, dblW5, dblW6, dblW7, dblW8
	
	With frm1
		dblW1 = UNICDbl(.txtW1.value)
		dblW2_VAL = UNICDbl(.txtW2_VAL.value)
		dblW3 = dblW1 * dblW2_VAL
		
		dblW4 = UNICDbl(.txtW4.value)
		dblW6 = UNICDbl(.txtW6.value)

		If (dblW4 - dblW3) > 0 Then
			dblW5 = dblW4 - dblW3
		Else	
			dblW5 = 0
		End If
		dblW7 = dblW5 + dblW6
		dblW8 = dblW4 - dblW5 - dblW6
		
		.txtW3.value = UNIFormatNumber(dblW3, ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
		.txtW5.value = UNIFormatNumber(dblW5, ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit) 
		.txtW7.value = UNIFormatNumber(dblW7, ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit) 
		.txtW8.value = UNIFormatNumber(dblW8, ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit) 
	End With

	lgBlnFlgChgValue= True ' ���濩�� 
End Sub

Sub GetFISC_DATE()	' ������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd, iGap
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	'W2�� ��� 
	iGap = DateDiff("m", CDate(lgF0), CDate(lgF1))+1
	
	ReDim lgMonth(1)
	If sRepType = "2" Then
		lgMonth(1) = "6/" & iGap	' ȭ��ǥ�ð� 
		lgMonth(0) = 6/iGap		' ��갪 
	Else
		lgMonth(1) = "12/" & iGap 	' ȭ��ǥ�ð� 
		lgMonth(0) = 12/iGap		' ��갪 
	End If
	
End Sub

'====================================== �� �Լ� =========================================

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1110110100101111")										<%'��ư ���� ���� %>
	  
	' �����Ѱ� 
	Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)
	
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

' -- ������ ����� 
Sub cboW2_onChange()
	With frm1
		.txtW2_VAL.value = lgW2(UNICDbl(.cboW2.value) - 1)
	End With
	Call SetHeadReCalc
End Sub

Sub txtW4_Change()
	If UNICDbl(frm1.txtW4.Value) < 0 Then
		Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "(4)ȸ�����", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
		.Value = 0
	End If
	Call SetHeadReCalc
End Sub

Sub txtW6_Change()
	If UNICDbl(frm1.txtW6.Value) < 0 Then
		Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "(6)�����Ѽ� ���뿡 ���� �ձݺ��ξ�", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
		.Value = 0
	End If
	Call SetHeadReCalc
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
	Dim dblSum, dblFiscYear
	Dim dblW9, dblW10, dblW11, dblW12, dblW13, dblW14, dblW15, dblW16, dblW17, dblW18
	Dim dblW19, dblW20, dblW21, dblW22, dblW23, dblW24, dblW25
	
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
			Case C_W9	' ������ ����� 
				dblFiscYear = UNICDbl(frm1.txtFISC_YEAR.text)
				.Col = C_W9	: .Row = Row	: dblW17 = UNiCDbl(.Value)
				If dblFiscYear - 5 > dblW17 Or dblFiscYear < dblW17 Then
					Call DisplayMsgBox("W40002", parent.VB_INFORMATION, "X", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.Value = ""
					Exit Sub
				End If
			Case C_W10, C_W11, C_W12, C_W13, C_W14, C_W15, C_W16
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.Value)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.Value = 0
				End If
				
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' �հ� 
				
				Call SetW17_W18(Row)
		End Select
	
		ggoSpread.Source = lgvspdData(Index)
		ggoSpread.UpdateRow .MaxRows
	ElseIf Index = TYPE_2 Then
		Select Case Col
			Case C_W9
				.Col = Col : .Row = Row : dblW9 = UNICDbl(.Value)	' �������� 
				
				Call GetW16(dblW9, dblW16, dblW17)
				
				If dblW16 = -1 Then
					Call DisplayMsgBox("W40001", parent.VB_INFORMATION, "X", "X")           '��: "(17)�Աݻ��Կ����� �߰��Ҽ������ϴ�."
					.Value = ""
					Exit Sub
				End If

				dblFiscYear = UNICDbl(frm1.txtFISC_YEAR.text)
						
				.Row = Row
				.Col = C_W19 : .Value = 0	: dblW19 = 0
				.Col = C_W20 : .Value = 0	: dblW20 = 0
				.Col = C_W21 : .Value = 0	: dblW21 = 0
							
				If (dblFiscYear - dblW9) = 3 Then
					.Col = C_W19
					dblW19 = UNICDbl(UNIFormatNumber( dblW16/3 * lgMonth(0) , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)) 
					.Value = dblW19
				ElseIf (dblFiscYear - dblW9) = 4 Then
					.Col = C_W20
					dblW20 = UNICDbl(UNIFormatNumber( dblW16/2 * lgMonth(0) , ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)) 
					.Value = dblW20
				ElseIf (dblFiscYear - dblW9) = 5 Then
					.Col = C_W21
					dblW21 = dblW16 * lgMonth(0)
					.Value = dblW21
				End If
		
				Call SetGridTYPE_2(Row)
			Case C_W19, C_W20, C_W21, C_W24
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.Value)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.Value = 0
				End If
				
				Call SetGridTYPE_2(Row)		
		End Select
		ggoSpread.Source = lgvspdData(Index)
		ggoSpread.UpdateRow .MaxRows
	End If
	
	End With
	
End Sub

' -- 2��° �׸��� 
Sub SetGridTYPE_2(Byval Row)
	Dim dblW9, dblW10, dblW11, dblW12, dblW13, dblW14, dblW15, dblW16, dblW17, dblW18
	Dim dblW19, dblW20, dblW21, dblW22, dblW23, dblW24, dblW25, iGrid1Row

	With lgvspdData(TYPE_2)
		
		.Row = Row
		.Col = C_W9	 : dblW9 = UNICDbl(.value)
		.Col = C_W19 : dblW19 = UNICDbl(.Value)
		.Col = C_W20 : dblW20 = UNICDbl(.Value)
		.Col = C_W21 : dblW21 = UNICDbl(.Value)
									
		' �հ躯�� 
		Call FncSumSheet(lgvspdData(TYPE_2), C_W19, 1, .MaxRows - 1, true, .MaxRows, C_W19, "V")	' �հ� 
		Call FncSumSheet(lgvspdData(TYPE_2), C_W20, 1, .MaxRows - 1, true, .MaxRows, C_W20, "V")	' �հ� 
		Call FncSumSheet(lgvspdData(TYPE_2), C_W21, 1, .MaxRows - 1, true, .MaxRows, C_W21, "V")	' �հ� 
					
		' W22 ���� 
		dblW22 = dblW19 + dblW20 + dblW21
		.Col = C_W22	: .Row = Row : .Value = dblW22
					
		Call FncSumSheet(lgvspdData(TYPE_2), C_W22, 1, .MaxRows - 1, true, .MaxRows, C_W22, "V")	' �հ� 
					
		' W23 ���� 
		iGrid1Row = GetRowByW9(TYPE_1, dblW9)	' �׸���1���� �ش� ������ ã�´�.
		If iGrid1Row > 0 Then	
			dblW17 = UNICDbl(GetGrid(TYPE_1, C_W17, iGrid1Row))
			dblW23 = dblW17 + dblW22
			.Col = C_W23	: .Row = Row : .Value = dblW23
		
			Call FncSumSheet(lgvspdData(TYPE_2), C_W23, 1, .MaxRows - 1, true, .MaxRows, C_W23, "V")	' �հ� 
		End If
		
		.Row = Row			
		.Col = C_W24	: dblW24 = UNICDbl(.Value)
		' W25 ���� 
		dblW25= dblW23 - dblW24
		.Col = C_W25	: .Value = dblW25
					
		Call FncSumSheet(lgvspdData(TYPE_2), C_W25, 1, .MaxRows - 1, true, .MaxRows, C_W25, "V")	' �հ�	
	End With
End Sub

' W9 (���Կ���)�� ���� ã�´� 
Function GetRowByW9(Byval pType, Byval pW9)
	Dim iMaxRows, iRow
	With lgvspdData(pType)
		iMaxRows = .MaxRows
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = C_W9 
			If CStr(.Value) = CStr(pW9) Then
				GetRowByW9 = iRow
				Exit Function
			End If
		Next
	End With
	GetRowByW9 = -1
End Function

Function GetGrid(Byval pType, Byval pCol, Byval pRow)
	With lgvspdData(pType)
		.Col = pCol : .Row = pRow : GetGrid = .value
	End With
End Function

' 2�� �׸��忡�� 1�� �׸����� ����Ÿ�� ã�Ƽ� W16�ݾ��� �����Ѵ� 
Sub GetW16(Byval pYear , Byref pdblW16, Byref pdblW17)
	Dim iRow, iMaxRows
	With lgvspdData(TYPE_1)
		iMaxRows = .MaxRows - 1
		.Col = C_W9
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			If UNICDbl(.Value) = pYear Then
				.Col = C_W16 : pdblW16 = UNICDbl(.Value)
				.Col = C_W17 : pdblW17 = UNICDbl(.Value)
				Exit Sub
			End If
		Next
		pdblW16 = -1 : pdblW17 = -1
	End With
End Sub


' �ܾ� �÷��� ����ɶ� ȣ��� 
Sub SetW17_W18(Row)
	Dim dblSum, dblW11, dblW12, dblW13, dblW14, dblW15, dblW16, dblW17, dblW18, iRow, dblW9
	
	With lgvspdData(TYPE_1)
		
		.Row = Row
		.Col = C_W11	: dblW11 = UNICDbl(.Value)	' ���� 
		.Col = C_W12	: dblW12 = UNICDbl(.Value)	' ���� 
		.Col = C_W13	: dblW13 = UNICDbl(.Value)	' ���� 
		
		.Col = C_W14	: dblW14 = UNICDbl(.Value)	' ���� 
		.Col = C_W15	: dblW15 = UNICDbl(.Value)	' �뺯 
		.Col = C_W16	: dblW16 = UNICDbl(.Value)	' ���� 
		
		.Col = C_W18	: dblW18 = dblW11 - dblW12 - dblW13				: .Value = dblW18
		.Col = C_W17	: dblW17 = dblW18 - dblW14 - dblW15 - dblW16	: .Value = dblW17

		Call FncSumSheet(lgvspdData(TYPE_1), C_W17, 1, .MaxRows - 1, true, .MaxRows, C_W17, "V")	' �հ� 
		Call FncSumSheet(lgvspdData(TYPE_1), C_W18, 1, .MaxRows - 1, true, .MaxRows, C_W18, "V")	' �հ� 
		
		If lgvspdData(TYPE_2).MaxRows > 0 Then
			dblW9 = GetGrid(TYPE_1, C_W9, Row)	' �������� ���Կ����� ���Ѵ�.
			iRow = GetRowByW9(TYPE_2, dblW9)
			If iRow > 0 Then Call vspdData_Change(TYPE_2, C_W9, iRow)	' �ձݻ��Կ����� �������� �߰ߵǸ� ..
		End If
	End With
	
End Sub

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
    Dim blnChange, i, sMsg, iRow
    
    blnChange = False
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    For i = TYPE_1 To TYPE_2
		With lgvspdData(i)
			If .MaxRows > 0 Then
				ggoSpread.Source = lgvspdData(i)
				If ggoSpread.SSCheckChange = True Then
					blnChange = True
				End If
			End If
		End With
	Next

    If lgBlnFlgChgValue = False And blnChange = False Then
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
	Dim dblW11, dblW12, dblW16, dblW14, dblW15, dblW13
	
	Verification = False
	
	With lgvspdData(TYPE_1)
		If .MaxRows > 0 Then
		
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
				Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, "������[(W14)+(W15)]", "(11)��λ� �غ�� �����ܾ�")                          <%'No data changed!!%>
				Exit Function
			End If

			'3. W11 < W16
			.Col = C_W16 : dblW16 = UNICDbl(.Value)
			If dblW11 < dblW16 Then
				Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, "������[(W16)]", "(11)��λ� �غ�� �����ܾ�")                          <%'No data changed!!%>
				Exit Function
			End If
		
			'4. W11 < W13
			.Col = C_W13 : dblW13 = UNICDbl(.Value)
			If dblW11 < dblW13 Then
				Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, "(13)�غ�� ���� �����", "(11)��λ� �غ�� �����ܾ�")                          <%'No data changed!!%>
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
			
			.Col = C_SEQ_NO : .Text = iRow	
			
			iRow = 2		: .Row = iRow
			.Col = C_SEQ_NO : .Text = SUM_SEQ_NO	
			.Col = C_W9	: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
			ggoSpread.SpreadLock C_W9, iRow, C_W16, iRow
		
		Else
				
			If iRow = .MaxRows Then	' -- ������ �հ��ٿ��� InsertRow�� �ϸ� ������ �߰��Ѵ�.
				ggoSpread.InsertRow iRow-1 , imRow 
				SetSpreadColor lgCurrGrid, iRow, iRow + imRow - 1

				Call SetDefaultVal(lgCurrGrid, iRow, imRow)
			Else
				ggoSpread.InsertRow ,imRow
				SetSpreadColor lgCurrGrid, iRow+1, iRow + imRow

				Call SetDefaultVal(lgCurrGrid, iRow+1, imRow)
			End If   
		End If
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
		For iCol = 3 To iMaxCols
			Call FncSumSheet(lgvspdData(pType), iCol, 1, .MaxRows - 1, true, .MaxRows, iCol, "V")	' �հ� 
		Next
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
	lgBlnFlgChgValue = False
	
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE
		    
	' �������� ���� : ���ߵǸ� ���ȴ�.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 ����üũ : �׸��� �� 
	If wgConfirmFlg = "N" Then
		'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1
		Call SetSpreadLock(TYPE_1)
		Call SetSpreadLock(TYPE_2)
		'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
		Call SetToolbar("1111111100011111")										<%'��ư ���� ���� %>
			
		'3. ������/������� ���� üũ(�������� ���� üũ)
		With frm1
			If .txtW2_VAL.value <> lgW2(1) Then
			End If
		End With
	Else
		
		'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
		Call SetToolbar("1110000000011111")										<%'��ư ���� ���� %>
	End If
	
	Call SetSpreadTotalLine ' - �հ���� �籸�� 
	
	lgBlnFlgChgValue= FALSE
	
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
	Call InitVariables 
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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:GetRef()">�ݾ׺ҷ�����</A>|<A href="vbscript:OpenRefMenu">�ҵ�ݾ��հ�ǥ��ȸ</A></TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w4103ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : ������ ������ ������ ũ�⿡ ���� ��ũ�ѹٰ� �����ǰ� �Ѵ� %>
						
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=15%>
                                   <table <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
									   <TR>
										   <TD width="100%" COLSPAN=9 CLASS="CLSFLD"><br>&nbsp;1. �ձݻ�۾�����</TD>
									   </TR>
									   <TR>
										   <TD CLASS="TD51" width="20%" ALIGN=CENTER>(1)���Ա�</TD>
										   <TD CLASS="TD51" width="20%" ALIGN=CENTER COLSPAN=2>(2)������</TD>
									       <TD CLASS="TD51" width="20%" ALIGN=CENTER COLSPAN=2>(3)�ѵ��� [(1)x(2)]</TD>
									       <TD CLASS="TD51" width="20%" ALIGN=CENTER>(4)ȸ�����</TD>
									       <TD CLASS="TD51" width="20%" ALIGN=CENTER>(5)�ѵ��ʰ��� [(4)-(3)]</TD>
									   </TR>
									  <TR>
											<TD width="20%"><script language =javascript src='./js/w4103ma1_txtW1_txtW1.js'></script></TD>
											<TD width="20%" COLSPAN=2><SELECT id="cboW2" name=cboW2 tag="25X2Z" Style="width : 100%; text-align: center"></SELECT><INPUT TYPE=HIDDEN ID="txtW2_VAL" NAME="txtW2_VAL"></TD>
											<TD width="20%" COLSPAN=2><script language =javascript src='./js/w4103ma1_txtW3_txtW3.js'></script></TD>
											<TD width="20%"><script language =javascript src='./js/w4103ma1_txtW4_txtW4.js'></script></TD>
											<TD width="20%"><script language =javascript src='./js/w4103ma1_txtW5_txtW5.js'></script></TD>
									  </TR>
									   <TR>
									       <TD width="30%" CLASS="TD51" ALIGN=CENTER COLSPAN=2>(6)�����Ѽ� ���뿡 ���� �ձݺ��ξ�</TD>
									       <TD width="28%" CLASS="TD51" ALIGN=CENTER COLSPAN=2>(7)�ձݺһ��� �� [(5)+(6)]</TD>
								           <TD width="22%" CLASS="TD51" ALIGN=CENTER COLSPAN=2>(8)�ձݻ��Ծ� [(4)-(5)-(6)]</TD>
									       <TD width="20%" CLASS="TD51" ALIGN=CENTER>�� ��</TD>
									  </TR>
									  <TR>
											<TD width="30%" COLSPAN=2><script language =javascript src='./js/w4103ma1_txtW6_txtW6.js'></script></TD>
											<TD width="28%" COLSPAN=2><script language =javascript src='./js/w4103ma1_txtW7_txtW7.js'></script></TD>
											<TD width="22%" COLSPAN=2><script language =javascript src='./js/w4103ma1_txtW8_txtW8.js'></script></TD>
											<TD width="20%"><INPUT TYPE=TEXT id="txtDESC1" name=txtDESC1 ALT="�� ��" tag="25X2Z" Style="width : 100%"></TD>
									  </TR>
								  </table>
								</TD>
							</TR>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=* VALIGN=TOP>
									<table <%=LR_SPACE_TYPE_20%> border="0" width="100%">
									   <TR>
										   <TD width="100%" HEIGHT=10 CLASS="CLSFLD"><br>&nbsp;2. �ͱݻ�۾� ����</TD>
									   </TR>
									   <TR>
										   <TD width="100%" HEIGHT=190><script language =javascript src='./js/w4103ma1_vspdData0_vspdData0.js'></script>
										  </TD>
									  </TR>
									   <TR>
										   <TD width="100%" HEIGHT=120><script language =javascript src='./js/w4103ma1_vspdData1_vspdData1.js'></script>
										  </TD>
									  </TR>
									  <TR>
										  <TD height=*>&nbsp;</TD>
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

