<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �� ���� ���� 
'*  3. Program ID           : W3129MA1
'*  4. Program Name         : W3129MA1.asp
'*  5. Program Desc         : ��20ȣ �����󰢺��������� �հ�ǥ 
'*  6. Modified date(First) : 2005/01/19
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

Const BIZ_MNU_ID = "w3131ma1"
Const BIZ_PGM_ID = "w3131mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID = "w3131OA1"
Const TYPE_1	= 0		' �׸��� �迭��ȣ. 
Const TYPE_2	= 1		

Dim C_SEQ_NO
Dim C_W1
Dim C_W2
Dim C_W3_CD
Dim C_W3
Dim C_W4_CD
Dim C_W4
Dim C_W5_CD
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

Dim IsOpenPop    
Dim gSelframeFlg   
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(1)
Dim StrSum

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()

	C_SEQ_NO			= 1

	C_W1				= 2
    C_W2				= 3
    C_W3_CD				= 4
    C_W3				= 5
    C_W4_CD				= 6
    C_W4				= 7
    C_W5_CD				= 8
    C_W5				= 9

    C_W6				= 10
    
    C_W7				= 2
    C_W8				= 3
    C_W9				= 4
    C_W10				= 5
    C_W11				= 6
    C_W12				= 7
    C_W13				= 8
    C_W14				= 9
    C_W15				= 10
    C_W16				= 11
    C_W17				= 12
    C_W18				= 13
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

	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1
		
    Call initSpreadPosVariables()  
	
	' 1�� �׸��� 
	With lgvspdData(TYPE_1)
			
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gAllowDragDropSpread   
    
		.ReDraw = false
    
		.MaxCols = C_W6 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols														'��: ����� �� Hidden Column
		.ColHidden = True    
		       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����", 10,,,10,1
		ggoSpread.SSSetEdit		C_W1,		"(1)�ڻ꺰", 20,,,100,1
		ggoSpread.SSSetDate		C_W2,		"(2)�򰡹��" & vbCrLf & "�Ű�����",	10,	2, Parent.gDateFormat,	-1
		ggoSpread.SSSetCombo	C_W3_CD,	"�ڵ�"    , 10, 2
		ggoSpread.SSSetCombo	C_W3,		"(3)�Ű���"   , 15, 2
		ggoSpread.SSSetCombo	C_W4_CD,	"�ڵ�"    , 10, 2 
		ggoSpread.SSSetCombo	C_W4,		"(4)�򰡹��"   , 15, 2
		ggoSpread.SSSetCombo	C_W5_CD,		"�ڵ�"    , 10, 2 
		ggoSpread.SSSetCombo	C_W5,		"(5)�� ��"  , 15, 2
		ggoSpread.SSSetEdit		C_W6,		"(6)�� ��", 10,,,50,1

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W3_CD,C_W3_CD,True)
		Call ggoSpread.SSSetColHidden(C_W4_CD,C_W4_CD,True)
		Call ggoSpread.SSSetColHidden(C_W5_CD,C_W5_CD,True)
		.rowheight(-1000) = 20	' ���� ������ 
	'Call InitSpreadComboBox2()
	
		.ReDraw = true

		Call SetSpreadLock(TYPE_1)

    End With   
    
	' 2�� �׸��� 
	With lgvspdData(TYPE_2)
			
		ggoSpread.Source = lgvspdData(TYPE_2)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_2,,parent.gAllowDragDropSpread   
    
		.ReDraw = false
 
		'����� 2�ٷ�    
		.ColHeaderRows = 3   
		       
		.MaxCols = C_W18 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols														'��: ����� �� Hidden Column
		.ColHidden = True    
		       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����", 10,,,10,1
		ggoSpread.SSSetEdit		C_W7,		"(7)����", 10,,,10,1
		ggoSpread.SSSetEdit		C_W8,		"(8)ǰ��", 7,,,10,1
		ggoSpread.SSSetEdit		C_W9,		"(9)�԰�", 7,,,10,1
		ggoSpread.SSSetEdit		C_W10,		"(10)����", 7,,,10,1
		ggoSpread.SSSetFloat	C_W11,		"(11)����", 7, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z" 
		ggoSpread.SSSetFloat	C_W12,		"(12)�ܰ�", 7, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec ,,,"Z" 
		ggoSpread.SSSetFloat	C_W13,		"(13)�ݾ�", 13, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec ,,,"Z" 
		ggoSpread.SSSetFloat	C_W14,		"(14)�ܰ�", 7, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"  
		ggoSpread.SSSetFloat	C_W15,		"(15)�ݾ�", 13, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"  
		ggoSpread.SSSetFloat	C_W16,		"(16)�ܰ�", 7, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"  
		ggoSpread.SSSetFloat	C_W17,		"(17)�ݾ�", 13, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"  
		ggoSpread.SSSetFloat	C_W18,		"(18)������", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec ,,,"Z" 

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
					
		' �׸��� ��� ��ħ ���� 
		ret = .AddCellSpan(C_SEQ_NO	, -1000, 1, 3)
		ret = .AddCellSpan(C_W7		, -1000, 1, 3)	
		ret = .AddCellSpan(C_W8		, -1000, 1, 3)	
		ret = .AddCellSpan(C_W9		, -1000, 1, 3)	
		ret = .AddCellSpan(C_W10	, -1000, 1, 3)	
		ret = .AddCellSpan(C_W11	, -1000, 1, 3)	
		ret = .AddCellSpan(C_W12	, -1000, 2, 1)
		ret = .AddCellSpan(C_W12	, -999 , 1, 2)
		ret = .AddCellSpan(C_W13	, -999 , 1, 2)
		ret = .AddCellSpan(C_W14	, -1000, 4, 1)	
		ret = .AddCellSpan(C_W14	, -999 , 2, 1)	
		ret = .AddCellSpan(C_W16	, -999 , 2, 1)	
		ret = .AddCellSpan(C_W18	, -1000, 1, 3)	
			
		 ' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W12	: .Text = "ȸ �� �� ��"
		.Col = C_W14	: .Text = "�� �� �� �� �� ��"
		.Col = C_W18	: .Text = "(18)������((15)�Ǵ� (15)�� (17)�� ū �ݾ� - (13))"
				
		' �ι�° ��� ��� ���� 
		.Row = -999	
		.Col = C_W12	: .Text = "(12)�ܰ�"
		.Col = C_W13	: .Text = "(13)�ݾ�"
		.Col = C_W14	: .Text = "�Ű���"
		.Col = C_W16	: .Text = "���Լ����"

		' ����° ��� ��� ���� 
		.Row = -998	
		.Col = C_W14	: .Text = "(14)�ܰ�"
		.Col = C_W15	: .Text = "(15)�ݾ�"
		.Col = C_W16	: .Text = "(16)�ܰ�"
		.Col = C_W17	: .Text = "(17)�ݾ�"
							
		.rowheight(-999) = 15	' ���� ������ 
		.rowheight(-998) = 15	' ���� ������	
		
		.ReDraw = true

		Call SetSpreadLock(TYPE_2)

    End With  
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub SetSpreadLock(pType)
    With lgvspdData(pType)
		ggoSpread.Source = lgvspdData(pType)
		.ReDraw = False
    
		If pType = TYPE_1 Then
			ggoSpread.SpreadLock C_W1, -1, C_W1
		ElseIf pType = TYPE_2 Then
			'ggoSpread.SpreadLock C_W13, -1, C_W13
			ggoSpread.SpreadLock C_W18, -1, C_W18
		End If

		.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With lgvspdData(TYPE_2)
		ggoSpread.Source = lgvspdData(TYPE_2)
		.ReDraw = False

		'ggoSpread.SSSetProtected C_W13, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_W18, pvStartRow, pvEndRow

		.ReDraw = True

    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO			= iCurColumnPos(1)
            C_DOC_DATE			= iCurColumnPos(2)
            C_DOC_AMT			= iCurColumnPos(3)
            C_DEBIT_CREDIT		= iCurColumnPos(4)
            C_DEBIT_CREDIT_NM	= iCurColumnPos(5)
            C_SUMMARY_DESC		= iCurColumnPos(6)
            C_COMPANY_NM		= iCurColumnPos(7)
            C_STOCK_RATE		= iCurColumnPos(8)
            C_ACQUIRE_AMT       = iCurColumnPos(9)
            C_COMPANY_TYPE		= iCurColumnPos(10)
            C_COMPANY_TYPE_NM	= iCurColumnPos(11)
            C_HOLDING_TERM		= iCurColumnPos(12)
            C_OWN_RGST_NO		= iCurColumnPos(13)
            C_CO_ADDR			= iCurColumnPos(14)
            C_REPRE_NM			= iCurColumnPos(15)
            C_STOCK_CNT			= iCurColumnPos(16)
    End Select    
End Sub

Sub InitData()
	Dim iMaxRows, iRow
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	iMaxRows = 6 ' �ϵ��ڵ��Ǵ� ��� 
	With lgvspdData(TYPE_1)
		.Redraw = False
		
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		ggoSpread.InsertRow , iMaxRows

		iRow = 0
		iRow = iRow + 1 : .Row = iRow
		.Col = C_SEQ_NO		: .value = iRow
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_SEQ_NO		: .value = iRow
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_SEQ_NO		: .value = iRow

		iRow = iRow + 1 : .Row = iRow
		.Col = C_SEQ_NO		: .value = iRow

		iRow = iRow + 1 : .Row = iRow
		.Col = C_SEQ_NO		: .value = iRow

		iRow = iRow + 1 : .Row = iRow
		.Col = C_SEQ_NO		: .value = iRow

		Call InitSpreadComboBox
		
		.Redraw = True
		
		Call InitData2
		
		Call SetSpreadLock(TYPE_1)
	End With	
End Sub

 ' -- DBQueryOk ������ �ҷ��ش�.
Sub InitData2()
	Dim iRow
	
	With lgvspdData(TYPE_1)
		.Redraw = False

		iRow = 0
		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1	: .value = " �� ǰ �� �� ǰ"
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1	: .value = " �� �� ǰ �� �� �� ǰ"
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1	: .value = " ��   ��   ��"

		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1	: .value = " ��   ��   ǰ"

		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1	: .value = " �� �� �� �� | ä ��"

		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1	: .value = " �� �� �� �� | �� Ÿ"

	End With
End Sub


Sub InitSpreadComboBox()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx

	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " MAJOR_CD='W1057' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		Call Spread_SetCombo(Replace(lgF0, chr(11), chr(9)), C_W3_CD, 1, 4)
		Call Spread_SetCombo(Replace(lgF1, chr(11), chr(9)), C_W3, 1, 4)
		Call Spread_SetCombo(Replace(lgF0, chr(11), chr(9)), C_W4_CD, 1, 4)
		Call Spread_SetCombo(Replace(lgF1, chr(11), chr(9)), C_W4, 1, 4)
	End If

	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " MAJOR_CD='W1058' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		Call Spread_SetCombo(Replace(lgF0, chr(11), chr(9)), C_W3_CD, 5, 6)
		Call Spread_SetCombo(Replace(lgF1, chr(11), chr(9)), C_W3, 5, 6)
		Call Spread_SetCombo(Replace(lgF0, chr(11), chr(9)), C_W4_CD, 5, 6)
		Call Spread_SetCombo(Replace(lgF1, chr(11), chr(9)), C_W4, 5, 6)
	End If
		  		  
	iCodeArr = vbTab & lgF0
    iNameArr = vbTab & lgF1
    
    iCodeArr =  "0" & chr(11) & "1"
    iNameArr =  "��" & chr(11) & "��"
     Call Spread_SetCombo(Replace(iCodeArr, chr(11), chr(9)), C_W5_CD,1, 6)
    Call Spread_SetCombo(Replace(iNameArr, chr(11), chr(9)), C_W5,1, 6)

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
		.TypeComboBoxMaxDrop = 3
		' Select the first item in the list
		'.TypeComboBoxCurSel = 0
		' Set the width to display the widest item in the list
		'.TypeComboBoxWidth = 1 
		.BlockMode = False
	End With

End Sub

'============================== ���۷��� �Լ�  ========================================
Function GetRef_OLD()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	' 2. �˾� 
	Dim arrRet, sParam, iRow, iActRow
	Dim arrParam(5), iMaxRows

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = sCoCd
	arrParam(1) = sFiscYear
	arrParam(2) = sRepType

	arrRet = window.showModalDialog("w3131ra1.asp", Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0, 0) = "" Then
		frm1.txtCO_CD.focus
	    Exit Function
	End If	

	
	With lgvspdData(TYPE_2)
		iActRow = UNICDbl(.ActiveRow)
		.Redraw = False
		
		lgBlnFlgChgValue = True
		ggoSpread.Source = lgvspdData(TYPE_2)
		iMaxRows = UBound(arrRet, 1)

		If iMaxRows > 1 Then
			Call FncInsertRow(1)
			Call FncInsertRow(iMaxRows-1)
		Else
			Call FncInsertRow(iMaxRows)
		End If
		
		For iRow = 1 To iMaxRows
			.Row = iActRow+iRow

			If lgIntFlgMode = parent.OPMD_UMODE Then
				ggoSpread.UpdateRow iRow
			End If
			
			.Col = C_W7		: .Value = arrRet(iRow-1, 0)
			.Col = C_W13	: .Value = arrRet(iRow-1, 1)
			
			' �հ���� ����Ѵ�.
			Call FncSumSheet(lgvspdData(TYPE_2), iRow, C_W3, C_W6, true, iRow, C_W2, "H")
		Next
			
		.Redraw = True
	End With
	
End Function

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

    arrRet = window.showModalDialog("w3131ra1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
	If arrRet(0, 0) = "" Then
	    Exit Function
	End If	

	
	With lgvspdData(TYPE_2)
		.Redraw = False
		
		lgBlnFlgChgValue = True
		ggoSpread.Source = lgvspdData(TYPE_2)
		ggoSpread.ClearSpreadData

		iMaxRows = UBound(arrRet, 1)
		
		Call FncInsertRow(1)
		If iMaxRows > 1 Then Call FncInsertRow(iMaxRows-1)
		For iRow = 0 To iMaxRows-1
			.Row = iRow+1
		
			'.Col = C_W1_CD	: .Value = arrRet(iRow, 3)
			.Col = C_W7		: .Value = arrRet(iRow, 5)
			.Col = C_W13	: .Value = arrRet(iRow, 2)
			Call vspdData_Change(TYPE_2, C_W13, iRow+1)
		Next
			
		.Redraw = True
	End With
	
End Function

Sub SetTotalRowLine()
	Dim iRow , ret
	With lgvspdData(TYPE_2)
		ggoSpread.Source = lgvspdData(TYPE_2)
		iRow = .MaxRows
		.Row = iRow
		.Col = C_W7		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
		ret = .AddCellSpan(C_W7	, iRow, 5, 1)
		ggoSpread.SpreadLock C_W7, iRow, C_W18, iRow
	End With
End Sub

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100000111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()

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
	If Index <> TYPE_1 Then Exit Sub
	
	With lgvspdData(TYPE_1)
		Select Case Col
			Case C_W3, C_W4 , C_W5
				.Col = Col	: .Row = Row
				iIdx = UNICDbl(.Value)
				.Col = Col -1
				.Value = iIdx
		End Select
	End With
End Sub

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, dblW11, dblW12, dblW13 , dblFiscYear, dblW26, dblW25, dblW24, dblW23, dblW22, dblW17, dblW15, dblW14
	Dim dblW27, dblW28, dblW29, dblW30, dblW31, dblW32, dblW33
	
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
	
	If Index = TYPE_1 Then
		Select Case Col
			Case C_W3, C_W4 , C_W5
				Call vspdData_ComboSelChange(TYPE_1, Col, Row)
		End Select
	ElseIf Index = TYPE_2 Then	'1�� �׸� 
	
		Select Case Col
			Case C_W11, C_W12, C_W14, C_W16
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.Value)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.Value = 0
				End If
				
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' �հ� 
				
				Call SetGridTYPE_2(Row, Col)
				
				ggoSpread.UpdateRow .MaxRows
			
			' �� ���� �ݾ��� ����� ��������� �����Ѵ�. 2005-02-21 ���� 
			Case C_W13, C_W15, C_W17
				Call SetGridTYPE_2_2(Row)
				
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' �հ� 
				
				ggoSpread.UpdateRow .MaxRows
				
		End Select
	End If
	
	End With
	
End Sub

Sub SetGridTYPE_2(Byval Row, Byval Col)
	Dim dblSum, dblW11, dblW12, dblW13, dblW14, dblW15, dblW16, dblW17, dblW18

	With lgvspdData(TYPE_2)
		.Row = Row
		
		.Col = C_W11 : dblW11 = UNICDbl(.Value)
		
		' W11, W12 ����� W13�ڵ���� 
		If Col = C_W11 Or Col = C_W12 Then
		
			.Col = C_W12 : dblW12 = UNICDbl(.Value)
				
			' W13 ���� 
			dblW13 = dblW11 * dblW12
			.Col = C_W13	: .Row = Row : .Value = dblW13
						
			Call FncSumSheet(lgvspdData(TYPE_2), C_W13, 1, .MaxRows - 1, true, .MaxRows, C_W13, "V")	' �հ� 
		Else
			.Col = C_W13 : dblW13 = UNICDbl(.Value)
		End If
		
		.Row = Row	' -- FncSumSheet �� Row�� ����� 
					
		' W11, W14 ����� W15�ڵ���� 
		If Col = C_W11 Or Col = C_W14 Then
		
			.Col = C_W14 : dblW14 = UNICDbl(.Value)
			dblW15 = dblW11 * dblW14
			.Col = C_W15	: .Row = Row : .Value = dblW15

			Call FncSumSheet(lgvspdData(TYPE_2), C_W15, 1, .MaxRows - 1, true, .MaxRows, C_W15, "V")	' �հ� 
		Else
			.Col = C_W15 : dblW15 = UNICDbl(.Value)
		End If
		
		.Row = Row	
		' W11, W16 ����� W17�ڵ���� 
		If Col = C_W11 Or Col = C_W16 Then
			.Col = C_W16	: dblW16 = UNICDbl(.Value)
			
			dblW17 = dblW11 * dblW16
			.Col = C_W17	: .Value = dblW17 
						
			Call FncSumSheet(lgvspdData(TYPE_2), C_W17, 1, .MaxRows - 1, true, .MaxRows, C_W17, "V")	' �հ�	
		Else
			.Col = C_W17 : dblW17 = UNICDbl(.Value)
		End If
		
		' W18 ����: (18)������ (15�Ǵ� 15�� (17)�� ū �ݾף�13)
		If dblW15 > dblW17 Then
			dblW18 = dblW15 - dblW13
		Else
			dblW18 = dblW17 - dblW13
		End If
		.Col = C_W18	: .Row = Row : .Value = dblW18 
		
		Call FncSumSheet(lgvspdData(TYPE_2), C_W18, 1, .MaxRows - 1, true, .MaxRows, C_W18, "V")	' �հ�	
		
	End With
End Sub

Sub SetGridTYPE_2_2(Row)
	Dim dblSum, dblW11, dblW12, dblW13, dblW14, dblW15, dblW16, dblW17, dblW18

	With lgvspdData(TYPE_2)
		.Row = Row
				
		.Col = C_W11	: .value = ""
		.Col = C_W12	: .value = ""
		.Col = C_W14	: .value = ""
		.Col = C_W16	: .value = ""		

		.Col = C_W13 : dblW13 = UNICDbl(.Value)
		.Col = C_W15 : dblW15 = UNICDbl(.Value)
		.Col = C_W17 : dblW17 = UNICDbl(.Value)
		
		' W18 ����: (18)������ (15�Ǵ� 15�� (17)�� ū �ݾף�13)
		If dblW15 > dblW17 Then
			dblW18 = dblW15 - dblW13
		Else
			dblW18 = dblW17 - dblW13
		End If
		.Col = C_W18	: .Row = Row : .Value = dblW18 
			
		Call FncSumSheet(lgvspdData(TYPE_2), C_W18, 1, .MaxRows - 1, true, .MaxRows, C_W18, "V")	' �հ�	
		
	End With
End Sub


Sub vspdData_Click(Index, ByVal Col, ByVal Row)
	lgCurrGrid = Index
	
	If Index = TYPE_1 Then Exit Sub
	
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

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    If lgBlnFlgChgValue = True Then
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

' -- �÷� ��� ���� 
Function GetColName(Byval pCol)
	With frm1.vspdData
		.Col = pCol	: .Row = -999
		GetColName = .Value
	End With
End Function

Function FncSave() 
    Dim blnChange, dblSum, iCol, i
    
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	
	For i = TYPE_1 To TYPE_2
		ggoSpread.Source = lgvspdData(i)
		If ggoSpread.SSCheckChange = True Then
			blnChange = True
		End If

		If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
		      Exit Function
		End If    
	Next
	
	If blnChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
	End If

<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
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

			.vspdData.Col = C_DOC_AMT
			.vspdData.Text = ""
    
			.vspdData.Col = C_COMPANY_NM
			.vspdData.Text = ""
			
			.vspdData.Col = C_STOCK_RATE
			.vspdData.Text = ""
			
			.vspdData.Col = C_ACQUIRE_AMT
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
    ggoSpread.Source = frm1.vspdData1	
    If frm1.vspdData1.MaxRows = 2 Then
		ggoSpread.EditUndo                                                 
	End If
	ggoSpread.EditUndo 
	
	Call CheckReCalc
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo, ret
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
 
	If lgCurrGrid = TYPE_1 Then lgCurrGrid = TYPE_2
	
	With lgvspdData(lgCurrGrid)	' ��Ŀ���� �׸��� 
			
		ggoSpread.Source = lgvspdData(lgCurrGrid)
			
		iRow = .ActiveRow
		.ReDraw = False
					
		If .MaxRows = 0 Then	' ù InsertRow�� 1��+�հ��� 

			iRow = 1
			ggoSpread.InsertRow , 2
			Call SetSpreadColor(iRow, iRow+1) 
			.Row = iRow		
			
			If lgCurrGrid = TYPE_1 Then
				.Col = C_SEQ_NO : .Text = iRow	
			
				iRow = 2		: .Row = iRow
				.Col = C_SEQ_NO : .Text = SUM_SEQ_NO	
				.Col = C_W7		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
				ret = .AddCellSpan(C_W7	, iRow, 5, 1)
				ggoSpread.SpreadLock C_W7, iRow, C_W18, iRow
			Else
				.Col = C_SEQ_NO : .Text = iRow	
			
				iRow = 2		: .Row = iRow
				.Col = C_SEQ_NO : .Text = SUM_SEQ_NO	
				.Col = C_W7		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
				ret = .AddCellSpan(C_W7	, iRow, 5, 1)
				ggoSpread.SpreadLock C_W7, iRow, C_W18, iRow
			End If		
		
		Else
				
			If iRow = .MaxRows Then	' -- ������ �հ��ٿ��� InsertRow�� �ϸ� ������ �߰��Ѵ�.
				ggoSpread.InsertRow iRow-1 , imRow 
				Call SetSpreadColor(iRow, iRow + imRow - 1)

				'If lgCurrGrid = TYPE_1 Then
					Call SetDefaultVal(lgCurrGrid, iRow, imRow)
				'End If
			Else
				ggoSpread.InsertRow ,imRow
				Call SetSpreadColor(iRow+1, iRow + imRow)

				'If lgCurrGrid = TYPE_1 Then
					Call SetDefaultVal(TYPE_2, iRow+1, imRow)
				'End If
			End If   
		End If
	End With
	

	'Call CheckW7Status(lgCurrGrid)	' ������ ���� üũ 

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
     
End Function

' �׸��忡 SEQ_NO, TYPE �ִ� ���� 
Function SetDefaultVal(Index, iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With lgvspdData(Index)	' ��Ŀ���� �׸��� 

	ggoSpread.Source = lgvspdData(Index)
	
	If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
		.Row = iRow
		MaxSpreadVal lgvspdData(Index), C_SEQ_NO, iRow
	Else
		iSeqNo = MaxSpreadVal(lgvspdData(Index), C_SEQ_NO, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i
			.Col = C_SEQ_NO : .Value = iSeqNo : iSeqNo = iSeqNo + 1
		Next
	End If
	End With
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

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData1 
    	.focus
    	ggoSpread.Source = frm1.vspdData1
    	lDelRows = ggoSpread.DeleteRow
    End With
    
    Call CheckReCalc
End Function

' ���� 
Function CheckReCalc()
	Dim dblSum
	
	With lgvspdData(TYPE_2)
		If .MaxRows = 0 Then Exit Function
		ggoSpread.Source = lgvspdData(TYPE_2)	
	
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W12, 1, .MaxRows - 1, true, .MaxRows, C_W12, "V")	' �հ� 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W13, 1, .MaxRows - 1, true, .MaxRows, C_W13, "V")	' �հ� 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W14, 1, .MaxRows - 1, true, .MaxRows, C_W14, "V")	' �հ� 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W15, 1, .MaxRows - 1, true, .MaxRows, C_W15, "V")	' �հ� 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W16, 1, .MaxRows - 1, true, .MaxRows, C_W16, "V")	' �հ� 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W17, 1, .MaxRows - 1, true, .MaxRows, C_W17, "V")	' �հ� 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W18, 1, .MaxRows - 1, true, .MaxRows, C_W18, "V")	' �հ� 
		
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

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
		
    FncExit = True
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
        strVal = strVal     & "&txtCurrGrid="        & lgCurrGrid      
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
	Call InitSpreadComboBox
	Call InitData2 ' �׸��� �ڻ걸�� ��� 
	
	lgIntFlgMode = parent.OPMD_UMODE
		    
	' �������� ���� : ���ߵǸ� ���ȴ�.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 ����üũ : �׸��� �� 
	If wgConfirmFlg = "N" Then
		'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1
		'Call SetSpreadLock()

		'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
		Call SetToolbar("1101111100000111")										<%'��ư ���� ���� %>

	Else
		
		'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
		Call SetToolbar("1100000000000111")										<%'��ư ���� ���� %>
	End If
	
	'frm1.vspdData.focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow, lCol, i
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel, lMaxRows, lMaxCols
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    For i = TYPE_1 To TYPE_2
    
		With lgvspdData(i)
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

		    document.all("txtSpread"&CStr(i)).value     =  strDel & strVal
			strDel = "" : strVal = ""

		End With
	Next
	
	frm1.txtMode.value        =  Parent.UID_M0002
	
	
	With lgvspdData(TYPE_2)
		.Col = C_W18
		.Row = .MaxRows
		frm1.txtSum.value = .value			
				
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' ���� ������ ���� ���� %>
	Call InitVariables
	Call FncNew()
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
					<TD WIDTH=* align=right><A href="vbscript:GetRef()">����ڻ� ��ȸ</A>| <A href="vbscript:OpenRefMenu">�ҵ�ݾ��հ�ǥ��ȸ</A></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="140">1. ����ڻ� �򰡹�� ���� 
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData0 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread0> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="*">2. ������ ��� 
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread0" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSum" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

