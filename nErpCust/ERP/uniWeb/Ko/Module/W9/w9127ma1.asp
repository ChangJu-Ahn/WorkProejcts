<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �ؿ��������θ��� 
'*  3. Program ID           : w9127ma1
'*  4. Program Name         : w9127ma1.asp
'*  5. Program Desc         : �ؿ��������θ��� 
'*  6. Modified date(First) : 2006/01/09
'*  7. Modified date(Last)  : 2007.03
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      :  lee wol san
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

Const BIZ_MNU_ID = "w9125ma1"
Const BIZ_PGM_ID = "w9127mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID = "w9127OA1"			' -- ���� : EBR�� A,B�� 4���� ���������� �ڿ� �ٿ��� 


Const TAB1 = 1																	'��: Tab�� ��ġ 
Const TAB2 = 2
Const TAB3 = 3
Const TAB4 = 4


' -- ���ĸ� 
Dim C_C_NM
Dim C_C_CD

' -- �÷� ���� 
Dim C_C01
Dim C_C02
Dim C_C03
Dim C_C04
Dim C_C05
Dim C_C06
Dim C_C07
Dim C_C08
Dim C_C09
Dim C_C10
Dim C_C11
Dim C_C12

' -- ������(����)
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
Dim C_W29
Dim C_W30
Dim C_W31
Dim C_W32
Dim C_W33
Dim C_W34
Dim C_W35
Dim C_W36
Dim C_W37
Dim C_W38
Dim C_W39
Dim C_W40
Dim C_W41
Dim C_W42
Dim C_W43
Dim C_W44
Dim C_W45
Dim C_W46
Dim C_W47
Dim C_W48
Dim C_W49
Dim C_W50
Dim C_W51
Dim C_W52
Dim C_W53



Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()

	' -- ���ĸ� 
	C_C_NM			= 1
	C_C_CD			= 2

	' -- �÷� ���� 
	C_C01			= 3
	C_C02			= 4
	C_C03			= 5
	C_C04			= 6
	C_C05			= 7
	C_C06			= 8
	C_C07			= 9
	C_C08			= 10
	C_C09			= 11
	C_C10			= 12
	C_C11			= 13
	C_C12			= 14
	
	' -- ������(����)
	C_W1			= 1
	C_W2			= 2
	C_W3			= 3
	C_W4			= 4
	C_W5			= 5
	C_W6			= 6
	C_W7			= 7
	C_W8			= 8
	C_W9			= 9
	C_W10			= 10
	C_W11			= 11
	C_W12			= 12
	C_W13			= 13
	C_W14			= 14
	C_W15			= 15
	C_W16			= 16
	C_W17			= 17
	C_W18			= 18
	C_W19			= 19
	C_W20			= 20
	C_W21			= 21
	C_W22			= 22
	C_W23			= 23
	C_W24			= 24
	C_W25			= 25
	C_W26			= 26
	C_W27			= 27
	
	
	C_W28			= 1
	C_W29			= 2 
	C_W30			= 3 
	C_W31			= 4 
	C_W32			= 5 
	C_W33			= 6 
	C_W34			= 7 
	C_W35			= 8 
	C_W36			= 9 
	C_W37			= 10
	C_W38			= 11
	C_W39			= 12
	C_W40			= 13
	C_W41			= 14
	C_W42			= 15
	C_W43			= 16
	C_W44			= 17
	C_W45			= 18
	C_W46			= 19
	C_W47			= 20
	C_W48			= 21
	C_W49			= 22 
	C_W50			= 23 
	C_W51			= 24 
	C_W52			= 25 
	C_W53			= 26 
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
	Dim ret, i
	
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("8","12","0")	' -- �ݾ� 15�ڸ� ���� : 
	
	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
    
    .EditEnterAction = 2
    
	.ReDraw = false

    .MaxCols = C_C12 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols														'��: ����� �� Hidden Column
	.ColHidden = True    

  	'����� 2�ٷ�    
	.ColHeaderRows = 3
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData

    ggoSpread.SSSetEdit		C_C_NM,		"", 22
    ggoSpread.SSSetEdit		C_C_CD,		"", 15, 2
     
    ggoSpread.SSSetFloat	C_C01, "01", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C02, "02", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C03, "03", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C04, "04", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C05, "05", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C06, "06", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C07, "07", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C08, "08", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C09, "09", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C10, "10", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C11, "11", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C12, "12", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec

	.Rowheight(-999) = 12	
	.Rowheight(-998) = 12	
	
	' �׸��� ��� ��ħ 
	ret = .AddCellSpan(C_C_NM	, -1000, 1, 3)
	.Col = C_C_NM	: .Row = -1000	: .TypeVAlign = 2	: .TypeHAlign = 2	: .Text = "��  ��"
	.Col = C_C_CD	: .Row = -999	: .TypeVAlign = 2	: .TypeHAlign = 2	: .Text = "(4)�������θ�"
	.Col = C_C_CD	: .Row = -998	: .TypeVAlign = 2	: .TypeHAlign = 2	: .Text = "(5)�������ΰ�����ȣ"
	

	ggoSpread.SSSetSplit2(1)
	' �׸��� ��� ��ħ ���� 

	Call ggoSpread.SSSetColHidden(C_C01, .MaxCols,True)

	.ReDraw = true

    End With   
    

	With frm1.vspdData2
	
	ggoSpread.Source = frm1.vspdData2
   'patch version
    ggoSpread.Spreadinit "V20061222",,parent.gForbidDragDropSpread    
    
    .EditEnterAction = 2
    
	.ReDraw = false

    .MaxCols = C_C12 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols														'��: ����� �� Hidden Column
	.ColHidden = True    

  	'����� 2�ٷ�    
	.ColHeaderRows = 3
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData

    ggoSpread.SSSetEdit		C_C_NM,		"", 22
    ggoSpread.SSSetEdit		C_C_CD,		"", 15, 2
     
    ggoSpread.SSSetFloat	C_C01, "01", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C02, "02", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C03, "03", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C04, "04", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C05, "05", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C06, "06", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C07, "07", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C08, "08", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C09, "09", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C10, "10", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C11, "11", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_C12, "12", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec

	.Rowheight(-999) = 12	
	.Rowheight(-998) = 12	

	' �׸��� ��� ��ħ 
	ret = .AddCellSpan(C_C_NM	, -1000, 1, 3)
	.Col = C_C_NM	: .Row = -1000	: .TypeVAlign = 2	: .TypeHAlign = 2	: .Text = "��  ��"
	.Col = C_C_CD	: .Row = -999	: .TypeVAlign = 2	: .TypeHAlign = 2	: .Text = "(4)�������θ�"
	.Col = C_C_CD	: .Row = -998	: .TypeVAlign = 2	: .TypeHAlign = 2	: .Text = "(5)�������ΰ�����ȣ"

	ggoSpread.SSSetSplit2(1)
	' �׸��� ��� ��ħ ���� 

	Call ggoSpread.SSSetColHidden(C_C01, .MaxCols,True)

	.ReDraw = true

    End With   
        
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitSpreadComboBox()

		  
End Sub


Sub SetSpreadLock()
	Dim i
	
    With frm1.vspdData

    .ReDraw = False
    
    ggoSpread.Source = frm1.vspdData
    
    ggoSpread.SpreadLock C_C_NM, -1, C_C_CD
    
    ggoSpread.SpreadLock C_C01,  C_W1, C_C12, C_W1
    ggoSpread.SpreadLock C_C01,  C_W9, C_C12, C_W9
    ggoSpread.SpreadLock C_C01,  C_W15, C_C12, C_W15
    ggoSpread.SpreadLock C_C01,  C_W22, C_C12, C_W22
    ggoSpread.SpreadLock C_C01,  C_W24, C_C12, C_W24

    .ReDraw = True

    End With

    With frm1.vspdData2

    .ReDraw = False
    
    ggoSpread.Source = frm1.vspdData2
    
    ggoSpread.SpreadLock C_C_NM, -1, C_C_CD
    
    ggoSpread.SpreadLock C_C01,  C_W28, C_C12, C_W28
    ggoSpread.SpreadLock C_C01,  C_W31, C_C12, C_W31
    ggoSpread.SpreadLock C_C01,  C_W34, C_C12, C_W34
    ggoSpread.SpreadLock C_C01,  C_W41, C_C12, C_W41
    ggoSpread.SpreadLock C_C01,  C_W45, C_C12, C_W45
    ggoSpread.SpreadLock C_C01,  C_W48, C_C12, C_W48

    .ReDraw = True

    End With
End Sub


Sub SetColorGrid(Byval pCol, Byval pBoolean)
	Dim i
	
    With frm1.vspdData

    .ReDraw = False

	'For i = C_C01 To C_C12 Step 2
		If .ColHidden = False Then
		
			If pBoolean Then
				ggoSpread.SSSetRequired pCol, C_W6, C_W12
			Else
				ggoSpread.SpreadUnLock pCol, C_W6, pCol, C_W12
			End If
		End If
	'Next

    .ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
          
    End Select    
End Sub

Sub InitData()
	Dim iMaxRows, iRow, ret
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	With frm1.vspdData
		.Redraw = False
		
		ggoSpread.Source = frm1.vspdData
		
		ggoSpread.InsertRow , C_W26+1	 ' �ϵ��ڵ��Ǵ� ��� 

		' -- ���� ������ 
		'.Rowheight(C_W7) = 20		' �������θ� 
		
		.Redraw = True
		
		Call InitData_Tab1
		
	End With	
	
	With frm1.vspdData2
		.Redraw = False
		
		ggoSpread.Source = frm1.vspdData2
		'200703 4���߰� 
		ggoSpread.InsertRow , C_W53' �ϵ��ڵ��Ǵ� ��� 

		' -- ���� ������ 
		'.Rowheight(C_W7) = 20		' �������θ� 
		
		.Redraw = True
		
		Call InitData_Tab2
		
	End With	

	Call SetSpreadLock	
	
	' -- �������θ�/�������ΰ�����ȣ �ҷ����� 
	Dim sWhere, arrCol, arrCol2, iCol
	
	sWhere = " CO_CD=" & FilterVar(frm1.txtCO_CD.value,"''","S") & vbCrLf
	sWhere = sWhere & " AND FISC_YEAR=" & FilterVar(frm1.txtFISC_YEAR.Text,"''","S") & vbCrLf
	sWhere = sWhere & " AND REP_TYPE=" & FilterVar(frm1.cboREP_TYPE.value,"''","S") & vbCrLf
	sWhere = sWhere & " AND W6 <> ''"
	
	call CommonQueryRs(" W7, W8 "," TB_A125 ", sWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If lgF0 <> "" Then
		arrCol = Split(lgF0, Chr(11))
		arrCol2 = Split(lgF1, Chr(11))
		
		With frm1.vspdData
		
		.Col = C_C01	
		For iCol = 0 To UBound(arrCol)-1
			.Row = -999	: .Text = arrCol(iCol)
			.Row = -998		: .Text = arrCol2(iCol)
			.ColHidden = False
			.Col = .Col + 1
		Next
		
		End With

		With frm1.vspdData2
		
		.Col = C_C01	
		For iCol = 0 To UBound(arrCol)-1
			.Row = -999	: .Text = arrCol(iCol)
			.Row = -998		: .Text = arrCol2(iCol)
			.ColHidden = False
			.Col = .Col + 1
		Next
		
		End With
				
	End If
	
End Sub

 ' -- DBQueryOk ������ �ҷ��ش�.
Sub InitData_Tab1()
	Dim iRow  , iCol
	dim iRowTmp
	With frm1.vspdData
		.Redraw = False

		iRow = 0
		iRow = iRow + 1 : .Row = iRow    :  .TypeVAlign = 2
		.Col = C_C_NM	: .value = " I. �� �� �� ��"
	  
	   '200703
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (1)���ݰ�����"
		
		
		
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (2)����ä��(Ư��������)"
		
		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (3)����ä��(�� Ÿ)"
        
        
        iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (4)����ڻ�"	

		iRow = iRow + 1 : .Row = iRow    :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (5)��������"

		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (6)�뿩��(Ư��������)"

		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (7)�뿩��(�� Ÿ)"

		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (8)�����ڻ�"
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "    1.���� �� ���๰"
		
		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "    2.�����ġ,������ݱ�"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "    3.�� Ÿ"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (1)�����ڻ�"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (2)�� Ÿ"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = " II. �� ä �� ��"


		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (1)����ä��(Ư��������)"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (2)����ä��(�� Ÿ)"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (3)���Ա�(Ư��������)"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (4)���Ա�(�� Ÿ)"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (5)�����ޱ�"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (6)�� Ÿ"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = " III. �� �� �� �� ��"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (1)�� �� ��"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (2)��Ÿ �ں���"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "    1.�ں��׿���"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "    2.�����׿���"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "    3.�� Ÿ"

	iRowTmp = C_W1
		For iRow = C_W1 To C_W26+1
		
			if iRow=2 then
				.Col = C_C_CD	: .Row = iRow
				.Text = "50"
				iRow=iRow + 1
			END IF
			
				.Col = C_C_CD	: .Row = iRow
				.Text = Right("0" & iRowTmp, 2)

			iRowTmp= iRowTmp + 1
		Next
		
		.Redraw = True
	End With
End Sub

 ' -- DBQueryOk ������ �ҷ��ش�.
Sub InitData_Tab2()
	Dim iRow  , iCol,iRowTmp

	With frm1.vspdData2
		.Redraw = False

		iRow = 0
		iRow = iRow + 1 : .Row = iRow    :  .TypeVAlign = 2
		.Col = C_C_NM	: .value = " I. �� �� ��"
	  
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (1)������ ���� ����"
		
		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (2)��Ÿ����"
        
        
        iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_C_NM	: .value = " II. �� �� �� ��"	
        
          iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (1) �������κ��� ����"	
		
		  iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (2) ��Ÿ����"	
		
		
		iRow = iRow + 1 : .Row = iRow    :.TypeVAlign = 2
		.Col = C_C_NM	: .value = " III. �Ǹź�� �Ϲݰ�����"

		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (1)�޿� (�����İ�����)"

		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (2)�޿� (����ä������)"

		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (3)�� �� ��"
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (4)�������ߺ�"
		
		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (5)��ջ󰢺�"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (6)�� Ÿ"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = " IV. �� �� �� �� ��"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (1)���ڼ���"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (2)������"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (3)�� Ÿ"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = " V. �� �� �� �� ��"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (1)���ں��"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (2)�� Ÿ"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = " VI. Ư �� �� ��"
		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (1)ä��������"
		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = "  (2)��Ÿ"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = " VII. Ư �� �� ��"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = " VIII. �� �� ��"

		iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		.Col = C_C_NM	: .value = " IX. �� �� �� �� ��"

		'iRow = iRow + 1 : .Row = iRow		:.TypeVAlign = 2
		'.Col = C_C_NM	: .value = " ���� ���޾�"
		iRowTmp = 1
		
		For iRow = 1 To frm1.vspdData2.maxRows
		
			.Col = C_C_CD	: .Row = iRow
			.Text = Right("0" & (iRowTmp+26), 2)
			
			if iRow=5 then	.Text ="51"
			if iRow=6 then	.Text ="52" 
			if iRow=22 then	.Text ="53" 
			if iRow=23 then	.Text ="54"
			if  iRow=5 or  iRow=6 or  iRow=22 or  iRow=23 then
			else
			iRowTmp= iRowTmp + 1
			end if
			
		Next
		exit sub
		
		For iRow = C_W27 To C_W49
			.Col = C_C_CD	: .Row = iRow
			
			
			if iRow=C_W27+5 then	.Text ="51" : iRow=iRow + 1
			if iRow=C_W27+6 then	.Text ="52" : iRow=iRow + 1
			if iRow=C_W27+22 then	.Text ="53" : iRow=iRow + 1
			if iRow=23 then	.Text ="54" : iRow=iRow + 1
			if iRow=C_W27+5 or iRow=C_W27+6 or iRow=C_W27+22 or iRow=C_W27+23 then
			else
			.Text = Right("0" & (iRowTmp+26), 2)
			end if
			iRowTmp= iRowTmp + 1
		Next
		
		
		
		.Redraw = True
	End With
End Sub

Sub SetText4Grid(Byval pCol, Byval pRow, Byval pData)
	With frm1.vspdData
		.Col = pCol : .Row = pRow : .Text = pDAta
	End With
End Sub

Sub SetText4Grid2(Byval pCol, Byval pRow, Byval pData)
	With frm1.vspdData2
		.Col = pCol : .Row = pRow : .Text = pDAta
	End With
End Sub

Sub SetValue4Grid(Byval pCol, Byval pRow, Byval pData)
	With frm1.vspdData
		.Col = pCol : .Row = pRow : .Value = pDAta
	End With
End Sub

Sub SetValue4Grid2(Byval pCol, Byval pRow, Byval pData)
	With frm1.vspdData2
		.Col = pCol : .Row = pRow : .Value = pDAta
	End With
End Sub

' -- mb �ܿ��� 05 �̻� ����Ÿ ����� ����� 
Sub ShowColumn(Byval pCol)
	With frm1.vspdData
		.Col = pCol	: .ColHidden = False
	End With	

	With frm1.vspdData2
		.Col = pCol	: .ColHidden = False
	End With	
End Sub
'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	
End Function


'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	'Call InitData()
	CAll ClickTab1

    Call FncQuery
    
End Sub

'====================================== �� �Լ� =========================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1

End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2

End Function

'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
	dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6   ,IntRetCD, i, blnData, dblSum
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
    
	lgBlnFlgChgValue = True
    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    With frm1.vspdData
    
    
    

' ggoSpread.SpreadLock C_C01,  C_W1, C_C12, C_W1
    'ggoSpread.SpreadLock C_C01,  C_W9, C_C12, C_W9
    
    'ggoSpread.SpreadLock C_C01,  C_W15, C_C12, C_W15
    'ggoSpread.SpreadLock C_C01,  C_W22, C_C12, C_W22
    'ggoSpread.SpreadLock C_C01,  C_W24, C_C12, C_W24
    
    Select Case Row
    
		Case C_W2, C_W3, C_W4, C_W5, C_W6, C_W7, C_W8,C_W9, C_W13, C_W14	' -- �ڻ��Ѱ� �� 
			.Col = Col		: dblSum = 0
			.Row = C_W2		: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W3		: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W4		: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W5		: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W6		: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W7		: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W8		: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W9		: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W13	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W14	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W1		: .Value = dblSum
			
		
		Case C_W12, C_W10, C_W11		' -- �����ڻ� 
			.Col = Col		: dblSum = 0

			.Row = C_W10	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W11	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W12	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W9		: .Value = dblSum
    
			Call vspdData_Change(Col, C_W8)

		Case  C_W16, C_W17, C_W18, C_W19, C_W20,C_W21	' -- ��ä�Ѱ� 
			.Col = Col		: dblSum = 0
			
			.Row = C_W16	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W17	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W18	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W19	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W20	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W21	: dblSum = dblSum + UNICDbl(.value)
			
			.Row = C_W15	: .Value = dblSum

		Case C_W23, C_W24	' -- �ں����Ѱ� 
			.Col = Col		: dblSum = 0
			
			.Row = C_W23	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W24	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W22	: .Value = dblSum
    
		Case C_W25, C_W26, C_W26 + 1		' -- ��Ÿ�ں��� 
			.Col = Col		: dblSum = 0
			.Row = C_W25	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W26	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W26 + 1	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W24	: .Value = dblSum
    
			Call vspdData_Change(Col, C_W23)

    End Select
    
    End With

End Sub

Sub vspdData2_Change(ByVal Col , ByVal Row )
	dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6   ,IntRetCD, i, blnData, dblSum
    Frm1.vspdData2.Row = Row
    Frm1.vspdData2.Col = Col
    
	lgBlnFlgChgValue = True
    If Frm1.vspdData2.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData2.text) < UNICDbl(Frm1.vspdData2.TypeFloatMin) Then
         Frm1.vspdData2.text = Frm1.vspdData2.TypeFloatMin
      End If
    End If

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row
    
    With frm1.vspdData2
    
    Select Case Row
    
		Case C_W29,C_W30	' -- ����� �� 
			.Col = Col		: dblSum = 0
			
			.Row = C_W29	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W30	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W28	: .Value = dblSum
		
			Call vspdData2_Change(Col, C_W28)
	
	Case C_W33, C_W32		' -- ������� 
			.Col = Col		: dblSum = 0
			

			.Row = C_W32	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W33	: dblSum = dblSum + UNICDbl(.value)
			
			.Row = C_W31	: .Value = dblSum
			Call vspdData2_Change(Col, C_W31)
					
	Case C_W35, C_W36, C_W37, C_W38, C_W39, C_W40		' -- �Ǹź���Ϲݰ����� 
			.Col = Col		: dblSum = 0
			
			.Row = C_W35	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W36	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W37	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W39	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W38	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W40	: dblSum = dblSum + UNICDbl(.value)
			
			.Row = C_W34	: .Value = dblSum
			Call vspdData2_Change(Col, C_W34)

		Case C_W42, C_W43, C_W44	' -- �����ܼ��� 
			.Col = Col		: dblSum = 0
			.Row = C_W42	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W43	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W44	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W41	: .Value = dblSum

			Call vspdData2_Change(Col, C_W41)
			
		Case C_W46, C_W47	' -- �����ܺ�� 
			.Col = Col		: dblSum = 0
			.Row = C_W46	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W47	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W45	: .Value = dblSum
    
			Call vspdData2_Change(Col, C_W45)
		
					
		Case C_W49, C_W50	' -- Ư������ 
			.Col = Col		: dblSum = 0
			.Row = C_W49	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W50	: dblSum = dblSum + UNICDbl(.value)
			.Row = C_W48	: .Value = dblSum
    
			Call vspdData2_Change(Col, C_W48)
				
		Case C_W28, C_W31, C_W34, C_W41, C_W45, C_W48, C_W51, C_W52		' -- �������� 
			.Col = Col		: dblSum = 0
			.Row = C_W28	: dblSum = dblSum + UNICDbl(.value) '����� 
			.Row = C_W31	: dblSum = dblSum - UNICDbl(.value) '������� 
			.Row = C_W34	: dblSum = dblSum - UNICDbl(.value) '�ǰ����� 
			.Row = C_W41	: dblSum = dblSum + UNICDbl(.value) '�����ܼ��� 
			.Row = C_W45	: dblSum = dblSum - UNICDbl(.value) '�����ܺ�� 
			.Row = C_W48	: dblSum = dblSum + UNICDbl(.value) 'Ư������ 
			.Row = C_W51	: dblSum = dblSum - UNICDbl(.value) 'Ư���ս� 
			.Row = C_W52	: dblSum = dblSum - UNICDbl(.value) '���μ� 
			
			.Row = C_W53	: .Value = dblSum
    
    End Select
    
    End With

End Sub
Sub ChkRequired()
	Dim iCol, iRow, blnData
	
	With frm1.vspdData
	
	For iCol = C_C01 To C_C12 Step 2
		.Col = iCol

		blnData = False
				
		For iRow = C_W6 To C_W12
			.Row = iRow
			If Trim(.Text) <> "" Then blnData = True
		Next
				
		If blnData Then
			ggoSpread.SSSetRequired		iCol		, C_W6	,C_W12
			ggoSpread.SSSetRequired		iCol		, C_W21	,C_W21
		Else
			ggoSpread.SpreadUnLock		iCol		,-1	, iCol
		End If
		
	Next
	
	End With
End Sub


Sub vspdData_Click(ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("1101011111") 

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

End Sub    

' -- �׸���1 �˾� ��ư Ŭ�� 
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    
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

    Call SetToolbar("1100100000000111")

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
   
    
    Call InitData()

    CALL DBQuery()
    
End Function

' -- �÷� ��� ���� 
Function GetColName(Byref pGrid, Byval pCol)
	With pGrid
		.Col = pCol	: .Row = -1000
		GetColName = .Value
	End With
End Function

Function FncSave() 
    Dim blnChange, dblSum, iCol, iRow
    
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
	
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		blnChange = True
    End If

    ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
		blnChange = True
    End If
	
'	If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
'	      Exit Function
'	End If    
	
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
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 

End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
           
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
	
    ggoSpread.Source = frm1.vspdData	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function FncDelete()
Dim iRow 
Dim IntRetCd

    frm1.vspdData.AddSelection C_W1, -1, C_W1, -1

    If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("800442", parent.VB_YES_NO, "X", "X")			    <%'%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '��: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call FncDeleteRow
       
    Call FncSave
    
   lgBlnFlgChgValue = True
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
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim iCol,i
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	
	If lgIntFlgMode <> parent.OPMD_UMODE  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE

		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 ����üũ : �׸��� �� 
		If wgConfirmFlg <>"Y" Then
			
			'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
			Call SetToolbar("1101100000000111")										<%'��ư ���� ���� %>

		Else
			
		End If
		For i=1 to frm1.vspdData.MaxRows
			frm1.vspdData.Row=i
			frm1.vspdData.Col =0
			frm1.vspdData.Text=""		
		Next
		For i=1 to frm1.vspdData2.MaxRows
			frm1.vspdData2.Row=i
			frm1.vspdData2.Col =0
			frm1.vspdData2.Text=""		
		Next
	Else
		Call SetToolbar("1101100000000111")										<%'��ư ���� ���� %>
	End If

	lgBlnFlgChgValue = False
    
	'Call SetSpreadLock(TYPE_1)
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow, lCol   
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel, lMaxRows, lMaxCols, arrVal(11)
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	lGrpCnt = 0
	
	For lCol = C_C01 To C_C12

		If lgIntFlgMode = parent.OPMD_CMODE  Then
			strVal = "C"  &  Parent.gColSep
		Else
			strVal = "U"  &  Parent.gColSep
		End If
		
		With frm1.vspdData
			.Col = lCol
			.Row = -1000	' -- ��� 
			strVal = strVal & Trim(.Text)  &  Parent.gColSep ' -- �÷���ȣ 
			
			For lRow = 1 To .MaxRows
               .Row = lRow

				Select Case lRow
					Case C_W26
						strVal = strVal & UNICDbl(.Value) &  Parent.gColSep 
					Case Else
						strVal = strVal & UNICDbl(.Value) &  Parent.gColSep 
				End Select
			Next

		End With

		With frm1.vspdData2
			.Col = lCol
			
			For lRow = 1 To .MaxRows
               .Row = lRow

				Select Case lRow
					Case .MaxRows
						strVal = strVal & UNICDbl(.Value) &  Parent.gRowSep		' <-- �ʵ� ������ 
					Case Else
						strVal = strVal & UNICDbl(.Value) &  Parent.gColSep 
				End Select
			
			Next
			
		End With

		arrVal(lGrpCnt) = strVal
		lGrpCnt = lGrpCnt + 1		  

    Next

    frm1.txtSpread.value        =  Join(arrVal, "")
	frm1.txtMode.value        =  Parent.UID_M0002

	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           

End Function


Function DbSaveOk()													        <%' ���� ������ ���� ���� %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
	Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
    Call MainQuery()
End Function

'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function


function FncBtnPrint1(strPrintType)
	dim sWhere,sMaxSeq,vArrSeq
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE
	Dim StrUrl  , i

	Dim intCnt,IntRetCD
	
	sWhere = "CO_CD=" & FilterVar("<%=wgCO_CD%>", "''", "S") & vbCrLf
	sWhere = sWhere & " AND FISC_YEAR=" & FilterVar("<%=wgFISC_YEAR%>", "''", "S") & vbCrLf
	sWhere = sWhere & " AND REP_TYPE=" & FilterVar("<%=wgREP_TYPE%>", "''", "S") & vbCrLf
	sWhere = sWhere & " AND ISNULL(A.W6,'') <> ''"

	if  CommonQueryRs("distinct  seq_no "," TB_A125 A ",sWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) then
		
	else
	lgF0 = "1" & chr(11) '��ȭ���� ����� �� �ֵ��� ��.
	end if

	vArrSeq = split(lgF0,chr(11))
	
	for i=0 to uBound(vArrSeq)-1
		StrUrl=""
			Call SetPrintCond(varCo_Cd, varFISC_YEAR, varREP_TYPE) 
			StrUrl = StrUrl & "varCo_Cd|"			& varCo_Cd
			StrUrl = StrUrl & "|varFISC_YEAR|"		& varFISC_YEAR
			StrUrl = StrUrl & "|varREP_TYPE|"       & varREP_TYPE
			StrUrl = StrUrl & "|varseq_no|"       & vArrSeq(i)

			 ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")

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
     	next 
end function
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>����������ǥ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�����Ͱ�꼭</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>����:  ��</TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
						</DIV>

						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint1('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

